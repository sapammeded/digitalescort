// ============================================
// CLOUDFLARE WORKER - VMS v3.9
// FITUR: KV Idempotency + Point-level Dedup + Batch Registry + Retry-Safe ACK
// ============================================

const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwrJLe8zCHTXuSsNEeo_oyVUMGyFVjq3Br3nvLql425Q2Ckh3isMwyAqthrYF6qLWnkIQ/exec';

export default {
    async fetch(request, env, ctx) {
        const corsHeaders = {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type',
            'Access-Control-Max-Age': '86400',
        };
        
        if (request.method === 'OPTIONS') {
            return new Response(null, { headers: corsHeaders });
        }
        
        const url = new URL(request.url);
        
        // Health check
        if (request.method === 'GET') {
            return new Response(JSON.stringify({
                status: 'online',
                version: '3.9',
                timestamp: Date.now(),
                features: ['kv-idempotency', 'point-dedup', 'batch-registry', 'retry-safe']
            }), { headers: { 'Content-Type': 'application/json', ...corsHeaders } });
        }
        
        // Sync endpoint
        if (request.method === 'POST' && (url.pathname === '/sync' || url.pathname === '/api/v1/sync')) {
            try {
                const payload = await request.json();
                const { points, deviceId, batchId, timestamp } = payload;
                
                if (!points || !Array.isArray(points) || points.length === 0) {
                    return new Response(JSON.stringify({
                        success: false,
                        error: 'No points provided',
                        code: 'EMPTY_PAYLOAD'
                    }), { status: 400, headers: { 'Content-Type': 'application/json', ...corsHeaders } });
                }
                
                // ========== KV BATCH IDEMPOTENCY ==========
                if (!env.VMS_KV) {
                    console.warn('[WORKER] KV namespace not bound, idempotency disabled');
                } else {
                    const processedKey = `processed:${batchId}`;
                    const existing = await env.VMS_KV.get(processedKey);
                    
                    if (existing) {
                        console.log(`[WORKER] Duplicate batch rejected: ${batchId}`);
                        return new Response(JSON.stringify({
                            success: true,
                            duplicate: true,
                            message: 'Batch already processed',
                            batchId: batchId,
                            previouslyProcessedAt: parseInt(existing)
                        }), { status: 200, headers: { 'Content-Type': 'application/json', ...corsHeaders } });
                    }
                }
                
                // Filter valid coordinates
                let validPoints = points.filter(p => 
                    typeof p.lat === 'number' && 
                    typeof p.lng === 'number' &&
                    Math.abs(p.lat) <= 90 &&
                    Math.abs(p.lng) <= 180
                ).map(p => ({
                    lat: p.lat,
                    lng: p.lng,
                    alt: p.alt || 0,
                    accuracy: p.accuracy || 0,
                    heading: p.heading || 0,
                    speed: p.speed || 0,
                    timestamp: p.timestamp || Date.now(),
                    pointHash: p.pointHash || `${p.timestamp || Date.now()}_${p.lat}_${p.lng}`,
                    seq: p.seq || 0
                }));
                
                if (validPoints.length === 0) {
                    return new Response(JSON.stringify({
                        success: false,
                        error: 'No valid coordinates',
                        code: 'INVALID_COORDINATES'
                    }), { status: 400, headers: { 'Content-Type': 'application/json', ...corsHeaders } });
                }
                
                // ========== POINT-LEVEL DEDUP (P0) ==========
                const dedupedPoints = [];
                if (env.VMS_KV) {
                    for (const p of validPoints) {
                        const pointKey = `pt:${deviceId || 'unknown'}:${p.pointHash}`;
                        const exists = await env.VMS_KV.get(pointKey);
                        if (!exists) {
                            dedupedPoints.push(p);
                            await env.VMS_KV.put(pointKey, '1', { expirationTtl: 604800 }); // 7 hari
                        } else {
                            console.log(`[WORKER] Duplicate point rejected: ${pointKey}`);
                        }
                    }
                } else {
                    dedupedPoints.push(...validPoints);
                }
                
                if (dedupedPoints.length === 0) {
                    console.log(`[WORKER] All points already processed, marking batch as processed`);
                    if (env.VMS_KV) {
                        await env.VMS_KV.put(`processed:${batchId}`, Date.now().toString(), { expirationTtl: 86400 * 30 });
                    }
                    return new Response(JSON.stringify({
                        success: true,
                        duplicate: true,
                        message: 'All points already processed',
                        batchId: batchId,
                        totalPoints: points.length,
                        dedupedPoints: 0
                    }), { status: 200, headers: { 'Content-Type': 'application/json', ...corsHeaders } });
                }
                
                console.log(`[WORKER] Processing batch ${batchId}: ${dedupedPoints.length}/${points.length} unique points`);
                
                // Backup ke KV (retensi 7 hari)
                if (env.VMS_KV) {
                    const backupKey = `backup:${batchId}`;
                    await env.VMS_KV.put(backupKey, JSON.stringify({
                        batchId,
                        deviceId,
                        points: dedupedPoints,
                        timestamp: Date.now()
                    }), { expirationTtl: 604800 });
                }
                
                // Forward ke Google Apps Script dengan timeout
                const controller = new AbortController();
                const timeoutId = setTimeout(() => controller.abort(), 25000);
                
                const gsResponse = await fetch(GOOGLE_SCRIPT_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json', 'Connection': 'close' },
                    body: JSON.stringify({
                        points: dedupedPoints,
                        deviceId: deviceId || 'vms-mobile',
                        batchId: batchId,
                        timestamp: Date.now(),
                        source: 'cloudflare-worker-v3.9',
                        totalOriginal: points.length,
                        dedupedCount: dedupedPoints.length
                    }),
                    signal: controller.signal
                }).catch(err => {
                    console.error('[WORKER] GAS fetch error:', err);
                    return null;
                });
                
                clearTimeout(timeoutId);
                
                let gsResult = { success: false, error: 'GAS timeout or error' };
                if (gsResponse && gsResponse.ok) {
                    try {
                        const text = await gsResponse.text();
                        gsResult = JSON.parse(text);
                    } catch(e) {
                        gsResult = { success: false, error: 'Invalid GAS response' };
                    }
                }
                
                // ========== MARK BATCH AS PROCESSED (setelah GAS sukses) ==========
                if (gsResult.success === true && env.VMS_KV) {
                    await env.VMS_KV.put(`processed:${batchId}`, Date.now().toString(), { expirationTtl: 86400 * 30 });
                    console.log(`[WORKER] Batch ${batchId} marked as processed`);
                }
                
                return new Response(JSON.stringify({
                    success: gsResult.success === true,
                    batchId: batchId,
                    received: points.length,
                    valid: validPoints.length,
                    deduped: validPoints.length - dedupedPoints.length,
                    unique: dedupedPoints.length,
                    timestamp: Date.now(),
                    duplicate: false,
                    gasResponse: gsResult,
                    version: '3.9'
                }), {
                    status: 200,
                    headers: { 'Content-Type': 'application/json', ...corsHeaders }
                });
                
            } catch (error) {
                console.error('[WORKER] Fatal error:', error);
                return new Response(JSON.stringify({
                    success: false,
                    error: error.message,
                    code: 'WORKER_ERROR',
                    timestamp: Date.now()
                }), {
                    status: 500,
                    headers: { 'Content-Type': 'application/json', ...corsHeaders }
                });
            }
        }
        
        return new Response(JSON.stringify({ error: 'Not found', code: 'NOT_FOUND' }), {
            status: 404,
            headers: { 'Content-Type': 'application/json', ...corsHeaders }
        });
    }
};