// ============================================
// GOOGLE APPS SCRIPT - VMS v3.9
// FITUR: Row-level Dedup + CacheService + Batch Append + Lock Optimization
// ============================================

// ========== KONFIGURASI ==========
const CONFIG = {
    SHEET_ID: '1IlKDIHpMdA7FJwLFEpJP-V4iDnAXnpE9YKFJRkGZ6uI',
    SHEET_NAME: 'VMS_Tracking_Data',
    HEADERS: [
        'Timestamp',   // A: Kolom 1
        'Latitude',    // B: Kolom 2
        'Longitude',   // C: Kolom 3
        'Accuracy',    // D: Kolom 4
        'Heading',     // E: Kolom 5
        'Speed',       // F: Kolom 6
        'DeviceID',    // G: Kolom 7
        'BatchID',     // H: Kolom 8
        'ServerTime',  // I: Kolom 9
        'PointHash'    // J: Kolom 10 (P0: Point-level dedup)
    ]
};

// ========== DO POST - MAIN ENTRY ==========
function doPost(e) {
    const startTime = Date.now();
    
    try {
        let data;
        if (e.postData && e.postData.contents) {
            data = JSON.parse(e.postData.contents);
        } else if (e.parameter && e.parameter.payload) {
            data = JSON.parse(e.parameter.payload);
        } else {
            return createResponse(false, 'No data received', null);
        }
        
        const { points, deviceId, batchId, timestamp, totalOriginal, dedupedCount } = data;
        
        if (!points || !Array.isArray(points) || points.length === 0) {
            return createResponse(false, 'No valid points', null);
        }
        
        console.log(`[GAS] Processing batch ${batchId} with ${points.length} unique points (from ${totalOriginal || points.length} original)`);
        console.log(`[GAS] Sheet ID: ${CONFIG.SHEET_ID}`);
        
        // ========== CACHESERVICE ROW-LEVEL DEDUP (P0) ==========
        const cache = CacheService.getScriptCache();
        const rows = [];
        let insertedCount = 0;
        let duplicateCount = 0;
        
        for (const point of points) {
            const pointHash = point.pointHash || `${point.timestamp || timestamp}_${point.lat}_${point.lng}`;
            const cacheKey = `pt:${pointHash}`;
            
            // Cek apakah point sudah pernah diproses (6 jam retention)
            if (cache.get(cacheKey)) {
                duplicateCount++;
                console.log(`[GAS] Duplicate point skipped: ${pointHash}`);
                continue;
            }
            
            rows.push([
                point.timestamp || timestamp || Date.now(),  // 1. Timestamp
                point.lat,                                    // 2. Latitude
                point.lng,                                    // 3. Longitude
                point.accuracy || 0,                          // 4. Accuracy
                point.heading || 0,                           // 5. Heading
                point.speed || 0,                             // 6. Speed
                deviceId || 'unknown',                        // 7. DeviceID
                batchId || 'unknown',                         // 8. BatchID
                startTime,                                    // 9. ServerTime
                pointHash                                     // 10. PointHash (P0)
            ]);
            
            // Cache point hash untuk 6 jam (mencegah duplicate dari batch berbeda)
            cache.put(cacheKey, '1', 21600);
            insertedCount++;
        }
        
        if (rows.length === 0) {
            console.log(`[GAS] All points duplicate, skipping write`);
            return createResponse(true, `All ${points.length} points already exist (duplicate)`, {
                batchId: batchId,
                received: points.length,
                inserted: 0,
                duplicates: duplicateCount,
                processingTime: Date.now() - startTime
            });
        }
        
        // ========== LOCK OPTIMIZATION (short lock duration) ==========
        const lock = LockService.getScriptLock();
        let lockAcquired = false;
        
        try {
            // Short lock timeout (10 detik, bukan 30)
            lock.waitLock(10000);
            lockAcquired = true;
            
            // AKSES SHEET
            const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
            let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
            
            if (!sheet) {
                sheet = ss.insertSheet(CONFIG.SHEET_NAME);
                sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
                sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setFontWeight('bold');
                sheet.setFrozenRows(1);
                console.log(`[GAS] Created new sheet: ${CONFIG.SHEET_NAME}`);
            }
            
            // BATCH INSERT (setValues, bukan appendRow per row)
            if (rows.length > 0) {
                const lastRow = sheet.getLastRow();
                sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
                console.log(`[GAS] Inserted ${rows.length} rows at row ${lastRow + 1}`);
            }
            
            // Cleanup: jaga performa spreadsheet (max 100.000 baris)
            const totalRows = sheet.getLastRow();
            if (totalRows > 100000) {
                const rowsToDelete = totalRows - 80000;
                sheet.deleteRows(2, rowsToDelete);
                console.log(`[GAS] Cleaned up ${rowsToDelete} old rows`);
            }
            
            const stats = {
                inserted: rows.length,
                duplicates: duplicateCount,
                batchId: batchId,
                deviceId: deviceId,
                sheetName: CONFIG.SHEET_NAME,
                sheetId: CONFIG.SHEET_ID,
                processingTime: Date.now() - startTime
            };
            
            return createResponse(true, `Inserted ${rows.length} unique points (${duplicateCount} duplicates skipped)`, stats);
            
        } catch (lockError) {
            console.error('[GAS] Lock error:', lockError);
            return createResponse(false, `Lock error: ${lockError.message}`, null);
        } finally {
            if (lockAcquired) lock.releaseLock();
        }
        
    } catch (error) {
        console.error('[GAS] Fatal error:', error);
        return createResponse(false, error.toString(), null);
    }
}

// ========== DO GET - HEALTH CHECK ==========
function doGet() {
    try {
        const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
        const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
        
        let response = {
            status: 'active',
            timestamp: Date.now(),
            version: '3.9',
            sheetId: CONFIG.SHEET_ID,
            sheetName: CONFIG.SHEET_NAME,
            headers: CONFIG.HEADERS,
            features: ['row-level-dedup', 'cache-service', 'batch-append']
        };
        
        if (sheet) {
            response.sheetExists = true;
            response.totalRows = sheet.getLastRow() - 1;
            response.lastUpdate = sheet.getLastRow() > 1 ? sheet.getRange(sheet.getLastRow(), 1).getValue() : null;
            response.sheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}`;
        } else {
            response.sheetExists = false;
            response.message = `Sheet "${CONFIG.SHEET_NAME}" will be created on first POST`;
        }
        
        return ContentService
            .createTextOutput(JSON.stringify(response, null, 2))
            .setMimeType(ContentService.MimeType.JSON);
            
    } catch (error) {
        return ContentService
            .createTextOutput(JSON.stringify({
                status: 'error',
                error: error.toString(),
                sheetId: CONFIG.SHEET_ID,
                timestamp: Date.now()
            }, null, 2))
            .setMimeType(ContentService.MimeType.JSON);
    }
}

// ========== HELPER FUNCTIONS ==========
function createResponse(success, message, data) {
    const response = {
        success: success,
        message: message,
        timestamp: Date.now(),
        sheetId: CONFIG.SHEET_ID,
        sheetName: CONFIG.SHEET_NAME,
        version: '3.9',
        data: data || null
    };
    
    return ContentService
        .createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
}

// ========== VALIDASI INSTALASI ==========
function validateInstallation() {
    try {
        const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
        let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
        
        if (!sheet) {
            sheet = ss.insertSheet(CONFIG.SHEET_NAME);
            sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
            sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setFontWeight('bold');
            sheet.setFrozenRows(1);
            return { 
                valid: true, 
                message: 'Sheet created successfully',
                sheetId: CONFIG.SHEET_ID,
                sheetUrl: `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}`,
                headers: CONFIG.HEADERS 
            };
        }
        
        const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        // Cek apakah perlu migrasi header (tambah kolom PointHash)
        if (currentHeaders.length < CONFIG.HEADERS.length) {
            sheet.getRange(1, CONFIG.HEADERS.length, 1, 1).setValue('PointHash');
            console.log('[GAS] Migrated sheet: added PointHash column');
        }
        
        return {
            valid: true,
            message: 'Installation valid',
            sheetId: CONFIG.SHEET_ID,
            sheetUrl: `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}`,
            totalRows: sheet.getLastRow() - 1,
            currentHeaders: currentHeaders,
            expectedHeaders: CONFIG.HEADERS
        };
    } catch (error) {
        return { 
            valid: false, 
            error: error.toString(),
            sheetId: CONFIG.SHEET_ID,
            help: 'Make sure the script has access to this spreadsheet'
        };
    }
}

// ========== RESET SHEET ==========
function resetSheet() {
    try {
        const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
        let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
        if (sheet) ss.deleteSheet(sheet);
        
        sheet = ss.insertSheet(CONFIG.SHEET_NAME);
        sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
        sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setFontWeight('bold');
        sheet.setFrozenRows(1);
        
        return { 
            success: true, 
            message: `Sheet "${CONFIG.SHEET_NAME}" reset successfully`,
            sheetId: CONFIG.SHEET_ID,
            sheetUrl: `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}`,
            headers: CONFIG.HEADERS 
        };
    } catch (error) {
        return { success: false, error: error.toString() };
    }
}

// ========== GET SHEET URL ==========
function getSheetUrl() {
    return {
        sheetId: CONFIG.SHEET_ID,
        sheetUrl: `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}`,
        editUrl: `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/edit`
    };
}

// ========== CLEAN OLD CACHE (Opsional) ==========
function cleanOldCache() {
    console.log('[GAS] Cache auto-expires after 6 hours, no manual cleanup needed');
    return { status: 'ok', message: 'Cache auto-expires after 6 hours' };
}