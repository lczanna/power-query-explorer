// XPress9 DataModel Decoder — Extracts M code from V3 .pbix DataModel files
// Requires xpress9.js (Emscripten module) to be loaded first
// Usage: const queries = await extractFromDataModel(arrayBuffer, fileName);
(function(){
'use strict';

// XPRESS9_WASM_B64 will be injected by the build script
const XPRESS9_WASM_B64 = '%%XPRESS9_WASM_B64%%';

let _xp9Module = null;

async function getXpress9() {
    if (_xp9Module) return _xp9Module;
    if (typeof Xpress9Module === 'undefined') throw new Error('XPress9 WASM module not loaded');
    const bin = Uint8Array.from(atob(XPRESS9_WASM_B64), c => c.charCodeAt(0));
    _xp9Module = await Xpress9Module({ wasmBinary: bin.buffer });
    return _xp9Module;
}

// ═══ XPress9 Decompression ═══

async function decompressXpress9(dataModelBuf) {
    const mod = await getXpress9();
    const buf = new Uint8Array(dataModelBuf);
    const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);

    // Read UTF-16LE signature (first 102 bytes)
    let sig = '';
    for (let i = 0; i < 102; i += 2) {
        const ch = buf[i] | (buf[i+1] << 8);
        if (ch === 0) break;
        sig += String.fromCharCode(ch);
    }

    const isMultiThread = sig.includes('multithreaded');
    let offset = 102;
    const chunks = [];

    if (isMultiThread) {
        // Multi-threaded format: 5x uint64 metadata after signature
        const readU64 = (o) => Number(view.getBigUint64(o, true));
        const mainChunks = readU64(offset); offset += 8;
        const prefixChunks = readU64(offset); offset += 8;
        const prefixThreads = readU64(offset); offset += 8;
        const mainThreads = readU64(offset); offset += 8;
        /* const chunkSize = */ readU64(offset); offset += 8;

        // Decompress thread groups: prefix threads, then main threads
        const groups = [];
        for (let t = 0; t < prefixThreads; t++) groups.push(prefixChunks);
        for (let t = 0; t < mainThreads; t++) groups.push(mainChunks);

        for (const nBlocks of groups) {
            // Each thread gets a fresh decoder session
            mod.ccall('xpress9_free', null, [], []);
            if (!mod.ccall('xpress9_init', 'number', [], [])) throw new Error('XPress9 init failed');
            for (let b = 0; b < nBlocks && offset + 8 <= buf.length; b++) {
                offset = decompressBlock(mod, buf, view, offset, chunks);
            }
        }
    } else {
        // Single-threaded: sequential blocks
        if (!mod.ccall('xpress9_init', 'number', [], [])) throw new Error('XPress9 init failed');
        while (offset + 8 <= buf.length) {
            const prevOffset = offset;
            offset = decompressBlock(mod, buf, view, offset, chunks);
            if (offset === prevOffset) break;
        }
    }

    mod.ccall('xpress9_free', null, [], []);

    // Concatenate chunks
    const totalLen = chunks.reduce((s, c) => s + c.length, 0);
    const result = new Uint8Array(totalLen);
    let pos = 0;
    for (const chunk of chunks) { result.set(chunk, pos); pos += chunk.length; }
    return result;
}

function decompressBlock(mod, buf, view, offset, chunks) {
    const uncompSize = view.getUint32(offset, true);
    const compSize = view.getUint32(offset + 4, true);
    offset += 8;
    if (compSize === 0 || uncompSize === 0 || offset + compSize > buf.length) return offset;

    const srcPtr = mod._malloc(compSize);
    const dstPtr = mod._malloc(uncompSize);
    mod.HEAPU8.set(buf.subarray(offset, offset + compSize), srcPtr);

    const result = mod.ccall('xpress9_decompress', 'number',
        ['number', 'number', 'number', 'number'],
        [srcPtr, compSize, dstPtr, uncompSize]);

    if (result > 0) {
        chunks.push(new Uint8Array(mod.HEAPU8.buffer.slice(dstPtr, dstPtr + result)));
    }

    mod._free(srcPtr);
    mod._free(dstPtr);
    return offset + compSize;
}

// ═══ ABF Parser — extract metadata.sqlitedb from decompressed DataModel ═══

function extractSQLiteFromABF(data) {
    // BackupLogHeader: UTF-16LE XML at offset 72, one page (4096 bytes)
    let headerXml = '';
    for (let i = 72; i < 4096 - 1; i += 2) {
        const ch = data[i] | (data[i+1] << 8);
        if (ch === 0) break;
        headerXml += String.fromCharCode(ch);
    }

    const getXmlVal = (xml, tag) => { const m = xml.match(new RegExp('<' + tag + '>(.*?)</' + tag + '>')); return m ? m[1] : null; };
    const vdOffset = parseInt(getXmlVal(headerXml, 'm_cbOffsetHeader'));
    const vdSize = parseInt(getXmlVal(headerXml, 'DataSize'));
    const errorCode = getXmlVal(headerXml, 'ErrorCode') === 'true';
    if (!vdOffset || !vdSize) throw new Error('Invalid ABF BackupLogHeader');

    // VirtualDirectory: UTF-8 XML
    const vdText = new TextDecoder('utf-8').decode(data.subarray(vdOffset, vdOffset + vdSize));
    const vdFiles = {};
    const vdRe = /<BackupFile><Path>(.*?)<\/Path><Size>(\d+)<\/Size><m_cbOffsetHeader>(\d+)<\/m_cbOffsetHeader>/g;
    let m, lastPath = null;
    while ((m = vdRe.exec(vdText)) !== null) {
        vdFiles[m[1]] = { size: parseInt(m[2]), offset: parseInt(m[3]) };
        lastPath = m[1];
    }

    // BackupLog: last VD entry, UTF-16LE with BOM
    if (!lastPath) throw new Error('Empty VirtualDirectory');
    const logEntry = vdFiles[lastPath];
    let logBytes = data.subarray(logEntry.offset, logEntry.offset + logEntry.size);
    if (errorCode && logBytes.length > 4) logBytes = logBytes.subarray(0, logBytes.length - 4);

    let logText;
    if (logBytes[0] === 0xFF && logBytes[1] === 0xFE) logText = new TextDecoder('utf-16le').decode(logBytes.subarray(2));
    else if (logBytes.length > 1 && logBytes[1] === 0) logText = new TextDecoder('utf-16le').decode(logBytes);
    else logText = new TextDecoder('utf-8').decode(logBytes);

    // Find metadata.sqlitedb via BackupLog StoragePath -> VD entry mapping
    const bfRe = /<BackupFile>[\s\S]*?<Path>(.*?)<\/Path>[\s\S]*?<StoragePath>(.*?)<\/StoragePath>/g;
    while ((m = bfRe.exec(logText)) !== null) {
        if (m[1].toLowerCase().endsWith('metadata.sqlitedb') && vdFiles[m[2]]) {
            const e = vdFiles[m[2]];
            return data.slice(e.offset, e.offset + e.size);
        }
    }
    throw new Error('metadata.sqlitedb not found in ABF structure');
}

// ═══ Minimal SQLite Reader — reads only what we need for M code ═══

function readMCodeFromSQLite(dbBuf) {
    const buf = new Uint8Array(dbBuf);
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);

    // Verify magic
    const magic = new TextDecoder('ascii').decode(buf.subarray(0, 16));
    if (!magic.startsWith('SQLite format 3')) throw new Error('Invalid SQLite database');

    const pageSize = dv.getUint16(16) || 65536;
    const reserved = buf[20];
    const usableSize = pageSize - reserved;

    function getPage(num) { return buf.subarray((num - 1) * pageSize, num * pageSize); }

    // Read a SQLite varint (MSB first, 7 bits per byte, up to 9 bytes)
    function readVarint(data, pos) {
        let result = 0;
        for (let i = 0; i < 8; i++) {
            const b = data[pos + i];
            result = result * 128 + (b & 0x7f);
            if (!(b & 0x80)) return { v: result, n: i + 1 };
        }
        // 9th byte: full 8 bits
        result = result * 256 + data[pos + 8];
        return { v: result, n: 9 };
    }

    // Read a big-endian int from buffer
    function readBE(data, pos, len) {
        let val = 0;
        for (let i = 0; i < len; i++) val = val * 256 + data[pos + i];
        return val;
    }

    // Read payload from a leaf table B-tree cell, handling overflow
    function readCellPayload(page, cellPtr) {
        const { v: payloadLen, n: n1 } = readVarint(page, cellPtr);
        const { v: rowid, n: n2 } = readVarint(page, cellPtr + n1);
        let hdrStart = cellPtr + n1 + n2;

        // Calculate local payload size (for leaf table B-tree: maxLocal = usableSize - 35)
        const maxLocal = usableSize - 35;
        const minLocal = ((usableSize - 12) * 32 / 255 | 0) - 23;
        let localSize;
        if (payloadLen <= maxLocal) {
            localSize = payloadLen;
        } else {
            localSize = minLocal + ((payloadLen - minLocal) % (usableSize - 4));
            if (localSize > maxLocal) localSize = minLocal;
        }

        const payload = new Uint8Array(payloadLen);
        payload.set(page.subarray(hdrStart, hdrStart + Math.min(localSize, payloadLen)));

        if (localSize < payloadLen) {
            // Read overflow pages
            let overflowPageNum = readBE(page, hdrStart + localSize, 4);
            let written = localSize;
            while (overflowPageNum !== 0 && written < payloadLen) {
                const oPage = getPage(overflowPageNum);
                overflowPageNum = readBE(oPage, 0, 4);
                const avail = Math.min(usableSize - 4, payloadLen - written);
                payload.set(oPage.subarray(4, 4 + avail), written);
                written += avail;
            }
        }

        return { payload, rowid };
    }

    // Parse a record from payload bytes
    function parseRecord(payload) {
        const { v: headerLen, n: hb } = readVarint(payload, 0);
        const types = [];
        let pos = hb;
        while (pos < headerLen) {
            const { v: st, n: sn } = readVarint(payload, pos);
            types.push(st);
            pos += sn;
        }

        const values = [];
        let dPos = headerLen;
        for (const st of types) {
            if (st === 0) { values.push(null); }
            else if (st >= 1 && st <= 6) {
                const lens = [0, 1, 2, 3, 4, 6, 8];
                const len = lens[st];
                let val = 0;
                for (let i = 0; i < len; i++) val = val * 256 + payload[dPos + i];
                // Sign extend for negative values
                if (len > 0 && payload[dPos] & 0x80) val -= (1 << (len * 8));
                values.push(val);
                dPos += len;
            }
            else if (st === 7) {
                const f64 = new DataView(payload.buffer, payload.byteOffset + dPos, 8).getFloat64(0, false);
                values.push(f64);
                dPos += 8;
            }
            else if (st === 8) { values.push(0); }
            else if (st === 9) { values.push(1); }
            else if (st >= 12 && st % 2 === 0) {
                const len = (st - 12) / 2;
                values.push(payload.slice(dPos, dPos + len));
                dPos += len;
            }
            else if (st >= 13 && st % 2 === 1) {
                const len = (st - 13) / 2;
                values.push(new TextDecoder('utf-8').decode(payload.subarray(dPos, dPos + len)));
                dPos += len;
            }
            else { values.push(null); }
        }
        return values;
    }

    // Traverse a B-tree and collect all rows
    function readTable(rootPage) {
        const rows = [];
        function traverse(pageNum) {
            const page = getPage(pageNum);
            const hdrOff = pageNum === 1 ? 100 : 0;
            const pageType = page[hdrOff];

            if (pageType === 0x0d) {
                // Leaf table B-tree page
                const numCells = (page[hdrOff + 3] << 8) | page[hdrOff + 4];
                for (let i = 0; i < numCells; i++) {
                    const cellPtr = (page[hdrOff + 8 + i*2] << 8) | page[hdrOff + 8 + i*2 + 1];
                    const { payload, rowid } = readCellPayload(page, cellPtr);
                    try { rows.push({ rowid, values: parseRecord(payload) }); } catch(e) { /* skip corrupt */ }
                }
            } else if (pageType === 0x05) {
                // Interior table B-tree page
                const numCells = (page[hdrOff + 3] << 8) | page[hdrOff + 4];
                const rightChild = readBE(page, hdrOff + 8, 4);
                for (let i = 0; i < numCells; i++) {
                    const cellPtr = (page[hdrOff + 12 + i*2] << 8) | page[hdrOff + 12 + i*2 + 1];
                    const childPage = readBE(page, cellPtr, 4);
                    traverse(childPage);
                }
                traverse(rightChild);
            }
        }
        traverse(rootPage);
        return rows;
    }

    // 1. Parse sqlite_master (page 1)
    const masterRows = readTable(1);
    const tableInfo = {};
    for (const row of masterRows) {
        const [type, name, , rootpage] = row.values;
        if (type === 'table' && typeof name === 'string') {
            tableInfo[name] = rootpage;
        }
    }

    // 2. Read Table names (ID -> Name mapping)
    const tableNames = {};
    if (tableInfo['Table']) {
        const tableRows = readTable(tableInfo['Table']);
        for (const row of tableRows) {
            // Table schema: ID(0), ModelID(1), Name(2), ...
            const id = row.values[0];
            const name = row.values[2];
            if (id != null && typeof name === 'string') tableNames[id] = name;
        }
    }

    // 3. Read Partition.QueryDefinition (M code for tables)
    // Partition names can include internal IDs (e.g. Table-<guid>), so we normalize
    // to one entry per table using the user-facing table name.
    const queries = [];
    const partitionByTable = new Map();
    const guidSuffixRe = /-[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    const systemTableNameRe = /^(?:LocalDateTable_|DateTableTemplate_)/i;

    function partitionScore(partName, tableName, queryDef) {
        let score = 0;
        if (typeof partName === 'string' && partName.trim()) {
            const cleanPartName = partName.replace(guidSuffixRe, '');
            if (cleanPartName === tableName) score += 4;
            if (!guidSuffixRe.test(partName)) score += 1;
        } else {
            score += 1;
        }
        if (typeof queryDef === 'string') score += Math.min(queryDef.length, 10000) / 10000;
        return score;
    }

    function isSystemModelTableName(name) {
        return typeof name === 'string' && systemTableNameRe.test(name);
    }

    if (tableInfo['Partition']) {
        const partRows = readTable(tableInfo['Partition']);
        for (const row of partRows) {
            // Partition schema: ID(0), TableID(1), Name(2), Description(3), DataSourceID(4), QueryDefinition(5), ...
            const tableId = row.values[1];
            const partName = row.values[2];
            const queryDef = row.values[5];
            if (typeof queryDef === 'string' && queryDef.trim().length > 0) {
                const partBaseName = (typeof partName === 'string' && partName.trim())
                    ? partName.replace(guidSuffixRe, '')
                    : '';
                const tableName = tableNames[tableId] || partBaseName || ('Table_' + tableId);
                if (isSystemModelTableName(tableName)) continue;
                const candidate = {
                    name: tableName,
                    tableName: tableName,
                    mCode: queryDef,
                    _score: partitionScore(partName, tableName, queryDef)
                };
                const existing = partitionByTable.get(tableName);
                if (!existing || candidate._score > existing._score) {
                    partitionByTable.set(tableName, candidate);
                }
            }
        }
    }

    for (const q of partitionByTable.values()) {
        queries.push({ name: q.name, tableName: q.tableName, mCode: q.mCode });
    }

    // 4. Read Expression table (shared M expressions / parameters)
    if (tableInfo['Expression']) {
        const exprRows = readTable(tableInfo['Expression']);
        for (const row of exprRows) {
            // Expression schema: ID(0), ModelID(1), Name(2), Description(3), Kind(4), Expression(5), ...
            const name = row.values[2];
            const expr = row.values[5];
            if (typeof expr === 'string' && expr.trim().length > 0) {
                if (isSystemModelTableName(name)) continue;
                queries.push({
                    name: name || 'Expression',
                    tableName: null,
                    mCode: expr
                });
            }
        }
    }

    return queries;
}

// ═══ Main entry point ═══

window.extractFromDataModel = async function(dataModelArrayBuffer, fileName) {
    // Decompress XPress9
    const decompressed = await decompressXpress9(dataModelArrayBuffer);

    // Extract metadata.sqlitedb from ABF structure
    const sqliteData = extractSQLiteFromABF(decompressed);

    // Read M code from SQLite
    const mEntries = readMCodeFromSQLite(sqliteData);

    // Return in the format expected by the app
    return mEntries.map(e => {
        const name = (typeof e.name === 'string' && e.name.trim()) ? e.name : 'Query';
        const code = typeof e.mCode === 'string' ? e.mCode : '';
        const dependencies = (typeof stripCS === 'function' && typeof findDeps === 'function')
            ? findDeps(stripCS(code), name)
            : [];
        const externalRefs = typeof findExternalFileRefs === 'function'
            ? findExternalFileRefs(code)
            : [];
        return {
            name,
            code,
            fileName: fileName,
            dependencies,
            externalRefs
        };
    });
};

})();
