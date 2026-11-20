// Export XLSX autonome : OOXML dans une archive ZIP sans dépendance externe.
(function () {
    "use strict";

    const ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const encoder = new TextEncoder();

    function xml(value) {
        return String(value ?? "")
            .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }

    function columnLetter(index) {
        let result = "";
        for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
            result = String.fromCharCode(65 + (n - 1) % 26) + result;
        }
        return result;
    }

    function uniqueHeaders(columns) {
        const used = new Set();
        return columns.map(function (column, index) {
            const base = String(column?.name || "").trim() || "Colonne " + (index + 1);
            let name = base.slice(0, 250);
            for (let suffix = 2; used.has(name.toLowerCase()); suffix++) {
                name = base.slice(0, 245) + " (" + suffix + ")";
            }
            used.add(name.toLowerCase());
            return name;
        });
    }

    function styleFor(type, summary) {
        const base = { number: 2, currency: 3, percentage: 4, date: 5 }[type] || 0;
        return summary ? (base ? base + 6 : 7) : base;
    }

    function makeCell(address, value, type, summary, formula) {
        const style = styleFor(type, summary);
        const styleAttr = style ? ' s="' + style + '"' : "";
        if (formula) {
            const clean = String(formula).replace(/^=/, "");
            return '<c r="' + address + '"' + styleAttr + '><f>' + xml(clean) + '</f><v/></c>';
        }
        if (value === null || value === undefined || value === "") return "";
        if (typeof value === "number" && Number.isFinite(value)) {
            return '<c r="' + address + '"' + styleAttr + '><v>' + value + '</v></c>';
        }
        if (typeof value === "boolean") {
            return '<c r="' + address + '" t="b"' + styleAttr + '><v>' + (value ? 1 : 0) + '</v></c>';
        }
        return '<c r="' + address + '" t="inlineStr"' + styleAttr +
            '><is><t xml:space="preserve">' + xml(value) + '</t></is></c>';
    }

    function styles() {
        const formats = ['0.##', '#,##0.00 "€"', '0.00%', 'dd/mm/yyyy'];
        const numFmts = formats.map((format, i) =>
            '<numFmt numFmtId="' + (164 + i) + '" formatCode="' + xml(format) + '"/>').join("");
        const xfs = [
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
            ...[164, 165, 166, 167].map(id => '<xf numFmtId="' + id + '" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'),
            '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>',
            ...[164, 165, 166, 167].map(id => '<xf numFmtId="' + id + '" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/>')
        ];
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<styleSheet xmlns="' + ns + '"><numFmts count="4">' + numFmts + '</numFmts>' +
            '<fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b/><sz val="11"/><name val="Aptos"/></font></fonts>' +
            '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill></fills>' +
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
            '<cellXfs count="' + xfs.length + '">' + xfs.join("") + '</cellXfs>' +
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    }

    function crc32(bytes) {
        let crc = -1;
        for (const byte of bytes) {
            crc ^= byte;
            for (let bit = 0; bit < 8; bit++) {
                crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
            }
        }
        return (crc ^ -1) >>> 0;
    }

    function zipFiles(files) {
        const chunks = [];
        const directory = [];
        let offset = 0;
        for (const [path, contents] of Object.entries(files)) {
            const name = encoder.encode(path);
            const data = encoder.encode(contents);
            const crc = crc32(data);
            const local = new Uint8Array(30 + name.length);
            const view = new DataView(local.buffer);
            view.setUint32(0, 0x04034b50, true);
            view.setUint16(4, 20, true);
            view.setUint16(6, 0x0800, true);
            view.setUint32(14, crc, true);
            view.setUint32(18, data.length, true);
            view.setUint32(22, data.length, true);
            view.setUint16(26, name.length, true);
            local.set(name, 30);
            chunks.push(local, data);

            const central = new Uint8Array(46 + name.length);
            const c = new DataView(central.buffer);
            c.setUint32(0, 0x02014b50, true);
            c.setUint16(4, 20, true);
            c.setUint16(6, 20, true);
            c.setUint16(8, 0x0800, true);
            c.setUint32(16, crc, true);
            c.setUint32(20, data.length, true);
            c.setUint32(24, data.length, true);
            c.setUint16(28, name.length, true);
            c.setUint32(42, offset, true);
            central.set(name, 46);
            directory.push(central);
            offset += local.length + data.length;
        }
        const directorySize = directory.reduce((size, entry) => size + entry.length, 0);
        const end = new Uint8Array(22);
        const e = new DataView(end.buffer);
        e.setUint32(0, 0x06054b50, true);
        e.setUint16(8, directory.length, true);
        e.setUint16(10, directory.length, true);
        e.setUint32(12, directorySize, true);
        e.setUint32(16, offset, true);
        return new Blob([...chunks, ...directory, end], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        });
    }

    function createWorkbookBlob(data, helpers) {
        if (!Array.isArray(data?.columns) || data.columns.length === 0) {
            throw new Error("Le tableau ne contient aucune colonne.");
        }
        const columns = data.columns;
        const headers = uniqueHeaders(columns);
        const rows = Array.isArray(data.rows) && data.rows.length ? data.rows : [[], [], []];
        const lastColumn = columnLetter(columns.length - 1);
        let sheetLastColumn = lastColumn;
        const lastDataRow = rows.length + 1;
        const tableRef = 'A1:' + lastColumn + lastDataRow;
        const sheetRows = [];
        sheetRows.push('<row r="1">' + headers.map((name, i) =>
            '<c r="' + columnLetter(i) + '1" t="inlineStr" s="1"><is><t>' + xml(name) + '</t></is></c>').join("") + '</row>');
        rows.forEach(function (row, rowIndex) {
            const rowNumber = rowIndex + 2;
            const cells = columns.map(function (column, i) {
                const formula = column.formula && helpers.normalizeExcelFormula(column.formula, { row: rowNumber });
                const type = formula && ["text", "status", "boolean"].includes(column.type) ? "" : column.type;
                const value = helpers.coerceExcelValue(Array.isArray(row) ? row[i] : "", column.type);
                return makeCell(columnLetter(i) + rowNumber, value, type, false, formula);
            }).join("");
            sheetRows.push('<row r="' + rowNumber + '">' + cells + '</row>');
        });
        let lastRow = lastDataRow;
        const summaries = Array.isArray(data.summaryFormulas) ? data.summaryFormulas : [];
        summaries.forEach(function (summary, index) {
            const target = columns.findIndex(column => column.name === summary.targetColumn);
            if (target < 0 || !summary.formula) return;
            const rowNumber = rows.length + 2 + index;
            // Si le résultat est en A, placer son libellé en B pour garder les deux.
            const labelColumn = target === 0 ? 'B' : 'A';
            if (target === 0 && columns.length === 1) sheetLastColumn = 'B';
            const label = makeCell(labelColumn + rowNumber, summary.label || 'Indicateur', "", true);
            const formula = helpers.normalizeExcelFormula(summary.formula, { lastRow: lastDataRow });
            const value = makeCell(columnLetter(target) + rowNumber, "", columns[target].type, true, formula);
            sheetRows.push('<row r="' + rowNumber + '">' + (target === 0 ? value + label : label + value) + '</row>');
            lastRow = Math.max(lastRow, rowNumber);
        });
        const widths = headers.map((header, i) => {
            const width = Math.min(60, Math.max(14, header.length + 3));
            return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + width + '" customWidth="1"/>';
        }).join("");
        const sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<worksheet xmlns="' + ns + '" xmlns:r="' + relNs + '">' +
            '<dimension ref="A1:' + sheetLastColumn + lastRow + '"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
            '<sheetFormatPr defaultRowHeight="15"/><cols>' + widths + '</cols><sheetData>' + sheetRows.join("") + '</sheetData>' +
            '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>';
        const tableName = 'TableIA_' + Date.now().toString().slice(-10);
        const table = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<table xmlns="' + ns + '" id="1" name="' + tableName + '" displayName="' + tableName + '" ref="' + tableRef + '" totalsRowShown="0">' +
            '<autoFilter ref="' + tableRef + '"/><tableColumns count="' + columns.length + '">' +
            headers.map((name, i) => '<tableColumn id="' + (i + 1) + '" name="' + xml(name) + '"/>').join("") +
            '</tableColumns><tableStyleInfo name="TableStyleMedium7" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';
        const title = String(data.title || 'TableIA').replace(/[\\\/:?*\[\]]/g, ' ').trim().slice(0, 31) || 'TableIA';
        return zipFiles({
            '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>',
            '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
            'xl/workbook.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="' + ns + '" xmlns:r="' + relNs + '"><sheets><sheet name="' + xml(title) + '" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>',
            'xl/_rels/workbook.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
            'xl/styles.xml': styles(),
            'xl/worksheets/sheet1.xml': sheet,
            'xl/worksheets/_rels/sheet1.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
            'xl/tables/table1.xml': table
        });
    }

    window.TableIAXlsx = { createWorkbookBlob };
})();
