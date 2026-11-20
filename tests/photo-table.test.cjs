const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const app = fs.readFileSync(path.join(root, 'app.js'), 'utf8');
const functions = app.slice(app.indexOf('function normalizeExcelFormula('), app.indexOf('function excelNumberFormatForType('));
const context = { Date, Number, String, Intl, console };
vm.createContext(context);
vm.runInContext(functions, context);

const photoResponse = {
    title: 'Dépenses',
    columns: [{ name: 'Nom', type: 'text' }, { name: 'Montant', type: 'text' }],
    rows: [['Alice', '12,50 €'], ['Bob', '1.234,56 €'], ['Total', '1.247,06 €']]
};

test('photo avec noms et montants : colonnes, types et somme sans doublon', () => {
    const result = context.structurePhotoTable(photoResponse);
    assert.equal(result.columns[0].name, 'Nom');
    assert.equal(result.columns[1].type, 'currency');
    assert.equal(result.rows.length, 2);
    assert.equal(result.summaryFormulas.length, 1);
    assert.equal(result.summaryFormulas[0].formula, '=SUM(B2:B{lastRow})');
    assert.equal(context.coerceExcelValue(result.rows[1][1], 'currency'), 1234.56);
    assert.equal(context.structurePhotoTable(result).summaryFormulas.length, 1);
    assert.equal(photoResponse.rows.length, 3);

    const existing = context.structurePhotoTable({
        ...photoResponse,
        summaryFormulas: [{ label: 'Total', targetColumn: 'Montant', formula: '=SUM(B2:B4)' }]
    });
    assert.equal(existing.summaryFormulas.length, 1);
    assert.equal(existing.summaryFormulas[0].formula, '=SUM(B2:B{lastRow})');
});

test('photo sans montant : aucune somme inventée', () => {
    const result = context.structurePhotoTable({
        columns: [{ name: 'Nom', type: 'text' }], rows: [['Alice'], ['Bob']]
    });
    assert.equal(result.summaryFormulas, undefined);
});

test('aperçu : la ligne Total affiche la somme calculée', () => {
    const nodes = [];
    const makeElement = tag => ({
        tag, children: [], style: {}, textContent: '', className: '',
        appendChild(child) { this.children.push(child); },
        set innerHTML(value) { this.children = []; }
    });
    context.document = { createElement: makeElement };
    context.textAnalyzeButton = {};
    context.tableContainer = makeElement('div');
    context.resultSection = { style: {} };
    context.insertExcelButton = {};
    vm.runInContext('let currentTableData;\n' + app.slice(app.indexOf('function displayAiTable('), app.lastIndexOf('if (isWebConversation) {')), context);
    context.displayAiTable(context.structurePhotoTable(photoResponse));
    const table = context.tableContainer.children[0];
    const footer = table.children[table.children.length - 1];
    assert.equal(footer.className, 'summary-row');
    assert.equal(footer.children[0].textContent, 'Total');
    assert.match(footer.children[1].textContent, /1\s?247,06\s?€/);
});

test('XLSX : données numériques et formule de somme sous le tableau', async () => {
    const browser = { window: {}, Blob, TextEncoder, Uint8Array, DataView, Date };
    vm.runInNewContext(fs.readFileSync(path.join(root, 'xlsx-export.js'), 'utf8'), browser);
    const data = context.structurePhotoTable(photoResponse);
    const blob = browser.window.TableIAXlsx.createWorkbookBlob(data, {
        coerceExcelValue: context.coerceExcelValue,
        normalizeExcelFormula: context.normalizeExcelFormula
    });
    const bytes = new Uint8Array(await blob.arrayBuffer());
    if (process.env.TABLEIA_TEST_XLSX) fs.writeFileSync(process.env.TABLEIA_TEST_XLSX, bytes);
    const view = new DataView(bytes.buffer);
    const entries = {};
    for (let offset = 0; view.getUint32(offset, true) === 0x04034b50;) {
        const size = view.getUint32(offset + 18, true);
        const nameLength = view.getUint16(offset + 26, true);
        const start = offset + 30 + nameLength;
        const name = new TextDecoder().decode(bytes.slice(offset + 30, start));
        entries[name] = new TextDecoder().decode(bytes.slice(start, start + size));
        offset = start + size;
    }
    assert.match(entries['xl/tables/table1.xml'], /ref="A1:B3"/);
    assert.match(entries['xl/worksheets/sheet1.xml'], /<c r="B2" s="3"><v>12\.5<\/v><\/c>/);
    assert.match(entries['xl/worksheets/sheet1.xml'], /<c r="B3" s="3"><v>1234\.56<\/v><\/c>/);
    assert.match(entries['xl/worksheets/sheet1.xml'], /<row r="4">/);
    assert.match(entries['xl/worksheets/sheet1.xml'], /<f>SUM\(B2:B3\)<\/f>/);
});
