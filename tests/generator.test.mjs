import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import ts from 'typescript';
import { Workbook, ExcelScript, formulas } from './excel-mock.mjs';

const product = JSON.parse(fs.readFileSync(new URL('../package.json',import.meta.url))).name;
const forward = product === 'the-architect';
const source = fs.readFileSync(new URL('../src/main.ts',import.meta.url),'utf8');
function load(text=source) {
  const js=ts.transpileModule(text,{compilerOptions:{target:ts.ScriptTarget.ES2020,module:ts.ModuleKind.None}}).outputText;
  const context=vm.createContext({ExcelScript,console:{log(){}}});
  vm.runInContext(js,context);
  return context.main;
}
const main=load();
const cases={
  classification:[['Rand',5,3],['Linear',4,''],['ReLU','',''],['Linear',3,''],['Softmax','',''],['CELoss','','']],
  regression:[['Rand',3,2],['Linear',4,''],['Tanh','',''],['Linear',2,''],['MSELoss','','']],
  normalization:[['Rand',3,2],['Linear',4,''],['LayerNorm','',''],['GELU','',''],['Linear',2,''],['MSELoss','','']],
  image:[['Image',3,2],['Flatten','',''],['Linear',3,''],['Softmax','',''],['CELoss','','']],
  tinyDigits:[['tinyDigits','',2],['Flatten','',''],['Linear',4,''],['ReLU','',''],['Linear',10,''],['Softmax','',''],['CELoss','','']],
};
for(const [name,layers] of Object.entries(cases)) {
  test(`${name}: complete generation and rerun`,()=>{
    const book=new Workbook(layers,2);
    // An unrelated worksheet must survive cleanup.
    book.addWorksheet('notes_epilogue');
    for(let run=0;run<2;run++) {
      main(book);
      assert.ok(book.getWorksheet('Blueprint'));
      assert.ok(book.getWorksheet('notes_epilogue'));
      const entries=[...book.getWorksheet('Blueprint').cells.values()];
      const text=entries.map(v=>v.value??v.formula??'').join('\n');
      assert.ok(formulas(book).some(([,v])=>v.formula.includes('MMULT(')));
      assert.doesNotMatch(text,/nn.Sequential|PyTorch Code|torchcode|code_layer_|optimizer|\.backward\(/);
      for(const sheet of book.getWorksheets()) {
        const output=[...sheet.cells.values()].map(v=>v.value??v.formula??'').join('\n');
        assert.doesNotMatch(output,/PyTorch Code|torchcode|code_layer_|nn.Sequential/);
      }
      if(forward) {
        assert.deepEqual([...book.sheets.keys()].sort(),['Architect','Blueprint','Helper','notes_epilogue','tinyDigits'].sort());
        assert.doesNotMatch(text,/Backprop|Weight Update|backward\(\)|optimizer|Training Loop/);
        assert.equal(book.getWorksheet('Architect').getRange('G5').getValue(),'');
      } else {
        for(const sheet of ['Blueprint_ep1','Blueprint_inference','Dashboard']) assert.ok(book.getWorksheet(sheet),sheet);
        assert.match(text,/Backprop/);
        assert.match(text,/Weight Update/);
        assert.ok(book.getWorksheet('Dashboard').events.includes('chart'));
        const epFormula=formulas(book,'Blueprint_ep1').map(([,v])=>v.formula).join('\n');
        assert.match(epFormula,/=Blueprint!newW_/);
        assert.equal(book.getWorksheet('Architect').getRange('G5').getValue(),'Backpropagation');
      }
    }
  });
}
test('source and distributed Office Scripts stay in sync',()=>{
  for(const name of ['main','ClearButton','RandomInputs']) {
    const body=fs.readFileSync(new URL(`../src/${name}.ts`,import.meta.url),'utf8');
    const packaged=JSON.parse(fs.readFileSync(new URL(`../scripts/${name}.osts`,import.meta.url)));
    assert.equal(packaged.body,body);
  }
});
test('clear removes generated sheets while preserving the product mode and unrelated sheets',()=>{
  const book=new Workbook(cases.classification,2);
  main(book);
  book.addWorksheet('MyBlueprintNotes');
  const clear=load(fs.readFileSync(new URL('../src/ClearButton.ts',import.meta.url),'utf8'));
  clear(book);
  assert.ok(book.getWorksheet('MyBlueprintNotes'));
  assert.equal(book.getWorksheet('Blueprint'),undefined);
  assert.equal(book.getWorksheet('Dashboard'),undefined);
  assert.equal(book.getWorksheet('Architect').getRange('G5').getValue(),forward ? '' : 'Backpropagation');
});
test('random architecture tool preserves the mode and generates a terminal loss',()=>{
  const book=new Workbook(cases.classification);
  const random=load(fs.readFileSync(new URL('../src/RandomInputs.ts',import.meta.url),'utf8'));
  random(book);
  const layers=book.getWorksheet('Architect').getRange('B5:B50').getValues().map(row=>row[0]).filter(Boolean);
  assert.ok(['CELoss','MSELoss'].includes(layers.at(-1)));
  assert.equal(book.getWorksheet('Architect').getRange('G5').getValue(),forward ? '' : 'Backpropagation');
});
if(forward) {
  test('works without a loss and cannot enable training through old controls',()=>{
    const book=new Workbook([['Rand',3,2],['Linear',2,''],['ReLU','','']],99);
    book.getWorksheet('Architect').getRange('G5').setValue(true);
    main(book);
    assert.equal(book.sheets.size,4);
    assert.ok(book.getWorksheet('Blueprint'));
    assert.doesNotMatch(source,/backpropFlag|\.backward\(|torch\.optim|addWorksheet\("Dashboard"\)/);
  });
} else {
  test('invalid settings fail before deleting prior generated results',()=>{
    for(const epochs of [0,-1,1.5,201,'bad']) {
      const book=new Workbook(cases.classification,epochs);
      const previous=book.addWorksheet('Dashboard');
      assert.throws(()=>main(book),/whole epoch count/);
      assert.equal(book.getWorksheet('Dashboard'),previous);
    }
    const book=new Workbook([['Rand',3,2],['Linear',2,'']]);
    const previous=book.addWorksheet('Dashboard');
    assert.throws(()=>main(book),/final layer/);
    assert.equal(book.getWorksheet('Dashboard'),previous);
  });
  test('one epoch still creates the dashboard and no extra epoch sheet',()=>{
    const book=new Workbook(cases.classification,1);
    main(book);
    assert.ok(book.getWorksheet('Dashboard'));
    assert.equal(book.getWorksheet('Blueprint_ep1'),undefined);
  });
}
