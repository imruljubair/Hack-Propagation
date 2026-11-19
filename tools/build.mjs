import fs from 'node:fs';
import ts from 'typescript';

const metadata = JSON.parse(fs.readFileSync(new URL('./office-script.json', import.meta.url)));
for (const name of ['main', 'RandomInputs', 'ClearButton']) {
  const body = fs.readFileSync(new URL(`../src/${name}.ts`, import.meta.url), 'utf8');
  const result = ts.transpileModule(body, {
    compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.None },
    reportDiagnostics: true,
  });
  const errors = (result.diagnostics || []).filter(d => d.category === ts.DiagnosticCategory.Error);
  if (errors.length) throw new Error(ts.formatDiagnosticsWithColorAndContext(errors, {
    getCanonicalFileName: p => p, getCurrentDirectory: () => process.cwd(), getNewLine: () => '\n',
  }));
  fs.mkdirSync(new URL('../scripts/', import.meta.url), { recursive: true });
  fs.writeFileSync(new URL(`../scripts/${name}.osts`, import.meta.url), JSON.stringify({
    ...metadata, body, description: `${metadata.description} Script: ${name}.`,
  }, null, 2) + '\n');
  console.log(`Built scripts/${name}.osts`);
}
