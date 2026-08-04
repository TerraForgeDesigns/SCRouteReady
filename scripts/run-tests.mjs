import { mkdir, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';

const testOutDir = resolve('.test-dist');

await rm(testOutDir, { recursive: true, force: true });
await mkdir(testOutDir, { recursive: true });
await writeFile(resolve(testOutDir, 'package.json'), '{"type":"commonjs"}\n');

const compileResult = spawnSync(process.execPath, [
  resolve('node_modules/typescript/bin/tsc'),
  '-p',
  resolve('tsconfig.test.json'),
], {
  stdio: 'inherit',
});

if (compileResult.status !== 0) {
  process.exit(compileResult.status ?? 1);
}

const testResult = spawnSync(process.execPath, ['--test', resolve('.test-dist/tests/pullSheetGenerator.test.js')], {
  stdio: 'inherit',
});

process.exit(testResult.status ?? 1);
