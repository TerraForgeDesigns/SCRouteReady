import { mkdir } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { build } from 'esbuild';

await mkdir('.test-dist', { recursive: true });

await build({
  entryPoints: ['./tests/pullSheetGenerator.test.ts'],
  bundle: true,
  format: 'esm',
  platform: 'node',
  outfile: '.test-dist/pullSheetGenerator.test.mjs',
});

const result = spawnSync(process.execPath, ['--test', '.test-dist/pullSheetGenerator.test.mjs'], {
  stdio: 'inherit',
});

process.exit(result.status ?? 1);
