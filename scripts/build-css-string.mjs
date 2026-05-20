// Pre-compiles src/styles/index.scss into a TypeScript constant export so the
// Application Customizer can inject the brand CSS manually via document.head
// + a <style id="phil-brand"> tag. This bypasses sp-css-loader's themable-
// styles pipeline, which silently drops non-themable `:root` declarations on
// modern SharePoint Online (see git log for Prompt 4 attempt 1).
//
// Wired into npm scripts (`prestart`, `prebuild`, `postinstall`) so a fresh
// clone + `npm ci` always produces the generated file before anything else
// needs to read it.
//
// Output is intentionally minified — saves ~700 bytes off the customizer
// bundle since the entire stylesheet is now a JS string literal.

// Use sass-embedded (the Rust-backed Sass implementation) since it's already
// installed as a transitive dep of @rushstack/heft-sass-plugin. Avoids adding
// a redundant top-level dep on the Dart-JS `sass` package.
import { compile } from 'sass-embedded';
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, '..');

const SCSS_ENTRY = resolve(ROOT, 'src', 'styles', 'index.scss');
const TS_OUT = resolve(
  ROOT,
  'src',
  'extensions',
  'phillipsBrand',
  'generated',
  'phillipsBrandCss.ts'
);

const result = compile(SCSS_ENTRY, { style: 'compressed' });

const banner = `// GENERATED FILE — DO NOT EDIT.
// Produced by scripts/build-css-string.mjs from src/styles/index.scss.
// Regenerate by running \`npm run build:css-string\` (or \`npm start\` /
// \`npm run build\`, both of which depend on it). This file is gitignored.

`;

const body = `export const PHIL_BRAND_CSS: string = ${JSON.stringify(result.css)};\n`;

await mkdir(dirname(TS_OUT), { recursive: true });
await writeFile(TS_OUT, banner + body, 'utf8');

const relativeOut = TS_OUT.replace(ROOT, '').replace(/\\/g, '/').replace(/^\//, '');
console.log(`[build-css-string] wrote ${relativeOut} (${result.css.length} bytes of CSS)`);
