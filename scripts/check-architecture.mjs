import { readFile, readdir, mkdtemp, writeFile, rm } from 'node:fs/promises';
import { dirname, resolve, relative, extname, sep, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { tmpdir } from 'node:os';
import { isBuiltin } from 'node:module';
import { spawnSync } from 'node:child_process';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const failures = [];
const fail = (file, message) => failures.push(`${relative(root, file)}: ${message}`);
const pure = new Set(['contracts', 'core', 'writer', 'sheets', 'slides', 'render', 'i18n']);
const allowed = {
  contracts: [], core: ['contracts'], writer: ['contracts', 'core'], sheets: ['contracts', 'core'], slides: ['contracts', 'core'],
  'format-common': ['contracts', 'core'], cfb: ['contracts', 'format-common'],
  ooxml: ['contracts', 'core', 'format-common', 'writer', 'sheets', 'slides'],
  'legacy-doc': ['contracts', 'cfb', 'format-common', 'writer'], 'legacy-xls': ['contracts', 'cfb', 'format-common', 'sheets'], 'legacy-ppt': ['contracts', 'cfb', 'format-common', 'slides'],
  render: ['contracts', 'core', 'writer', 'sheets', 'slides'], pdf: ['contracts', 'render'],
  'format-service': ['contracts', 'core', 'format-common', 'ooxml', 'cfb', 'legacy-doc', 'legacy-xls', 'legacy-ppt', 'writer', 'sheets', 'slides'],
  i18n: [], ui: ['contracts', 'core', 'format-common', 'writer', 'sheets', 'slides', 'render', 'format-service', 'i18n', 'pdf'], desktop: ['contracts', 'core', 'ui', 'i18n'],
};
const approvedExternal = { ooxml: ['fflate', '@xmldom/xmldom'], desktop: ['electron'], pdf: ['pdfjs-dist'] };
const exactVersion = /^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/;
const packages = new Map();
const rootManifest = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'));
for (const base of ['packages', 'apps']) {
  for (const entry of await readdir(join(root, base), { withFileTypes: true })) {
    if (!entry.isDirectory()) continue;
    const directory = join(root, base, entry.name);
    let manifest;
    try { manifest = JSON.parse(await readFile(join(directory, 'package.json'), 'utf8')); } catch { fail(directory, 'Missing or invalid workspace package.json'); continue; }
    if (manifest.name !== `@office-ex/${entry.name}`) fail(directory, 'Workspace name must match its directory');
    if (packages.has(manifest.name)) fail(directory, 'Duplicate workspace package name');
    if (!Object.hasOwn(allowed, entry.name)) fail(directory, 'New package needs an explicit architecture rule');
    if (base === 'packages' && manifest.exports?.['.'] !== './src/index.ts') fail(directory, 'Public package entry must be src/index.ts');
    packages.set(manifest.name, { name: entry.name, directory, manifest, imports: new Set() });
  }
}
for (const [name, version] of Object.entries({ ...rootManifest.dependencies, ...rootManifest.devDependencies })) {
  if (!exactVersion.test(version)) fail(join(root, 'package.json'), `${name} must use an exact dependency version`);
}
let lock;
try { lock = JSON.parse(await readFile(join(root, 'package-lock.json'), 'utf8')); } catch { fail(root, 'The root npm lockfile is missing or invalid'); }
if (lock && lock.lockfileVersion !== 3) fail(join(root, 'package-lock.json'), 'Expected npm lockfile v3');

async function sourceFiles(directory) {
  const result = [];
  let entries;
  try { entries = await readdir(directory, { withFileTypes: true }); } catch { return result; }
  for (const entry of entries) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) result.push(...await sourceFiles(path));
    else if (/\.(?:[cm]?[jt]sx?)$/.test(entry.name)) result.push(path);
  }
  return result;
}

// Import scanner skips comments and literal contents; non-literal dynamic loads are prohibited.
// It is a dependency linter, not an executable-content security sandbox.
function tokens(source) {
  const result = [];
  for (let index = 0; index < source.length;) {
    const character = source[index];
    if (/\s/.test(character)) { index++; continue; }
    if (source.startsWith('//', index)) { const end = source.indexOf('\n', index + 2); index = end < 0 ? source.length : end; continue; }
    if (source.startsWith('/*', index)) { const end = source.indexOf('*/', index + 2); index = end < 0 ? source.length : end + 2; continue; }
    if (character === '"' || character === "'" || character === '`') {
      const quote = character; const start = ++index;
      while (index < source.length && source[index] !== quote) { if (source[index] === '\\') index += 2; else index++; }
      result.push({ kind: quote === '`' ? 'template' : 'string', value: source.slice(start, index), position: start }); index++; continue;
    }
    if (/[A-Za-z_$]/.test(character)) {
      const start = index++; while (index < source.length && /[A-Za-z0-9_$]/.test(source[index])) index++;
      result.push({ kind: 'id', value: source.slice(start, index), position: start }); continue;
    }
    result.push({ kind: 'punctuation', value: character, position: index++ });
  }
  return result;
}
function moduleReferences(source, file) {
  const list = tokens(source); const references = [];
  for (let index = 0; index < list.length; index++) {
    const token = list[index]; if (token.kind !== 'id') continue;
    if ((token.value === 'eval' && list[index + 1]?.value === '(') || (token.value === 'new' && list[index + 1]?.value === 'Function')) fail(file, 'Dynamic executable strings are forbidden');
    if (token.value === 'require' || (token.value === 'import' && list[index + 1]?.value === '(')) {
      if (list[index + 1]?.value !== '(') continue;
      const specifier = list[index + 2];
      if (specifier?.kind !== 'string' || list[index + 3]?.value !== ')') fail(file, 'Dynamic module loads must use one literal specifier');
      else references.push({ specifier: specifier.value, typeOnly: false });
    } else if ((token.value === 'import' || token.value === 'export') && list[index + 1]?.value !== '.') {
      if (list[index + 1]?.kind === 'string' && token.value === 'import') references.push({ specifier: list[index + 1].value, typeOnly: false });
      else {
        for (let next = index + 1; next < list.length && next < index + 500; next++) {
          if (list[next].value === ';' || (next > index + 1 && list[next].kind === 'id' && ['import', 'export', 'function', 'class', 'const', 'let'].includes(list[next].value))) break;
          if (list[next].value === 'from' && list[next + 1]?.kind === 'string') { references.push({ specifier: list[next + 1].value, typeOnly: list[index + 1]?.value === 'type' }); break; }
        }
      }
    }
  }
  return references;
}
const pureFiles = [];
let fileCount = 0;
for (const workspace of packages.values()) {
  const dependencies = { ...workspace.manifest.dependencies, ...workspace.manifest.devDependencies };
  for (const [name, version] of Object.entries(dependencies)) {
    if (!exactVersion.test(version)) fail(workspace.directory, `${name} must use an exact dependency version`);
    if (name.startsWith('@office-ex/')) {
      const target = packages.get(name);
      if (!target) fail(workspace.directory, `Unknown workspace dependency ${name}`);
      else if (target.manifest.version !== version) fail(workspace.directory, `${name} version differs from the workspace version`);
      if (!allowed[workspace.name]?.includes(name.slice('@office-ex/'.length))) fail(workspace.directory, `Forbidden dependency ${name}`);
      workspace.imports.add(name);
    } else if (!(approvedExternal[workspace.name] ?? []).includes(name)) fail(workspace.directory, `External dependency ${name} needs explicit review`);
  }
  const files = await sourceFiles(join(workspace.directory, 'src')); fileCount += files.length;
  if (!files.length) fail(workspace.directory, 'Workspace has no source files');
  if (pure.has(workspace.name)) pureFiles.push(...files.filter((file) => !file.endsWith('.d.ts')));
  for (const file of files) {
    const source = await readFile(file, 'utf8');
    for (const { specifier, typeOnly } of moduleReferences(source, file)) {
      if (specifier.includes('\\') || specifier.includes('\0')) { fail(file, 'Escaped module specifiers are forbidden'); continue; }
      if (specifier.startsWith('.')) {
        const target = resolve(dirname(file), specifier); const local = relative(workspace.directory, target);
        if (local === '..' || local.startsWith(`..${sep}`) || local.startsWith(sep)) fail(file, 'Relative import crosses a workspace boundary');
      } else if (specifier.startsWith('@office-ex/')) {
        if (!packages.has(specifier)) fail(file, `Use an existing workspace public API, not ${specifier}`);
        else if (!dependencies[specifier]) fail(file, `Undeclared workspace dependency ${specifier}`);
        if (workspace.name === 'ui' && specifier === '@office-ex/format-service' && !typeOnly && !file.endsWith(`${sep}format-worker.ts`)) fail(file, 'Runtime parsers must stay in the dedicated format worker');
        if (workspace.name === 'desktop' && specifier === '@office-ex/ui' && !typeOnly) fail(file, 'Desktop main may import only UI protocol types');
      } else if (isBuiltin(specifier)) {
        if (workspace.name !== 'desktop') fail(file, `Node builtin ${specifier} is forbidden in this package`);
      } else {
        const name = specifier.startsWith('@') ? specifier.split('/').slice(0, 2).join('/') : specifier.split('/')[0];
        if (!(approvedExternal[workspace.name] ?? []).includes(name)) fail(file, `Unapproved platform/external import ${specifier}`);
        if (!dependencies[name] && !(workspace.name === 'desktop' && rootManifest.devDependencies?.[name])) fail(file, `Undeclared external dependency ${name}`);
      }
    }
  }
}

const complete = new Set(); const visiting = new Set();
function visit(name, path = []) {
  if (visiting.has(name)) { fail(packages.get(name)?.directory ?? root, `Workspace dependency cycle: ${[...path, name].join(' -> ')}`); return; }
  if (complete.has(name)) return;
  visiting.add(name);
  for (const dependency of packages.get(name)?.imports ?? []) visit(dependency, [...path, name]);
  visiting.delete(name); complete.add(name);
}
for (const name of packages.keys()) visit(name);

// A separate compiler pass removes all DOM and Node ambient declarations from the domain.
// TextEncoder/TextDecoder are neutral byte codecs implemented by both target runtimes.
const temporary = await mkdtemp(join(tmpdir(), 'office-ex-architecture-'));
try {
  const globals = join(temporary, 'neutral-globals.d.ts');
  await writeFile(globals, 'declare class TextEncoder { encode(input?: string): Uint8Array<ArrayBuffer>; }\ndeclare class TextDecoder { constructor(label?: string, options?: {fatal?: boolean;ignoreBOM?: boolean}); decode(input?: ArrayBufferView | ArrayBuffer, options?: {stream?: boolean}): string; }\n');
  const config = join(temporary, 'tsconfig.json');
  await writeFile(config, JSON.stringify({ compilerOptions: { target: 'ES2022', lib: ['ES2022'], module: 'NodeNext', moduleResolution: 'NodeNext', strict: true, noEmit: true, types: [], skipLibCheck: true, verbatimModuleSyntax: true }, files: [...pureFiles, globals] }));
  const check = spawnSync(process.execPath, [join(root, 'node_modules/typescript/bin/tsc'), '--project', config], { encoding: 'utf8', cwd: root });
  if (check.status !== 0) failures.push(`Pure-domain typecheck failed (no DOM/Node globals):\n${check.stdout}${check.stderr}${check.error?.message ?? ''}`);
} finally { await rm(temporary, { recursive: true, force: true }); }
if (failures.length) { console.error(failures.join('\n')); process.exitCode = 1; }
else console.log(`Architecture verified: ${packages.size} workspaces, ${fileCount} source files; public APIs, declared dependencies, acyclic graph, pure domain.`);
