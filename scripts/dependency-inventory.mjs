import { readFile, readdir, mkdir, writeFile, copyFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { dirname, join, resolve, relative, sep } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const INVENTORY = join(ROOT, 'docs/DEPENDENCIES.lock.json');
const PERMISSIVE = new Set(['MIT', 'Apache-2.0', 'ISC', 'BSD-2-Clause', 'BSD-3-Clause']);
const RUNTIME_LIBRARIES = new Set(['fflate', '@xmldom/xmldom', 'pdfjs-dist']);
const FORBIDDEN_ENGINES = /(?:onlyoffice|libreoffice|collabora|univer|officegen|exceljs|pptxgenjs)/i;
const SOURCE_OVERRIDES = {
  electron: 'https://github.com/electron/electron', fflate: 'https://github.com/101arrowz/fflate', '@xmldom/xmldom': 'https://github.com/xmldom/xmldom',
  typescript: 'https://github.com/microsoft/TypeScript', esbuild: 'https://github.com/evanw/esbuild', tsx: 'https://github.com/privatenumber/tsx',
  'playwright-core': 'https://github.com/microsoft/playwright', '@types/node': 'https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/node',
  fsevents: 'https://github.com/fsevents/fsevents', 'pdfjs-dist': 'https://github.com/mozilla/pdf.js', '@napi-rs/canvas':'https://github.com/Brooooooklyn/canvas',
};
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const stable = (value) => `${JSON.stringify(value, null, 2)}\n`;
const sortedEntries = (value) => JSON.stringify(Object.entries(value ?? {}).sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0));
const fail = (message) => { throw new Error(`Dependency audit: ${message}`); };
async function maybeJson(path) { try { return JSON.parse(await readFile(path, 'utf8')); } catch (error) { if (error.code === 'ENOENT') return undefined; throw error; } }
function repositoryUrl(value) {
  let url = typeof value === 'string' ? value : value?.url;
  if (!url) return undefined;
  if (/^[\w.-]+\/[\w.-]+$/.test(url)) url = `https://github.com/${url}`;
  return url.replace(/^git\+/, '').replace(/^git:\/\//, 'https://').replace(/\.git$/, '');
}
function packageName(path, item) { return item.name ?? path.slice(path.lastIndexOf('node_modules/') + 'node_modules/'.length); }
function noticeDirectory(name, version) { return join('notices', `${name.replace(/^@/, '').replaceAll('/', '__')}@${version}`); }
function locateDependency(packages, importer, name) {
  let current = importer;
  while (true) {
    const candidate = current ? `${current}/node_modules/${name}` : `node_modules/${name}`;
    if (packages[candidate]) return candidate;
    if (!current) return undefined;
    const index = current.lastIndexOf('/node_modules/');
    current = index >= 0 ? current.slice(0, index) : '';
  }
}
async function topLevelNotices(directory) {
  let entries; try { entries = await readdir(directory, { withFileTypes: true }); } catch (error) { if (error.code === 'ENOENT') return []; throw error; }
  return entries.filter((entry) => entry.isFile() && /^(?:licen[cs]e|copying|notice)(?:[._-].*)?$/i.test(entry.name)).map((entry) => entry.name).sort();
}

async function inspectLock() {
  const lockBytes = await readFile(join(ROOT, 'package-lock.json')); const lock = JSON.parse(lockBytes);
  if (lock.lockfileVersion !== 3 || !lock.packages) fail('Expected npm lockfile v3');
  const rootPackage = await maybeJson(join(ROOT, 'package.json'));
  if (!rootPackage || rootPackage.license !== 'MIT') fail('Application license must remain MIT');
  for (const field of ['dependencies', 'devDependencies', 'optionalDependencies']) if (sortedEntries(rootPackage[field]) !== sortedEntries(lock.packages['']?.[field])) fail(`Root ${field} differs from lock`);
  const direct = new Set(Object.keys({ ...rootPackage.dependencies, ...rootPackage.devDependencies }));
  const productionRoots = new Set();
  for (const [path, item] of Object.entries(lock.packages)) {
    if (!path || path.startsWith('node_modules/')) continue;
    const installedWorkspace = await maybeJson(join(ROOT, path, 'package.json'));
    if (!installedWorkspace || installedWorkspace.version !== item.version) fail(`Workspace manifest does not match lock: ${path}`);
    for (const field of ['dependencies', 'devDependencies', 'optionalDependencies']) if (sortedEntries(installedWorkspace[field]) !== sortedEntries(item[field])) fail(`Workspace ${field} differs from lock: ${path}`);
    for (const [name, version] of Object.entries(item.dependencies ?? {})) {
      if (installedWorkspace.dependencies?.[name] !== version) fail(`Workspace dependency differs from lock: ${path}/${name}`);
      if (!name.startsWith('@office-ex/')) { productionRoots.add(name); direct.add(name); }
    }
  }
  if (productionRoots.size !== RUNTIME_LIBRARIES.size || [...productionRoots].some((name) => !RUNTIME_LIBRARIES.has(name))) fail('Production external dependencies changed; review their purpose and licenses first');
  for (const [name, version] of Object.entries({ ...rootPackage.dependencies, ...rootPackage.devDependencies })) {
    if (!/^\d+\.\d+\.\d+(?:-[\w.-]+)?$/.test(version)) fail(`Direct dependency is not pinned: ${name}`);
    if (lock.packages[`node_modules/${name}`]?.version !== version) fail(`Direct dependency differs from lock: ${name}`);
  }
  const inventory = []; const notices = []; const installed = new Map();
  for (const [path, item] of Object.entries(lock.packages).sort(([a], [b]) => a < b ? -1 : 1)) {
    if (!path.startsWith('node_modules/') || item.link) continue;
    const name = packageName(path, item);
    if (!/^(?:@[A-Za-z0-9_.-]+\/)?[A-Za-z0-9_.-]+$/.test(name) || name === '..' || !/^\d+\.\d+\.\d+(?:-[\w.-]+)?$/.test(item.version)) fail(`Invalid locked package identity: ${path}`);
    if (path.split('/').some((part) => part === '..' || part === '.') || path.includes('\\')) fail(`Unsafe locked package path: ${path}`);
    if (FORBIDDEN_ENGINES.test(name)) fail(`Office engine/converter dependency is forbidden: ${name}`);
    const specialToolLicense = name === 'graceful-fs' && item.version === '4.2.12' && item.dev === true && item.license === 'BlueOak-1.0.0';
    if (!PERMISSIVE.has(item.license) && !specialToolLicense) fail(`License needs explicit review: ${name}@${item.version}: ${item.license ?? 'missing'}`);
    const pdfNodeOptional = (name === '@napi-rs/canvas' || /^@napi-rs\/canvas-(?:android|darwin|linux|win32)-/.test(name)) && item.version === '1.2.0' && item.optional === true && item.license === 'MIT';
    if (!item.dev && !RUNTIME_LIBRARIES.has(name) && !pdfNodeOptional) fail(`Unexpected production transitive dependency: ${name}`);
    if (!item.integrity?.startsWith('sha512-') || !item.resolved?.startsWith('https://registry.npmjs.org/')) fail(`Missing pinned registry origin/integrity: ${name}`);
    const local = await maybeJson(join(ROOT, path, 'package.json'));
    if (!local && !item.optional) fail(`Required installed package is missing: ${name}`);
    if (local) {
      if (local.name !== name || local.version !== item.version || local.license !== item.license) fail(`Installed package metadata differs from lock: ${name}`);
      installed.set(name, { path, item, local });
    }
    const origin = SOURCE_OVERRIDES[name] ?? (name.startsWith('@esbuild/') ? SOURCE_OVERRIDES.esbuild : name.startsWith('@typescript/typescript-') ? SOURCE_OVERRIDES.typescript : repositoryUrl(local?.repository));
    const dependencies = [];
    for (const [dependency, requested] of Object.entries({ ...item.dependencies, ...item.optionalDependencies }).sort(([a], [b]) => a < b ? -1 : 1)) {
      const resolvedPath = locateDependency(lock.packages, path, dependency);
      const optional = Object.hasOwn(item.optionalDependencies ?? {}, dependency);
      if (!resolvedPath && !optional) fail(`Unresolved locked dependency: ${name} -> ${dependency}`);
      dependencies.push({ name: dependency, requested, optional, path: resolvedPath ?? null, version: resolvedPath ? lock.packages[resolvedPath].version ?? null : null });
    }
    inventory.push({ name, version: item.version, path, direct: direct.has(name), usage: pdfNodeOptional ? 'optional-pdf-node-adapter-not-bundled' : RUNTIME_LIBRARIES.has(name) ? 'runtime-library' : name === 'electron' ? 'runtime-host-package' : 'development-or-installation', license: item.license, optional: item.optional === true, os: item.os ?? null, cpu: item.cpu ?? null, origin: origin ?? item.resolved, tarball: item.resolved, integrity: item.integrity, dependencies });
    if (local) {
      const names = await topLevelNotices(join(ROOT, path));
      if(name==='pdfjs-dist')for(const directory of ['cmaps','standard_fonts','iccs'])for(const filename of await topLevelNotices(join(ROOT,path,directory)))names.push(directory+'/'+filename);
      if ((direct.has(name) || RUNTIME_LIBRARIES.has(name)) && !names.some((file) => /^licen[cs]e/i.test(file))) fail(`Required license text missing: ${name}`);
      for (const filename of names) {
        const source = `${path}/${filename}`; const bytes = await readFile(join(ROOT, source));
        if (bytes.length > 2 * 1024 * 1024) fail(`Unexpectedly large npm notice; review before copying: ${source}`);
        const destination = `${noticeDirectory(name, item.version).split(sep).join('/')}/${filename}`;
        notices.push({ name, version: item.version, source, file: destination, sha256: hash(bytes), byteLength: bytes.length, bytes });
      }
    }
  }
  return { lock, lockSha256: hash(lockBytes), rootPackage, packages: inventory, notices, installed };
}

/** Verify the pinned dependency inventory; optionally refresh it or stage redistributable notices. */
export async function auditDependencies({ write = false, stageDir } = {}) {
  const current = await inspectLock();
  if (write) {
    const inventory = { schemaVersion: 1, reviewedOn: '2026-11-25', lockSha256: current.lockSha256, policy: 'docs/DEPENDENCIES.md', packages: current.packages, noticeFiles: current.notices.map(({ bytes, ...item }) => item) };
    await writeFile(INVENTORY, stable(inventory));
    for (const notice of current.notices) { const target = join(ROOT, notice.file); await mkdir(dirname(target), { recursive: true }); await writeFile(target, notice.bytes); }
  } else {
    const previous = await maybeJson(INVENTORY);
    if (!previous || previous.lockSha256 !== current.lockSha256) fail('Committed inventory does not match package-lock.json; review changes and run --write');
    // Other platforms install different optional native tools. The locked metadata is common.
    if (previous.packages.length !== current.packages.length || previous.packages.some((pkg, index) => ['name', 'version', 'path', 'license', 'integrity', 'tarball', 'usage'].some((key) => pkg[key] !== current.packages[index]?.[key]))) fail('Committed package inventory differs from lock');
    for (const notice of previous.noticeFiles) {
      const saved = await readFile(join(ROOT, notice.file)); if (hash(saved) !== notice.sha256) fail(`Committed notice changed: ${notice.file}`);
      const packagePath = notice.source.slice(0, notice.source.lastIndexOf('/'));
      let source; try { source = await readFile(join(ROOT, notice.source)); } catch (error) {
        if (error.code !== 'ENOENT' || !current.lock.packages[packagePath]?.optional) throw error;
      }
      if (source && hash(source) !== notice.sha256) fail(`Installed notice differs from reviewed text: ${notice.source}`);
    }
  }
  const fontManifest=JSON.parse(await readFile(join(ROOT,'assets/fonts/manifest.json'),'utf8'));
  for(const font of fontManifest.files){if(!/^[A-Za-z0-9_.-]+$/.test(font.file)||font.license!=='OFL-1.1'||font.faces.some(face=>face.fsType!==0))fail('Unreviewed bundled font');const bytes=await readFile(join(ROOT,'assets/fonts',font.file));if(bytes.length!==font.bytes||hash(bytes)!==font.sha256)fail('Bundled font differs from reviewed original: '+font.file);}
  for(const notice of fontManifest.notices){if(!/^[A-Za-z0-9_.-]+$/.test(notice.file))fail('Invalid font notice path');const bytes=await readFile(join(ROOT,'assets/fonts',notice.file));if(bytes.length!==notice.bytes||hash(bytes)!==notice.sha256)fail('Font license notice differs from reviewed original: '+notice.file);}
  let nativeDistribution;
  if (stageDir) {
    const destination = resolve(ROOT, stageDir); const local = relative(ROOT, destination);
    if (local.startsWith(`..${sep}`) || local === '..' || !local) fail('Notice staging destination must be a directory inside the repository');
    await mkdir(destination, { recursive: true });
    await copyFile(join(ROOT, 'LICENSE'), join(destination, 'Office-EX-LICENSE'));
    await copyFile(join(ROOT, 'THIRD_PARTY_NOTICES.md'), join(destination, 'THIRD_PARTY_NOTICES.md'));
    for (const notice of current.notices) {
      const target = join(destination, notice.file); await mkdir(dirname(target), { recursive: true }); await writeFile(target, notice.bytes);
    }
    const nativeDirectory = join(ROOT, 'node_modules/electron/dist');
    const version = (await readFile(join(nativeDirectory, 'version'), 'utf8')).trim().replace(/^v/, '');
    if (version !== current.installed.get('electron')?.item.version) fail('Native Electron version differs from the pinned npm package');
    nativeDistribution = { version, platform: process.platform, arch: process.arch, notices: [] };
    for (const filename of ['LICENSE', 'LICENSES.chromium.html']) {
      const source = join(nativeDirectory, filename); const bytes = await readFile(source);
      if (!bytes.length) fail(`Native Electron notice is empty: ${filename}`);
      const target = join(destination, 'electron-distribution', filename); await mkdir(dirname(target), { recursive: true }); await copyFile(source, target);
      nativeDistribution.notices.push({ file: `electron-distribution/${filename}`, byteLength: bytes.length, sha256: hash(bytes) });
    }
    const fontDestination=join(destination,'fonts');await mkdir(fontDestination,{recursive:true});
    for(const file of ['manifest.json','NotoCore-COPYRIGHT.txt','NotoCJK-COPYRIGHT.txt'])await copyFile(join(ROOT,'assets/fonts',file),join(fontDestination,file));
    await copyFile(INVENTORY, join(destination, 'DEPENDENCIES.lock.json'));
    await writeFile(join(destination, 'electron-distribution.json'), stable(nativeDistribution));
  }
  const summary = { packages: current.packages.length, runtimeLibraries: [...RUNTIME_LIBRARIES], npmNoticeFiles: current.notices.length, lockSha256: current.lockSha256, ...(nativeDistribution ? { nativeDistribution } : {}) };
  return summary;
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const args = process.argv.slice(2);
  if (args.some((arg) => !['--write', '--stage'].includes(arg))) fail('Supported arguments: --write, --stage');
  const result = await auditDependencies({ write: args.includes('--write'), stageDir: args.includes('--stage') ? 'dist/licenses' : undefined });
  console.log(`Dependency audit passed: ${result.packages} locked npm packages, ${result.npmNoticeFiles} installed notice files; ${result.runtimeLibraries.join(', ')} runtime libraries${result.nativeDistribution ? `; Electron ${result.nativeDistribution.version} native notices staged` : ''}.`);
}
