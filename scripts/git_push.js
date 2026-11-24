const { execSync } = require('child_process');
const path = require('path');

const git = 'C:/Users/USER/.gemini/antigravity/scratch/mingit/cmd/git.exe';
const cwd = 'C:/Users/USER/.gemini/antigravity/scratch/municipio-escolar';

function run(cmd) {
  try {
    const out = execSync(`"${git}" ${cmd}`, { cwd, encoding: 'utf8' });
    console.log(cmd, '->', out.trim());
    return { ok: true, out };
  } catch (e) {
    console.log(cmd, 'FAILED:');
    if (e.stdout) console.log('STDOUT:', e.stdout.trim());
    if (e.stderr) console.log('STDERR:', e.stderr.trim());
    return { ok: false, err: e.message };
  }
}

console.log('1. Setting up Git configuration...');
run('config user.name "jmanosalvapinedo-hash"');
run('config user.email "jmanosalva@champagnat.edu.pe"');
run('config credential.helper wincred');

console.log('2. Initializing repository...');
run('init');
run('branch -M main');

console.log('3. Setting remote origin...');
run('remote remove origin');
run('remote add origin https://github.com/jmanosalvapinedo-hash/elecciones-champagnat-2026.git');

console.log('4. Staging files...');
run('add .');

console.log('5. Committing changes...');
run('commit -m "Elecciones Municipio Escolar 2026-2027 - San Marcelino Champagnat"');

console.log('6. Pushing to GitHub main...');
const pushResult = run('push -u origin main --force');
console.log('Push finished:', pushResult.ok);
