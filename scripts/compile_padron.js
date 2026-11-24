const fs = require('fs');
const path = require('path');

const dataDir = path.join(__dirname, '..', 'data');
const files = [
  path.join(dataDir, 'padron_staging.json'),
  path.join(dataDir, 'padron_staging_batch2.json'),
  path.join(dataDir, 'padron_staging_batch3.json'),
  path.join(dataDir, 'padron_staging_batch4.json'),
  path.join(dataDir, 'padron_staging_batch5.json')
];

let allStudents = [];

files.forEach(f => {
  if (fs.existsSync(f)) {
    const raw = JSON.parse(fs.readFileSync(f, 'utf8'));
    allStudents = allStudents.concat(raw);
  }
});

console.log(`Total crudo recopilado: ${allStudents.length} estudiantes.`);

// Remove accidental duplicates if any based on DNI
const seenDni = new Set();
const uniqueStudents = [];
for (const s of allStudents) {
  const dniClean = String(s.code || s.dni).trim();
  if (!seenDni.has(dniClean)) {
    seenDni.add(dniClean);
    uniqueStudents.push({
      dni: dniClean,
      name: s.name.trim(),
      grade: s.grade.trim(),
      section: s.section.trim(),
      status: s.status || 'DEFINITIVA'
    });
  }
}

console.log(`Total únicos: ${uniqueStudents.length} estudiantes.`);

// Grade order map for neat sorting
const gradeOrder = {
  '1ro de Secundaria': 1,
  '2do de Secundaria': 2,
  '3ro de Secundaria': 3,
  '4to de Secundaria': 4,
  '5to de Secundaria': 5
};

uniqueStudents.sort((a, b) => {
  const gA = gradeOrder[a.grade] || 99;
  const gB = gradeOrder[b.grade] || 99;
  if (gA !== gB) return gA - gB;
  if (a.section !== b.section) return a.section.localeCompare(b.section);
  return a.name.localeCompare(b.name);
});

// Assign 3-digit voting codes starting from 101
let startCode = 101;
const padronFinal = uniqueStudents.map((s, idx) => {
  const voteCode = String(startCode + idx);
  return {
    id: String(idx + 1),
    code: voteCode, // Código de 3 dígitos para votar
    dni: s.dni,     // DNI para guiarse
    name: s.name,
    grade: s.grade,
    section: s.section,
    status: s.status,
    hasVoted: false,
    votedAt: null
  };
});

fs.writeFileSync(path.join(dataDir, 'padron.json'), JSON.stringify(padronFinal, null, 2), 'utf8');
fs.writeFileSync(path.join(dataDir, 'votes.json'), JSON.stringify([], null, 2), 'utf8');

console.log(`¡Padrón generado exitosamente!`);
console.log(`Primer estudiante: Código [${padronFinal[0].code}] - ${padronFinal[0].name} (${padronFinal[0].grade} ${padronFinal[0].section})`);
console.log(`Último estudiante: Código [${padronFinal[padronFinal.length - 1].code}] - ${padronFinal[padronFinal.length - 1].name} (${padronFinal[padronFinal.length - 1].grade} ${padronFinal[padronFinal.length - 1].section})`);

// Breakdown
const counts = {};
padronFinal.forEach(s => {
  const key = `${s.grade} - ${s.section}`;
  counts[key] = (counts[key] || 0) + 1;
});
console.log('\nResumen por Salón:');
Object.keys(counts).sort().forEach(k => {
  console.log(`  * ${k}: ${counts[k]} alumnos`);
});
