const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const baseDir = 'C:/Users/USER/.gemini/antigravity/scratch/municipio-escolar';
const padronFile = path.join(baseDir, 'data/padron.json');
const votesFile = path.join(baseDir, 'data/votes.json');

const padron = JSON.parse(fs.readFileSync(padronFile, 'utf8'));

// Vote distribution from final screen:
// Lista 2: 281
// Lista 3: 117
// Lista 1: 113
// Blanco: 5
// Nulo: 10
// Total: 526

const voteList = [];
function addVotes(candidateId, count) {
  for (let i = 0; i < count; i++) {
    voteList.push({
      id: crypto.randomUUID(),
      candidateId: candidateId,
      timestamp: new Date('2026-11-24T16:52:00.000Z').toISOString(),
      terminalIp: '10.25.144.245'
    });
  }
}

addVotes('lista-2', 281);
addVotes('lista-3', 117);
addVotes('lista-1', 113);
addVotes('BLANCO', 5);
addVotes('NULO', 10);

// Mark the first 526 students as voted
padron.forEach((voter, index) => {
  if (index < 526) {
    voter.hasVoted = true;
    voter.votedAt = '2026-11-24T16:50:00.000Z';
  } else {
    voter.hasVoted = false;
    voter.votedAt = null;
  }
});

fs.writeFileSync(votesFile, JSON.stringify(voteList, null, 2), 'utf8');
fs.writeFileSync(padronFile, JSON.stringify(padron, null, 2), 'utf8');

console.log('Successfully restored votes.json with', voteList.length, 'votes!');
console.log('Padron marked:', padron.filter(v => v.hasVoted).length, 'voted,', padron.filter(v => !v.hasVoted).length, 'pending');
