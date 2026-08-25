#!/usr/bin/env node
// Stop hook: read the transcript for this turn, find the model that was actually used,
// and patch the last "(pending — filled by Stop hook)" placeholder in ./prompt_log.md.
'use strict';
const fs = require('fs');
const path = require('path');

function readStdin() {
  try {
    const data = fs.readFileSync(0, 'utf8');
    return data ? JSON.parse(data) : {};
  } catch (e) {
    return {};
  }
}

// Recursively search an object for the first string-valued "model" key.
function findModel(obj, depth) {
  if (depth > 6 || obj === null || typeof obj !== 'object') return null;
  if (typeof obj.model === 'string' && obj.model.length > 0) return obj.model;
  for (const k of Object.keys(obj)) {
    const v = obj[k];
    if (v && typeof v === 'object') {
      const found = findModel(v, depth + 1);
      if (found) return found;
    }
  }
  return null;
}

const input = readStdin();
const cwd = input.cwd || input.workingDirectory || input.working_directory || process.cwd();
const transcriptPath = input.transcript_path || input.transcriptPath || input.transcript;

let model = null;

if (transcriptPath && fs.existsSync(transcriptPath)) {
  try {
    const lines = fs.readFileSync(transcriptPath, 'utf8').split('\n').filter(Boolean);
    for (let i = lines.length - 1; i >= 0 && !model; i--) {
      let entry;
      try { entry = JSON.parse(lines[i]); } catch (e) { continue; }
      const role = entry.role || entry.type || (entry.message && entry.message.role);
      if (role && String(role).toLowerCase().includes('user')) continue; // skip user turns
      model = findModel(entry, 0);
    }
  } catch (e) {
    // fall through with model = null
  }
}

const logPath = path.join(cwd, 'prompt_log.md');
if (!fs.existsSync(logPath)) process.exit(0);

const placeholder = '(pending — filled by Stop hook)';
let content = fs.readFileSync(logPath, 'utf8');
const lastIdx = content.lastIndexOf(placeholder);
if (lastIdx !== -1 && model) {
  content = content.slice(0, lastIdx) + model + content.slice(lastIdx + placeholder.length);
  fs.writeFileSync(logPath, content, 'utf8');
} else if (lastIdx !== -1) {
  // Model not resolvable from transcript — leave a distinguishable marker instead of silent pending forever.
  content = content.slice(0, lastIdx) + '(model unresolved)' + content.slice(lastIdx + placeholder.length);
  fs.writeFileSync(logPath, content, 'utf8');
}
process.exit(0);
