#!/usr/bin/env node
// UserPromptSubmit hook: append prompt + timestamp to ./prompt_log.md (relative to hook cwd).
// Model name is filled in later by log-stop-model.js (Stop hook).
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

function pad(n) { return String(n).padStart(2, '0'); }
function timestamp() {
  const d = new Date();
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

const input = readStdin();
const cwd = input.cwd || input.workingDirectory || input.working_directory || process.cwd();
const sessionId = input.session_id || input.sessionId || 'unknown';
const prompt = input.prompt || input.user_prompt || input.userPrompt || input.message || '';

if (!prompt) {
  // Nothing to log (e.g. slash-command-only submit with no text) — exit quietly.
  process.exit(0);
}

const logPath = path.join(cwd, 'prompt_log.md');
let header = '';
if (!fs.existsSync(logPath)) {
  header = '# prompt_log.md\n\n전역 hook(UserPromptSubmit -> Stop -> PostToolUse:ExitPlanMode)이 자동으로 기록한다.\n\n---\n\n';
}

const block = `### ${timestamp()} (session ${sessionId})\n**Prompt:** ${prompt}\n**Model:** (pending — filled by Stop hook)\n\n`;

fs.appendFileSync(logPath, header + block, 'utf8');
process.exit(0);
