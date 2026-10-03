#!/usr/bin/env node
// UserPromptSubmit hook: append prompt + timestamp to the project's prompt_log.md.
// Model name is filled in later by log-stop-model.js (Stop hook).
//
// ⚠ 로그의 «자리»는 log-path.js 가 정한다 — 세 로깅 훅이 같은 답을 써야 한다.
//   cwd 에 그냥 적으면 배포 폴더 안에서 세션을 연 날 사용자 프롬프트 «원문»이 그 폴더에
//   쌓여 zip 에 딸려 나간다(2026-08-31 한 수업 폴더에서 실제로 발생). 근거는 log-path.js 머리말.
'use strict';
const fs = require('fs');
const path = require('path');

let resolveLogPath;
try {
  ({ resolveLogPath } = require('./log-path'));
} catch (e) {                              // fail-open — 훅이 프롬프트를 막으면 안 된다
  resolveLogPath = (d) => path.join(d, 'prompt_log.md');
}

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

const logPath = resolveLogPath(cwd);
let header = '';
if (!fs.existsSync(logPath)) {
  header = '# prompt_log.md\n\n전역 hook(UserPromptSubmit -> Stop -> PostToolUse:ExitPlanMode)이 자동으로 기록한다.\n\n---\n\n';
}

const block = `### ${timestamp()} (session ${sessionId})\n**Prompt:** ${prompt}\n**Model:** (pending — filled by Stop hook)\n\n`;

fs.appendFileSync(logPath, header + block, 'utf8');
process.exit(0);
