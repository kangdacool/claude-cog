#!/usr/bin/env node
// PostToolUse (matcher: ExitPlanMode) hook: once a plan is approved, append the full plan
// file content to the project's prompt_log.md so plans are traceable alongside prompts.
//
// ⚠ 자리 계산은 log-path.js 로 통일 — 세 로깅 훅이 같은 파일을 봐야 한다.
'use strict';
const fs = require('fs');
const path = require('path');
const os = require('os');

let resolveLogPath;
try {
  ({ resolveLogPath } = require('./log-path'));
} catch (e) {                              // fail-open
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

function findPlanPathInText(text) {
  if (!text) return null;
  const m = String(text).match(/[A-Za-z]:\\[^\r\n"]*?\.md/);
  return m ? m[0] : null;
}

const input = readStdin();
const cwd = input.cwd || input.workingDirectory || input.working_directory || process.cwd();

// 1) Try to find an explicit plan path in the tool response/input text.
let planPath = findPlanPathInText(JSON.stringify(input.tool_response || {}))
  || findPlanPathInText(JSON.stringify(input.tool_input || {}));

// 2) Fallback: most recently modified .md file in the plans directory.
if (!planPath || !fs.existsSync(planPath)) {
  const plansDir = path.join(os.homedir(), '.claude', 'plans');
  try {
    const files = fs.readdirSync(plansDir)
      .filter((f) => f.endsWith('.md'))
      .map((f) => {
        const p = path.join(plansDir, f);
        return { p, mtime: fs.statSync(p).mtimeMs };
      })
      .sort((a, b) => b.mtime - a.mtime);
    if (files.length > 0) planPath = files[0].p;
  } catch (e) {
    // no plans directory — nothing to log
  }
}

if (!planPath || !fs.existsSync(planPath)) process.exit(0);

let planContent;
try {
  planContent = fs.readFileSync(planPath, 'utf8');
} catch (e) {
  process.exit(0);
}

const logPath = resolveLogPath(cwd);
const block = `### PLAN — ${timestamp()}\n(source: ${planPath})\n\n${planContent}\n\n`;
fs.appendFileSync(logPath, block, 'utf8');
process.exit(0);
