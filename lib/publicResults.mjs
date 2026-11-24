// Public export policy only. Never evaluate, normalize/rethreshold scores or
// access a live product. Historical source files are read, never written.
import crypto from 'node:crypto';
import {promptToEndDuration} from './duration.mjs';

export const PUBLIC_RESULTS_SCHEMA='autosheetbench-public-results/v1';
export const REVIEW_SCHEMA='autosheetbench-public-review/v1';
export const hash=value=>crypto.createHash('sha256').update(value).digest('hex');
export const finite=value=>typeof value==='number'&&Number.isFinite(value)?value:null;
export const iso=value=>typeof value==='string'&&Number.isFinite(Date.parse(value))?new Date(value).toISOString():null;

const patterns=[
  ['embedded-data',/\bdata:[^\s"')\]]+/gi],
  ['secret',/\b(?:sk-(?:or-|proj-)?[a-zA-Z0-9_-]{12,}|AIza[a-zA-Z0-9_-]{20,}|gh[pousr]_[a-zA-Z0-9_]{15,}|(?:Bearer|api[_ -]?key|access[_ -]?token|password|secret)\s*[:=]\s*["']?[^\s"',;]{6,})/gi],
  ['email',/\b[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi],
  ['url',/\b(?:https?:\/\/|file:\/\/|plugin:\/\/|gs:\/\/)[^\s<>"')\]]+/gi],
  ['local-path',/(?:\/(?:Users|home|private|var|tmp|Volumes)\/[^\s<>"']+|[A-Z]:\\(?:Users|Documents)\\[^\s<>"']+)/gi],
  ['private-id',/\b(?:spreadsheetId|conversationId|agentId|accountId|requestId|sessionId)\s*["']?\s*[:=]\s*["']?[a-zA-Z0-9_-]{6,}/gi],
  ['ssn',/\b\d{3}-\d{2}-\d{4}\b/g],
  ['phone',/\b(?:phone|tel(?:ephone)?|mobile)\s*[:=]\s*\+?[\d() .-]{8,}/gi],
];

export function sanitizePublicText(value,{limit=12000}={}) {
  let text=String(value??'').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,'');
  const redactions={};
  for(const [name,pattern] of patterns) {
    text=text.replace(pattern,()=>{redactions[name]=(redactions[name]||0)+1;return `[${name} removed]`;});
  }
  const truncated=text.length>limit;
  return {text:truncated?text.slice(0,limit)+'\n[excerpt truncated]':text,redactions,truncated};
}

export function sensitiveImageText(text) {
  const result=sanitizePublicText(text,{limit:Number.MAX_SAFE_INTEGER});
  return Object.keys(result.redactions);
}

export function latestEvaluation(evaluations) {
  return [...evaluations].sort((a,b)=>Date.parse(iso(b.value.updatedAt)||'1970-01-01')-Date.parse(iso(a.value.updatedAt)||'1970-01-01')
    || a.file.localeCompare(b.file))[0]||null;
}

export function resultTimestamp(metadata) {
  for(const key of ['finishedAt','endedAt','completedAt','startedAt']) {
    const value=iso(metadata[key]);if(value)return value;
  }
  return null;
}

export function selectLatestResults(candidates) {
  const selected=new Map();
  const sorted=[...candidates].sort((a,b)=>{
    const time=Date.parse(resultTimestamp(b.metadata)||'1970-01-01')-Date.parse(resultTimestamp(a.metadata)||'1970-01-01');
    // Same execution copied into a later snapshot: prefer its latest saved eval,
    // then a stable source ID. Never prefer a pass or an earlier complete eval.
    return time||Date.parse(iso(b.evaluation?.value.updatedAt)||'1970-01-01')-Date.parse(iso(a.evaluation?.value.updatedAt)||'1970-01-01')
      ||a.sourceId.localeCompare(b.sourceId);
  });
  for(const candidate of sorted) {
    const key=`${candidate.configuration.id}/${candidate.taskId}`;
    if(!selected.has(key))selected.set(key,candidate);
  }
  return [...selected.values()].sort((a,b)=>a.configuration.id.localeCompare(b.configuration.id)||a.taskId.localeCompare(b.taskId));
}

export function savedEvaluation(candidate) {
  const raw=candidate.evaluation?.value||null,metadata=candidate.metadata;
  const manual=metadata.manualScore===0||metadata.manualScore===100?metadata.manualScore:null;
  const stored=finite(raw?.effectiveScore);
  const staleManual=manual===null&&(raw?.manualScore===0||raw?.manualScore===100);
  // No read-time rubric alignment, quality threshold updates or score rebuild.
  // If a cleared historical override obscures the automatic result, show the
  // stored score separately and leave the public effective result unresolved.
  const score=manual??(staleManual?null:stored);
  const passed=manual!==null?manual===100:staleManual?null:typeof raw?.passed==='boolean'?raw.passed:null;
  const incomplete=manual!==null?false:raw?raw.incomplete===true||score===null||passed===null:true;
  return {score,passed,incomplete,storedScore:stored,storedPassed:typeof raw?.passed==='boolean'?raw.passed:null,
    manualScore:manual,manualAdjudicatedAt:iso(metadata.manualOverrideUpdatedAt),staleManualOverride:staleManual,
    ...(manual!==null?{manualReason:sanitizePublicText(metadata.manualOverrideReason||'',{limit:1800})}:{}),
    evaluatedAt:iso(raw?.updatedAt),schemaVersion:typeof raw?.schemaVersion==='string'?sanitizePublicText(raw.schemaVersion,{limit:120}).text:null,
    judge:{model:sanitizePublicText(raw?.judgeProfile?.model||'Not recorded').text,
      reasoning:sanitizePublicText(raw?.judgeProfile?.reasoningEffort||'Not recorded').text,
      profileId:typeof raw?.judgeProfile?.judgeProfileId==='string'?hash(raw.judgeProfile.judgeProfileId).slice(0,16):null},
    summary:sanitizePublicText(raw?.judge?.summary||raw?.summary||'',{limit:2400}),
    missingEvidence:Array.isArray(raw?.missingEvidence)?raw.missingEvidence.map(v=>sanitizePublicText(v,{limit:100}).text):[],
    sourceHash:candidate.evaluation?.sha256||null};
}

// `permissionWait`: the measured permission-card wait of this run (lib/permissionWaits.mjs).
export function publicTiming(metadata,{permissionWait=null}={}) {
  const policy=metadata.timingPolicy||metadata.durationMeasurement||metadata.timingSource||'legacy-unspecified';
  const source=sanitizePublicText(policy,{limit:160}).text;
  const verified=metadata.timingPolicy==='recording-turn-sum-v1'&&metadata.timingStatus==='verified';
  const seconds=finite(metadata.durationSeconds)??(finite(metadata.durationMs)!==null?metadata.durationMs/1000:null);
  // Autosheet's backend times each agent turn from prompt submission to its
  // final message; on recorded runs it matches the recording's submit-to-finish
  // span within 0.1 s. A timed-out run keeps the runner's stop, since the
  // backend can keep counting after it.
  let agentSeconds=finite(metadata.agentDurationMs)!==null?metadata.agentDurationMs/1000:null;
  if(agentSeconds!==null&&metadata.timedOut===true&&seconds!==null&&seconds<agentSeconds)agentSeconds=seconds;
  // The displayed duration, computed exactly as the website computes it.
  const promptToEnd=promptToEndDuration(metadata,{permissionWait});
  return {seconds,policy:source,recordingVerified:verified,agentSeconds,
    promptToEndSeconds:promptToEnd?promptToEnd.ms/1000:null,promptToEndSource:promptToEnd?.source||null,
    ...(promptToEnd?.boundSeconds!==undefined?{promptToEndBoundSeconds:promptToEnd.boundSeconds}:{}),
    ...(promptToEnd?.permissionWaitSeconds!==undefined?{permissionWaitSeconds:promptToEnd.permissionWaitSeconds}:{}),
    status:sanitizePublicText(metadata.timingStatus||'historical-unverified',{limit:120}).text,
    finishedInTime:typeof metadata.finishedInTime==='boolean'?metadata.finishedInTime:null};
}

// Autosheet in Excel ended some replies with its conversation ID on lines of its
// own; it is dropped, as build_hf_public_tree.py drops it from v1.0.0's answers.
const ANSWER_ID_FOOTER=/\nID:\n[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}\b/g;

export function extractPublicAnswers(entries) {
  const byTurn=new Map();
  let currentTurn=0;
  for(const entry of entries||[]) {
    if(entry?.role==='user'){currentTurn=Number.isInteger(entry.turn)?entry.turn:currentTurn+1;continue;}
    if(entry?.role!=='assistant'||entry.final===false||['reasoning','analysis','thinking','tool'].includes(entry.kind))continue;
    if(typeof entry.text!=='string'||!entry.text.trim())continue;
    const turn=Number.isInteger(entry.turn)&&entry.turn>0?entry.turn:Math.max(1,currentTurn);
    byTurn.set(turn,{turn,...sanitizePublicText(entry.text.replace(ANSWER_ID_FOOTER,''))});
  }
  return [...byTurn.values()].sort((a,b)=>a.turn-b.turn);
}

export function imageOrigin(metadata,file,kind) {
  if(/(?:recording|video|tab-frames|review-samples|layout|viewport|source-view|native-sheet-view|recovered-sheet)/i.test(file))return null;
  const renderer=kind==='chart'?metadata.capture?.chartRenderer:metadata.capture?.documentRenderer;
  const reviewed=kind==='chart'?metadata.focusedChartEvidence?.source:metadata.reviewedSheetEvidence?.source;
  if(/recording|video|pdf|libreoffice/i.test(reviewed||''))return null;
  const expected=kind==='chart'?'excel-copy-picture-chart':'excel-copy-picture-range';
  if(metadata.platform!=='google_sheets'&&renderer===expected)return expected;
  return null;
}

export function validatePublicDataset(data) {
  if(data?.schemaVersion!==PUBLIC_RESULTS_SCHEMA||!Array.isArray(data.results)||!Array.isArray(data.configurations)||!Array.isArray(data.tasks))throw Error('Unsupported public results file');
  const configs=new Set(data.configurations.map(x=>x.id)),tasks=new Set(data.tasks.map(x=>x.id)),seen=new Set();
  for(const row of data.results) {
    const key=`${row.configurationId}/${row.taskId}`;
    if(!configs.has(row.configurationId)||!tasks.has(row.taskId)||seen.has(key))throw Error('Invalid/duplicate public task × configuration');
    seen.add(key);
    if(row.evaluation?.score!==null&&(!Number.isFinite(row.evaluation?.score)||row.evaluation.score<0||row.evaluation.score>100))throw Error('Invalid stored public score');
    for(const image of row.images||[])if(!/^assets\/[a-f0-9]{64}\.png$/.test(image.file))throw Error('Unsafe public image reference');
  }
  return data;
}
