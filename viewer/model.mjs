// Dependency-free public-viewer model; bundled as a classic script for file://.
export function validateDataset(data) {
  if(data?.schemaVersion!=='autosheetbench-public-results/v1'||!Array.isArray(data.results)||!Array.isArray(data.configurations)||!Array.isArray(data.tasks))throw Error('Choose an AutosheetBench public snapshot or evaluator report');
  const tasks=new Set(data.tasks.map(c=>c.id)),configs=new Set(data.configurations.map(c=>c.id)),seen=new Set();
  if(data.results.length>100000)throw Error('Results file exceeds viewer limits');
  for(const row of data.results) {
    const key=`${row.configurationId}/${row.taskId}`;
    if(!tasks.has(row.taskId)||!configs.has(row.configurationId)||seen.has(key))throw Error('Unknown or duplicate task × configuration');
    seen.add(key);
    if(row.evaluation?.score!==null&&(!Number.isFinite(row.evaluation?.score)||row.evaluation.score<0||row.evaluation.score>100))throw Error('Invalid result score');
    for(const image of row.images||[]){
      const legacy=/^assets\/[a-f0-9]{64}\.png$/.test(image.file);
      const local=typeof image.file==='string'&&image.file.match(/^runs\/([a-zA-Z0-9][a-zA-Z0-9._-]*)\/tasks\/([a-zA-Z0-9][a-zA-Z0-9._-]*)\/artifacts\/screenshots\/[0-9]{2,}-(?:chart|range)\.png$/);
      if(!legacy&&!(local&&local[2]===row.taskId&&!local[1].includes('..')&&!local[2].includes('..')))throw Error('Unsafe screenshot path');
    }
  }
  return data;
}

export function importReport(report) {
  if(report?.schemaVersion!=='atomic-evaluation-report/v1'||!Array.isArray(report.evaluations))return validateDataset(report);
  const config={id:'config-local-import',label:'Imported local evaluation',platform:'local',model:'See original report',reasoning:'Not grouped',product:'Local report',connector:''};
  const tasks=report.evaluations.map(item=>({id:item.taskId,title:item.taskId,category:'Imported',difficulty:'Not recorded'}));
  const results=report.evaluations.map(item=>({id:`import-${item.taskId}`,taskId:item.taskId,configurationId:config.id,resultAt:null,
    executionStatus:'not recorded',evaluation:{score:Number.isFinite(item.effectiveScore)?item.effectiveScore:null,passed:typeof item.passed==='boolean'?item.passed:null,
      incomplete:item.incomplete===true,storedScore:item.effectiveScore??null,manualScore:null,evaluatedAt:item.updatedAt||null,judge:{model:item.judgeProfile?.model||'Not recorded'},summary:{text:''}},
    timing:{seconds:null,policy:'not supplied by evaluator report',recordingVerified:false,status:'not recorded'},datasetInputMatches:null,
    answer:{status:'missing',turns:[]},images:[],withheldImages:[]}));
  return validateDataset({schemaVersion:'autosheetbench-public-results/v1',publicationStatus:'local-import',createdAt:new Date().toISOString(),snapshotAt:null,
    selectionPolicy:'one evaluator report imported locally',tasks,configurations:[config],results,review:{limitations:['Imported locally; not submitted to a server. Answer/image evidence is not loaded from private report paths.']}});
}

export function rowStatus(row) {
  if(row.evaluation?.incomplete||typeof row.evaluation?.passed!=='boolean')return 'incomplete';
  return row.evaluation.passed?'pass':'fail';
}

export function filterRows(data,filters={}) {
  const tasks=new Map(data.tasks.map(c=>[c.id,c])),configs=new Map(data.configurations.map(c=>[c.id,c]));
  const query=String(filters.query||'').trim().toLowerCase();
  return data.results.filter(row=>{
    const c=tasks.get(row.taskId),config=configs.get(row.configurationId);
    return (!filters.platform||filters.platform==='all'||config.platform===filters.platform)
      &&(!filters.configuration||filters.configuration==='all'||config.id===filters.configuration)
      &&(!filters.category||filters.category==='all'||c.category===filters.category)
      &&(!filters.status||filters.status==='all'||rowStatus(row)===filters.status)
      &&(!query||`${c.id} ${c.title} ${config.label}`.toLowerCase().includes(query));
  });
}

export function summarize(data,rows,{timingPolicy,includeMismatched=false}={}) {
  const finite=value=>typeof value==='number'&&Number.isFinite(value);
  const mean=values=>values.length?values.reduce((a,b)=>a+b,0)/values.length:null;
  const groups=data.configurations.map(config=>{
    const all=rows.filter(row=>row.configurationId===config.id),compatible=includeMismatched?all:all.filter(row=>row.datasetInputMatches!==false);
    const scored=compatible.filter(row=>!row.evaluation?.incomplete&&finite(row.evaluation?.score)&&typeof row.evaluation?.passed==='boolean');
    const timed=compatible.filter(row=>row.timing?.policy===timingPolicy&&finite(row.timing.seconds)&&row.timing.seconds>=0);
    return {...config,rows:all.length,scored:scored.length,score:mean(scored.map(row=>row.evaluation.score)),
      passed:scored.filter(row=>row.evaluation.passed).length,failed:scored.filter(row=>!row.evaluation.passed).length,
      incomplete:all.filter(row=>rowStatus(row)==='incomplete').length,
      seconds:mean(timed.map(row=>row.timing.seconds)),timed:timed.length,recordingVerified:timed.filter(row=>row.timing.recordingVerified).length,
      inputMismatches:all.filter(row=>row.datasetInputMatches===false).length,unknownInputs:all.filter(row=>row.datasetInputMatches===null).length};
  }).filter(config=>config.rows>0);
  return groups.sort((a,b)=>(b.score??-1)-(a.score??-1)||a.label.localeCompare(b.label));
}

export function formatSeconds(value) {
  if(typeof value!=='number'||!Number.isFinite(value))return 'Not recorded';
  return value<60?`${value.toFixed(1)}s`:`${Math.floor(value/60)}m ${(value%60).toFixed(0).padStart(2,'0')}s`;
}

// Historical measurements are admitted only when their saved provenance names
// recording analysis. Never relabel runner clocks or transcript timestamps.
const RECORDING_POLICIES=new Set(['recording-turn-sum-v1','recording-final-message-completion',
  'recording-protocol-timeout-cutoff','recording-runner-submit-to-busy-label-end-sum-of-turns',
  'recording-send-to-terminal-sum-of-turns','recording-stop-control-minus-question-wait',
  'recording-submit-to-feedback-controls','visual-completion-backward-from-final-frame',
  // Add-in stop-button span, scripts/audit_addin_button_durations.py.
  'recording-submit-to-finish']);
export function recordingTiming(row) {
  const t=row.timing||{};
  const review=t.recordingReview;
  if(review?.method==='saved-prompt-to-completion-video-review'&&
    /^[a-f0-9]{64}$/.test(review.recordingSha256||'')&&review.sourceMetadataSha256===row.sourceResultHash&&
    /^[a-f0-9]{64}$/.test(review.sourceMetadataSha256||'')&&Number.isFinite(review.seconds)&&review.seconds>0)return review.seconds;
  if(t.policy==='recording-turn-sum-v1'&&t.recordingVerified!==true)return null;
  return RECORDING_POLICIES.has(t.policy)&&Number.isFinite(t.seconds)&&t.seconds>=0?t.seconds:null;
}

// Displayed duration: prompt submission to end, computed by
// lib/duration.mjs promptToEndDuration exactly as the website
// does. Snapshots exported before it fall back to agent time, then a recording
// measurement. Runner clocks and transcript timestamps never count.
export function publicDuration(row) {
  if(row.timing&&Object.hasOwn(row.timing,'promptToEndSeconds'))return Number.isFinite(row.timing.promptToEndSeconds)?row.timing.promptToEndSeconds:null;
  const agent=row.timing?.agentSeconds;
  return Number.isFinite(agent)&&agent>=0?agent:recordingTiming(row);
}

export function answerAvailability(row) {
  const n=row.answer?.turns?.length||0;
  return n?`${n} assistant excerpt${n===1?'':'s'}`:row.answer?.status==='withheld'?'Assistant text withheld':'Assistant text unavailable';
}

export function expectedImages(spec) {
  const compare=spec.compare||{};
  return ['screenshotRange','sheetScreenshots','chartScreenshots','captureAllSheets','captureNewSheets',
    'captureAllCharts','captureChartsFromNewSheets','captureChartsFromSheets'].some(key=>
      Array.isArray(compare[key])?compare[key].length>0:Boolean(compare[key]));
}

export function configurationKey(c) {
  const connector=c.connector||null,p=`${c.product} ${c.label}`.toLowerCase();
  // The label carries the model's name, so Autosheet with a Gemini model is
  // recognized as Autosheet before the name Gemini picks Gemini in Sheets.
  const autosheet=/autosheet in (excel|sheets)|gpt for (excel|sheets|work)/.test(p);
  const productKey=/claude/.test(p)?'claude':/chatgpt/.test(p)?'chatgpt':/copilot/.test(p)?'copilot':
    !autosheet&&/gemini/.test(p)?'gemini':c.platform==='google_sheets'?'gpt-for-sheets':'gpt-for-excel';
  const normalize=value=>String(value||'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const model=productKey==='gemini'&&/not exposed/i.test(c.model)?'Auto':c.model?.split(' + ')[0];
  const reasoning=/not exposed/i.test(c.reasoning)?'Auto':String(c.reasoning).replace(/ thinking on$/i,' thinking');
  return [productKey,c.platform,connector,normalize(model),normalize(reasoning)].join('/');
}

export function viewDataset(data,curation={}) {
  const ids=curation.configurationIds||data.configurations.map(c=>c.id);
  return {...data,configurations:ids.flatMap(id=>{
    const c=data.configurations.find(c=>c.id===id);if(!c)return [];
    const website=curation.websiteProfiles?.[id];
    const name=website?.name||website?.product||c.product||c.label;
    return [{...c,...(website?{label:[name,website.model,website.reasoning].filter((v,i,a)=>i<2||v!==a[1]).join(' · '),product:name,model:website.model,
      reasoning:website.reasoning,websiteProfile:website}:{} )}];
  }),tasks:data.tasks.map(c=>({...c,...curation.taskPresentation?.[c.id]})),results:data.results.filter(r=>ids.includes(r.configurationId))};
}

export function recordingSummary(data,rows,{includeMismatched=false}={}) {
  // Reuse the saved-score selection while supplying only displayed durations.
  const recordingRows=rows.map(r=>({...r,timing:{...r.timing,seconds:publicDuration(r),policy:'recording-only'}}));
  return summarize(data,recordingRows,{timingPolicy:'recording-only',includeMismatched});
}

// Adapt the frozen public schema to the EXISTING inspector results component.
// Judge fields here carry the already-saved effective result. No judge runs and
// no verdict/threshold inference is performed by this adapter.
export function inspectorMatrix(data,{timingPolicy,recordingOnly=false,includeMismatched=false,forCharts=true,recordings={}}={}) {
  const configurations=data.configurations.map(c=>({...c,runConfigId:c.id,external:true,readOnly:true,
    displayModel:c.model,reasoningEffort:c.reasoning,executionKind:c.platform,source:c.product}));
  const rows=data.tasks.map(c=>({taskId:c.id,logicalTaskId:c.id,title:c.title,family:c.category,category:c.category,imageExpected:c.imageExpected,
    platform:'cross_platform',applicability:'cross_platform',source:'real',external:true,byConfig:{}}));
  const byId=new Map(rows.map(row=>[row.taskId,row]));
  for(const r of data.results) {
    const admitted=!forCharts||includeMismatched||r.datasetInputMatches!==false;
    const complete=!r.evaluation.incomplete&&typeof r.evaluation.passed==='boolean'&&typeof r.evaluation.score==='number';
    const score=!forCharts?r.evaluation.score:admitted&&complete?r.evaluation.score:null;
    const seconds=recordingOnly?publicDuration(r):r.timing?.policy===timingPolicy?r.timing.seconds:null;
    const timed=admitted&&Number.isFinite(seconds)&&seconds>=0;
    byId.get(r.taskId).byConfig[r.configurationId]={external:true,readOnly:true,runAvailable:true,taskId:r.taskId,runConfigId:r.configurationId,
      status:'completed',evaluationStatus:complete?'evaluated':'recorded',incomplete:r.evaluation.incomplete===true,outputAvailable:true,
      effectiveScore:score,automaticScore:score,judgeScore:score,
      passed:!forCharts?r.evaluation.passed:admitted&&complete?r.evaluation.passed:null,judgePassed:admitted&&complete?r.evaluation.passed:null,
      totalElapsedMs:timed?seconds*1000:null,recordingElapsedMs:timed?seconds*1000:null,
      publicDurationMs:recordingOnly?(seconds===null?null:seconds*1000):typeof r.timing?.seconds==='number'?r.timing.seconds*1000:null,publicTimingPolicy:r.timing?.policy,
      publicRecordingOnly:recordingOnly,publicDurationSource:r.timing?.promptToEndSource||null,publicDurationBoundSeconds:r.timing?.promptToEndBoundSeconds??null,publicAnswerText:answerAvailability(r),publicImagesExpected:byId.get(r.taskId).imageExpected,
      publicAnswerTurns:r.answer?.turns?.length||0,publicImages:r.images?.length||0,
      recordingUrl:recordings[r.id]?.sourceResultHash===r.sourceResultHash&&recordings[r.id]?.status==='included'?recordings[r.id].file:null,
      datasetInputMatches:r.datasetInputMatches,resultId:r.id};
  }
  return {lane:'atomic',configurations,rows};
}
