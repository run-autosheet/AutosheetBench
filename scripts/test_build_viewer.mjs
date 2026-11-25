import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import vm from 'node:vm';
import {hash,publicTiming} from '../lib/publicResults.mjs';
import {validateDataset,importReport,filterRows,summarize,rowStatus,inspectorMatrix,recordingTiming,publicDuration,recordingSummary,expectedImages,answerAvailability,configurationKey,viewDataset} from '../viewer/model.mjs';
import {websiteProfiles} from '../viewer/website-model.mjs';
import {buildPublicInspector} from './build_viewer.mjs';
import {savedEvaluationDetails,buildEvaluationDetails} from './public_evaluation_details.mjs';
import {readReviewedRecordings} from './public_recordings_bundle.mjs';

const dataset=()=>({schemaVersion:'autosheetbench-public-results/v1',createdAt:'2026-11-12',publicationStatus:'local-review-draft',
  tasks:[{id:'a',title:'Task A',category:'charts'},{id:'b',title:'Task B',category:'formulas'}],
  configurations:[{id:'excel',label:'Excel',platform:'microsoft_excel'},{id:'sheets',label:'Sheets',platform:'google_sheets'}],
  results:[{taskId:'a',configurationId:'excel',evaluation:{score:100,passed:true,incomplete:false},timing:{seconds:10,policy:'one'},datasetInputMatches:true,images:[]},
    {taskId:'b',configurationId:'excel',evaluation:{score:0,passed:false,incomplete:false},timing:{seconds:300,policy:'two'},datasetInputMatches:false,images:[]},
    {taskId:'a',configurationId:'sheets',evaluation:{score:null,passed:null,incomplete:true},timing:{seconds:30,policy:'one'},datasetInputMatches:null,images:[]}]});

test('viewer validates unique cells, scores and strictly local content-addressed images',()=>{
  assert.equal(validateDataset(dataset()).results.length,3);
  const dupe=dataset();dupe.results.push(dupe.results[0]);assert.throws(()=>validateDataset(dupe),/duplicate/);
  const unsafe=dataset();unsafe.results[0].images=[{file:'https://private.test/pixel.png'}];assert.throws(()=>validateDataset(unsafe),/Unsafe/);
  const bad=dataset();bad.results[0].evaluation.score=200;assert.throws(()=>validateDataset(bad),/score/);
});

test('visible filters, score denominators, mismatch exclusions and single-policy timing stay explicit',()=>{
  const data=dataset();assert.equal(filterRows(data,{query:'task a',platform:'microsoft_excel'}).length,1);
  assert.equal(filterRows(data,{status:'fail'}).length,1);assert.equal(rowStatus(data.results[2]),'incomplete');
  let stats=summarize(data,data.results,{timingPolicy:'one'});assert.equal(stats[0].score,100);assert.equal(stats[0].scored,1);assert.equal(stats[0].seconds,10);assert.equal(stats[0].timed,1);assert.equal(stats[0].inputMismatches,1);
  stats=summarize(data,data.results,{timingPolicy:'one',includeMismatched:true});assert.equal(stats[0].score,50);assert.equal(stats[0].seconds,10);
  assert.equal(stats.find(c=>c.id==='sheets').score,null);assert.equal(stats.find(c=>c.id==='sheets').unknownInputs,1);
});

test('own evaluator report imports locally without loading private evidence paths',()=>{
  const imported=importReport({schemaVersion:'atomic-evaluation-report/v1',evaluations:[{taskId:'a',effectiveScore:80,passed:true,artifacts:{outputWorkbook:'/private/workbook.xlsx'},images:['https://tracking.test/pixel'],transcript:'SECRET'}]});
  assert.equal(imported.publicationStatus,'local-import');assert.equal(imported.results[0].evaluation.score,80);
  assert(!JSON.stringify(imported).includes('SECRET'));assert(!JSON.stringify(imported).includes('/private'));assert.equal(imported.results[0].images.length,0);
});

test('shared inspector adapter preserves saved verdicts and separates timing populations',()=>{
  const data=dataset();data.results[0].evaluation={score:17.125,passed:true,incomplete:false};
  data.results[1].evaluation={score:94,passed:false,incomplete:false};
  const matrix=inspectorMatrix(data,{timingPolicy:'one',forCharts:false});
  assert.equal(matrix.rows[0].byConfig.excel.effectiveScore,17.125);
  assert.equal(matrix.rows[0].byConfig.excel.passed,true,'never infer a pass threshold');
  assert.equal(matrix.rows[1].byConfig.excel.passed,false);
  assert.equal(matrix.rows[1].byConfig.excel.publicDurationMs,300000,'matrix retains original duration');
  assert.equal(matrix.rows[1].byConfig.excel.totalElapsedMs,null,'other timing policy is excluded from charts');
  assert.equal(matrix.rows[0].byConfig.sheets.judgeScore,null);
  assert.equal(matrix.rows[0].byConfig.sheets.incomplete,true);
  assert(matrix.configurations.every(c=>c.readOnly));
  const charts=inspectorMatrix(data,{timingPolicy:'one'});
  assert.equal(charts.rows[1].byConfig.excel.judgeScore,null,'known input mismatch excluded only from summaries');
  assert.equal(charts.rows[0].byConfig.excel.judgePassed,true);
  assert.equal(charts.rows[0].byConfig.excel.totalElapsedMs,10000);
  const included=inspectorMatrix(data,{timingPolicy:'two',includeMismatched:true});
  assert.equal(included.rows[1].byConfig.excel.judgeScore,94);
  assert.equal(included.rows[1].byConfig.excel.judgePassed,false);
  assert.equal(included.rows[1].byConfig.excel.totalElapsedMs,300000);
});

test('recording-only presentation rejects clocks and unverified new endpoints without altering originals',()=>{
  const data=dataset();data.results[0].timing.policy='recording-send-to-terminal-sum-of-turns';
  assert.equal(recordingTiming(data.results[0]),10);
  assert.equal(recordingTiming(data.results[1]),null);
  for(const policy of ['legacy-unspecified','native-conversation-timestamps','sum-of-turn-submit-to-completion'])assert.equal(recordingTiming({timing:{policy,seconds:12}}),null);
  assert.equal(recordingTiming({timing:{policy:'recording-turn-sum-v1',seconds:12}}),null);
  assert.equal(recordingTiming({timing:{policy:'recording-turn-sum-v1',seconds:12,recordingVerified:true}}),12);
  const original=JSON.stringify(data),summary=recordingSummary(data,data.results);
  assert.equal(summary[0].seconds,10);assert.equal(summary[0].timed,1);
  const matrix=inspectorMatrix(data,{recordingOnly:true,forCharts:false});
  assert.equal(matrix.rows[1].byConfig.excel.publicDurationMs,null);
  assert.equal(data.results[1].timing.seconds,300);assert.equal(JSON.stringify(data),original);
});

test('image expectations come from capture targets; absent final answers never imply zero executed turns',()=>{
  assert.equal(expectedImages({compare:{}}),false);
  assert.equal(expectedImages({compare:{sheetScreenshots:[],captureAllCharts:false}}),false);
  assert.equal(expectedImages({compare:{screenshotRange:'Data!A1:D9'}}),true);
  assert.equal(expectedImages({compare:{captureChartsFromNewSheets:true}}),true);
  assert.equal(answerAvailability({answer:{status:'missing',turns:[]}}),'Assistant text unavailable');
  assert.equal(answerAvailability({answer:{status:'withheld',turns:[]}}),'Assistant text withheld');
  assert.equal(answerAvailability({answer:{turns:[{turn:3,text:'Done'}]}}),'1 assistant excerpt');
});

test('saved checkpoints keep explicit verdicts and exclude raw payloads, workbook values and secrets',()=>{
  const raw={judgeConfig:{controlPoints:[{id:'cp_1',label:'Readable dates',weight:2}]},judge:{raw:'PRIVATE RAW',controlPointScores:[{id:'cp_1',passed:false,score:100,reason:'Contact person@example.com; /Users/private/run.xlsx'}]},
    deterministic:{report:{candidate_workbook:'/Users/private/run.xlsx',dimensions:[{dimension:'values',applicable:true,checks_run:2,checks_passed:1,checks:[{kind:'value',cell:'A1',sheet:'Data',passed:false,actual:'PRIVATE PATIENT',expected:'PRIVATE EXPECTED'},{kind:'value',passed:true}]}]}}};
  const details=savedEvaluationDetails(raw,'bound-hash');
  assert.equal(details.checkpoints[0].passed,false);assert.equal(details.checkpoints[0].score,100,'no inferred threshold or normalization');
  assert.equal(details.checkpoints[0].label,'Readable dates');assert.equal(details.checkpoints[0].weight,2);
  assert(details.checkpoints[0].reason.text.includes('[email removed]'));
  assert.equal(details.deterministic[0].checks[0].label,'value · Data!A1');
  assert(!JSON.stringify(details).includes('PRIVATE'));assert(!JSON.stringify(details).includes('/Users/'));
  raw.judge.controlPointScores[0].reason='api_key=abcdef1234567890';
  assert(savedEvaluationDetails(raw,'hash').checkpoints[0].reason.text.includes('withheld'));
  delete raw.judge.controlPointScores;
  assert.equal(savedEvaluationDetails(raw,'hash').checkpoints[0].passed,null,'a definition alone is not a passed checkpoint');
});

test('checkpoint sidecar binds exact frozen evaluation hashes and never follows a changed latest evaluation',t=>{
  const root=fs.mkdtempSync(path.join(fs.realpathSync(os.tmpdir()),'atomic-checkpoints-'));t.after(()=>fs.rmSync(root,{recursive:true,force:true}));
  const runs=path.join(root,'runs'),file=path.join(runs,'run-a/tasks/a/evaluations/saved.json');fs.mkdirSync(path.dirname(file),{recursive:true});
  const bytes=JSON.stringify({judge:{controlPointScores:[{id:'cp_1',label:'Saved criterion',passed:true,score:73,reason:'Saved reason'}]}});fs.writeFileSync(file,bytes);
  const row={id:'r',taskId:'a',configurationId:'config',sourceId:'run-a',evaluation:{sourceHash:hash(bytes)}};
  const inventory={schemaVersion:'autosheetbench-public-review/v1',createdAt:'frozen',runsRoot:runs,rows:[row],sources:[{file,sha256:hash(bytes)}]};
  const data={snapshotAt:'frozen',results:[row]};
  const sidecar=buildEvaluationDetails({inventory,data,configurationIds:['config']});
  assert.equal(sidecar.counts.withCheckpoints,1);assert.equal(sidecar.results.r.checkpoints[0].score,73);assert.equal(fs.readFileSync(file,'utf8'),bytes);
  fs.writeFileSync(file,JSON.stringify({judge:{controlPointScores:[{label:'NEW changed evidence'}]}}));
  const changed=buildEvaluationDetails({inventory,data,configurationIds:['config']});assert.equal(changed.results.r.status,'unavailable');assert(!JSON.stringify(changed).includes('NEW'));
  assert.throws(()=>buildEvaluationDetails({inventory,data:{...data,snapshotAt:'other'},configurationIds:['config']}),/does not match/);
});

test('website configuration identity and presentation do not select a different saved attempt',()=>{
  const data=dataset();data.configurations[0]={...data.configurations[0],product:'Claude for Excel add-in',label:'Claude in Excel',model:'Opus 5.5',reasoning:'Extended thinking on'};
  assert.equal(configurationKey(data.configurations[0]),configurationKey({product:'Claude in Excel',platform:'microsoft_excel',model:'Opus 5.5',reasoning:'Extended thinking'}));
  const curated={configurationIds:['excel'],websiteProfiles:{excel:{name:'Claude in Excel',model:'Opus 5.5',reasoning:'Extended thinking'}},taskPresentation:{a:{title:'Website title',imageExpected:false}}};
  const view=viewDataset(data,curated);assert.equal(view.configurations.length,1);assert.equal(view.results.length,2);
  assert.deepEqual(view.results,data.results.filter(r=>r.configurationId==='excel'));assert.equal(view.tasks[0].imageExpected,false);
  const noName=viewDataset(data,{websiteProfiles:{excel:{product:'Claude in Excel',model:'Opus 5.5',reasoning:'High'}}});
  assert.equal(noName.configurations[0].product,'Claude in Excel');assert(noName.configurations[0].label.startsWith('Claude in Excel'));
  data.results[0].evaluation.score=42;
  const profiles=websiteProfiles(data);assert.equal(profiles[0].score,100,'website charts show saved pass/fail; partial scores remain in the raw row');
  assert.equal(data.results[0].evaluation.score,42);
  data.configurations[1].websiteProfile={product:'Claude Code · Autosheet MCP',productKey:'claude',connector:'autosheet',platform:'google_sheets',model:'Opus 5.5 + Opus 5.5',reasoning:'Medium'};
  const mcp=websiteProfiles(data).find(p=>p.id==='sheets');
  assert.equal(mcp.mcp,'Autosheet');assert.equal(mcp.mcpModel,'Opus 5.5');
  assert.equal(mcp.name,'Claude Code + Autosheet MCP');
  assert.equal(data.tasks[0].title,'Task A');
});

test('shared-component standalone build verifies public allowlist, no evaluation routes and escaped data',async t=>{
  const root=fs.mkdtempSync(path.join(fs.realpathSync(os.tmpdir()),'atomic-viewer-'));t.after(()=>fs.rmSync(root,{recursive:true,force:true}));
  const bundle=path.join(root,'bundle');fs.mkdirSync(bundle);const data=dataset();data.tasks[0].title='</script><script>alert(1)</script> & \u2028';
  const files={'results.json':JSON.stringify(data),'results.jsonl':data.results.map(r=>JSON.stringify(r)).join('\n'),'results.csv':'task_id\n'};
  for(const [file,bytes] of Object.entries(files))fs.writeFileSync(path.join(bundle,file),bytes);
  fs.writeFileSync(path.join(bundle,'private-review.json'),'DO NOT COPY');
  fs.writeFileSync(path.join(bundle,'manifest.json'),JSON.stringify({schemaVersion:'autosheetbench-public-bundle/v1',files:Object.fromEntries(Object.entries(files).map(([k,v])=>[k,{sha256:hash(v)}]))}));
  const output=path.join(root,'viewer');const built=await buildPublicInspector({bundle,output,curation:{title:'Curated snapshot',configurationIds:['excel']}});assert.equal(built.results,3);
  assert.equal(built.curation.title,'Curated snapshot');assert.deepEqual(built.curation.configurationIds,['excel']);
  assert(!fs.existsSync(path.join(output,'private-review.json')));
  const js=fs.readFileSync(path.join(output,'data.js'),'utf8');assert(!js.includes('</script>'));assert(js.includes('\\u003c'));
  const sandbox={window:{}};vm.runInNewContext(js,sandbox);assert.equal(sandbox.window.ATOMIC_PUBLIC_RESULTS.tasks[0].title,data.tasks[0].title);
  vm.runInNewContext(fs.readFileSync(path.join(output,'model.js'),'utf8'),sandbox);assert.equal(sandbox.window.PublicResultsModel.validateDataset(sandbox.window.ATOMIC_PUBLIC_RESULTS).results.length,3);
  const viewer=fs.readFileSync(path.join(output,'viewer.js'),'utf8'),html=fs.readFileSync(path.join(output,'index.html'),'utf8');
  assert(!/\/api\/benchmark-inspector\/(evaluate|archive|manual-score|clear-evaluation)/.test(viewer));
  assert(html.includes("connect-src 'none'"));
  assert(viewer.includes('Task pass rate'));assert(viewer.includes('Chart sample sizes by configuration'));
  assert(!viewer.includes('Chart timing policy'));
  for(const notice of ['Public inputs:', 'The public Auto Format New Table input uses synthetic operational IDs', 'Public workbook metadata has been cleaned.', 'The three job tasks mask recruiter emails'])assert(!viewer.includes(notice),'revision history must not appear as permanent banners');
  assert(viewer.includes('Input matches the task input'));
  if(fs.existsSync(path.resolve('data/atomic-input-revision.json'))){
    assert(html.includes('input-revision.js?v='));
    const revision=JSON.parse(fs.readFileSync(path.join(output,'input-revision.json'),'utf8'));
    assert.equal(revision.id,'atomic-contact-privacy-v2');
    assert.equal(revision.files.length,13);
    assert(!revision.files.some(f=>f.asset));
  }
  assert(!viewer.includes('"/icons/'),'every website logo/platform icon is bundled, including Excel');
  assert(html.includes(`viewer.js?v=${hash(viewer).slice(0,16)}`));
  const manifest=JSON.parse(fs.readFileSync(path.join(output,'manifest.json')));for(const [file,info] of Object.entries(manifest.files))assert.equal(hash(fs.readFileSync(path.join(output,file))),info.sha256);
  assert.equal(manifest.renderer.renderer,'website-charts-and-shared-inspector/v2');
  assert(manifest.sourceComponents['viewer/components/InspectorResultCell.js']);
  await assert.rejects(buildPublicInspector({bundle,output}),/exists/);
  fs.appendFileSync(path.join(bundle,'results.csv'),'tampered');await assert.rejects(buildPublicInspector({bundle,output:path.join(root,'bad')}),/checksum/);
});


test('recording loader accepts only reviewed, frozen, content-addressed derivatives',t=>{
  const root=fs.mkdtempSync(path.join(fs.realpathSync(os.tmpdir()),'atomic-recording-bundle-'));t.after(()=>fs.rmSync(root,{recursive:true,force:true}));
  const bytes=Buffer.from('redacted video fixture'),sha=hash(bytes),file=`videos/${sha}.mp4`,sourceResultHash='frozen-metadata';
  const data={snapshotAt:'frozen',results:[{id:'r',taskId:'a',configurationId:'c',sourceResultHash,evaluation:{score:87,passed:true},timing:{seconds:12,policy:'recording-turn-sum-v1',recordingVerified:true}}]};
  const row={id:'r',taskId:'a',configurationId:'c',sourceResultHash,file,sha256:sha,redacted:true,review:{continuousSourceAndOutputFrames:true,layoutReviewed:true,captureScopeVerified:true,captureMode:'window',outputOcrErrors:0,residualPatternMatches:0},timing:{allPresentationTimestampsIdentical:true,frameCountIdentical:true,effectiveEndTimeIdentical:true}};
  const report={schemaVersion:'autosheetbench-redacted-recordings/v2',snapshotAt:'frozen',originalsIncluded:false,included:1,withheld:0,recordings:[row],withheldRecordings:[]};
  fs.mkdirSync(path.join(root,'videos'));fs.writeFileSync(path.join(root,file),bytes);
  const write=()=>{fs.writeFileSync(path.join(root,'recordings.json'),JSON.stringify(report));fs.writeFileSync(path.join(root,'manifest.json'),JSON.stringify({schemaVersion:'autosheetbench-redacted-recording-bundle/v2',sourceResultsSha256:'snapshot-hash',files:{[file]:{sha256:sha},'recordings.json':{sha256:hash(JSON.stringify(report))}}}));};write();
  const parsed=readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'});assert.equal(parsed.results.r.file,`recordings/${sha}.mp4`);assert.equal(parsed.counts.included,1);
  assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'new-results'}),/frozen results/);
  const frozenData=JSON.stringify(data);row.durationSeconds=2.25;row.timing.effectiveEndTimeIdentical=false;
  const padding={schemaVersion:'autosheetbench-viewing-tail-padding/v1',sourceEffectiveEndSeconds:1.75,sourceFinalFrameSeconds:2,viewingEffectiveEndSeconds:2.25,extraPlaybackSeconds:.5,sourceEndPrecedesFinalFrame:true,allFrameTimestampsPreserved:true,benchmarkTimingChanged:false};
  row.review.viewingTailPadding=padding;write();assert.equal(readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}).results.r.review.viewingTailPadding.extraPlaybackSeconds,.5);assert.equal(JSON.stringify(data),frozenData);
  const roundedPadding={...padding,viewingEffectiveEndSeconds:2.2500005,extraPlaybackSeconds:.5000005};
  row.review.viewingTailPadding=roundedPadding;row.durationSeconds=2.250001;write();
  assert.equal(readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}).counts.included,1);
  row.durationSeconds=2.250002;write();assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/required derivative review/);
  row.durationSeconds=2.25;row.review.viewingTailPadding=padding;
  for(const bad of [{...padding,extraPlaybackSeconds:3},{...padding,benchmarkTimingChanged:true},{...padding,allFrameTimestampsPreserved:false},{...padding,sourceFinalFrameSeconds:1},{...padding,viewingEffectiveEndSeconds:2.2}]){row.review.viewingTailPadding=bad;write();assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/required derivative review/);}
  const finalFrame={schemaVersion:'autosheetbench-viewing-tail-padding/v1',sourceEffectiveEndSeconds:1.75,sourceFinalFrameSeconds:1.75,viewingEffectiveEndSeconds:2.25,extraPlaybackSeconds:.5,sourceEndsAtFinalFrame:true,allFrameTimestampsPreserved:true,benchmarkTimingChanged:false};
  row.review.viewingTailPadding=finalFrame;write();assert.equal(readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}).counts.included,1);
  for(const bad of [{...finalFrame,sourceFinalFrameSeconds:2},{...finalFrame,extraPlaybackSeconds:3,viewingEffectiveEndSeconds:4.75},{...finalFrame,sourceEndsAtFinalFrame:false}]){row.review.viewingTailPadding=bad;write();assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/required derivative review/);}
  row.review.viewingTailPadding=padding;row.timing.frameCountIdentical=false;write();assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/required derivative review/);
  delete row.review.viewingTailPadding;delete row.durationSeconds;Object.assign(row.timing,{frameCountIdentical:true,effectiveEndTimeIdentical:true});write();

  row.review.captureScopeVerified=false;write();assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/required derivative review/);
  Object.assign(row.review,{legacyViewingCopy:true,spreadsheetContinuityReviewed:true,captureBinding:'content-reviewed-viewing-copy',continuityReviewSha256:'a'.repeat(64),contentApprovalSha256:'b'.repeat(64)});write();
  assert.equal(readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}).results.r.review.captureScopeVerified,false);
  delete row.review.continuityReviewSha256;write();assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/required derivative review/);
  for(const key of ['legacyViewingCopy','spreadsheetContinuityReviewed','captureBinding','contentApprovalSha256'])delete row.review[key];
  row.review.captureScopeVerified=true;row.review.captureMode='window-crop';write();assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/required derivative review/);
  row.review.captureMode='window';write();
  row.review.sourceFixtureIdentityVerified=false;write();assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/explicit content-review clearance/);
  row.review.recordingContentOnlyClearance=true;write();assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/explicit content-review clearance/);
  Object.assign(row.review,{spreadsheetContinuityReviewed:true,continuityReviewSha256:'a'.repeat(64),contentApprovalSha256:'b'.repeat(64)});write();
  assert.equal(readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}).results.r.review.sourceFixtureIdentityVerified,false);
  row.review.sourceFixtureIdentityVerified=true;write();assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/explicit content-review clearance/);
  for(const key of ['sourceFixtureIdentityVerified','recordingContentOnlyClearance','spreadsheetContinuityReviewed','continuityReviewSha256','contentApprovalSha256'])delete row.review[key];write();
  row.review.layoutReviewed=false;write();assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/required derivative review/);
  row.review.layoutReviewed=true;row.sourceResultHash='new-metadata';write();assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/required derivative review/);
  row.sourceResultHash=sourceResultHash;write();fs.writeFileSync(path.join(root,file),'changed');assert.throws(()=>readReviewedRecordings({bundle:root,data,resultsHash:'snapshot-hash'}),/checksum mismatch/);
});

test('displayed duration prefers agent time, else a recording measurement',()=>{
  assert.equal(publicDuration({timing:{agentSeconds:7.1,policy:'recording-final-message-completion',seconds:37}}),7.1);
  assert.equal(publicDuration({timing:{agentSeconds:null,policy:'recording-final-message-completion',seconds:37}}),37);
  assert.equal(publicDuration({timing:{agentSeconds:null,policy:'observed-native-completion-upper-bound',seconds:218}}),null);
  assert.equal(publicTiming({durationSeconds:300.019,agentDurationMs:322796,timedOut:true}).agentSeconds,300.019);
  assert.equal(publicTiming({durationSeconds:37,agentDurationMs:7100}).agentSeconds,7.1);
});
