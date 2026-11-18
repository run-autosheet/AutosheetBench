// Call from an active awaited Computer Use session after execution/export.
// The controller measures the visible native range/chart; this helper takes a
// NEW tab screenshot and crops that target. It never reads recording frames.
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
import {fileURLToPath} from 'node:url';
import {imageHash, validateTargetImage} from '../lib/targetedCapture.mjs';
import {captureMetadata, withCaptureSession} from './capture_session.mjs';

function captureIdentityUrl(value) {
  const url=new URL(value),hash=new URLSearchParams(url.hash.slice(1));
  // Sheets consumes this navigation hint after the measured selection is set.
  // Keep every other URL component, including the workbook and worksheet IDs.
  url.searchParams.delete('range');hash.delete('range');
  url.hash=hash.toString();return url.href;
}

// Load native image libraries only when processing pixels. Awaited browser
// runtimes can provide a local subprocess processor without importing Sharp.
export async function processTargetScreenshotWithNode(bytes, options) {
  const directory=fs.mkdtempSync(path.join(os.tmpdir(),'asb-target-pixels-'));
  try {
    const input=path.join(directory,'source'),output=path.join(directory,'crop.png');
    fs.writeFileSync(input,bytes);
    const script=fileURLToPath(new URL('./process_target_screenshot.mjs',import.meta.url));
    const {stdout}=await promisify(execFile)('node',[script,input,output,JSON.stringify(options)],{maxBuffer:1024*1024});
    return {...JSON.parse(stdout),cropped:fs.readFileSync(output)};
  } finally {fs.rmSync(directory,{recursive:true,force:true});}
}

export async function processTargetScreenshot(bytes, {clip, viewport}, {hostProcess=false}={}) {
  let sharp;
  try {sharp=(await import('sharp')).default;}
  catch(error) {
    // The awaited browser runtime accepts JS modules but not Sharp's package
    // JSON import. Process these already-captured pixels in the host runtime.
    if(!hostProcess&&String(error.message).includes('Unsupported import specifier'))
      return processTargetScreenshotWithNode(bytes,{clip,viewport});
    throw error;
  }
  const details = await sharp(bytes).metadata();
  const sx = details.width / viewport.width, sy = details.height / viewport.height;
  if (Math.abs(sx - sy) > 0.02) throw Error('Screenshot and native viewport scales disagree');
  const pixelCrop = {left: Math.floor(clip.x*sx), top: Math.floor(clip.y*sy),
    width: Math.ceil((clip.x+clip.width)*sx)-Math.floor(clip.x*sx),
    height: Math.ceil((clip.y+clip.height)*sy)-Math.floor(clip.y*sy)};
  const sourceExtension = {png:'png',jpeg:'jpg',webp:'webp'}[details.format];
  if (!sourceExtension) throw Error('Fresh screenshot has an unsupported image format');
  return {cropped: await sharp(bytes).extract(pixelCrop).png().toBuffer(), pixelCrop, sourceExtension};
}

export async function captureTargetedSheetsImages(tab, plan, {measureTarget, verifySource,
  processScreenshot = processTargetScreenshot} = {}) {
  if (typeof measureTarget !== 'function') throw Error('A controller must measure each visible native range/chart target');
  const {taskDir, taskId, spreadsheetId, outputWorkbookPath, sourceCheck} = plan;
  const metadataFile = path.join(taskDir, 'metadata.json');
  const metadata = captureMetadata(taskDir, taskId);
  if (metadata.spreadsheetId !== spreadsheetId) throw Error('Capture belongs to another spreadsheet');
  if (plan.retryMissing && typeof verifySource !== 'function') throw Error('Sheets retry requires a fresh native revision check');
  if (verifySource) await verifySource();
  if (sourceCheck?.driveMetadata?.id !== spreadsheetId || !sourceCheck.modifiedBeforeOriginalExport ||
      !(Date.parse(sourceCheck.driveMetadata.modified_time) <= Date.parse(sourceCheck.originalExportedAt))) {
    throw Error('Verify the native spreadsheet is unchanged since the exact output export');
  }
  if (!Array.isArray(plan.targets) || !plan.targets.length) throw Error('Specific range/chart targets are required');
  const manifest = await withCaptureSession({taskDir, taskId, workbook: outputWorkbookPath, platform: 'google_sheets',
    targets: plan.targets, retryMissing: plan.retryMissing, captureRole: plan.captureRole || 'runner'}, async (target, stem) => {
      if (!['document','chart'].includes(target.kind)) throw Error('Target kind must be document or chart');
      if(target.chartIdentityError)throw Error(target.chartIdentityError);
      const kind = target.kind;
      const source = kind === 'chart' ? 'native-google-sheets-chart-screenshot' : 'native-google-sheets-range-screenshot';
      validateTargetImage({...target, kind, source, captureRole: 'runner', file: 'new-target.png', capturedAt: new Date().toISOString()}, 'google_sheets');
      if (target.sheetId === undefined) throw Error(`Missing native sheet mapping: ${target.sheetName}`);
      const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit?gid=${target.sheetId}#gid=${target.sheetId}`
        + (target.range || target.anchorRange ? `&range=${encodeURIComponent(target.range || target.anchorRange)}` : '');
      await tab.goto(url);
      const actual = new URL(await tab.url());
      if (actual.pathname.split('/')[3] !== spreadsheetId ||
          new URLSearchParams(actual.hash.slice(1)).get('gid') !== String(target.sheetId)) throw Error('Capture target changed');
      const observed = await measureTarget(tab, target);
      const chartTitle=value=>String(value||'').split('  Press')[0].replace(/\s+/g,' ').trim();
      if (observed.spreadsheetId !== spreadsheetId || String(observed.sheetId) !== String(target.sheetId)
          || observed.sheetName !== target.sheetName || observed.targetViewed !== true
          || (kind === 'document' && observed.range !== target.range)
          || (kind === 'chart' && (observed.chartName !== target.chartName || observed.chartIndex !== target.chartIndex))) {
        throw Error('Observed native range/chart does not match the requested target');
      }
      if(kind==='chart' && target.nativeChartTitle && chartTitle(observed.nativeChartLabel)!==chartTitle(target.nativeChartTitle)) {
        throw Error('Observed native chart title differs from its verified workbook identity');
      }
      const {clip, viewport} = observed;
      const valid = value => typeof value === 'number' && Number.isFinite(value);
      if (!clip || !viewport || ![clip.x,clip.y,clip.width,clip.height,viewport.width,viewport.height].every(valid)
          || clip.x < 0 || clip.y < 0 || clip.width <= 0 || clip.height <= 0
          || clip.x + clip.width > viewport.width || clip.y + clip.height > viewport.height) throw Error('Complete target bounds must be visible');
      // Only a fresh screenshot API is used. This is deliberately independent
      // of any tab-frames directory, video sample, clipboard or recording file.
      const bytes = await tab.screenshot({fullPage: false});
      if (captureIdentityUrl(await tab.url()) !== captureIdentityUrl(actual.href)) throw Error('Target changed during fresh screenshot capture');
      // Sidebars can finish opening/closing between measurement and screenshot.
      // Reject those pixels rather than registering an occluded target crop.
      const after=await measureTarget(tab,target);
      if(after.targetViewed!==true||after.spreadsheetId!==observed.spreadsheetId||
         String(after.sheetId)!==String(observed.sheetId)||after.sheetName!==observed.sheetName||
         after.range!==observed.range||after.chartName!==observed.chartName||after.chartIndex!==observed.chartIndex||after.nativeChartLabel!==observed.nativeChartLabel||
         JSON.stringify(after.clip)!==JSON.stringify(clip)||JSON.stringify(after.viewport)!==JSON.stringify(viewport)) {
        throw Error('Native target layout changed during fresh screenshot capture');
      }
      const {cropped, pixelCrop: crop, sourceExtension} = await processScreenshot(bytes, {clip, viewport});
      if (!ArrayBuffer.isView(cropped) || cropped.BYTES_PER_ELEMENT !== 1 || !cropped.length
          || !['png','jpg','webp'].includes(sourceExtension)) {
        throw Error('Screenshot processor did not return a supported nonempty image');
      }
      const file = `${stem}.png`;
      const sourceScreenshot = `${stem}.source.${sourceExtension}`;
      const sourceFile = path.join(taskDir, sourceScreenshot), outputFile = path.join(taskDir, file);
      fs.writeFileSync(sourceFile, bytes, {flag: 'wx'});
      fs.writeFileSync(outputFile, cropped, {flag: 'wx'});
      if (verifySource) await verifySource();
      return {kind, file, source, clip, viewport, pixelCrop: crop,
        ...(kind==='chart' && observed.nativeChartLabel?{nativeChartLabel:observed.nativeChartLabel}:{}),
        sourceScreenshot, sourceScreenshotSha256: imageHash(sourceFile),
        capturedAt: new Date().toISOString(), sha256: imageHash(outputFile)};
  });
  const {images, errors} = manifest;
  if (!plan.retryMissing) {
    metadata.artifacts = {...metadata.artifacts, documentImages: images.filter(i => i.kind==='document').map(i=>i.file),
      chartImages: images.filter(i=>i.kind==='chart').map(i=>i.file)};
    metadata.outputImagePolicy = 'targeted-output-images/v1';
    metadata.capture = {source: 'fresh-targeted-native-google-sheets', capturedAt: manifest.capturedAt, imageManifest: path.basename(manifest.manifestFile), errors};
    metadata.reviewedSheetEvidence = {source: metadata.capture.source, capturedAt: manifest.capturedAt,
      images: images.filter(item=>item.kind==='document')};
    metadata.focusedChartEvidence = {source: metadata.capture.source, capturedAt: manifest.capturedAt,
      images: images.filter(item=>item.kind==='chart')};
    fs.writeFileSync(metadataFile, JSON.stringify(metadata,null,2)+'\n');
  }
  return manifest;
}
