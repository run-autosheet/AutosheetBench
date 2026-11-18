// Public runner/capture coordination. Never remove another process's lock or
// guess that a lock is stale. Node and the Python launcher share this filename.
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import {randomUUID} from 'node:crypto';

export function nativeSessionLock() {
  const directory=process.env.AUTOSHEETBENCH_NATIVE_LOCK_DIR || os.tmpdir();
  return path.join(directory,`autosheetbench-public-excel-${process.getuid?.() || 0}.lock`);
}

export function assertNoActiveBatch(taskDir) {
  for(let directory=path.resolve(taskDir);;directory=path.dirname(directory)) {
    if(fs.existsSync(path.join(directory,'.public-run.lock'))) {
      throw Error('Screenshot retry blocked: the parent public batch is active or locked');
    }
    if(directory===path.dirname(directory)) break;
  }
}

function exclusiveLock(file,owner) {
  const token=randomUUID();
  fs.writeFileSync(file,JSON.stringify({owner,token,pid:process.pid})+'\n',{flag:'wx'});
  return ()=>{
    // Do not remove a replaced lock. Interrupted owners need explicit operator
    // recovery after verifying that the process/session has stopped.
    const current=JSON.parse(fs.readFileSync(file,'utf8'));
    if(current.token!==token) throw Error('Capture lock ownership changed; refusing to remove it');
    fs.unlinkSync(file);
  };
}

export function acquireCaptureLocks({taskDir,platform,retryMissing=false,captureRole='runner'}) {
  const release=[];
  try {
    if(retryMissing) {
      assertNoActiveBatch(taskDir);
      // Acquire the same lock as the launcher: a check alone would race with
      // the next batch start. Manual tasks use the documented tasks/<id> layout.
      if(path.basename(path.dirname(taskDir))==='tasks') {
        release.push(exclusiveLock(path.join(path.dirname(path.dirname(taskDir)),'.public-run.lock'),'capture-retry'));
      }
    }
    if(platform==='microsoft_excel') {
      const file=nativeSessionLock();
      const token=process.env.AUTOSHEETBENCH_EXCEL_SESSION_TOKEN;
      let borrowed=false;
      if(!retryMissing && captureRole==='runner' && token && fs.existsSync(file)) {
        borrowed=JSON.parse(fs.readFileSync(file,'utf8')).token===token;
      }
      if(!borrowed) {
        try {release.push(exclusiveLock(file,'native-capture'));}
        catch(error) {
          if(error.code==='EEXIST') throw Error('Native Excel session is busy; screenshot capture is deferred');
          throw error;
        }
      }
    }
    return ()=>{for(const close of release.reverse()) close();};
  } catch(error) {
    for(const close of release.reverse()) close();
    throw error;
  }
}
