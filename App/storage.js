// Metadata and samples are updated together in one IndexedDB transaction.
export class RecordingStore {
  async open() {
    this.db = await new Promise((resolve,reject)=>{
      const r=indexedDB.open('turtle-recordings',1);
      r.onupgradeneeded=()=>{
        r.result.createObjectStore('sessions',{keyPath:'id'});
        const samples=r.result.createObjectStore('samples',{keyPath:['sessionId','seq']});
        samples.createIndex('sessionId','sessionId');
      };
      r.onsuccess=()=>resolve(r.result);
      r.onerror=()=>reject(r.error);
      r.onblocked=()=>reject(new Error('Close other TURTLE tabs and reload.'));
    });
    this.db.onversionchange=()=>this.db.close();
  }
  transact(stores,mode,fn){
    return new Promise((resolve,reject)=>{
      const tx=this.db.transaction(stores,mode);let result;
      tx.oncomplete=()=>resolve(typeof result==='function'?result():result);
      tx.onerror=()=>reject(tx.error||new Error('Storage operation failed.'));
      tx.onabort=()=>reject(tx.error||new Error('Storage operation was interrupted.'));
      try{result=fn(tx);}catch(error){tx.abort();reject(error);}
    });
  }
  create(session){return this.transact(['sessions'],'readwrite',tx=>{tx.objectStore('sessions').add(session);});}
  append(session,row){return this.transact(['sessions','samples'],'readwrite',tx=>{
    tx.objectStore('samples').add({...row,sessionId:session.id,seq:session.count});
    tx.objectStore('sessions').put(session);
  });}
  finish(session){return this.transact(['sessions'],'readwrite',tx=>{tx.objectStore('sessions').put(session);});}
  list(){return this.transact(['sessions'],'readonly',tx=>{const r=tx.objectStore('sessions').getAll();return()=>r.result.sort((a,b)=>b.startedAt.localeCompare(a.startedAt));});}
  rows(id){return this.transact(['samples'],'readonly',tx=>{const r=tx.objectStore('samples').index('sessionId').getAll(id);return()=>r.result.sort((a,b)=>a.seq-b.seq);});}
  remove(id){return this.transact(['sessions','samples'],'readwrite',tx=>{
    tx.objectStore('sessions').delete(id);
    const r=tx.objectStore('samples').index('sessionId').openCursor(IDBKeyRange.only(id));
    r.onsuccess=()=>{const cursor=r.result;if(cursor){cursor.delete();cursor.continue();}};
  });}
}

// Separate rows from the preamble so periodic appends never repeat metadata.
export function csvRows(rows) {
  return rows.map(r=>[r.elapsed.toFixed(3),r.values[0]??'',r.values[1]??''].join(',')+'\r\n').join('');
}
export function sessionCSV(session, rows) {
  function cell(value){let s=String(value??'');if(/^[\s]*[=+\-@]/.test(s))s="'"+s;return '"'+s.replaceAll('"','""')+'"';}
  const headers=['Elapsed Time (s)',...Array.from({length:2},(_,i)=>
    `Thermocouple ${i+1}${session.names?.[i]?' - '+session.names[i]:''} (°C)`)];
  // BOM helps Excel recognize UTF-8 channel names and the degree symbol.
  return '\uFEFF'+['Date Created,'+cell(session.startedAt)+',',
    'Session Name,'+cell(session.name)+',',',,',',,',headers.map(cell).join(',')].join('\r\n')+'\r\n'+csvRows(rows);
}
