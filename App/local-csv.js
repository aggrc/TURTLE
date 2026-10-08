import {sessionCSV,csvRows} from './storage.js?v=8';

// A folder grant lets us recover from a locked CSV without another file picker.
// Every continuation is a complete recording, so users never need to merge files.
export class LocalCSV {
  constructor(directory, session) {
    this.directory=directory; this.session=session; this.handle=null;
    this.count=0; this.offset=0; this.lastSaved=null; this.pending=0;
    this.error=null; this.chain=Promise.resolve(); this.initialized=false;
    this.revision=0; this.notice=''; this.savedMetadata=null;
  }
  async write(handle, text, append) {
    let stream;
    try {
      if(append && (await handle.getFile()).size!==this.offset)
        throw new Error('The CSV was changed outside TURTLE.');
      const bytes=new TextEncoder().encode(text);
      stream=await handle.createWritable({keepExistingData:append,mode:'exclusive'});
      await stream.seek(append?this.offset:0);
      await stream.write(bytes);
      const size=(append?this.offset:0)+bytes.length;
      await stream.truncate(size);await stream.close();
      return size;
    } catch(error) {if(stream)try{await stream.abort();}catch{}throw error;}
  }
  save(rows,session=this.session) {
    const metadata={...session,names:[...session.names]};
    const metadataKey=JSON.stringify([metadata.name,metadata.names,metadata.startedAt]);
    const snapshot=rows.slice();this.pending++;
    const operation=this.chain.then(async()=>{
      const append=this.initialized&&metadataKey===this.savedMetadata;
      if(append && snapshot.length<=this.count && !this.error)return;
      try {
        if(!this.handle)this.handle=await this.directory.getFileHandle('recording.csv',{create:true});
        const text=append?csvRows(snapshot.slice(this.count)):sessionCSV(metadata,snapshot);
        try{this.offset=await this.write(this.handle,text,append);}
        catch(error){
          // Permission loss or a full disk cannot be solved by rotating a file.
          if(['NotAllowedError','SecurityError','QuotaExceededError','AbortError'].includes(error.name))throw error;
          const name=`recording-continuation-${String(++this.revision).padStart(3,'0')}.csv`;
          const next=await this.directory.getFileHandle(name,{create:true});
          this.offset=await this.write(next,sessionCSV(metadata,snapshot),false);
          this.handle=next;
          this.notice='Previous CSV locked or changed. Saving the complete recording to '+name+'.';
        }
        this.session=metadata;this.savedMetadata=metadataKey;this.count=snapshot.length;this.initialized=true;this.lastSaved=new Date();this.error=null;
      }catch(error){this.error=error;throw error;}
    });
    this.chain=operation.catch(()=>{});
    return operation.finally(()=>{this.pending--;});
  }
}
