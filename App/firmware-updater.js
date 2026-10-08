export const firmwareModuleURL=import.meta.url;
// STK500v1 for classic 16 MHz Nano / ATmega328P. No erase/fuse/bootloader commands.
export const FLASH_LIMIT=30720; // Preserve the larger, 2 KB old Nano bootloader.
export const PAGE_SIZE=128;
export const NANO_SIGNATURE=[0x1e,0x95,0x0f];
const sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms));

export function parseIntelHex(text) {
  if(typeof text!=='string'||text.length>200000)throw new Error('Firmware file is empty or too large.');
  const bytes=new Uint8Array(FLASH_LIMIT).fill(255),used=new Uint8Array(FLASH_LIMIT);
  let base=0,end=0,total=0,eof=false;
  for(const [lineIndex,raw] of text.replace(/^\uFEFF/,'').split(/\r?\n/).entries()) {
    const line=raw.trim();if(!line)continue;
    const fail=reason=>{throw new Error(`HEX line ${lineIndex+1}: ${reason}`);};
    if(eof)fail('data after end-of-file record.');
    if(!/^:(?:[0-9a-fA-F]{2})+$/.test(line))fail('invalid Intel HEX syntax.');
    const record=Uint8Array.from(line.slice(1).match(/../g),n=>parseInt(n,16));
    const length=record[0];if(record.length!==length+5)fail('incorrect record length.');
    if(record.reduce((sum,n)=>sum+n,0)%256!==0)fail('checksum mismatch.');
    const address=record[1]*256+record[2],kind=record[3],data=record.slice(4,-1);
    if(kind===0) {
      const start=base+address;
      if(start+length>FLASH_LIMIT)fail('firmware overlaps the reserved Nano bootloader or exceeds 30 KB. Use the .hex WITHOUT bootloader.');
      for(let i=0;i<length;i++){
        if(used[start+i])fail('overlapping data records.');
        used[start+i]=1;bytes[start+i]=data[i];
      }
      if(length){end=Math.max(end,start+length);total+=length;}
    }else if(kind===1){if(length!==0||address!==0)fail('invalid end-of-file record.');eof=true;}
    else if(kind===2||kind===4){if(length!==2||address!==0)fail('invalid extended address record.');base=(data[0]*256+data[1])*(kind===2?16:65536);}
    else if(kind===3||kind===5){if(length!==4||address!==0)fail('invalid start address record.');}
    else fail('unsupported record type.');
  }
  if(!eof||!total||!used[0]||!used[1]||(bytes[0]===255&&bytes[1]===255))throw new Error('Choose a complete application HEX with a reset vector and end-of-file record.');
  const image=bytes.slice(0,Math.ceil(end/PAGE_SIZE)*PAGE_SIZE);
  return {bytes:image,dataBytes:total,pageCount:image.length/PAGE_SIZE};
}

function deadline(promise,ms,label) {
  let timer;
  return Promise.race([promise,new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error(label)),ms);})]).finally(()=>clearTimeout(timer));
}

// One continuously active reader owns the stream; timed-out reads never remain
// queued to consume responses belonging to later commands.
export class SerialBootloader {
  constructor(port) {
    this.port=port;this.reader=port.readable.getReader();this.writer=port.writable.getWriter();
    this.queue=[];this.error=null;this.notify=null;this.closing=false;
    this.pump=this.readLoop();
  }
  async readLoop(){
    try{
      while(!this.closing){
        const {value,done}=await this.reader.read();
        if(done){if(!this.closing)this.error=new Error('USB connection closed.');break;}
        if(value){this.queue.push(...value);if(this.queue.length>8192)throw new Error('Unexpected serial data during update.');}
        this.notify?.();
      }
    }catch(e){if(!this.closing)this.error=e;}
    finally{this.notify?.();this.reader.releaseLock();}
  }
  async take(count,timeout){
    const until=Date.now()+timeout;
    while(this.queue.length<count){
      if(this.error)throw this.error;
      if(this.closing)throw new Error('USB update connection closed.');
      const remaining=until-Date.now();if(remaining<=0)throw new Error('Bootloader response timed out.');
      await new Promise((resolve,reject)=>{
        const timer=setTimeout(()=>{this.notify=null;reject(new Error('Bootloader response timed out.'));},remaining);
        this.notify=()=>{clearTimeout(timer);this.notify=null;resolve();};
      });
    }
    return Uint8Array.from(this.queue.splice(0,count));
  }
  async command(command,payload=[],length=0){
    if(this.error)throw this.error;
    // A new request is issued only after the previous complete response.
    await deadline(this.writer.write(Uint8Array.from([command,...payload,0x20])),2500,'USB write timed out.');
    const result=await this.take(length+2,2000);
    if(result[0]!==0x14||result.at(-1)!==0x10)throw new Error('Unexpected bootloader response.');
    return result.slice(1,-1);
  }
  async reset(delay){
    await this.port.setSignals({dataTerminalReady:false,requestToSend:false});await sleep(50);
    await this.port.setSignals({dataTerminalReady:true,requestToSend:true});await sleep(100);
    await this.port.setSignals({dataTerminalReady:false,requestToSend:false});
    await sleep(delay);this.queue.length=0;
  }
  async close(){
    this.closing=true;this.notify?.();
    try{await deadline(this.reader.cancel(),2500,'USB reader did not close.');}catch{}
    await deadline(this.pump,3000,'USB connection did not close.');
    try{await deadline(this.writer.abort(),2500,'USB writer did not close.');}catch{}
    this.writer.releaseLock();
    await deadline(this.port.close(),3000,'USB port is still busy. Unplug and reconnect TURTLE before retrying.');
  }
}

export async function flashNano(port,image,onProgress=()=>{},options={}) {
  // Validate again at this boundary, even if callers bypass the file picker.
  if(!(image?.bytes instanceof Uint8Array)||!image.bytes.length||image.bytes.length>FLASH_LIMIT||image.bytes.length%PAGE_SIZE)
    throw new Error('Invalid Nano application image.');
  const mode=options.mode||'auto';
  if(!['auto','modern','old'].includes(mode))throw new Error('Unknown bootloader setting.');
  const rates=mode==='modern'?[115200]:mode==='old'?[57600]:[115200,57600];
  const makeTransport=options.makeTransport||(p=>new SerialBootloader(p));
  const delays=options.resetDelays||[100,350];
  let transport=null,baud=null,changed=false,lastFailure=null;
  try{
    for(const rate of rates){
      for(const delay of delays){
        onProgress({stage:'detecting',percent:0,text:`Detecting ${rate===115200?'modern':'old'} Nano bootloader…`});
        let opened=false;
        try{
          await port.open({baudRate:rate,dataBits:8,stopBits:1,parity:'none',flowControl:'none',bufferSize:4096});opened=true;
          transport=makeTransport(port);
          await transport.reset(delay);
          await transport.command(0x30); // GET_SYNC
          const signature=await transport.command(0x75,[],3); // READ_SIGN
          if(!NANO_SIGNATURE.every((n,i)=>signature[i]===n)){
            const e=new Error('The selected device is not an ATmega328P Nano. No firmware was written.');e.wrongChip=true;throw e;
          }
          baud=rate;break;
        }catch(e){
          lastFailure=e;
          if(transport){const closing=transport;transport=null;await closing.close();}
          else if(opened)await port.close();
          if(e.wrongChip)throw e;
        }
      }
      if(baud)break;
    }
    if(!transport)throw new Error('Could not reach either Nano bootloader. Close Arduino IDE/Serial Monitor and retry. If needed, choose the bootloader manually. '+(lastFailure?.message||''));
    onProgress({stage:'detected',percent:0,text:`ATmega328P found · ${baud===115200?'modern':'old'} bootloader (${baud} baud).`});
    await transport.command(0x50); // ENTER_PROGMODE
    const pages=image.bytes.length/PAGE_SIZE;
    for(let page=0;page<pages;page++){
      const address=page*PAGE_SIZE,word=address/2;
      await transport.command(0x55,[word&255,word>>8]); // LOAD_ADDRESS (words)
      changed=true;
      await transport.command(0x64,[0,PAGE_SIZE,0x46,...image.bytes.slice(address,address+PAGE_SIZE)]); // PROG_PAGE flash
      onProgress({stage:'writing',percent:Math.round((page+1)/pages*70),text:`Uploading firmware · ${page+1} of ${pages} pages`});
    }
    for(let page=0;page<pages;page++){
      const address=page*PAGE_SIZE,word=address/2;
      await transport.command(0x55,[word&255,word>>8]);
      const actual=await transport.command(0x74,[0,PAGE_SIZE,0x46],PAGE_SIZE); // READ_PAGE flash
      if(actual.length!==PAGE_SIZE||actual.some((byte,i)=>byte!==image.bytes[address+i]))throw new Error(`Verification failed at flash address 0x${address.toString(16)}. Reconnect and retry the update.`);
      onProgress({stage:'verifying',percent:70+Math.round((page+1)/pages*29),text:`Verifying firmware · ${page+1} of ${pages} pages`});
    }
    let restartConfirmed=true;
    try{await transport.command(0x51);}catch{restartConfirmed=false;} // LEAVE_PROGMODE, some bridges reset before ACK
    const closing=transport;transport=null;await closing.close();
    onProgress({stage:'complete',percent:100,text:restartConfirmed?'Firmware uploaded and verified. Click Connect USB to resume.':'Firmware verified. Unplug and reconnect TURTLE, then click Connect USB.'});
    return {baud,verified:true,restartConfirmed};
  }catch(e){
    if(changed)e.message+=' Keep the same firmware file and retry; Arduino IDE can also restore the application.';
    throw e;
  }finally{if(transport)await transport.close();}
}
