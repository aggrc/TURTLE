import {parseIntelHex,flashNano,firmwareModuleURL} from './firmware-updater.js?v=8';
import {LocalCSV} from './local-csv.js?v=8';
import {parsePacket,createLineDecoder,configCommand,faultLabel} from './protocol.js?v=8';
import {RecordingStore,sessionCSV} from './storage.js?v=8';
const $=id=>document.getElementById(id), clock=()=>performance.now(), store=new RecordingStore();
let port=null,reader=null,readTask=null,connected=false,busy=false,closing=false,writing=false;
let latestFaults=[null,null],deviceSeq=null,deviceMs=null,deviceElapsed=0,recordDeviceStart=null,droppedPackets=0;
let pendingAck=null,commandToken=0,settingsConfirmed=false;
let lastPacket=0,latest=[null,null],trace=[],requested={type:'T',rate:1};
let names=['',''],sessions=[],current=null,rows=[],recording=false,recordStart=0,elapsed=0;
let ready=false,storageError=false,saving=0,transition=false,saveChain=Promise.resolve(),downloadedFallback=false;
let releaseSession=null,localFile=null,csvTimer=null;
let updating=false,firmwareLoading=false,firmwareImage=null,firmwareName='',firmwareLoadToken=0;
const supportsFiles=window.isSecureContext&&'showDirectoryPicker' in window;
function clearCSVTimer(){clearInterval(csvTimer);csvTimer=null;}
async function saveLocalCSV(){
  if(!localFile)return true;
  try{const work=localFile.save(rows,current);render();await work;if(localFile.notice)message(localFile.notice);render();return true;}
  catch(e){message('Local CSV update failed. Recording remains in the browser backup. Check folder permissions and disk space, then retry saving. After stopping, you can download the browser backup.',true);render();return false;}
}
async function holdSession(id){
  if(!navigator.locks)throw new Error('This browser does not support safe session locking.');
  return new Promise((resolve,reject)=>{navigator.locks.request('turtle-session-'+id,{ifAvailable:true},lock=>{
    if(!lock){resolve(null);return;}
    return new Promise(release=>resolve(release));
  }).catch(reject);});
}
const supportsSerial=window.isSecureContext&&'serial' in navigator;
function message(text,error=false){$('message').textContent=text;$('message').classList.toggle('error',error);}
function fresh(){return connected&&settingsConfirmed&&lastPacket>0&&clock()-lastPacket<Math.max(5000,(requested.rate+3)*1000);}
function duration(ms){const n=Math.floor(ms/1000);return`${String(Math.floor(n/60)).padStart(2,'0')}:${String(n%60).padStart(2,'0')}`;}
function render(){
  $('firmware-file').disabled=updating||firmwareLoading;
  $('bootloader-mode').disabled=updating;
  $('update-firmware').disabled=!supportsSerial||connected||recording||busy||transition||writing||updating||firmwareLoading||!firmwareImage;
  $('firmware-help').textContent=recording?'Stop and save your recording, then disconnect to update.':connected?'Disconnect TURTLE above before updating its firmware.':'Classic Arduino Nano / ATmega328P only. Keep the USB cable connected until verification finishes.';
  $('connect').hidden=connected;$('disconnect').hidden=!connected;
  $('connect').disabled=!supportsSerial||busy||transition||updating;$('disconnect').disabled=busy||writing||transition;
  $('record').disabled=transition||busy||(!recording&&(!supportsFiles||!ready||storageError||writing||!fresh()||!latest.some(v=>v!==null)));
  $('record').textContent=transition?'Saving…':recording?'Stop & save recording':'Start new recording';$('record').classList.toggle('danger',recording);
  $('reset-faults').disabled=!connected||busy||writing||transition;
  $('apply').disabled=!connected||busy||recording||writing||transition;
  for(const id of ['type','rate'])$(id).disabled=recording||transition||busy||writing;
  for(const id of ['session-name','name-0','name-1'])$(id).disabled=transition||busy;
  $('connection-title').textContent=connected?'TURTLE connected':'Connect your TURTLE';
  $('feed-status').textContent=recording?'RECORDING':trace.length?'STOPPED':'READY';
  $('chart-subtitle').textContent=recording?'Recording · Latest 300 samples':trace.length?'Recording stopped · Graph frozen':'Start a recording to plot temperatures';
  $('save-now').hidden=!localFile;
  $('save-now').disabled=transition||!!localFile?.pending;
  $('save-now').textContent=localFile?.error?'Retry CSV save':'Save CSV now';
  $('file-status').textContent=localFile ? localFile.error ? 'CSV save failed · Browser backup available' : localFile.pending ? 'Updating '+(localFile.handle?.name||'recording.csv')+'…' : `${(localFile.handle?.name||'recording.csv')} · ${localFile.count.toLocaleString()} samples saved${localFile.lastSaved?' · '+localFile.lastSaved.toLocaleTimeString():''}` : supportsFiles ? 'Choose a folder when you start. Saves every 10 seconds; automatically continues in a new CSV if the active file is locked.' : 'Recording requires local folder access. Open this site in desktop Chrome or Edge.';
  $('file-status').classList.toggle('save-error',!!localFile?.error);
  $('count').textContent=(current?.count||0).toLocaleString();
  $('elapsed').textContent=duration(recording?clock()-recordStart:elapsed);
  for(let i=0;i<2;i++){
    $('temp-'+i).textContent=latest[i]===null?'—':latest[i].toFixed(2);
    $('state-'+i).textContent=!connected?'Offline':!lastPacket?'Awaiting data':!fresh()?'Stale':latest[i]===null?'Fault / disconnected':'Connected';
    $('note-'+i).textContent=!lastPacket?'No measurements received':!fresh()?'Last reading · not live':faultLabel(latestFaults[i]);
    $('legend-'+i).textContent=(trace.length&&current?current.names:names)[i]?`TC ${i+1} · ${(trace.length&&current?current.names:names)[i]}`:`TC ${i+1}`;
  }
  $('last-packet').textContent=lastPacket?`Last reading ${Math.floor((clock()-lastPacket)/1000)}s ago${droppedPackets?' · '+droppedPackets+' missed packets':''}`:'No data yet';
  $('storage-status').textContent=storageError?'Automatic saving unavailable. Download the latest CSV before closing this tab.':!ready?'Opening saved recordings…':saving?'Saving samples…':recording?'Browser backup up to date.':'Browser backup ready.';
  $('storage-status').classList.toggle('save-error',storageError);
}
function failStorage(error){
  storageError=true;downloadedFallback=false;clearCSVTimer();void saveLocalCSV();
  if(recording){recording=false;elapsed=clock()-recordStart;if(current){current={...current,status:'save-error',endedAt:new Date().toISOString(),durationMs:elapsed};}}
  message('Recording stopped: browser storage is unavailable or full. Download the latest CSV to keep all readings from this run.',true);
  releaseSession?.();releaseSession=null;
  render();renderSessions();console.error('TURTLE storage:',error);
}
function enqueueSave(fn){
  saving++;
  saveChain=saveChain.then(()=>{if(storageError)return;return fn();}).catch(failStorage).finally(()=>{saving--;render();});
  return saveChain;
}
function ingest(line){
  const ack=line.match(/^ACK:(\d+),([BEJKNRST]),([0-5])$/);
  if(ack){
    if(pendingAck&&Number(ack[1])===pendingAck.token){
      const pending=pendingAck;pendingAck=null;clearTimeout(pending.timer);
      if(ack[2]!==pending.type||Number(ack[3])!==pending.rate)pending.reject(new Error('Device settings did not match.'));
      else {requested={type:ack[2],rate:Number(ack[3])};settingsConfirmed=true;pending.resolve();}
    }
    return;
  }
  if(line.startsWith('ERR:')){if(pendingAck){const p=pendingAck;pendingAck=null;clearTimeout(p.timer);p.reject(new Error(line));}return;}
  const packet=parsePacket(line);if(!packet)return;
  if(packet.legacy||!settingsConfirmed)return;
  if(packet.type!==requested.type||packet.rate!==requested.rate){
    settingsConfirmed=false;void stopRecording('Device settings changed');message('Device settings changed. Send settings again before recording.',true);return;
  }
  if(deviceSeq!==null){
    const step=(packet.seq-deviceSeq)>>>0;
    if(step===0)return; // ignore duplicate packets
    const dt=(packet.ms-deviceMs)>>>0;
    if(step>0x7fffffff||dt>0x7fffffff){
      settingsConfirmed=false;void stopRecording('Device restarted');message('TURTLE restarted. Send settings again to continue.',true);return;
    }
    droppedPackets+=step-1;deviceElapsed+=dt;
  }
  deviceSeq=packet.seq;deviceMs=packet.ms;
  const values=packet.values;latest=values;latestFaults=packet.faults;lastPacket=clock();
  if(recording){
    if(recordDeviceStart===null)recordDeviceStart=deviceElapsed;
    const row={utc:new Date().toISOString(),elapsed:(deviceElapsed-recordDeviceStart)/1000,values:[...values],faults:[...latestFaults],deviceMs:packet.ms,sequence:packet.seq};rows.push(row);
    trace.push({t:row.elapsed,values:[...values]});if(trace.length>300)trace.shift();
    current={...current,count:rows.length,durationMs:lastPacket-recordStart};
    const snapshot={...current,names:[...current.names]};enqueueSave(()=>store.append(snapshot,row));
    if(rows.length>=100000)void stopRecording('Sample limit reached');
  }
  if(recording)draw();render();
}
function draw(){
  const svg=$('graph'),NS='http://www.w3.org/2000/svg';svg.replaceChildren();
  function node(tag,attrs,text){const n=document.createElementNS(NS,tag);for(const[k,v]of Object.entries(attrs))n.setAttribute(k,String(v));if(text!==undefined)n.textContent=text;svg.append(n);}
  const vals=trace.flatMap(p=>p.values).filter(v=>v!==null);$('chart-empty').hidden=vals.length>0;
  let lo=vals.length?Math.min(...vals):0,hi=vals.length?Math.max(...vals):40;const pad=Math.max(2,(hi-lo)*.12);lo-=pad;hi+=pad;
  const start=trace[0]?.t||0,end=Math.max(start+10,trace.at(-1)?.t||10),x=t=>60+(t-start)/(end-start)*720,y=v=>250-(v-lo)/(hi-lo)*220;
  for(let i=0;i<=4;i++){const yy=30+i*55;node('line',{x1:60,y1:yy,x2:780,y2:yy,stroke:'#e6edf2'});node('text',{x:50,y:yy+5,'text-anchor':'end'},(hi-i*(hi-lo)/4).toFixed(1));const t=start+(end-start)*i/4;node('text',{x:x(t),y:278,'text-anchor':'middle'},t.toFixed(0)+'s');}
  node('text',{x:12,y:16},'°C');
  for(let c=0;c<2;c++){let d='',pen=false;for(const p of trace){if(p.values[c]===null){pen=false;continue;}d+=`${pen?'L':'M'}${x(p.t).toFixed(2)},${y(p.values[c]).toFixed(2)} `;pen=true;}node('path',{d,stroke:c?'#356ac3':'#11675a','stroke-width':2.5,fill:'none'});const p=trace.at(-1);if(p&&p.values[c]!==null)node('circle',{cx:x(p.t),cy:y(p.values[c]),r:4,fill:c?'#356ac3':'#11675a'});}
}
async function sendSettings(reset=false){
  if(!port?.writable||writing)throw new Error('Device is not ready.');
  const type=reset?requested.type:$('type').value,rate=reset?requested.rate:Number($('rate').value);
  const token=commandToken=commandToken%65535+1,command=configCommand(type,rate,token);
  writing=true;settingsConfirmed=false;render();
  // Applying the same configuration explicitly clears both fault latches.
  const response=new Promise((resolve,reject)=>{
    const timer=setTimeout(()=>{pendingAck=null;reject(new Error('No settings confirmation. Install TURTLE firmware v4 and reconnect.'));},3000);
    pendingAck={token,type,rate,resolve,reject,timer};
  });
  response.catch(()=>{});
  try{
    const writer=port.writable.getWriter();
    try{await writer.write(new TextEncoder().encode(command));}finally{writer.releaseLock();}
    await response;
    if(!reset){lastPacket=0;deviceSeq=null;deviceMs=null;deviceElapsed=0;}
  }catch(e){if(pendingAck){clearTimeout(pendingAck.timer);pendingAck=null;}throw e;}
  finally{writing=false;render();}
}
async function readLoop(activePort){
  const decoder=createLineDecoder(ingest);let failure=null;
  try{reader=activePort.readable.getReader();while(!closing){const{value,done}=await reader.read();if(done)break;if(value)decoder.push(value);}}
  catch(e){if(!closing)failure=e;}finally{reader?.releaseLock();reader=null;}
  if(!closing){connected=false;busy=true;render();await stopRecording('Connection lost');try{await activePort.close();}catch{}if(port===activePort)port=null;busy=false;message('Device disconnected. '+(storageError?'Download the latest CSV before closing.':'Saved recordings are available below.')+(failure?' Reconnect to continue.':''),true);render();}
}
async function connect(){
  if(!supportsSerial||busy||connected||transition||updating)return;
  busy=true;render();let selected;
  try{
    selected=await navigator.serial.requestPort();await selected.open({baudRate:9600,dataBits:8,stopBits:1,parity:'none',flowControl:'none'});
    port=selected;connected=true;settingsConfirmed=false;deviceSeq=null;deviceMs=null;deviceElapsed=0;droppedPackets=0;closing=false;latest=[null,null];lastPacket=0;draw();render();
    try{await port.setSignals({dataTerminalReady:true,requestToSend:true});}catch{}
    readTask=readLoop(selected);message('Connecting to TURTLE…');await new Promise(r=>setTimeout(r,2000));
    if(port!==selected||!connected)return;await sendSettings();message('Connected. Start a new recording when you are ready.');
  }catch(e){if(port)await disconnect(false);else if(selected?.readable){try{await selected.close();}catch{}}message(e.name==='NotFoundError'?'No device selected.':`Could not connect: ${e.message}. Close the Python app or serial monitor and try again.`,e.name!=='NotFoundError');}
  finally{busy=false;render();}
}
async function disconnect(showMessage=true){
  busy=true;closing=true;render();await stopRecording('Disconnected');
  try{if(reader)await reader.cancel();if(readTask)await readTask;if(port)await port.close();}catch{}
  finally{port=null;reader=null;readTask=null;connected=false;busy=false;closing=false;render();if(showMessage&&!localFile?.error)message(storageError?'Disconnected. Download the latest CSV before closing.':'Disconnected. Your saved recordings are available below.',storageError);}
}
async function startRecording(){
  if(recording||transition||busy||writing||!supportsFiles||!ready||storageError||!fresh()||!latest.some(v=>v!==null))return;
  if(localFile?.error&&!downloadedFallback){message('Save or download the previous recording before starting another.',true);return;}
  transition=true;render();
  const startedAt=new Date().toISOString();
  const session={id:crypto.randomUUID(),name:$('session-name').value.trim()||`Recording · ${new Date().toLocaleString()}`,startedAt,names:[...names],...requested,count:0,durationMs:0,status:'open'};
  let nextFile=null;
  {
    try{
      const root=await window.showDirectoryPicker({mode:'readwrite',id:'turtle-recordings'});
      const folder=await root.getDirectoryHandle(`TURTLE_${startedAt.replace(/[:.]/g,'-')}_${session.id.slice(0,8)}`,{create:true});
      nextFile=new LocalCSV(folder,session);
      await nextFile.save([]);
    }catch(e){transition=false;message(e.name==='AbortError'?'Recording canceled. No new samples recorded.':'Could not open the recording folder. Choose another folder and allow access to start recording.',e.name!=='AbortError');render();return;}
  }
  try{
    releaseSession=await holdSession(session.id);if(!releaseSession)throw new Error('Session is already in use.');
    await store.create(session);localFile=nextFile;current=session;rows=[];trace=[];draw();elapsed=0;recordStart=clock();recordDeviceStart=null;droppedPackets=0;downloadedFallback=false;
    if(!fresh()||!connected){current={...current,status:'Connection lost',endedAt:new Date().toISOString()};await store.finish(current);releaseSession?.();releaseSession=null;await refreshSessions();return;}
    recording=true;clearCSVTimer();if(localFile)csvTimer=setInterval(()=>{if(!localFile.pending)void saveLocalCSV();},10000);$('session-name').value=current.name;message('Recording. CSV updates every 10 seconds; browser backup saves each sample.');await refreshSessions();
  }catch(e){failStorage(e);}finally{transition=false;render();}
}
async function stopRecording(reason='Completed'){
  if(!recording)return;
  recording=false;clearCSVTimer();transition=true;elapsed=clock()-recordStart;
  current={...current,status:reason,endedAt:new Date().toISOString(),durationMs:elapsed};const snapshot={...current};render();
  const fileSaved=await saveLocalCSV();
  await enqueueSave(()=>store.finish(snapshot));
  releaseSession?.();releaseSession=null;
  transition=false;
  if(!storageError&&fileSaved){message(reason==='Completed'?'Recording saved. The graph stays frozen until your next recording.':`Recording saved · ${reason.toLowerCase()}.`,reason!=='Completed');await refreshSessions();}
  render();renderSessions();
}
function download(session,data){
  const csv=sessionCSV(session,data),url=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'})),a=document.createElement('a');
  a.href=url;a.download=`TURTLE_${session.name.replace(/[^a-zA-Z0-9_-]/g,'_').slice(0,50)}_${session.id.slice(0,8)}.csv`;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
  if(current?.id===session.id)downloadedFallback=true;
}
async function exportSession(session){
  if(session.id===current?.id&&(recording||transition))return;
  try{const data=current?.id===session.id?rows:await store.rows(session.id);download(session,data);message('CSV download requested. Your saved recording remains available.');}catch(e){message('Could not download this recording. Please try again.',true);}
}
async function refreshSessions(){
  if(!ready)return;
  try{sessions=await store.list();renderSessions();render();}catch(e){message('Could not read saved recordings. Try refreshing the list.',true);}
}
function renderSessions(){
  const list=$('sessions-list');list.replaceChildren();let display=[...sessions];
  if(current){display=display.filter(s=>s.id!==current.id);display.unshift(current);}
  $('sessions-empty').hidden=display.length>0;
  for(const s of display){
    const active=s.id===current?.id&&(recording||transition),line=document.createElement('article');line.className='session-row';
    const info=document.createElement('div'),title=document.createElement('h3');title.textContent=s.name;info.append(title);
    const details=document.createElement('p');details.textContent=active?`${new Date(s.startedAt).toLocaleString()} · Recording in progress`:`${new Date(s.startedAt).toLocaleString()} · ${s.count.toLocaleString()} samples · ${duration(s.durationMs)}`;info.append(details);
    const channels=document.createElement('p');channels.className='small';channels.textContent=`TC 1: ${s.names[0]||'Unnamed'} · TC 2: ${s.names[1]||'Unnamed'}`;info.append(channels);
    const state=document.createElement('span');state.className='session-state';state.textContent=active?'Recording / saving':s.status==='save-error'?'Download now · save incomplete':s.status==='open'?'Unfinished session · may be active in another tab':s.status;info.append(state);
    const actions=document.createElement('div');actions.className='session-actions';const exportButton=document.createElement('button');exportButton.className='secondary';exportButton.textContent='Download CSV';exportButton.hidden=active;exportButton.disabled=active||transition||!s.count;exportButton.onclick=()=>exportSession(s);actions.append(exportButton);
    const del=document.createElement('button');del.className='text-button';del.textContent='Delete';del.disabled=active||s.status==='save-error';
    del.onclick=async()=>{
      if(!window.confirm(`Delete “${s.name}” from this browser? Download a CSV first if you need a backup.`))return;
      del.disabled=true;let release;
      try{release=await holdSession(s.id);if(!release){message('This recording is active in another tab. Stop it there before deleting.',true);del.disabled=false;return;}await store.remove(s.id);if(current?.id===s.id){current=null;rows=[];trace=[];draw();localFile=null;elapsed=0;}await refreshSessions();}catch(e){del.disabled=false;message('Could not delete the recording. Nothing else was removed.',true);}finally{release?.();}
    };actions.append(del);line.append(info,actions);list.append(line);
  }
}
async function init(){
  try{const saved=JSON.parse(localStorage.getItem('turtle-channel-names')||'null');if(Array.isArray(saved)&&saved.length===2)names=saved.map(n=>typeof n==='string'?n.slice(0,60):'');}catch{}
  names.forEach((n,i)=>{$('name-'+i).value=n;});draw();render();
  try{await store.open();ready=true;await refreshSessions();}catch(e){storageError=true;message('This browser cannot save recordings. Allow site storage and reload to record.',true);}render();
  if(!supportsSerial)message('Open this page in desktop Chrome or Edge to connect TURTLE by USB.',true);
}
function updateRecordingNames(){
  if(!recording||!current)return;
  current={...current,name:$('session-name').value.trim()||`Recording · ${new Date(current.startedAt).toLocaleString()}`,names:[...names]};
  const snapshot={...current,names:[...current.names]};
  enqueueSave(()=>store.finish(snapshot));
  renderSessions();
  message('Names updated. CSV headers will update on the next automatic save, or choose Save CSV now.');
}
$('session-name').addEventListener('input',()=>{updateRecordingNames();render();});
for(let i=0;i<2;i++)$('name-'+i).addEventListener('input',()=>{names[i]=$('name-'+i).value.trim().slice(0,60);try{localStorage.setItem('turtle-channel-names',JSON.stringify(names));}catch{message('Channel names could not be remembered by this browser.',true);}updateRecordingNames();render();});
$('connect').addEventListener('click',connect);$('disconnect').addEventListener('click',()=>{if(!recording||window.confirm('A recording is in progress. Stop recording, save the remaining samples, and disconnect?'))void disconnect();});
$('record').addEventListener('click',()=>{
  if(recording){if(window.confirm('Are you sure you want to stop recording? The remaining samples will be saved before stopping.'))void stopRecording();}
  else void startRecording();
});
$('reset-faults').addEventListener('click',()=>sendSettings(true).then(()=>message('Fault latches reset. Waiting for fresh measurements.')).catch(e=>{void stopRecording('Fault reset not confirmed');message(e.message,true);}));
$('apply').addEventListener('click',()=>sendSettings().then(()=>message('Settings sent.')).catch(e=>message('Could not send settings. '+e.message,true)));
$('save-now').addEventListener('click',()=>void saveLocalCSV());
$('refresh-sessions').addEventListener('click',refreshSessions);
window.addEventListener('beforeunload',e=>{if(updating||recording||saving||transition||localFile?.pending||(localFile?.error&&!downloadedFallback)||(storageError&&rows.length&&!downloadedFallback)){e.preventDefault();e.returnValue='';}});
setInterval(()=>{if(recording&&!writing&&!fresh())void stopRecording('No recent measurements');render();},500);
void init();

// Firmware is validated and loaded before the click that opens the USB picker.
async function loadFirmwareFile(file){
  if(!file||updating||firmwareLoading)return;
  const token=++firmwareLoadToken;
  firmwareLoading=true;firmwareImage=null;render();
  try{
    if(!/\.hex$/i.test(file.name)||file.size>200000)throw new Error('Choose the exported TURTLE application .hex file (without bootloader).');
    const image=parseIntelHex(await file.text());
    if(token!==firmwareLoadToken)return;
    firmwareImage=image;firmwareName=file.name;
    $('firmware-status').textContent=`Ready: ${file.name} · ${image.dataBytes.toLocaleString()} firmware bytes`;
    $('firmware-status').classList.toggle('save-error',false);
  }catch(e){$('firmware-status').textContent=e.message;$('firmware-status').classList.toggle('save-error',true);}
  finally{firmwareLoading=false;render();}
}
async function loadBundledFirmware(){
  const token=firmwareLoadToken;
  try{
    const base=new URL('./firmware/',firmwareModuleURL);
    const response=await fetch(new URL('release.json',base),{cache:'no-store'});
    if(!response.ok)throw new Error("Included firmware could not be loaded.");
    const release=await response.json();if(!release.file)throw new Error("No firmware is bundled with this release.");
    if(release.target!=='atmega328p-nano-16mhz'||!/^[a-zA-Z0-9_.-]+\.hex$/.test(release.file)||!/^[0-9a-f]{64}$/i.test(release.sha256))throw new Error('Bundled firmware metadata is invalid.');
    const result=await fetch(new URL(release.file,base),{cache:'no-store'});
    if(!result.ok)throw new Error('Bundled firmware could not be downloaded.');
    const bytes=await result.arrayBuffer();if(bytes.byteLength>200000)throw new Error('Bundled firmware is too large.');
    const digest=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes)),b=>b.toString(16).padStart(2,'0')).join('');
    if(digest!==release.sha256.toLowerCase())throw new Error('Bundled firmware integrity check failed.');
    const image=parseIntelHex(new TextDecoder().decode(bytes));
    if(token!==firmwareLoadToken)return;
    firmwareImage=image;firmwareName=String(release.version||release.file);
    $('firmware-status').textContent=`Ready: ${firmwareName} · ${image.dataBytes.toLocaleString()} firmware bytes`;
  }catch(e){if(token===firmwareLoadToken)$('firmware-status').textContent=e.message+' You can choose an exported TURTLE .hex file instead.';}
  finally{render();}
}
async function updateFirmware(){
  if(updating||firmwareLoading||!firmwareImage||!supportsSerial||connected||recording||busy||transition||writing)return;
  if(!window.confirm(`Install ${firmwareName} on a classic ATmega328P Nano? This replaces its current program. Keep USB connected until verification finishes.`))return;
  updating=true;render();$('firmware-progress').hidden=false;$('firmware-progress').value=0;
  $('firmware-status').classList.toggle('save-error',false);
  let wakeLock=null;
  try{
    // No vendor filter: official FTDI and clone CH340 USB bridges are both allowed.
    const target=await navigator.serial.requestPort();
    try{wakeLock=await navigator.wakeLock?.request('screen');}catch{}
    await flashNano(target,firmwareImage,progress=>{
      $('firmware-status').textContent=progress.text;$('firmware-progress').value=progress.percent;
    },{mode:$('bootloader-mode').value});
    latest=[null,null];lastPacket=0;settingsConfirmed=false;
    message('Firmware upload verified. Connect USB to resume temperature readings.');
  }catch(e){
    $('firmware-status').textContent=e.name==='NotFoundError'?'Update canceled. No device selected.':e.message;
    $('firmware-status').classList.toggle('save-error',e.name!=='NotFoundError');
  }finally{try{await wakeLock?.release();}catch{}updating=false;render();}
}
$('firmware-file').addEventListener('change',()=>void loadFirmwareFile($('firmware-file').files[0]));
$('update-firmware').addEventListener('click',()=>void updateFirmware());
void loadBundledFirmware();
