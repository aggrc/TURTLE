export function parseStatus(line) {
  const start = line.indexOf('STATUS:');
  if (start < 0) return null;
  const channels = [null, null];
  const seen = new Set();
  for (const field of line.slice(start + 7).trim().split(',')) {
    if (!field.trim()) continue;
    const match = field.trim().match(/^T([12]):\s*(.+)$/);
    if (!match || seen.has(match[1])) return null;
    seen.add(match[1]);
    const value = match[2].trim();
    if (value === 'Not Connected') continue;
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(value)) return null;
    const n = Number(value);
    if (!Number.isFinite(n)) return null;
    channels[Number(match[1]) - 1] = n;
  }
  return seen.size === 2 ? channels : null;
}

export function createLineDecoder(onLine) {
  const decoder = new TextDecoder();
  let pending = '';
  return {
    push(bytes) {
      pending += decoder.decode(bytes, {stream: true});
      let index;
      while ((index = pending.indexOf('\n')) >= 0) {
        const line = pending.slice(0, index).replace(/\r$/, '');
        pending = pending.slice(index + 1);
        onLine(line);
      }
      if (pending.length > 8192) { pending = ''; throw new Error('Serial line exceeded 8 KB. Check device and baud rate.'); }
    }
  };
}

export function settingsCommands(type, rate) {
  if (!['K','J','T','E','N','S','R','B'].includes(type) || ![0,1,2,3,4,5].includes(Number(rate))) throw new Error('Invalid device settings.');
  return [`TYPE:${type};`, `RATE:${Number(rate)};`];
}

export function toCSV(rows, source, type, rate) {
  return 'source,thermocouple_type_requested,sample_interval_s_requested,timestamp_utc,elapsed_s,tc1_c,tc2_c\r\n' + rows.map(r => [source,type,rate,r.utc,r.elapsed.toFixed(3),r.values[0] ?? '',r.values[1] ?? ''].join(',')).join('\r\n');
}

// v4 data includes device time, sequence, applied settings and per-channel faults.
export function parsePacket(line) {
  if(!line.startsWith('DATA:')) {
    const values=parseStatus(line);
    return values?{values,faults:[null,null],legacy:true}:null;
  }
  const p=line.slice(5).split(',');
  if(p.length!==8 || !/^\d+$/.test(p[0]) || !/^\d+$/.test(p[1]) ||
     !/^[BEJKNRST]$/.test(p[2]) || !/^[0-5]$/.test(p[3]))return null;
  const seq=Number(p[0]),ms=Number(p[1]);
  if(seq>0xffffffff||ms>0xffffffff)return null;
  const values=[],faults=[];
  for(const i of [4,6]) {
    if(!/^\d+$/.test(p[i+1]))return null;
    const f=Number(p[i+1]);if(f>511)return null;
    if(p[i]!=='NA'&&!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(p[i]))return null;
    const value=p[i]==='NA'?null:Number(p[i]);
    if((value!==null&&!Number.isFinite(value)) || (value===null&&f===0))return null;
    values.push(f?null:value);faults.push(f);
  }
  return {values,faults,seq,ms,type:p[2],rate:Number(p[3]),legacy:false};
}
export function configCommand(type,rate,token) {
  settingsCommands(type,rate);
  if(!Number.isInteger(token)||token<1||token>65535)throw new Error('Invalid command token.');
  return `CONFIG:${type},${rate},${token};`;
}
export function faultLabel(mask) {
  if(mask===null)return 'No fault details (older firmware)';
  const labels=[[256,'Converter communication / timeout'],[128,'Cold junction out of range'],[64,'Thermocouple out of range'],[32,'Cold junction high'],[16,'Cold junction low'],[8,'Temperature high'],[4,'Temperature low'],[2,'Input voltage fault'],[1,'Thermocouple disconnected']];
  return labels.filter(([bit])=>mask&bit).map(([,label])=>label).join(' · ')||'Temperature · degrees Celsius';
}
