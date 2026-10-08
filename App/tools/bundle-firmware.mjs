// Node.js 20+. Run against extracted app assets before uploading the package.
import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL,fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
const [hexPath,version='TURTLE v4']=process.argv.slice(2);
if(!hexPath)throw new Error('Usage: node tools/bundle-firmware.mjs /path/to/TURTLE.ino.hex "TURTLE v4"');
const root=fileURLToPath(new URL('../',import.meta.url));
const {parseIntelHex}=await import(pathToFileURL(path.join(root,'firmware-updater.js')));
const bytes=await fs.readFile(hexPath);const image=parseIntelHex(bytes.toString('utf8'));
await fs.mkdir(path.join(root,'firmware'),{recursive:true});
await fs.writeFile(path.join(root,'firmware','turtle.hex'),bytes);
await fs.writeFile(path.join(root,'firmware','release.json'),JSON.stringify({version,file:'turtle.hex',sha256:createHash('sha256').update(bytes).digest('hex'),target:'atmega328p-nano-16mhz'},null,2)+'\n');
console.log(`Bundled ${version}: ${image.dataBytes} bytes. Repackage and upload the complete app/plugin.`);
