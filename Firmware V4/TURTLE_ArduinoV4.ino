// TURTLE v4: two MAX31856 devices, hardware SPI, CS pins unchanged.
// Only the Arduino SPI library is required. Serial: 9600 baud, 8N1.
#include <Arduino.h>
#include <SPI.h>
#include <string.h>

const uint8_t CS_PINS[2] = {10, 9};
// Normally off, open-circuit detection enabled, interrupt/latching faults,
// cold junction enabled, 60 Hz rejection. FAULTCLR is effective in this mode.
const uint8_t CR0_BASE = 0x14;
const uint8_t FAULT_CLEAR = 0x02, ONE_SHOT = 0x40;
const uint16_t DEVICE_ERROR = 256; // Software fault, separate from SR bits.
const uint32_t CONVERSION_TIMEOUT_MS = 500;
char tcType = 'T';
uint8_t tcCode = 7, rateSeconds = 1;
uint32_t cycleStart = 0, lastPoll = 0, sequence = 0;
bool converting = false, firstCycle = true, extendedProtocol = false;
bool channelDone[2] = {false, false};
float temperatures[2] = {0, 0};
uint16_t faults[2] = {DEVICE_ERROR, DEVICE_ERROR};
char command[64];
uint8_t commandLength = 0;
bool commandOverflow = false;
uint32_t lastCommandByte = 0;
bool configPending = false;
char pendingType = 'T';
uint8_t pendingRate = 1;
uint16_t pendingToken = 0;
bool pendingAck = false;

uint8_t readRegister(uint8_t cs, uint8_t address) {
  SPI.beginTransaction(SPISettings(1000000, MSBFIRST, SPI_MODE1));
  digitalWrite(cs, LOW);
  SPI.transfer(address & 0x7F);
  uint8_t value = SPI.transfer(0);
  digitalWrite(cs, HIGH);
  SPI.endTransaction();
  return value;
}
void writeRegister(uint8_t cs, uint8_t address, uint8_t value) {
  SPI.beginTransaction(SPISettings(1000000, MSBFIRST, SPI_MODE1));
  digitalWrite(cs, LOW);
  SPI.transfer(address | 0x80);
  SPI.transfer(value);
  digitalWrite(cs, HIGH);
  SPI.endTransaction();
}
void clearFaults(uint8_t cs) {
  // Preserve the chosen mode/filter; never disable open-circuit detection.
  writeRegister(cs, 0, CR0_BASE | FAULT_CLEAR);
}
void configureChannel(uint8_t cs) {
  writeRegister(cs, 0, CR0_BASE); // normally off before changing type
  writeRegister(cs, 1, tcCode);  // one sample, chosen thermocouple type
  writeRegister(cs, 2, 0);       // unmask all fault sources
  clearFaults(cs);
}
int8_t typeCode(char type) {
  switch(type) {
    case 'B': return 0; case 'E': return 1; case 'J': return 2;
    case 'K': return 3; case 'N': return 4; case 'R': return 5;
    case 'S': return 6; case 'T': return 7; default: return -1;
  }
}
void acknowledge(uint16_t token) {
  Serial.print(F("ACK:")); Serial.print(token); Serial.print(',');
  Serial.print(tcType); Serial.print(','); Serial.println(rateSeconds);
}
void applyPending() {
  if (!configPending || converting) return;
  tcType = pendingType; tcCode = (uint8_t)typeCode(tcType);
  rateSeconds = pendingRate;
  for (uint8_t i=0; i<2; ++i) configureChannel(CS_PINS[i]);
  configPending = false; firstCycle = true;
  if (pendingAck) acknowledge(pendingToken);
}
void handleCommand(const char *text) {
  const size_t length=strlen(text);
  bool validConfig=length>=12 && length<=16 && !strncmp(text,"CONFIG:",7) &&
    typeCode(text[7])>=0 && text[8]==',' && text[9]>='0' && text[9]<='5' && text[10]==',';
  uint32_t token=0;
  if(validConfig)for(size_t i=11;i<length;++i){
    if(text[i]<'0'||text[i]>'9'){validConfig=false;break;}
    token=token*10+(uint8_t)(text[i]-'0');
  }
  validConfig=validConfig && token>0 && token<=65535UL;
  if (validConfig) {
    if (configPending) { Serial.println(F("ERR:BUSY")); return; }
    pendingType=text[7]; pendingRate=(uint8_t)(text[9]-'0'); pendingToken=(uint16_t)token;
    pendingAck=true; configPending=true; extendedProtocol=true;
  } else if (!strcmp(text,"RESET")) {
    if (configPending) { Serial.println(F("ERR:BUSY")); return; }
    pendingType=tcType; pendingRate=rateSeconds; pendingToken=0;
    pendingAck=true; configPending=true; // safely clear between conversions
  } else if (!strncmp(text,"TYPE:",5) && strlen(text)==6 && typeCode(text[5])>=0) {
    pendingType=text[5]; if(!configPending)pendingRate=rateSeconds;
    pendingAck=false; configPending=true;
  } else if (!strncmp(text,"RATE:",5) && strlen(text)==6 && text[5]>='0' && text[5]<='5') {
    pendingRate=(uint8_t)(text[5]-'0'); if(!configPending)pendingType=tcType;
    pendingAck=false; configPending=true;
  } else Serial.println(F("ERR:COMMAND"));
}
void serviceSerial() {
  // A missing delimiter must not stall acquisition or grow the heap.
  if ((commandLength || commandOverflow) && (uint32_t)(millis()-lastCommandByte)>1000) {
    commandLength=0; commandOverflow=false;
  }
  uint8_t budget=64;
  while (budget-- && Serial.available()>0) {
    char c=(char)Serial.read(); lastCommandByte=millis();
    if(c==';') {
      if(commandOverflow)Serial.println(F("ERR:COMMAND_TOO_LONG"));
      else if(commandLength){command[commandLength]='\0';handleCommand(command);}
      commandLength=0;commandOverflow=false;
    } else if(c!='\r' && c!='\n') {
      if(commandLength<sizeof(command)-1 && !commandOverflow)command[commandLength++]=c;
      else commandOverflow=true;
    }
  }
}
void startCycle(uint32_t now) {
  cycleStart=now; lastPoll=now; converting=true; firstCycle=false;
  for(uint8_t i=0;i<2;++i) {
    channelDone[i]=false; faults[i]=0;
    // Clear latched faults BEFORE a new conversion, then assess its NEW status.
    // Retrying configuration also recovers a converter after its power returns.
    configureChannel(CS_PINS[i]);
    if(readRegister(CS_PINS[i],1)!=tcCode ||
       (readRegister(CS_PINS[i],0)&0xBD)!=CR0_BASE) {
      faults[i]=DEVICE_ERROR;channelDone[i]=true;continue;
    }
    writeRegister(CS_PINS[i],0,CR0_BASE|ONE_SHOT);
  }
}
void readChannel(uint8_t i) {
  uint8_t cs=CS_PINS[i];
  // Read temperature and fault status together after this conversion completes.
  SPI.beginTransaction(SPISettings(1000000,MSBFIRST,SPI_MODE1));
  digitalWrite(cs,LOW);SPI.transfer(0x0C);
  uint32_t raw=(uint32_t)SPI.transfer(0)<<16;
  raw|=(uint32_t)SPI.transfer(0)<<8;raw|=SPI.transfer(0);
  faults[i]=SPI.transfer(0);
  digitalWrite(cs,HIGH);SPI.endTransaction();
  int32_t signedRaw=(raw&0x800000UL)?(int32_t)raw-16777216L:(int32_t)raw;
  temperatures[i]=signedRaw/4096.0f;
  channelDone[i]=true;
  // Preserve faults[i] for this packet, but reset the hardware latch now.
  clearFaults(cs);
}
void sendPacket(uint32_t now) {
  if(extendedProtocol) {
    Serial.print(F("DATA:"));Serial.print(sequence++);Serial.print(',');Serial.print(now);
    Serial.print(',');Serial.print(tcType);Serial.print(',');Serial.print(rateSeconds);
    for(uint8_t i=0;i<2;++i){Serial.print(',');if(faults[i])Serial.print(F("NA"));else Serial.print(temperatures[i],3);Serial.print(',');Serial.print(faults[i]);}
    Serial.println();
  } else {
    Serial.print(F("STATUS:T1:"));
    if(faults[0])Serial.print(F("Not Connected"));else Serial.print(temperatures[0],2);
    Serial.print(F(",T2:"));
    if(faults[1])Serial.print(F("Not Connected"));else Serial.print(temperatures[1],2);
    Serial.println(',');
  }
}
void setup() {
  Serial.begin(9600);
  // Deselect BOTH devices before the first SPI transaction.
  for(uint8_t i=0;i<2;++i){digitalWrite(CS_PINS[i],HIGH);pinMode(CS_PINS[i],OUTPUT);}
  SPI.begin();
  for(uint8_t i=0;i<2;++i)configureChannel(CS_PINS[i]);
}
void loop() {
  serviceSerial();applyPending();
  uint32_t now=millis();
  uint32_t period=rateSeconds?(uint32_t)rateSeconds*1000UL:250UL;
  if(!converting && (firstCycle || (uint32_t)(now-cycleStart)>=period))startCycle(now);
  if(converting && (uint32_t)(now-lastPoll)>=5) {
    lastPoll=now;
    for(uint8_t i=0;i<2;++i)if(!channelDone[i]) {
      // Guard against returning the previous conversion on the triggering edge.
      if((uint32_t)(now-cycleStart)>=20 && !(readRegister(CS_PINS[i],0)&ONE_SHOT))readChannel(i);
      else if((uint32_t)(now-cycleStart)>=CONVERSION_TIMEOUT_MS) {
        faults[i]=DEVICE_ERROR;channelDone[i]=true;clearFaults(CS_PINS[i]);
      }
    }
    if(channelDone[0]&&channelDone[1]){converting=false;sendPacket(now);}
  }
}
