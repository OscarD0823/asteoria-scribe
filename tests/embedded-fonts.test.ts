import test from 'node:test';import assert from 'node:assert/strict';import fs from 'node:fs';
import {DOMParser} from '@xmldom/xmldom';import {unzipSync,zipSync,strToU8,strFromU8} from 'fflate';
import {importLocalFont} from '@prisma/core/local-fonts';import {setFontLoader} from '@prisma/core/fonts';
import {exportDocx} from '../src/document-export';import {importOfficeDocx} from '../src/document-import';
import {defaultLayout,plainText} from '../src/flow';
Object.assign(globalThis,{DOMParser});
setFontLoader(async file=>new Uint8Array(fs.readFileSync('public/fonts/'+file)));
async function fixture(){
 const fonts=await Promise.all(['Regular','Bold'].map(style=>importLocalFont(new Uint8Array(fs.readFileSync('public/fonts/LiberationSerif-'+style+'.ttf')))));
 const content={type:'doc',content:[{type:'paragraph',attrs:{fontFamily:fonts[0].value},content:[{type:'text',text:'Fuente original á ñ €'},{type:'text',text:' y negrita',marks:[{type:'bold'}]}]}]};
 return {fonts,bytes:await exportDocx({content,layout:defaultLayout(),fonts},'Fuentes',fonts)};
}
test('Word embedded font bytes, family, styles and original archive survive import and text edits',async()=>{
 const {fonts,bytes}=await fixture(),result=await importOfficeDocx(bytes);
 for(const font of fonts)assert(result.flow.fonts?.some(face=>face.hash===font.hash&&face.style===font.style));
 assert.equal(result.flow.importReport?.fontSubstitutions.length,0);
 assert(JSON.stringify(result.flow.content).includes(fonts[0].value));assert.match(plainText(result.flow.content),/á ñ €/);
 assert.deepEqual(await exportDocx(result.flow,'Reabierto',result.flow.fonts),bytes);
 const text=result.flow.content.content![0].content![0];text.text='Cambio con fuente original';
 const saved=unzipSync(await exportDocx(result.flow,'Editado',result.flow.fonts)),original=unzipSync(bytes);
 for(const name of Object.keys(original).filter(name=>name.startsWith('word/fonts/')))assert.deepEqual(saved[name],original[name]);
});
test('an external embedded-font relationship cannot fetch data and leaves a visible substitution warning',async()=>{
 const {bytes}=await fixture(),files=unzipSync(bytes),rel=files['word/_rels/fontTable.xml.rels'];
 files['word/_rels/fontTable.xml.rels']=strToU8(strFromU8(rel).replace(/Target="[^"]*"/g,'Target="https://invalid.example/font.ttf" TargetMode="External"'));
 const result=await importOfficeDocx(zipSync(files));assert.equal(result.flow.fonts?.length??0,0);
 assert(result.warnings.some(warning=>warning.includes('relación de fuente local no válida')));
 assert.match(plainText(result.flow.content),/Fuente original/);
});
test('theme font references resolve to the embedded original family instead of a default substitute',async()=>{
 const {bytes,fonts}=await fixture(),files=unzipSync(bytes);
 files['word/document.xml']=strToU8(strFromU8(files['word/document.xml']).replace(/<w:rFonts\b[^>]*\/>/g,'<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>'));
 files['word/theme/theme1.xml']=strToU8('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Fuentes"><a:majorFont><a:latin typeface="Liberation Serif"/></a:majorFont><a:minorFont><a:latin typeface="Liberation Serif"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>');
 const result=await importOfficeDocx(zipSync(files));
 assert(JSON.stringify(result.flow.content).includes(fonts[0].value));assert.equal(result.flow.importReport?.fontSubstitutions.length,0);
});
