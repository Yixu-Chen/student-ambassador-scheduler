(function(root){
  const enc=new TextEncoder(),dec=new TextDecoder('utf-8');
  function u16(v,o){return v.getUint16(o,true)}function u32(v,o){return v.getUint32(o,true)}
  async function unzip(buffer){
    const view=new DataView(buffer);let eocd=-1;for(let i=buffer.byteLength-22;i>=Math.max(0,buffer.byteLength-65557);i--){if(u32(view,i)===0x06054b50){eocd=i;break}}
    if(eocd<0)throw new Error('文件不是有效的 Excel 工作簿');
    const total=u16(view,eocd+10),central=u32(view,eocd+16),files={};let p=central;
    for(let n=0;n<total;n++){
      if(u32(view,p)!==0x02014b50)throw new Error('无法读取 Excel 文件目录');
      const method=u16(view,p+10),compressed=u32(view,p+20),nameLen=u16(view,p+28),extraLen=u16(view,p+30),commentLen=u16(view,p+32),local=u32(view,p+42);
      const name=dec.decode(new Uint8Array(buffer,p+46,nameLen));const localName=u16(view,local+26),localExtra=u16(view,local+28),start=local+30+localName+localExtra;const bytes=new Uint8Array(buffer,start,compressed);
      if(method===0)files[name]=bytes;
      else if(method===8){if(!root.DecompressionStream)throw new Error('当前浏览器不支持离线读取 Excel，请使用最新版 Microsoft Edge 或 Google Chrome');const stream=new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));files[name]=new Uint8Array(await new Response(stream).arrayBuffer())}
      p+=46+nameLen+extraLen+commentLen;
    }
    return files;
  }
  function xml(bytes){return new DOMParser().parseFromString(dec.decode(bytes),'application/xml')}
  function textOf(node){return Array.from(node.getElementsByTagName('t')).map(x=>x.textContent||'').join('')}
  function colIndex(ref){let n=0;for(const ch of ref.match(/^[A-Z]+/i)?.[0]||'A')n=n*26+ch.toUpperCase().charCodeAt(0)-64;return n-1}
  async function readFirstSheet(file){
    const files=await unzip(await file.arrayBuffer());if(!files['xl/workbook.xml'])throw new Error('缺少 Excel 工作簿信息');
    const wb=xml(files['xl/workbook.xml']),first=wb.getElementsByTagName('sheet')[0];if(!first)throw new Error('名单中没有工作表');const rid=first.getAttribute('r:id');
    const rel=xml(files['xl/_rels/workbook.xml.rels']);let target='worksheets/sheet1.xml';for(const r of rel.getElementsByTagName('Relationship'))if(r.getAttribute('Id')===rid)target=r.getAttribute('Target');
    target=target.replace(/^\//,'').replace(/^xl\//,'');const sheetFile='xl/'+target.replace(/^\.\//,'');if(!files[sheetFile])throw new Error('无法找到第一个工作表');
    const shared=[];if(files['xl/sharedStrings.xml'])for(const si of xml(files['xl/sharedStrings.xml']).getElementsByTagName('si'))shared.push(textOf(si));
    const doc=xml(files[sheetFile]),rows=[];for(const row of doc.getElementsByTagName('row')){const values=[];for(const c of row.getElementsByTagName('c')){const idx=colIndex(c.getAttribute('r')||'A1'),type=c.getAttribute('t'),v=c.getElementsByTagName('v')[0];let value='';if(type==='inlineStr')value=textOf(c);else if(v)value=type==='s'?shared[Number(v.textContent)]:v.textContent;values[idx]=value}rows.push(values)}return rows;
  }
  function esc(v){return String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
  function colName(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
  function sheetXml(rows){let body='';rows.forEach((row,ri)=>{let cells='';row.forEach((v,ci)=>{if(v===undefined||v===null||v==='')return;const r=colName(ci+1)+(ri+1);if(typeof v==='number')cells+=`<c r="${r}"><v>${v}</v></c>`;else cells+=`<c r="${r}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`});body+=`<row r="${ri+1}">${cells}</row>`});return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData>${body}</sheetData></worksheet>`}
  const crcTable=(()=>{const t=[];for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=(c&1)?0xedb88320^(c>>>1):c>>>1;t[n]=c>>>0}return t})();
  function crc32(bytes){let c=0xffffffff;for(const b of bytes)c=crcTable[(c^b)&255]^(c>>>8);return(c^0xffffffff)>>>0}
  function write16(a,v){a.push(v&255,(v>>>8)&255)}function write32(a,v){a.push(v&255,(v>>>8)&255,(v>>>16)&255,(v>>>24)&255)}
  function zip(files){const local=[],central=[];let offset=0,count=0;for(const [name,content] of Object.entries(files)){const nb=enc.encode(name),data=enc.encode(content),crc=crc32(data),head=[];write32(head,0x04034b50);write16(head,20);write16(head,0x800);write16(head,0);write16(head,0);write16(head,0);write32(head,crc);write32(head,data.length);write32(head,data.length);write16(head,nb.length);write16(head,0);local.push(new Uint8Array(head),nb,data);const ch=[];write32(ch,0x02014b50);write16(ch,20);write16(ch,20);write16(ch,0x800);write16(ch,0);write16(ch,0);write16(ch,0);write32(ch,crc);write32(ch,data.length);write32(ch,data.length);write16(ch,nb.length);write16(ch,0);write16(ch,0);write16(ch,0);write16(ch,0);write32(ch,0);write32(ch,offset);central.push(new Uint8Array(ch),nb);offset+=head.length+nb.length+data.length;count++}const centralSize=central.reduce((n,x)=>n+x.length,0),end=[];write32(end,0x06054b50);write16(end,0);write16(end,0);write16(end,count);write16(end,count);write32(end,centralSize);write32(end,offset);write16(end,0);return new Blob([...local,...central,new Uint8Array(end)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
  function makeWorkbook(sheets){
    const content=sheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('');
    const names=sheets.map((s,i)=>`<sheet name="${esc(s.name.slice(0,31))}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('');
    const rels=sheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('');
    const files={
      '[Content_Types].xml':`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${content}</Types>`,
      '_rels/.rels':'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
      'xl/workbook.xml':`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names}</sheets></workbook>`,
      'xl/_rels/workbook.xml.rels':`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`
    };sheets.forEach((s,i)=>files[`xl/worksheets/sheet${i+1}.xml`]=sheetXml(s.rows));return zip(files)
  }
  function downloadWorkbook(sheets,filename){const a=document.createElement('a');a.href=URL.createObjectURL(makeWorkbook(sheets));a.download=filename;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
  root.XlsxLite={readFirstSheet,makeWorkbook,downloadWorkbook};
})(typeof self!=='undefined'?self:this);
