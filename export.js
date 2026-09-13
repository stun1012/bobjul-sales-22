// Minimal XLSX writer: inline text cells, stored ZIP entries, no remote dependencies.
function buildSalesWorkbook(rows) {
  const xml=v=>String(v??'').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const headers=['번호','회사명','이름','취급제품','벤더사명','제품 개요','레퍼런스'];
  const values=[headers,...rows.map(r=>[r.rowIndex+1,r.company,r.name,r.product,r.vendor,r.overview,r.references])];
  const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const sheet='<worksheet xmlns="'+ns+'"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>'+[8,22,16,26,22,52,45].map((w,i)=>'<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+w+'" customWidth="1"/>').join('')+'</cols><sheetData>'+values.map((row,i)=>'<row r="'+(i+1)+'">'+row.map((v,j)=>{const ref=String.fromCharCode(65+j)+(i+1);return typeof v==='number'?'<c r="'+ref+'" s="0"><v>'+v+'</v></c>':'<c r="'+ref+'" s="'+(i===0?1:0)+'" t="inlineStr"><is><t xml:space="preserve">'+xml(v)+'</t></is></c>'}).join('')+'</row>').join('')+'</sheetData><autoFilter ref="A1:G'+values.length+'"/></worksheet>';
  const files={
    '[Content_Types].xml':'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    '_rels/.rels':'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml':'<workbook xmlns="'+ns+'" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="22기 영업" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels':'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    'xl/styles.xml':'<styleSheet xmlns="'+ns+'"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF163B40"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
    'xl/worksheets/sheet1.xml':sheet
  };
  const enc=new TextEncoder(),parts=[],central=[];let offset=0;
  const crc=bytes=>{let c=0xffffffff;for(const b of bytes){c^=b;for(let j=0;j<8;j++)c=(c>>>1)^((c&1)?0xedb88320:0)}return (c^0xffffffff)>>>0};
  const pack=(size,entries)=>{const a=new Uint8Array(size),v=new DataView(a.buffer);for(const [p,n,w] of entries)w===2?v.setUint16(p,n,true):v.setUint32(p,n,true);return a};
  for(const [name,body] of Object.entries(files)){
    const nameBytes=enc.encode(name),bytes=enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+body),sum=crc(bytes);
    const local=pack(30,[[0,0x04034b50,4],[4,20,2],[12,33,2],[14,sum,4],[18,bytes.length,4],[22,bytes.length,4],[26,nameBytes.length,2]]);
    const dir=pack(46,[[0,0x02014b50,4],[4,20,2],[6,20,2],[14,33,2],[16,sum,4],[20,bytes.length,4],[24,bytes.length,4],[28,nameBytes.length,2],[42,offset,4]]);
    parts.push(local,nameBytes,bytes);central.push(dir,nameBytes);offset+=local.length+nameBytes.length+bytes.length;
  }
  const centralSize=central.reduce((n,b)=>n+b.length,0),count=Object.keys(files).length;
  return new Blob([...parts,...central,pack(22,[[0,0x06054b50,4],[8,count,2],[10,count,2],[12,centralSize,4],[16,offset,4]])],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function downloadSalesExcel(){
  const q=document.querySelector('#global-search').value.trim().toLocaleLowerCase();
  const rows=(data.directory||[]).map((r,rowIndex)=>({...r,rowIndex})).filter(r=>['company','name','product','vendor','overview','references'].some(k=>String(r[k]||'').toLocaleLowerCase().includes(q)));
  if(!rows.length){showToast('내보낼 등록 정보가 없습니다');return}
  const url=URL.createObjectURL(buildSalesWorkbook(rows)),a=document.createElement('a');
  a.href=url;a.download='22기영업_'+new Date().toLocaleDateString('sv-SE')+'.xlsx';document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),10000);
  showToast(rows.length+'건을 엑셀로 다운로드했습니다');
}
document.querySelector('#export-btn').addEventListener('click',downloadSalesExcel);
