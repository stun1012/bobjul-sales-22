// Workbook parsing stays in the browser; no upload endpoint is used.
const importFields=[['company','회사명'],['name','이름'],['product','취급제품'],['vendor','벤더사명'],['overview','제품 개요'],['references','레퍼런스']];
function inspectImport(matrix,existing){
 const normalize=v=>String(v??'').trim(),header=(matrix[0]||[]).map(v=>normalize(v).replace(/\s/g,''));
 const indexes=importFields.map(([,label])=>header.indexOf(label.replace(/\s/g,'')));
 if(indexes.some(i=>i<0))throw Error('첫 행에 회사명, 이름, 취급제품, 벤더사명, 제품 개요, 레퍼런스 열이 필요합니다. 양식을 사용해 주세요.');
 if(new Set(indexes).size!==6||importFields.some(([,label])=>header.filter(h=>h===label.replace(/\s/g,'')).length!==1))throw Error('열 제목이 중복되어 있습니다.');
 const signature=r=>JSON.stringify(importFields.map(([k])=>normalize(r[k])));
 const seen=new Set(existing.map(signature)),rows=[],errors=[];let duplicates=0;
 if(matrix.length>1001)throw Error('한 번에 최대 1,000행까지 등록할 수 있습니다.');
 matrix.slice(1).forEach((cells,i)=>{
 if(!cells.some(v=>normalize(v)))return;
 const row=Object.fromEntries(importFields.map(([key],j)=>[key,normalize(cells[indexes[j]])]));
 const missing=importFields.filter(([key])=>!row[key]).map(([,label])=>label);
 if(missing.length){errors.push((i+2)+'행: '+missing.join(', ')+' 입력 필요 (해당 사항이 없으면 - 입력)');return;}
 if(importFields.some(([key])=>row[key].length>10000)){errors.push((i+2)+'행: 셀당 10,000자 초과');return;}
 const sig=signature(row);if(seen.has(sig)){duplicates++;return;}seen.add(sig);rows.push(row);
 });return {rows,errors,duplicates};
}
(function(){
 let pending=[],loading;
 const dialog=document.createElement('dialog');dialog.id='import-dialog';
 dialog.innerHTML='<div class="dialog-head"><h3>엑셀 대량 등록</h3><button type="button" class="icon-btn" id="import-close" aria-label="닫기">×</button></div><p>첫 번째 시트 · 최대 1,000행 / 5MB · .xlsx, .xls</p><p>기존 자료는 유지하고 새 행을 추가합니다. 6개 항목이 모두 같은 중복 행은 제외합니다. 파일은 서버로 전송되지 않습니다.</p><button type="button" class="secondary" id="import-template">양식 다운로드</button><label class="import-file">엑셀 파일 선택<input type="file" id="import-file" accept=".xlsx,.xls"></label><p id="import-status" role="status">양식을 작성하거나 엑셀 다운로드 파일을 선택하세요.</p><div id="import-preview"></div><div class="dialog-actions"><button type="button" class="secondary" id="import-cancel">취소</button><button type="button" class="primary" id="import-confirm" disabled>등록 확정</button></div>';
 document.body.append(dialog);
 const file=document.getElementById('import-file'),status=document.getElementById('import-status'),preview=document.getElementById('import-preview'),confirm=document.getElementById('import-confirm');
 const reset=()=>{pending=[];confirm.disabled=true;preview.replaceChildren();};
 document.getElementById('import-btn').onclick=()=>{reset();file.value='';status.textContent='양식을 작성하거나 엑셀 다운로드 파일을 선택하세요.';dialog.showModal();};
 document.getElementById('import-close').onclick=document.getElementById('import-cancel').onclick=()=>dialog.close();
 document.getElementById('import-template').onclick=()=>{const url=URL.createObjectURL(buildSalesWorkbook([])),a=document.createElement('a');a.href=url;a.download='22기영업_업로드양식.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);};
 function loadReader(){if(window.XLSX)return Promise.resolve();if(loading)return loading;loading=new Promise((resolve,reject)=>{const s=document.createElement('script');s.src='https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';s.onload=()=>window.XLSX?resolve():reject(Error('엑셀 라이브러리 로드 실패'));s.onerror=()=>{loading=null;s.remove();reject(Error('엑셀 라이브러리를 불러올 수 없습니다. 인터넷 연결을 확인하고 다시 시도하세요.'));};document.head.append(s);});return loading;}
 let sequence=0;
 file.onchange=async()=>{reset();const token=++sequence,f=file.files[0];if(!f)return;
 try{if(!/\.(xlsx|xls)$/i.test(f.name)||f.size>5*1024*1024)throw Error('.xlsx 또는 .xls 파일(5MB 이하)을 선택하세요.');
 status.textContent='파일 확인 중…';await loadReader();
 const book=XLSX.read(await f.arrayBuffer(),{type:'array',sheetRows:1002,cellHTML:false,cellFormula:false});
 if(token!==sequence||!dialog.open)return;
 const sheet=book.Sheets[book.SheetNames[0]];
 const result=inspectImport(XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false,blankrows:true}),data.directory||[]);
 status.textContent='추가 '+result.rows.length+'건 · 중복 제외 '+result.duplicates+'건 · 오류 '+result.errors.length+'건';
 if(result.errors.length){preview.textContent=result.errors.slice(0,30).join('\n')+(result.errors.length>30?'\n외 '+(result.errors.length-30)+'건':'');return;}
 pending=result.rows;confirm.disabled=!pending.length;
 preview.innerHTML='<p>미리보기 (최대 10건) · 오류가 없을 때만 등록됩니다.</p><div class="table-wrap"><table><thead><tr>'+importFields.map(([,label])=>'<th>'+label+'</th>').join('')+'</tr></thead><tbody>'+pending.slice(0,10).map(r=>'<tr>'+importFields.map(([key])=>'<td>'+escapeHtml(r[key])+'</td>').join('')+'</tr>').join('')+'</tbody></table></div>';
 }catch(error){if(token===sequence)status.textContent=error.message||'파일을 읽지 못했습니다. 양식을 확인해 주세요.';}};
 confirm.onclick=()=>{if(!pending.length)return;confirm.disabled=true;
 // Revalidate against records added since preview, and save atomically.
 const result=inspectImport([importFields.map(([,label])=>label),...pending.map(r=>importFields.map(([key])=>r[key]))],data.directory||[]);
 const next={...data,directory:[...(data.directory||[]),...result.rows.map(r=>({...r,id:crypto.randomUUID()}))]};
 try{localStorage.setItem(KEY,JSON.stringify(next));data=next;pending=[];dialog.close();document.getElementById('global-search').value='';renderDirectory();go('directory');showToast(result.rows.length+'건 등록되었습니다');}
 catch{status.textContent='저장 공간이 부족하거나 저장이 차단되었습니다. 기존 데이터는 변경하지 않았습니다.';confirm.disabled=false;}
 };
})();

