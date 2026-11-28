(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const state = { workbook:null, rows:[], filtered:[], sheetName:'', page:0, pageSize:25, colStart:0, colSize:20, fileName:'' };
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
  const cell = value => value instanceof Date ? value.toISOString().slice(0,10) : (value == null ? '' : String(value));
  const colName = n => { let s=''; for(n++;n;n=Math.floor((n-1)/26)) s=String.fromCharCode(65+(n-1)%26)+s; return s; };
  function status(message, type='') { $('load-status').textContent=message; $('load-status').className=`status ${type}`; }
  function showWorkbook(buffer, fileName) {
    if (!window.XLSX) throw new Error('엑셀 읽기 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.');
    const wb = XLSX.read(buffer,{type:'array',cellDates:true});
    if (!wb.SheetNames.length) throw new Error('엑셀 파일에 시트가 없습니다.');
    state.workbook=wb; state.fileName=fileName;
    $('sheet-select').innerHTML=wb.SheetNames.map((name,i)=>`<option value="${esc(name)}">${i+1}. ${esc(name)}</option>`).join('');
    $('explorer').hidden=false;
    $('workbook-summary').innerHTML = [
      ['파일명',fileName],['시트 수',`${wb.SheetNames.length}개`],['분류','선택한 원본 Excel']
    ].map(([k,v])=>`<div class="summary-card"><span>${k}</span><strong>${esc(v)}</strong></div>`).join('');
    $('sheet-search').value=''; state.page=0; state.colStart=0;
    renderSheet();
  }
  async function loadSelected() {
    const path=$('workbook-select').value;
    $('load-workbook').disabled=true; status('선택한 원본 엑셀 파일을 불러오는 중입니다…');
    try {
      const response=await fetch(path);
      if(!response.ok) throw new Error(`파일을 찾을 수 없습니다 (${response.status}). 저장소 data 폴더에 엑셀 파일이 있는지 확인하세요.`);
      const buffer=await response.arrayBuffer();
      showWorkbook(buffer,path.split('/').pop());
      status('엑셀 파일을 불러왔습니다. 시트를 선택해 내용을 확인하세요.','success');
    } catch(error) { status(error.message,'error'); }
    finally { $('load-workbook').disabled=false; }
  }
  function loadLocal(file) {
    if(!file) return;
    if(!/\.xlsx?$/i.test(file.name)){status('Excel 파일(.xlsx 또는 .xls)만 선택할 수 있습니다.','error');return;}
    status('선택한 파일을 읽는 중입니다…');
    const reader=new FileReader();
    reader.onload=()=>{try{showWorkbook(reader.result,file.name);status('내 컴퓨터의 엑셀 파일을 불러왔습니다.','success');}catch(e){status(e.message,'error');}};
    reader.onerror=()=>status('파일을 읽지 못했습니다.','error');
    reader.readAsArrayBuffer(file);
  }
  function renderSheet() {
    if(!state.workbook) return;
    state.sheetName=$('sheet-select').value;
    const ws=state.workbook.Sheets[state.sheetName];
    state.rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true,blankrows:true});
    const rowCount=state.rows.length, colCount=Math.max(0,...state.rows.map(r=>r.length));
    const meta=[state.rows[0]?.[0],state.rows[1]?.[0],state.rows[2]?.[0],state.rows[3]?.[0]].filter(v=>v!==''&&v!=null).map(cell).join(' · ');
    $('sheet-meta').textContent=`선택 시트: ${state.sheetName} | ${rowCount.toLocaleString('ko-KR')}행 × ${colCount.toLocaleString('ko-KR')}열${meta?` | 원본 표기: ${meta}`:''}`;
    state.page=0; state.colStart=0; applySearch();
  }
  function applySearch() {
    const q=$('sheet-search').value.trim().toLocaleLowerCase();
    state.filtered = !q ? state.rows.map((r,i)=>({r,i})) : state.rows.map((r,i)=>({r,i})).filter(({r})=>r.some(v=>cell(v).toLocaleLowerCase().includes(q)));
    state.page=0; renderTable();
  }
  function headingFor(col) {
    const code=cell(state.rows[4]?.[col]).trim();
    const name=cell(state.rows[5]?.[col]).trim();
    if(code && name) return `${code} · ${name}`;
    if(name) return name;
    if(code) return code;
    return `열 ${col+1} (${colName(col)})`;
  }
  function renderTable() {
    const totalCols=Math.max(0,...state.rows.map(r=>r.length));
    const start=state.colStart, end=Math.min(totalCols,start+state.colSize);
    $('column-range').textContent=totalCols?`열 ${start+1}–${end} / 전체 ${totalCols}열`:'열 없음';
    $('prev-columns').disabled=start===0;
    $('next-columns').disabled=end>=totalCols;
    const head=$('sheet-table').querySelector('thead'), body=$('sheet-table').querySelector('tbody');
    head.innerHTML=`<tr><th class="row-index">행</th>${Array.from({length:Math.max(0,end-start)},(_,k)=>`<th title="${esc(headingFor(start+k))}">${esc(headingFor(start+k))}</th>`).join('')}</tr>`;
    const pages=Math.max(1,Math.ceil(state.filtered.length/state.pageSize));
    state.page=Math.min(state.page,pages-1);
    const pageRows=state.filtered.slice(state.page*state.pageSize,(state.page+1)*state.pageSize);
    body.innerHTML=pageRows.map(({r,i})=>`<tr><td class="row-index">${i+1}</td>${Array.from({length:Math.max(0,end-start)},(_,k)=>`<td title="${esc(cell(r[start+k]))}">${esc(cell(r[start+k]))}</td>`).join('')}</tr>`).join('') || `<tr><td colspan="${Math.max(1,end-start+1)}">검색 결과가 없습니다.</td></tr>`;
    $('row-count').textContent=`검색 결과 ${state.filtered.length.toLocaleString('ko-KR')}행 · ${pageRows.length}행 표시`;
    $('page-label').textContent=`${state.page+1} / ${pages}`;
    $('prev-page').disabled=state.page===0; $('next-page').disabled=state.page>=pages-1;
  }
  function downloadCSV() {
    const rows=state.filtered.map(({r})=>r);
    const csv='\uFEFF'+rows.map(r=>r.map(v=>`"${cell(v).replace(/"/g,'""')}"`).join(',')).join('\r\n');
    const url=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8;'}));
    const a=document.createElement('a');a.href=url;a.download=`${state.sheetName.replace(/[\\/:*?"<>|]/g,'_')}_원본자료.csv`;a.click();URL.revokeObjectURL(url);
  }
  $('load-workbook').addEventListener('click',loadSelected);
  $('local-file').addEventListener('change',e=>loadLocal(e.target.files[0]));
  $('sheet-select').addEventListener('change',renderSheet);
  $('sheet-search').addEventListener('input',applySearch);
  $('prev-columns').addEventListener('click',()=>{state.colStart=Math.max(0,state.colStart-state.colSize);renderTable();});
  $('next-columns').addEventListener('click',()=>{state.colStart+=state.colSize;renderTable();});
  $('prev-page').addEventListener('click',()=>{state.page=Math.max(0,state.page-1);renderTable();});
  $('next-page').addEventListener('click',()=>{state.page++;renderTable();});
  $('download-sheet').addEventListener('click',downloadCSV);
})();
