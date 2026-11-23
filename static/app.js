let headers=[], records=[];
const $=s=>document.querySelector(s);
async function api(url, opts={}) {
  const res=await fetch(url,{headers:{'Content-Type':'application/json'},...opts});
  const data=await res.json().catch(()=>({}));
  if(!res.ok) throw new Error(data.error||'Something went wrong');
  return data;
}
function toast(msg){const t=$('#toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),2400)}
function setSaving(msg){$('#saveState').textContent=msg}
async function load(){
  try{
    const d=await api('/api/data');headers=d.headers;records=d.records;
    $('#fileName').textContent=d.file;$('#recordCount').textContent=records.length;$('#columnCount').textContent=headers.length;
    render();
  }catch(e){toast(e.message)}
}
function render(){
  const q=$('#search').value.trim().toLowerCase();
  const filtered=records.filter(r=>r.values.some(v=>String(v??'').toLowerCase().includes(q)));
  $('#resultCount').textContent=`Showing ${filtered.length} of ${records.length} records`;
  const thead=$('#dataTable thead'),tbody=$('#dataTable tbody');
  thead.innerHTML='<tr>'+headers.map((h,i)=>`<th><div class="th-inner"><span>${esc(h)}</span><button class="col-menu" title="Column options" data-col="${i+1}">⋯</button></div></th>`).join('')+'<th>Actions</th></tr>';
  tbody.innerHTML='';
  if(!filtered.length){tbody.innerHTML=`<tr><td class="empty" colspan="${headers.length+1}">No matching records found.</td></tr>`;return}
  filtered.forEach(r=>{
    const tr=document.createElement('tr');
    r.values.forEach((v,i)=>{
      const td=document.createElement('td');td.className='editable';td.textContent=v??'';td.title='Double-click to edit';
      td.addEventListener('dblclick',()=>editCell(td,r._row,i+1,v??''));tr.appendChild(td);
    });
    const action=document.createElement('td');action.innerHTML=`<div class="row-actions"><button title="Delete record">⌫</button></div>`;
    action.querySelector('button').addEventListener('click',async()=>{if(confirm('Delete this record from the Excel file?')){try{await api(`/api/rows/${r._row}`,{method:'DELETE'});await load();toast('Record deleted from Excel')}catch(e){toast(e.message)}}});
    tr.appendChild(action);tbody.appendChild(tr);
  });
  thead.querySelectorAll('.col-menu').forEach(b=>b.addEventListener('click',()=>columnMenu(Number(b.dataset.col))));
}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function editCell(td,row,col,value){
  if(td.querySelector('input'))return;
  const input=document.createElement('input');input.className='cell-input';input.value=value;td.textContent='';td.appendChild(input);input.focus();input.select();
  let done=false;
  async function save(){if(done)return;done=true;const next=input.value;if(next===String(value)){td.textContent=value;return}
    setSaving('Saving to Excel…');
    try{await api('/api/cells',{method:'PUT',body:JSON.stringify({row,col,value:next})});await load();setSaving('All changes saved');toast('Cell updated in Excel')}
    catch(e){td.textContent=value;setSaving('Save failed');toast(e.message)}
  }
  input.addEventListener('keydown',e=>{if(e.key==='Enter')save();if(e.key==='Escape'){done=true;td.textContent=value}});
  input.addEventListener('blur',save);
}
function openModal(title,fields,callback,buttonLabel='Save'){
  $('#modalTitle').textContent=title;$('#saveModal').textContent=buttonLabel;
  $('#modalFields').innerHTML=fields.map((h,i)=>`<div class="field"><label>${esc(h)}</label><input data-index="${i}" placeholder="Enter ${esc(h.toLowerCase())}"></div>`).join('');
  $('#modal').classList.add('open');
  const close=()=>$('#modal').classList.remove('open');
  $('#closeModal').onclick=close;$('#cancelModal').onclick=close;
  $('#saveModal').onclick=async()=>{const vals=[...$('#modalFields').querySelectorAll('input')].map(x=>x.value);try{await callback(vals);close();await load()}catch(e){toast(e.message)}};
  $('#modalFields input')?.focus();
}
$('#addRecord').onclick=()=>openModal('Add new record',headers,async vals=>{await api('/api/rows',{method:'POST',body:JSON.stringify({values:vals})});toast('New record saved to Excel')},'Add record');
$('#addColumn').onclick=()=>openModal('Add a column',['New column name'],async vals=>{await api('/api/columns',{method:'POST',body:JSON.stringify({name:vals[0]})});toast('Column added to Excel')},'Add column');
async function columnMenu(col){
  const action=prompt(`Column: ${headers[col-1]}\nType R to rename or D to delete:`);
  if(!action)return;
  if(action.toLowerCase()==='r'){
    const name=prompt('Enter the new column name:',headers[col-1]);if(!name)return;
    try{await api(`/api/columns/${col}`,{method:'PUT',body:JSON.stringify({name})});await load();toast('Column renamed in Excel')}catch(e){toast(e.message)}
  }else if(action.toLowerCase()==='d'){
    if(confirm(`Delete column "${headers[col-1]}" and all its cell values?`)){
      try{await api(`/api/columns/${col}`,{method:'DELETE'});await load();toast('Column deleted from Excel')}catch(e){toast(e.message)}
    }
  }
}
$('#search').addEventListener('input',render);
load();
