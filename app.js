import { createClient } from "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/+esm";
import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

const SUPABASE_URL="https://ccqhvmavmgngevihtcnf.supabase.co";
const SUPABASE_KEY=["sb_publishable_","ajVwLJ-Lb2tHeLFspj7gUQ_jZ_R9ide"].join("");
const supabase=createClient(SUPABASE_URL,SUPABASE_KEY);

const state={user:null,workbook:null,sheets:[],sheet:null,cells:new Map(),history:[],locks:[],selected:null,undo:[],redo:[],dirty:false};
const $=id=>document.getElementById(id);
const esc=v=>String(v??"").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;");
const colName=n=>{let s="";n++;while(n){s=String.fromCharCode(65+(n-1)%26)+s;n=Math.floor((n-1)/26)}return s};
const key=(r,c)=>r+":"+c;

function authMessage(m){$("authMessage").textContent=m||""}
function showApp(){ $("authScreen").classList.add("hidden");$("appScreen").classList.remove("hidden") }
function showAuth(){ $("appScreen").classList.add("hidden");$("authScreen").classList.remove("hidden") }

let signup=false;
$("loginTab").onclick=()=>{signup=false;$("loginTab").classList.add("active");$("signupTab").classList.remove("active");$("authName").classList.add("hidden");$("authBtn").textContent="Sign In"};
$("signupTab").onclick=()=>{signup=true;$("signupTab").classList.add("active");$("loginTab").classList.remove("active");$("authName").classList.remove("hidden");$("authBtn").textContent="Create Account"};
$("authBtn").onclick=async()=>{
  authMessage("Working...");
  const email=$("authEmail").value.trim(),password=$("authPassword").value,name=$("authName").value.trim();
  if(!email||!password)return authMessage("Email and password are required.");
  const r=signup
    ? await supabase.auth.signUp({email,password,options:{data:{display_name:name}}})
    : await supabase.auth.signInWithPassword({email,password});
  if(r.error)return authMessage(r.error.message);
  if(signup&&!r.data.session)return authMessage("Check your email to verify the account, then sign in.");
  if(r.data.user) await start(r.data.user);
};
$("signOutBtn").onclick=async()=>{await supabase.auth.signOut();location.reload()};

async function ensureProfile(user){
  await supabase.from("excel_profiles").upsert({user_id:user.id,email:user.email,display_name:user.user_metadata?.display_name||user.email?.split("@")[0]},{onConflict:"user_id"});
}
async function loadWorkbook(){
  const wid=location.pathname.match(/^\/w\/([0-9a-f-]{20,})$/i)?.[1]; const {data:members,error}=await supabase.from("excel_workbook_members").select("workbook_id,role,excel_workbooks(*)").eq("user_id",state.user.id);
  if(error)throw error;
  if(wid){ const m=members?.find(x=>x.workbook_id===wid); if(!m) throw new Error("You do not have access to this workbook."); state.workbook=m.excel_workbooks; } else if(!members?.length){
    const {data,error:e}=await supabase.from("excel_workbooks").insert({name:"My Workbook",owner_id:state.user.id,source_type:"new"}).select().single();
    if(e)throw e;
    state.workbook=data;
  }else if(!state.workbook) state.workbook=members[0].excel_workbooks;
  const {data:sheets,error:e}=await supabase.from("excel_worksheets").select("*").eq("workbook_id",state.workbook.id).order("sort_order");
  if(e)throw e;
  if(!sheets.length){
    const {data:s,error:se}=await supabase.from("excel_worksheets").insert({workbook_id:state.workbook.id,name:"Sheet1",sort_order:0}).select().single();
    if(se)throw se; state.sheets=[s]; 
  }else state.sheets=sheets;
  state.sheet=state.sheets[0];
  $("userBadge").textContent=state.user.user_metadata?.display_name||state.user.email;
  $("workbookTitle").textContent=state.workbook.name;
  await loadSheet();
}
async function loadSheet(){
  const [{data:cells,error:e},{data:locks,error:l}]=await Promise.all([
    supabase.from("excel_cells").select("*").eq("worksheet_id",state.sheet.id),
    supabase.from("excel_cell_locks").select("*").eq("worksheet_id",state.sheet.id)
  ]);
  if(e)throw e;if(l)throw l;
  state.cells.clear();cells.forEach(x=>state.cells.set(key(x.row_index,x.col_index),x));
  state.locks=locks||[];render();
}
function isLocked(r,c){return state.locks.some(x=>x.row_start<=r&&x.row_end>=r&&x.col_start<=c&&x.col_end>=c)}
function selectedCell(){return state.selected||{r:0,c:0}}
function selectCell(r,c){state.selected={r,c};$("nameBox").value=colName(c)+(r+1);const x=state.cells.get(key(r,c));$("formulaBar").value=x?.formula??x?.value??"";render()}
function render(){
  const maxR=Math.max(50,...[...state.cells.values()].map(x=>x.row_index+1)),maxC=Math.max(12,...[...state.cells.values()].map(x=>x.col_index+1));
  let h='<table><colgroup><col style="width:46px">';
  for(let c=0;c<maxC;c++)h+='<col style="width:120px">';
  h+='</colgroup><thead><tr><th class="corner"></th>';
  for(let c=0;c<maxC;c++)h+=`<th data-col="${c}">${colName(c)}</th>`;h+='</tr></thead><tbody>';
  for(let r=0;r<maxR;r++){h+=`<tr><th class="row-number" data-row="${r}">${r+1}</th>`;
    for(let c=0;c<maxC;c++){const x=state.cells.get(key(r,c)),v=x?.value??"",sel=state.selected?.r===r&&state.selected?.c===c,locked=isLocked(r,c);
      h+=`<td data-row="${r}" data-col="${c}" class="${sel?"selected ":""}${locked?"locked ":""}${state.history.some(h=>h.action==="cell_edit"&&h.worksheet_id===state.sheet.id&&h.row_index===r&&h.col_index===c)?"changed":""}" style="${fmt(x?.format)}">${esc(v)}</td>`}
    h+='</tr>'}h+='</tbody></table>';spreadsheet.innerHTML=h;
  spreadsheet.querySelectorAll("td").forEach(td=>{td.onclick=()=>selectCell(+td.dataset.row,+td.dataset.col);td.ondblclick=()=>edit(td)});
  spreadsheet.querySelectorAll("th[data-col]").forEach(th=>th.onclick=()=>selectColumn(+th.dataset.col));
  spreadsheet.querySelectorAll(".row-number").forEach(th=>th.onclick=()=>selectRow(+th.dataset.row));
  renderTabs(); $("openSupabase").onclick=async()=>{const {data}=await supabase.from("excel_workbook_members").select("workbook_id,role,excel_workbooks(*)").eq("user_id",state.user.id);const list=(data||[]).map(x=>x.excel_workbooks.name+" — "+x.role+" — "+location.origin+"/w/"+x.workbook_id).join("\n");alert(list||"No workbooks found.");}; $("openDrive").onclick=()=>alert("Google Drive requires Google OAuth/Picker credentials. The XLSX editor is ready; connect a Google Cloud OAuth client to enable Drive import.");renderHistory();
}
function fmt(f){f=f||{};return Object.entries(f).map(([k,v])=>({bold:"font-weight:bold",italic:"font-style:italic",underline:"text-decoration:underline",color:"color:"+v,bg:"background:"+v,align:"text-align:"+v,size:"font-size:"+v+"px"}[k]||"")).filter(Boolean).join(";")}
function renderTabs(){$("sheetTabs").innerHTML=state.sheets.map(s=>`<div class="sheet-tab ${s.id===state.sheet.id?"active":""}" data-id="${s.id}">${esc(s.name)}</div>`).join("");$("sheetTabs").querySelectorAll(".sheet-tab").forEach(t=>t.onclick=async()=>{state.sheet=state.sheets.find(s=>s.id===t.dataset.id);await loadSheet()})}
function renderHistory(){
  $("historyList").innerHTML=state.history.slice(0,100).map(x=>`<div class="history-item"><b>${esc(x.action)}</b> · ${esc(x.email||"User")}<br>${esc(x.detail||"")}<br><small>${new Date(x.created_at).toLocaleString()}</small></div>`).join("");
}
async function loadHistory(){
  const {data}=await supabase.from("excel_edit_history").select("*").eq("workbook_id",state.workbook.id).order("created_at",{ascending:false}).limit(100);
  state.history=data||[];renderHistory();
}
async function saveCell(r,c,value,format=null){
  if(isLocked(r,c))return alert("This cell is locked.");
  const old=state.cells.get(key(r,c)),payload={worksheet_id:state.sheet.id,row_index:r,col_index:c,value:String(value??""),formula:null,format:format??old?.format??{},updated_by:state.user.id};
  const {data,error}=await supabase.from("excel_cells").upsert(payload,{onConflict:"worksheet_id,row_index,col_index"}).select().single();
  if(error)return alert(error.message);
  state.cells.set(key(r,c),data);state.dirty=true;state.undo.push({r,c,old:old?.value??"",value:data.value});state.redo=[];
  render();
}
function edit(td){
  const r=+td.dataset.row,c=+td.dataset.col;if(isLocked(r,c))return alert("This cell is locked.");
  const old=state.cells.get(key(r,c))?.value??"";td.contentEditable="true";td.focus();
  const done=async()=>{td.contentEditable="false";if(td.textContent!==String(old))await saveCell(r,c,td.textContent)};
  td.onblur=done;td.onkeydown=e=>{if(e.key==="Enter"){e.preventDefault();done()}if(e.key==="Escape"){td.textContent=old;td.blur()}};
}
function selectRow(r){state.selected={r,c:0};state.selection={type:"row",r};render()}
function selectColumn(c){state.selected={r:0,c};state.selection={type:"column",c};render()}
async function addLock(type){
  const s=state.selected||{r:0,c:0},range=state.selection?.type==="row"?{rs:s.r,re:s.r,cs:0,ce:16383}:state.selection?.type==="column"?{rs:0,re:1048575,cs:s.c,ce:s.c}:{rs:s.r,re:s.r,cs:s.c,ce:s.c};
  const {error}=await supabase.from("excel_cell_locks").insert({workbook_id:state.workbook.id,worksheet_id:state.sheet.id,lock_type:type,row_start:range.rs,row_end:range.re,col_start:range.cs,col_end:range.ce,locked_by:state.user.id});
  if(error)return alert(error.message);await loadSheet();
}
$("lockCellBtn").onclick=()=>addLock("cell");$("lockRowBtn").onclick=()=>addLock("row");$("lockColBtn").onclick=()=>addLock("column");
$("unlockBtn").onclick=async()=>{const s=selectedCell();const l=state.locks.filter(x=>x.row_start<=s.r&&x.row_end>=s.r&&x.col_start<=s.c&&x.col_end>=s.c);for(const x of l)await supabase.from("excel_cell_locks").delete().eq("id",x.id);await loadSheet()};
$("historyBtn").onclick=()=>{$("historyPanel").classList.remove("hidden");loadHistory()}; $("shareBtn").onclick=async()=>{const email=prompt("Email address to invite:");if(!email)return;const role=confirm("OK = Editor, Cancel = Viewer")?"editor":"viewer";const {data,error}=await supabase.functions.invoke("excelshare-invite",{body:{email,role,workbook_id:state.workbook.id,redirect_to:location.origin}});alert(error?.message||data?.error||(data?.ok?"Invitation sent.":"Unable to send invitation."))};$("closeHistory").onclick=()=>$("historyPanel").classList.add("hidden");
$("boldBtn").onclick=()=>formatSelected("bold",true);$("italicBtn").onclick=()=>formatSelected("italic",true);$("underlineBtn").onclick=()=>formatSelected("underline",true);
$("fontColor").onchange=e=>formatSelected("color",e.target.value);$("cellColor").onchange=e=>formatSelected("bg",e.target.value);$("fontSize").onchange=e=>formatSelected("size",e.target.value);
$("alignLeft").onclick=()=>formatSelected("align","left");$("alignCenter").onclick=()=>formatSelected("align","center");$("alignRight").onclick=()=>formatSelected("align","right");
async function formatSelected(prop,value){const s=selectedCell();const x=state.cells.get(key(s.r,s.c));const f={...(x?.format||{}),[prop]:value};await saveCell(s.r,s.c,x?.value??"",f)}
$("formulaBar").onchange=()=>{const s=selectedCell();saveCell(s.r,s.c,$("formulaBar").value)}; $("nameBox").onchange=()=>{const m=$("nameBox").value.match(/^([A-Z]+)([0-9]+)$/i);if(!m)return;let n=0;for(const ch of m[1].toUpperCase())n=n*26+ch.charCodeAt(0)-64;selectCell(+m[2]-1,n-1)};
$("undoBtn").onclick=async()=>{const x=state.undo.pop();if(x){await saveCell(x.r,x.c,x.old);state.redo.push(x)}};$("redoBtn").onclick=async()=>{const x=state.redo.pop();if(x)await saveCell(x.r,x.c,x.value)};
$("wrapBtn").onclick=()=>spreadsheet.classList.toggle("wrap");
$("freezeBtn").onclick=()=>document.querySelector("thead").classList.toggle("frozen");
$("exportBtn").onclick=()=>exportXlsx();
$("openComputer").onchange=async e=>{const f=e.target.files[0];if(f)await importXlsx(f)};
async function importXlsx(file){
  const buf=await file.arrayBuffer(),wb=XLSX.read(buf,{type:"array"}),name=file.name.replace(/\.xlsx?$/i,"");
  const {data:w,error}=await supabase.from("excel_workbooks").insert({name,owner_id:state.user.id,original_filename:file.name,source_type:"computer"}).select().single();
  if(error)return alert(error.message);state.workbook=w;
  state.sheets=[];for(let i=0;i<wb.SheetNames.length;i++){const n=wb.SheetNames[i],{data:s}=await supabase.from("excel_worksheets").insert({workbook_id:w.id,name:n,sort_order:i}).select().single();state.sheets.push(s);
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,defval:""});const batch=[];rows.forEach((row,r)=>row.forEach((v,c)=>{if(v!=="")batch.push({worksheet_id:s.id,row_index:r,col_index:c,value:String(v),updated_by:state.user.id,format:{}})}));for(let i=0;i<batch.length;i+=500)await supabase.from("excel_cells").insert(batch.slice(i,i+500));
  }
  state.sheet=state.sheets[0];$("workbookTitle").textContent=name;history=[];await loadSheet();
}
async function exportXlsx(){
  const wb=XLSX.utils.book_new();for(const s of state.sheets){const {data}=await supabase.from("excel_cells").select("*").eq("worksheet_id",s.id);const maxR=Math.max(0,...(data||[]).map(x=>x.row_index)),maxC=Math.max(0,...(data||[]).map(x=>x.col_index));const a=Array.from({length:maxR+1},()=>Array(maxC+1).fill(""));(data||[]).forEach(x=>a[x.row_index][x.col_index]=x.value);XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(a),s.name.slice(0,31))}XLSX.writeFile(wb,(state.workbook.name||"ExcelShare")+".xlsx")}
async function start(user){state.user=user;showApp();try{await ensureProfile(user);await loadWorkbook();await loadHistory();subscribe()}catch(e){alert(e.message)}}
function subscribe(){
  supabase.channel("excelshare-"+state.workbook.id).on("postgres_changes",{event:"*",schema:"public",table:"excel_cells"},p=>{if(p.new?.worksheet_id===state.sheet?.id&&p.new?.updated_by!==state.user.id){state.cells.set(key(p.new.row_index,p.new.col_index),p.new);render()}}).subscribe();
}
supabase.auth.getSession().then(({data})=>{if(data.session)start(data.session.user)});
supabase.auth.onAuthStateChange((_e,s)=>{if(s&&!state.user)start(s.user)});
renderTabs();
