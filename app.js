const workbook={
Sales:[
["Date","Salesperson","Region","Product","Units","Revenue","Status"],
["2026-09-01","Raj","North","Product A",25,125000,"Active"],
["2026-09-02","Amit","West","Product B",18,99000,"Active"],
["2026-09-03","Neha","South","Product C",32,176000,"Active"],
["2026-09-04","Priya","East","Product A",21,105000,"Active"],
["2026-09-05","Raj","North","Product B",27,148500,"Active"],
["2026-09-06","Amit","West","Product C",15,82500,"Active"]
],
Customers:[
["Customer ID","Customer Name","City","State","Segment","Status"],
["C001","ABC Industries","Delhi","Delhi","Enterprise","Active"],
["C002","Sharma Traders","Pune","Maharashtra","SMB","Active"],
["C003","Mysore Engineering","Mysuru","Karnataka","Enterprise","Active"],
["C004","Eastern Supplies","Kolkata","West Bengal","SMB","Inactive"],
["C005","Coastal Metals","Chennai","Tamil Nadu","Enterprise","Active"],
["C006","Bharat Components","Hyderabad","Telangana","SMB","Active"]
],
Inventory:[
["SKU","Product","Category","Warehouse","Stock","Reorder Level"],
["P001","Product A","Industrial","Bengaluru",145,50],
["P002","Product B","Industrial","Pune",82,40],
["P003","Product C","Electrical","Chennai",36,45],
["P004","Product D","Electrical","Delhi",210,60],
["P005","Product E","Mechanical","Hyderabad",74,35],
["P006","Product F","Mechanical","Kolkata",28,30]
],
Employees:[
["Employee ID","Name","Department","Designation","Location","Joining Year"],
["E001","Raj Kumar","Sales","Manager","Delhi",2021],
["E002","Amit Shah","Sales","Executive","Pune",2023],
["E003","Neha Rao","Operations","Manager","Bengaluru",2020],
["E004","Priya Nair","Finance","Analyst","Chennai",2024],
["E005","Vikram Singh","IT","Engineer","Hyderabad",2022],
["E006","Anita Das","HR","Executive","Kolkata",2023]
],
Projects:[
["Project ID","Project Name","Manager","Priority","Start Date","Status"],
["PR001","Plant Automation","Raj Kumar","High","2026-07-01","In Progress"],
["PR002","ERP Upgrade","Neha Rao","Medium","2026-08-10","In Progress"],
["PR003","CRM Implementation","Vikram Singh","High","2026-06-15","Completed"],
["PR004","Warehouse Digitisation","Amit Shah","Medium","2026-09-01","In Progress"],
["PR005","Energy Monitoring","Priya Nair","Low","2026-09-10","Planning"],
["PR006","Quality Dashboard","Anita Das","High","2026-08-20","In Progress"]
]};

let currentSheet="Sales",selectedRow=1,selectedColumn=0,history=[],undoStack=[],redoStack=[];
const spreadsheet=document.getElementById("spreadsheet");
const sheetTabs=document.getElementById("sheetTabs");
const formulaBar=document.getElementById("formulaBar");
const nameBox=document.getElementById("nameBox");

function columnName(number){
  let result=""; number++;
  while(number>0){const remainder=(number-1)%26;result=String.fromCharCode(65+remainder)+result;number=Math.floor((number-1)/26)}
  return result;
}
function escapeHTML(value){return String(value??"").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;")}

function renderSheetTabs(){
  sheetTabs.innerHTML="";
  Object.keys(workbook).forEach(sheetName=>{
    const tab=document.createElement("div");
    tab.className="sheet-tab"+(sheetName===currentSheet?" active":"");
    tab.textContent=sheetName;
    tab.onclick=()=>{currentSheet=sheetName;selectedRow=1;selectedColumn=0;render()};
    sheetTabs.appendChild(tab);
  });
}

function render(){
  renderSheetTabs();
  const data=workbook[currentSheet];
  const columns=Math.max(12,...data.map(row=>row.length));
  const rows=Math.max(data.length,60);
  let html='<table><colgroup><col style="width:46px">';
  for(let c=0;c<columns;c++)html+='<col style="width:120px">';
  html+='</colgroup><thead><tr><th class="corner"></th>';
  for(let c=0;c<columns;c++)html+=`<th>${columnName(c)}</th>`;
  html+='</tr></thead><tbody>';
  for(let r=0;r<rows;r++){
    const row=data[r]||[];
    html+=`<tr><th class="row-number">${r+1}</th>`;
    for(let c=0;c<columns;c++){
      const value=row[c]??"",key=`${currentSheet}!${r},${c}`;
      const changed=history.some(item=>item.key===key);
      const selected=r===selectedRow&&c===selectedColumn;
      html+=`<td data-row="${r}" data-column="${c}" class="${changed?"changed ":""}${selected?"selected":""}">${escapeHTML(value)}</td>`;
    }
    html+="</tr>";
  }
  html+="</tbody></table>";
  spreadsheet.innerHTML=html;
  spreadsheet.querySelectorAll("td").forEach(cell=>{
    cell.onclick=()=>selectCell(Number(cell.dataset.row),Number(cell.dataset.column));
    cell.ondblclick=()=>editCell(cell);
  });
  updateFormulaBar();
}

function selectCell(row,column){
  selectedRow=row;selectedColumn=column;
  nameBox.value=columnName(column)+(row+1);
  updateFormulaBar();render();
}
function updateFormulaBar(){
  const data=workbook[currentSheet],value=data[selectedRow]?.[selectedColumn]??"";
  formulaBar.value=value;
}

function editCell(cell){
  const row=Number(cell.dataset.row),column=Number(cell.dataset.column);
  const oldValue=workbook[currentSheet][row]?.[column]??"";
  cell.contentEditable="true";cell.classList.add("editing");cell.focus();
  const finish=()=>{
    cell.contentEditable="false";cell.classList.remove("editing");
    const newValue=cell.textContent;
    if(newValue!==String(oldValue))commitChange(row,column,oldValue,newValue);
  };
  cell.onblur=finish;
  cell.onkeydown=event=>{
    if(event.key==="Enter"){event.preventDefault();finish()}
    if(event.key==="Escape"){cell.textContent=oldValue;cell.contentEditable="false";cell.classList.remove("editing")}
  };
}

function commitChange(row,column,oldValue,newValue){
  if(!workbook[currentSheet][row])workbook[currentSheet][row]=[];
  workbook[currentSheet][row][column]=newValue;
  const change={
    key:`${currentSheet}!${row},${column}`,sheet:currentSheet,row,column,
    oldValue,newValue,user:"You",time:new Date().toLocaleString()
  };
  history.unshift(change);undoStack.push(change);redoStack=[];
  renderHistory();render();
}

function renderHistory(){
  const container=document.getElementById("historyList");
  container.innerHTML=history.slice(0,100).map(change=>`
    <div class="history-item">
      <div class="history-cell">${change.sheet}!${columnName(change.column)}${change.row+1}</div>
      <span class="old-value">${escapeHTML(change.oldValue)}</span> → 
      <span class="new-value">${escapeHTML(change.newValue)}</span><br>
      ${change.user} · ${change.time}
    </div>`).join("");
}

document.getElementById("undoBtn").onclick=()=>{
  const change=undoStack.pop();if(!change)return;
  workbook[change.sheet][change.row][change.column]=change.oldValue;
  redoStack.push(change);render();
};
document.getElementById("redoBtn").onclick=()=>{
  const change=redoStack.pop();if(!change)return;
  workbook[change.sheet][change.row][change.column]=change.newValue;
  undoStack.push(change);render();
};
formulaBar.onchange=()=>{
  const oldValue=workbook[currentSheet][selectedRow]?.[selectedColumn]??"";
  const newValue=formulaBar.value;
  if(String(oldValue)!==newValue)commitChange(selectedRow,selectedColumn,oldValue,newValue);
};
nameBox.onchange=()=>{
  const match=nameBox.value.match(/^([A-Z]+)([0-9]+)$/i);if(!match)return;
  let column=0;for(const ch of match[1].toUpperCase())column=column*26+ch.charCodeAt(0)-64;
  selectCell(Number(match[2])-1,column-1);
};
document.getElementById("wrapBtn").onclick=()=>spreadsheet.classList.toggle("wrap");
document.getElementById("historyBtn").onclick=()=>document.getElementById("historyPanel").classList.remove("hidden");
document.getElementById("closeHistory").onclick=()=>document.getElementById("historyPanel").classList.add("hidden");
document.getElementById("boldBtn").onclick=()=>document.execCommand("bold");
document.getElementById("italicBtn").onclick=()=>document.execCommand("italic");
document.getElementById("freezeBtn").onclick=()=>{
  const row=document.querySelector("tbody tr");if(!row)return;
  row.querySelectorAll("td,th").forEach(cell=>{
    if(cell.classList.contains("row-number"))return;
    cell.style.position="sticky";cell.style.top="27px";cell.style.zIndex="3";
  });
};
document.getElementById("exportBtn").onclick=()=>{
  alert("Excel export will be connected with ExcelJS/SheetJS in the next build.");
};
render();renderHistory();