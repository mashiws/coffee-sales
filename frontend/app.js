
const api = (url, opt={}) => fetch(url, opt).then(async r => {
  if(!r.ok){ let t = await r.text(); throw new Error(t); }
  const ct = r.headers.get("content-type") || "";
  return ct.includes("application/json") ? r.json() : r;
});
let drinks=[], cups=[], sales=[], trash=[], editingDrink=null, editingCup=null;

const today = new Date().toISOString().slice(0,10);
saleDate.value = today; reportDate.value = today;

document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>{
  document.querySelectorAll(".tab,.page").forEach(x=>x.classList.remove("active"));
  b.classList.add("active"); document.getElementById(b.dataset.tab).classList.add("active");
});

function money(v){ return Number(v||0).toFixed(2); }
function timeOnly(s){ return (s||"").replace("T"," ").slice(11,16); }
function opt(list, text){ return list.map(x=>`<option value="${x.id}">${text(x)}</option>`).join(""); }

async function loadAll(){
  drinks = await api("/api/drinks");
  cups = await api("/api/cup-sizes");
  fillSelects();
  await loadSales();
  renderDicts();
}
function fillSelects(){
  saleDrink.innerHTML = `<option value="">Выберите напиток</option>` + opt(drinks, d=>`${d.name} — ${money(d.price)}`);
  saleCup.innerHTML = `<option value="">Выберите объём</option>` + opt(cups, c=>c.name);
  filterDrink.innerHTML = `<option value="">Все напитки</option>` + opt(drinks, d=>d.name);
  filterCup.innerHTML = `<option value="">Все объёмы</option>` + opt(cups, c=>c.name);
}
async function loadSales(){
  const d = saleDate.value;
  sales = await api(`/api/sales?date_=${d}`);
  trash = (await api(`/api/sales?date_=${d}&include_deleted=1`)).filter(x=>x.is_deleted);
  await loadReportForKpi(d);
  renderSales();
}
async function loadReportForKpi(d){
  const r = await api(`/api/reports/daily?date_=${d}`);
  kpiCups.textContent = r.total_cups; kpiRevenue.textContent = money(r.revenue); kpiLeader.textContent = r.leader;
}
function renderSales(){
  let q = searchSale.value.toLowerCase(), fd=filterDrink.value, fc=filterCup.value;
  let list = sales.filter(s => !s.is_deleted)
    .filter(s => !q || s.drink_name.toLowerCase().includes(q))
    .filter(s => !fd || String(s.drink_id)===fd)
    .filter(s => !fc || String(s.cup_size_id)===fc);
  salesTable.innerHTML = list.map(s=>`
    <tr><td>${timeOnly(s.sold_at)}</td><td>${s.drink_name}</td><td>${s.cup_name}</td><td>${s.quantity}</td><td>${money(s.price)}</td><td>${money(s.total)}</td>
    <td><button onclick="editSale(${s.id})">✏️</button> <button onclick="delSale(${s.id})">🗑</button></td></tr>`).join("");
  trashTable.innerHTML = trash.map(s=>`
    <tr><td>${timeOnly(s.sold_at)}</td><td>${s.drink_name}</td><td>${s.cup_name}</td><td>${s.quantity}</td>
    <td><button onclick="restoreSale(${s.id})">↩</button> <button onclick="forceSale(${s.id})">Удалить</button></td></tr>`).join("");
}
addSale.onclick = async ()=>{
  await api("/api/sales",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({drink_id:+saleDrink.value,cup_size_id:+saleCup.value,quantity:+saleQty.value})});
  saleDrink.value=""; saleCup.value=""; saleQty.value="";
  await loadSales();
};
async function delSale(id){ if(confirm("Удалить запись в корзину?")){ await api(`/api/sales/${id}`,{method:"DELETE"}); await loadSales(); } }
async function restoreSale(id){ await api(`/api/sales/${id}/restore`,{method:"POST"}); await loadSales(); }
async function forceSale(id){ if(confirm("Удалить окончательно?")){ await api(`/api/sales/${id}/force`,{method:"DELETE"}); await loadSales(); } }
function editSale(id){
  let s=sales.find(x=>x.id===id);
  modalTitle.textContent="Редактировать продажу";
  modalBody.innerHTML=`<div class="grid"><label>Напиток<select id="mDrink">${opt(drinks,d=>`${d.name} — ${money(d.price)}`)}</select></label><label>Объём<select id="mCup">${opt(cups,c=>c.name)}</select></label><label>Количество<input id="mQty" type="number" min="1" value="${s.quantity}"></label></div>`;
  mDrink.value=s.drink_id; mCup.value=s.cup_size_id;
  modalSave.onclick=async()=>{await api(`/api/sales/${id}`,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify({drink_id:+mDrink.value,cup_size_id:+mCup.value,quantity:+mQty.value})}); closeModal(); await loadSales();};
  modal.classList.remove("hidden");
}
function closeModal(){ modal.classList.add("hidden"); }
[saleDate, searchSale, filterDrink, filterCup].forEach(el=>el.oninput=()=> el===saleDate?loadSales():renderSales());
resetFilters.onclick=()=>{searchSale.value="";filterDrink.value="";filterCup.value="";renderSales();};

async function loadReport(){
  const r = await api(`/api/reports/daily?date_=${reportDate.value}`);
  repCups.textContent=r.total_cups; repRevenue.textContent=money(r.revenue); repLeader.textContent=r.leader;
  reportTable.innerHTML = r.items.map(i=>`<tr><td>${i.drink_name}</td><td>${i.cup_name}</td><td>${i.quantity}</td><td>${money(i.price)}</td><td>${money(i.total)}</td></tr>`).join("");
}
loadReport.onclick=loadReport;
xlsxReport.onclick=()=> location.href=`/api/reports/daily/export.xlsx?date_=${reportDate.value}`;
pdfReport.onclick=()=> location.href=`/api/reports/daily/export.pdf?date_=${reportDate.value}`;

function renderDicts(){
  drinksTable.innerHTML = drinks.map(d=>`<tr><td>${d.name}</td><td>${money(d.price)}</td><td><button onclick="editDrink(${d.id})">✏️</button> <button onclick="deleteDrink(${d.id})">🗑</button></td></tr>`).join("");
  cupsTable.innerHTML = cups.map(c=>`<tr><td>${c.name}</td><td><button onclick="editCup(${c.id})">✏️</button> <button onclick="deleteCup(${c.id})">🗑</button></td></tr>`).join("");
}
saveDrink.onclick=async()=>{
  let id=editingDrink; let method=id?"PUT":"POST"; let url=id?`/api/drinks/${id}`:"/api/drinks";
  await api(url,{method,headers:{"Content-Type":"application/json"},body:JSON.stringify({name:drinkName.value,price:+drinkPrice.value})});
  editingDrink=null; saveDrink.textContent="Добавить"; drinkName.value=""; drinkPrice.value=""; await loadAll();
};
function editDrink(id){ let d=drinks.find(x=>x.id===id); editingDrink=id; drinkName.value=d.name; drinkPrice.value=d.price; saveDrink.textContent="Сохранить"; }
async function deleteDrink(id){ if(confirm("Удалить напиток?")){ await api(`/api/drinks/${id}`,{method:"DELETE"}); await loadAll(); } }
saveCup.onclick=async()=>{
  let id=editingCup; let method=id?"PUT":"POST"; let url=id?`/api/cup-sizes/${id}`:"/api/cup-sizes";
  await api(url,{method,headers:{"Content-Type":"application/json"},body:JSON.stringify({name:cupName.value})});
  editingCup=null; saveCup.textContent="Добавить"; cupName.value=""; await loadAll();
};
function editCup(id){ let c=cups.find(x=>x.id===id); editingCup=id; cupName.value=c.name; saveCup.textContent="Сохранить"; }
async function deleteCup(id){ if(confirm("Удалить объём?")){ await api(`/api/cup-sizes/${id}`,{method:"DELETE"}); await loadAll(); } }
dictType.onchange=()=>{drinkPanel.classList.toggle("hidden", dictType.value!=="drinks"); cupPanel.classList.toggle("hidden", dictType.value!=="cups");};

exportBackup.onclick=()=> location.href="/api/backups/export";
importBackup.onclick=async()=>{
  if(!backupFile.files[0]) return alert("Выберите файл .zip");
  let fd=new FormData(); fd.append("file",backupFile.files[0]);
  let r=await api("/api/backups/import",{method:"POST",body:fd});
  backupResult.textContent=JSON.stringify(r,null,2);
  await loadAll();
};
loadAll(); loadReport();
