<!DOCTYPE html>
<html lang="it">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Inventario 库存管理</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <style>
    .tab-active{border-bottom:2px solid #2563eb;color:#2563eb}
  </style>
</head>
<body class="bg-gray-100 min-h-screen">
<div id="app" class="max-w-2xl mx-auto bg-white min-h-screen shadow">
  <!-- Header -->
  <div class="p-4 border-b sticky top-0 bg-white z-10">
    <h1 class="text-xl font-bold text-center">库存管理 <span class="text-gray-500 text-sm">Inventario</span></h1>
    <div class="flex justify-center gap-2 mt-2">
      <button onclick="lang='zh';render()" class="px-2 py-1 rounded text-sm border" :class="lang==='zh'?'bg-blue-500 text-white':''">中文</button>
      <button onclick="lang='en';render()" class="px-2 py-1 rounded text-sm border" :class="lang==='en'?'bg-blue-500 text-white':''">EN</button>
      <button onclick="lang='it';render()" class="px-2 py-1 rounded text-sm border" :class="lang==='it'?'bg-blue-500 text-white':''">IT</button>
    </div>
  </div>

  <!-- Tab导航 -->
  <div class="flex border-b text-sm">
    <div onclick="switchTab('stock')" class="flex-1 text-center py-2 tab-item cursor-pointer" data-tab="stock">库存</div>
    <div onclick="switchTab('in')" class="flex-1 text-center py-2 tab-item cursor-pointer" data-tab="in">入库</div>
    <div onclick="switchTab('out')" class="flex-1 text-center py-2 tab-item cursor-pointer" data-tab="out">出库</div>
    <div onclick="switchTab('check')" class="flex-1 text-center py-2 tab-item cursor-pointer" data-tab="check">盘点</div>
    <div onclick="switchTab('search')" class="flex-1 text-center py-2 tab-item cursor-pointer" data-tab="search">查询</div>
    <div onclick="switchTab('low')" class="flex-1 text-center py-2 tab-item cursor-pointer" data-tab="low">低库存</div>
    <div onclick="switchTab('setting')" class="flex-1 text-center py-2 tab-item cursor-pointer" data-tab="setting">工具</div>
  </div>

  <!-- 库存列表页 -->
  <div id="page-stock" class="page p-4 hidden">
    <div class="mb-4">
      <input id="stockSearch" type="text" placeholder="搜索条码/名称 / Cerca codice/nome" class="w-full border p-2 rounded" oninput="render()">
    </div>
    <div id="stockList"></div>
    <button onclick="openAdd()" class="fixed bottom-6 right-6 w-12 h-12 rounded-full bg-blue-500 text-white text-2xl shadow-lg">+</button>
  </div>

  <!-- 入库 -->
  <div id="page-in" class="page p-4 hidden">
    <h2 class="font-semibold mb-3">入库 / Entrata merce</h2>
    <div class="space-y-3">
      <div>
        <label>条码 Codice</label>
        <div class="flex gap-2">
          <input id="inCode" class="flex-1 border p-2 rounded">
          <button onclick="scanCode('inCode')" class="bg-gray-700 text-white px-3 rounded">扫码</button>
        </div>
      </div>
      <div><label>数量 Quantità</label><input id="inQty" type="number" min="1" value="1" class="w-full border p-2 rounded"></div>
      <button onclick="stockIn()" class="w-full bg-green-500 text-white py-2 rounded">确认入库 Conferma</button>
    </div>
  </div>

  <!-- 出库 -->
  <div id="page-out" class="page p-4 hidden">
    <h2 class="font-semibold mb-3">出库 / Uscita merce</h2>
    <div class="space-y-3">
      <div>
        <label>条码 Codice</label>
        <div class="flex gap-2">
          <input id="outCode" class="flex-1 border p-2 rounded">
          <button onclick="scanCode('outCode')" class="bg-gray-700 text-white px-3 rounded">扫码</button>
        </div>
      </div>
      <div><label>数量 Quantità</label><input id="outQty" type="number" min="1" value="1" class="w-full border p-2 rounded"></div>
      <button onclick="stockOut()" class="w-full bg-red-500 text-white py-2 rounded">确认出库 Conferma</button>
    </div>
  </div>

  <!-- 盘点页面（升级：按条码排序） -->
  <div id="page-check" class="page p-4 hidden">
    <h2 class="font-semibold mb-3">货架盘点补货清单 / Lista riordino scaffale <span class="text-sm text-gray-400">(按条码排序)</span></h2>
    <div id="checkList"></div>
    <button onclick="exportCheckExcel()" class="w-full mt-4 bg-blue-500 text-white py-2 rounded">导出Excel Esporta Excel</button>
  </div>

  <!-- 查询 -->
  <div id="page-search" class="page p-4 hidden">
    <h2 class="font-semibold mb-3">商品查询 / Ricerca articolo</h2>
    <div class="flex gap-2 mb-3">
      <input id="searchCode" placeholder="输入条码 / Inserisci codice" class="flex-1 border p-2 rounded">
      <button onclick="scanCode('searchCode')" class="bg-gray-700 text-white px-3 rounded">扫码</button>
    </div>
    <button onclick="searchItem()" class="w-full bg-blue-500 text-white py-2 rounded">搜索 Cerca</button>
    <div id="searchResult" class="mt-4"></div>
  </div>

  <!-- 低库存 -->
  <div id="page-low" class="page p-4 hidden">
    <h2 class="font-semibold mb-3">低库存提醒 / Scorte basse</h2>
    <div id="lowList"></div>
  </div>

  <!-- 工具页 -->
  <div id="page-setting" class="page p-4 hidden">
    <h2 class="font-semibold mb-3">工具 Strumenti</h2>
    <div class="space-y-3">
      <button onclick="exportAll()" class="w-full border py-2 rounded">导出全部库存 Esporta tutto</button>
      <button onclick="importExcel()" class="w-full border py-2 rounded">导入Excel Importa</button>
      <input type="file" id="importFile" accept=".xlsx,.xls" class="hidden" onchange="doImport()">
      <button onclick="clearAll()" class="w-full border py-2 rounded text-red-500">清空全部（谨慎！）Svuota tutto</button>
    </div>
  </div>
</div>

<!-- 添加/编辑弹窗 -->
<div id="editModal" class="fixed inset-0 bg-black/50 hidden z-50 flex items-center justify-center p-4">
  <div class="bg-white w-full max-w-md rounded p-4 max-h-[90vh] overflow-y-auto">
    <h3 class="font-bold mb-3" id="modalTitle">新增商品 Nuovo articolo</h3>
    <div class="space-y-3">
      <div>
        <label>条码 Codice *</label>
        <div class="flex gap-2">
          <input id="m_code" class="flex-1 border p-2 rounded">
          <button onclick="scanCode('m_code')" class="bg-gray-700 text-white px-3 rounded">扫码</button>
        </div>
      </div>
      <div><label>品牌 Marchio</label><input id="m_brand" class="w-full border p-2 rounded"></div>
      <div><label>商品名称 Nome prodotto</label><input id="m_name" class="w-full border p-2 rounded"></div>
      <div><label>规格 Dimensione</label><input id="m_spec" class="w-full border p-2 rounded"></div>
      <div><label>当前库存 Giacenza iniziale</label><input id="m_qty" type="number" min="0" value="0" class="w-full border p-2 rounded"></div>
      <div><label>最低预警库存 Soglia minima</label><input id="m_min" type="number" min="0" value="2" class="w-full border p-2 rounded"></div>
      <div>
        <label>图片 Foto</label>
        <input type="file" id="m_img" accept="image/*" class="w-full border p-2 rounded" onchange="previewImg()">
        <img id="imgPreview" class="mt-2 max-h-32 hidden">
      </div>
    </div>
    <div class="flex gap-2 mt-4">
      <button onclick="closeModal()" class="flex-1 border py-2 rounded">取消 Annulla</button>
      <button onclick="saveItem()" class="flex-1 bg-blue-500 text-white py-2 rounded">保存 Salva</button>
    </div>
  </div>
</div>

<script>
let lang = "zh";
let currentTab = "stock";
let editId = null;
let stock = JSON.parse(localStorage.getItem("stock")) || [];

// 语言包
const txt = {
  zh:{add:"新增商品",edit:"编辑商品",del:"删除",save:"保存",cancel:"取消",warn:"警告",confirm:"确认?"},
  en:{add:"Add item",edit:"Edit",del:"Delete",save:"Save",cancel:"Cancel",warn:"Warning",confirm:"Sure?"},
  it:{add:"Aggiungi articolo",edit:"Modifica",del:"Cancella",save:"Salva",cancel:"Annulla",warn:"Attenzione",confirm:"Sei sicuro?"}
}

function switchTab(tab){
  currentTab = tab;
  document.querySelectorAll(".page").forEach(p=>p.classList.add("hidden"));
  document.getElementById("page-"+tab).classList.remove("hidden");
  document.querySelectorAll(".tab-item").forEach(t=>t.classList.remove("tab-active"));
  document.querySelector(`[data-tab="${tab}"]`).classList.add("tab-active");
  render();
}

// ===== 核心：精准查找【完全匹配条码】 =====
function findByCode(code){
  return stock.find(i=>i.code === code.trim())
}
function findIndexById(id){
  return stock.findIndex(i=>i.id === id)
}

// 保存本地
function saveStorage(){
  localStorage.setItem("stock", JSON.stringify(stock))
}

// 弹窗
function openAdd(){
  editId=null;
  document.getElementById("modalTitle").innerText = txt[lang].add;
  document.getElementById("m_code").value="";
  document.getElementById("m_brand").value="";
  document.getElementById("m_name").value="";
  document.getElementById("m_spec").value="";
  document.getElementById("m_qty").value="0";
  document.getElementById("m_min").value="2";
  document.getElementById("imgPreview").classList.add("hidden");
  document.getElementById("editModal").classList.remove("hidden");
}
function openEdit(id){
  editId=id;
  const item = stock.find(x=>x.id===id);
  document.getElementById("modalTitle").innerText = txt[lang].edit;
  document.getElementById("m_code").value=item.code;
  document.getElementById("m_brand").value=item.brand;
  document.getElementById("m_name").value=item.name;
  document.getElementById("m_spec").value=item.spec;
  document.getElementById("m_qty").value=item.qty;
  document.getElementById("m_min").value=item.min;
  if(item.img){
    document.getElementById("imgPreview").src=item.img;
    document.getElementById("imgPreview").classList.remove("hidden");
  }else{
    document.getElementById("imgPreview").classList.add("hidden");
  }
  document.getElementById("editModal").classList.remove("hidden");
}
function closeModal(){
  document.getElementById("editModal").classList.add("hidden");
}

function previewImg(){
  const f = document.getElementById("m_img").files[0];
  if(!f) return;
  const r = new FileReader();
  r.onload = e=>{
    document.getElementById("imgPreview").src=e.target.result;
    document.getElementById("imgPreview").classList.remove("hidden");
  }
  r.readAsDataURL(f);
}

function saveItem(){
  const code = document.getElementById("m_code").value.trim();
  const brand = document.getElementById("m_brand").value.trim();
  const name = document.getElementById("m_name").value.trim();
  const spec = document.getElementById("m_spec").value.trim();
  const qty = Number(document.getElementById("m_qty").value);
  const min = Number(document.getElementById("m_min").value);
  const imgDom = document.getElementById("imgPreview");
  const img = imgDom.classList.contains("hidden") ? "" : imgDom.src;
  if(!code){ alert("条码不能为空 / Codice obbligatorio"); return; }

  // 新增时校验：条码不能重复
  if(editId === null){
    const exist = stock.some(x=>x.code === code);
    if(exist){
      alert("⚠️该条码已经存在！Codice già presente!");
      return;
    }
    stock.push({id:Date.now(),code,brand,name,spec,qty,min,img})
  }else{
    const idx = findIndexById(editId);
    stock[idx] = {...stock[idx],code,brand,name,spec,qty,min,img}
  }
  saveStorage();
  closeModal();
  render();
}

function delItem(id){
  if(!confirm(txt[lang].confirm)) return;
  stock = stock.filter(x=>x.id!==id);
  saveStorage();render();
}

// 入库
function stockIn(){
  const code = document.getElementById("inCode").value.trim();
  const qty = Number(document.getElementById("inQty").value);
  if(!code || qty<=0) return alert("填写条码和数量");
  const it = findByCode(code);
  if(!it) return alert("该条码不存在，先新增商品 / Codice non trovato");
  it.qty += qty;
  saveStorage();
  alert("入库成功！");
  document.getElementById("inCode").value="";
  document.getElementById("inQty").value="1";
  render();
}
// 出库
function stockOut(){
  const code = document.getElementById("outCode").value.trim();
  const qty = Number(document.getElementById("outQty").value);
  if(!code || qty<=0) return alert("填写条码和数量");
  const it = findByCode(code);
  if(!it) return alert("该条码不存在 / Codice non trovato");
  if(it.qty < qty) return alert("库存不足！ Giacenza insufficiente");
  it.qty -= qty;
  saveStorage();
  alert("出库成功！");
  document.getElementById("outCode").value="";
  document.getElementById("outQty").value="1";
  render();
}

// 查询（精准条码匹配）
function searchItem(){
  const code = document.getElementById("searchCode").value.trim();
  const wrap = document.getElementById("searchResult");
  wrap.innerHTML="";
  if(!code) return;
  const it = findByCode(code);
  if(!it){ wrap.innerHTML="<p class='text-red-500'>无此商品 / Articolo non trovato</p>";return; }
  wrap.innerHTML = itemHtml(it);
}

// 渲染单商品卡片
function itemHtml(it){
  return `
  <div class="border rounded p-3 mb-2">
    ${it.img ? `<img src="${it.img}" class="h-24 mb-2 object-contain">` : ""}
    <div><b>${it.brand}</b></div>
    <div>${it.name} ${it.spec}</div>
    <div class="text-sm text-gray-600">条码:${it.code}</div>
    <div class="text-sm">库存:${it.qty} | 预警:${it.min}</div>
    <div class="flex gap-2 mt-2">
      <button onclick="openEdit(${it.id})" class="text-blue-500 text-sm">${txt[lang].edit}</button>
      <button onclick="delItem(${it.id})" class="text-red-500 text-sm">${txt[lang].del}</button>
    </div>
  </div>
  `
}

// 主渲染函数
function render(){
  const searchKey = document.getElementById("stockSearch")?.value.trim() || "";
  let list = [...stock];

  // 库存页面：支持名称/条码搜索
  if(currentTab === "stock" && searchKey){
    list = list.filter(x=>x.code.includes(searchKey) || x.name.includes(searchKey) || x.brand.includes(searchKey))
  }
  if(currentTab === "stock"){
    document.getElementById("stockList").innerHTML = list.length ? list.map(itemHtml).join("") : "<p class='text-center text-gray-400'>暂无商品</p>";
  }

  // ✅盘点清单：按条码升序排序
  if(currentTab === "check"){
    let sortedList = [...stock].sort((a,b)=> a.code.localeCompare(b.code));
    let html = "";
    sortedList.forEach((it,idx)=>{
      html += `
      <div class="border-b py-2">
        <div>#${idx+1}</div>
        <div><b>${it.brand}</b></div>
        <div>${it.name} ${it.spec}</div>
        <div class="text-sm text-gray-500">条码 Codice: ${it.code}</div>
        <div class="text-sm">库存 Giacenza: ${it.qty}</div>
      </div>
      `
    })
    document.getElementById("checkList").innerHTML = html || "<p class='text-gray-400'>暂无库存</p>";
  }

  // 低库存
  if(currentTab === "low"){
    const low = stock.filter(x=>x.qty <= x.min);
    document.getElementById("lowList").innerHTML = low.length ? low.map(itemHtml).join("") : "<p class='text-center text-gray-400'>没有低库存商品 ✅</p>";
  }
}

// 扫码
function scanCode(inputId){
  alert("打开相机扫码，识别后自动填入；苹果Safari注意开启相机权限");
  const inputDom = document.getElementById(inputId);
  const code = prompt("手动输入条码 / Inserisci codice:");
  if(code) inputDom.value = code.trim();
}

// Excel导出盘点（导出文件同样按条码排序）
function exportCheckExcel(){
  let sortedList = [...stock].sort((a,b)=> a.code.localeCompare(b.code));
  const data = sortedList.map(it=>({
    "#":"",
    "品牌/Brand":it.brand,
    "条码/Codice":it.code,
    "商品名称/Nome prodotto":it.name + " " + it.spec,
    "库存/Giacenza":it.qty
  }))
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"盘点清单");
  XLSX.writeFile(wb,`盘点清单_${new Date().toLocaleDateString()}.xlsx`)
}
// 全部导出
function exportAll(){
  const ws = XLSX.utils.json_to_sheet(stock);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"AllStock");
  XLSX.writeFile(wb,`全部库存_${new Date().toLocaleDateString()}.xlsx`)
}
// 导入
function importExcel(){ document.getElementById("importFile").click(); }
function doImport(){
  const f = document.getElementById("importFile").files[0];
  const r = new FileReader();
  r.onload = e=>{
    const wb = XLSX.read(e.target.result,{type:"binary"});
    const arr = XLSX.utils.sheet_to_json(wb.SheetNames[0]);
    stock = arr.map(x=>({
      id:Date.now()+Math.random(),
      code:x.code||"",brand:x.brand||"",name:x.name||"",spec:x.spec||"",
      qty:Number(x.qty||0),min:Number(x.min||2),img:x.img||""
    }))
    saveStorage();render();alert("导入完成");
  }
  r.readAsBinaryString(f);
}
function clearAll(){
  if(!confirm(txt[lang].confirm))return;
  stock=[];saveStorage();render();
}

render();
</script>
</body>
</html>
