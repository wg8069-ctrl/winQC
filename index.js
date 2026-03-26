<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>偉剛異常通報</title>
<script src="https://static.line-scdn.net/liff/edge/versions/2.22.3/sdk.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:-apple-system,BlinkMacSystemFont,'Helvetica Neue',sans-serif;background:#f2f2f2;max-width:480px;margin:0 auto;min-height:100vh;}
.header{background:#00b900;padding:12px 16px;position:sticky;top:0;z-index:10;}
.header-title{color:#fff;font-size:16px;font-weight:600;}
.header-sub{color:#c8ffc8;font-size:12px;margin-top:2px;}
.progress{background:#009900;padding:8px 16px 10px;display:flex;}
.ps{flex:1;text-align:center;font-size:10px;color:#c8ffc8;}
.ps.active{color:#fff;font-weight:600;}
.ps.done{color:#90ee90;}
.pd{width:8px;height:8px;border-radius:50%;background:#006600;margin:0 auto 3px;}
.ps.active .pd{background:#fff;}
.ps.done .pd{background:#90ee90;}
.card{background:#fff;margin:12px;border-radius:14px;padding:16px;border:0.5px solid #e8e8e8;}
.card+.card{margin-top:0;}
.card-title{font-size:13px;color:#555;margin-bottom:12px;font-weight:600;}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:8px;}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;}
.g4{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:6px;}
.sb{padding:12px 6px;border:1.5px solid #e0e0e0;border-radius:10px;background:#fff;font-size:13px;color:#333;cursor:pointer;text-align:center;line-height:1.4;-webkit-tap-highlight-color:transparent;transition:all .15s;}
.sb:active{transform:scale(0.96);}
.sb.on{border-color:#00b900;background:#e8ffe8;color:#006600;font-weight:600;}
.sb.on-red{border-color:#e24b4a;background:#ffeaea;color:#a32d2d;font-weight:600;}
.sb.on-amber{border-color:#ba7517;background:#fff3cd;color:#633806;font-weight:600;}
.rb{padding:10px 4px;border:1.5px solid #e0e0e0;border-radius:8px;background:#fff;font-size:14px;font-weight:600;color:#333;cursor:pointer;text-align:center;-webkit-tap-highlight-color:transparent;transition:all .15s;}
.rb:active{transform:scale(0.96);}
.rb.on{border-color:#00b900;background:#e8ffe8;color:#006600;}
.ratio-result{background:#e8ffe8;border:1px solid #00b900;border-radius:8px;padding:10px;font-size:15px;color:#006600;font-weight:600;text-align:center;margin-top:10px;display:none;}
.photo-zone{border:2px dashed #ccc;border-radius:12px;padding:28px 16px;text-align:center;cursor:pointer;-webkit-tap-highlight-color:transparent;}
.photo-zone.on{border-color:#00b900;background:#f5fff5;}
.pi{font-size:36px;margin-bottom:8px;}
.pt{font-size:14px;color:#888;}
.photo-zone.on .pt{color:#006600;}
.nav{display:flex;gap:8px;margin:0 12px 16px;}
.btn-next{flex:1;background:#00b900;color:#fff;border:none;border-radius:12px;padding:16px;font-size:16px;font-weight:600;cursor:pointer;-webkit-tap-highlight-color:transparent;}
.btn-next:active{opacity:.85;}
.btn-next:disabled{background:#ccc;cursor:default;}
.btn-back{background:#fff;color:#555;border:1.5px solid #ddd;border-radius:12px;padding:16px;font-size:15px;cursor:pointer;white-space:nowrap;-webkit-tap-highlight-color:transparent;}
.skip{text-align:center;font-size:13px;color:#aaa;text-decoration:underline;cursor:pointer;margin:-8px 12px 16px;}
.step{display:none;}
.step.active{display:block;}
.tinput{width:100%;border:1.5px solid #ddd;border-radius:8px;padding:12px;font-size:16px;outline:none;margin-bottom:6px;}
.tinput:focus{border-color:#00b900;}
.recent{display:flex;flex-wrap:wrap;gap:6px;margin-top:6px;}
.rbtn{padding:6px 12px;border:1px solid #ddd;border-radius:14px;font-size:12px;color:#555;cursor:pointer;background:#f9f9f9;-webkit-tap-highlight-color:transparent;}
.rbtn:active{background:#e8ffe8;border-color:#00b900;}
.hint{font-size:11px;color:#aaa;margin-top:6px;}
.ratio-wrap{display:flex;gap:10px;align-items:flex-start;}
.ratio-col{flex:1;}
.ratio-label{font-size:12px;color:#888;margin-bottom:6px;font-weight:600;}
.ratio-div{font-size:28px;color:#bbb;padding-top:24px;}
.stbl{width:100%;font-size:14px;border-collapse:collapse;}
.stbl td{padding:8px 0;border-bottom:0.5px solid #f0f0f0;vertical-align:top;}
.stbl .lb{color:#888;width:85px;font-size:13px;}
.stbl .vl{color:#222;font-weight:600;}
.tag{display:inline-block;padding:3px 10px;border-radius:10px;font-size:12px;}
.tg-red{background:#ffeaea;color:#c0392b;}
.tg-amber{background:#fff3cd;color:#856404;}
.tg-green{background:#e8ffe8;color:#006600;}
.done{text-align:center;padding:40px 20px;}
.done .di{font-size:60px;margin-bottom:16px;}
.done .dt{font-size:20px;font-weight:600;color:#006600;margin-bottom:6px;}
.done .dn{font-size:24px;font-weight:600;color:#00b900;margin:12px 0;}
.done .ds{font-size:14px;color:#888;}
.nbox{background:#f0fff0;border:1px solid #c8ecc8;border-radius:12px;padding:14px;margin-top:20px;font-size:14px;color:#444;text-align:left;line-height:1.8;}
.loading{text-align:center;padding:40px;font-size:15px;color:#888;}
.err{background:#ffeaea;border:1px solid #e24b4a;border-radius:10px;padding:12px;margin:12px;font-size:14px;color:#a32d2d;display:none;}
</style>
</head>
<body>

<div class="header">
  <div class="header-title">偉剛異常通報</div>
  <div class="header-sub" id="hstep">步驟 1 / 6</div>
</div>
<div class="progress">
  <div class="ps active" id="ps1"><div class="pd"></div>單位</div>
  <div class="ps" id="ps2"><div class="pd"></div>異常</div>
  <div class="ps" id="ps3"><div class="pd"></div>比例</div>
  <div class="ps" id="ps4"><div class="pd"></div>照片</div>
  <div class="ps" id="ps5"><div class="pd"></div>判定</div>
  <div class="ps" id="ps6"><div class="pd"></div>確認</div>
</div>

<div class="err" id="errmsg"></div>

<!-- Step 1 -->
<div class="step active" id="s1">
  <div class="card">
    <div class="card-title">發生單位</div>
    <div class="g2" id="unit-g">
      <button class="sb" onclick="pick(this,'unit-g','unit')">🏭 本廠</button>
      <button class="sb" onclick="pick(this,'unit-g','unit')">🏭 二廠</button>
      <button class="sb" onclick="pick(this,'unit-g','unit')">🔍 品保</button>
      <button class="sb" onclick="pick(this,'unit-g','unit')">📦 倉管收料</button>
      <button class="sb" onclick="pick(this,'unit-g','unit')">🚚 外包</button>
      <button class="sb" onclick="pickOther(this,'unit-g','unit','unit-other-in')">✏️ 其他</button>
    </div>
    <input class="tinput" id="unit-other-in" placeholder="請輸入發生單位" style="display:none;margin-top:8px;" oninput="D.unit=this.value.trim();chk1()">
  </div>
  <div class="card">
    <div class="card-title">責任單位</div>
    <div class="g3" id="resp-g">
      <button class="sb" onclick="pick(this,'resp-g','resp')">CR<br><small style="color:#888;font-size:11px">成型</small></button>
      <button class="sb" onclick="pick(this,'resp-g','resp')">AS<br><small style="color:#888;font-size:11px">組裝</small></button>
      <button class="sb" onclick="pick(this,'resp-g','resp')">QC<br><small style="color:#888;font-size:11px">品保</small></button>
      <button class="sb" onclick="pick(this,'resp-g','resp')">WH<br><small style="color:#888;font-size:11px">倉庫</small></button>
      <button class="sb" onclick="pick(this,'resp-g','resp')">供應商</button>
      <button class="sb" onclick="pickOther(this,'resp-g','resp','resp-other-in')">✏️ 其他</button>
    </div>
    <input class="tinput" id="resp-other-in" placeholder="請輸入責任單位" style="display:none;margin-top:8px;" oninput="D.resp=this.value.trim();chk1()">
  </div>
  <div class="nav">
    <button class="btn-next" id="n1" onclick="go(2)" disabled>下一步 →</button>
  </div>
</div>

<!-- Step 2 -->
<div class="step" id="s2">
  <div class="card">
    <div class="card-title">異常狀況（可多選）</div>
    <div class="g2" id="anom-g">
      <button class="sb" onclick="multi(this)">外觀不良<br><small style="color:#888;font-size:11px">刮傷/髒污</small></button>
      <button class="sb" onclick="multi(this)">斷差<br><small style="color:#888;font-size:11px">段差/錯位</small></button>
      <button class="sb" onclick="multi(this)">尺寸異常<br><small style="color:#888;font-size:11px">超出公差</small></button>
      <button class="sb" onclick="multi(this)">組裝困難<br><small style="color:#888;font-size:11px">鎖不緊/卡住</small></button>
      <button class="sb" onclick="multi(this)">功能失效<br><small style="color:#888;font-size:11px">測試不通過</small></button>
      <button class="sb" onclick="multi(this)">來料不良<br><small style="color:#888;font-size:11px">進料異常</small></button>
      <button class="sb" onclick="multi(this)">標示錯誤</button>
      <button class="sb" onclick="multiOther(this)">✏️ 其他</button>
    </div>
    <input class="tinput" id="anom-other-in" placeholder="請輸入異常狀況說明" style="display:none;margin-top:8px;" oninput="updateAnomOther(this.value)">
  </div>
  <div class="card">
    <div class="card-title">品名 <span style="color:#aaa;font-size:11px;font-weight:400">（不知道可以不填）</span></div>
    <input class="tinput" id="prod-in" placeholder="例：WC4-795B-CR（可略過）" oninput="D.prod=this.value.trim();chk2()">
    <div class="recent" id="recent-list"></div>
    <div class="hint">最近使用品名，點一下帶入</div>
  </div>
  <div class="nav">
    <button class="btn-back" onclick="go(1)">← 返回</button>
    <button class="btn-next" id="n2" onclick="go(3)" disabled>下一步 →</button>
  </div>
</div>

<!-- Step 3 -->
<div class="step" id="s3">
  <div class="card">
    <div class="card-title">訂單數量</div>
    <div class="g4" id="qty-g">
      <button class="rb" onclick="pickR(this,'qty-g','qty')">200</button>
      <button class="rb" onclick="pickR(this,'qty-g','qty')">500</button>
      <button class="rb" onclick="pickR(this,'qty-g','qty')">1000</button>
      <button class="rb" onclick="pickR(this,'qty-g','qty')">1200</button>
      <button class="rb" onclick="pickR(this,'qty-g','qty')">2000</button>
      <button class="rb" onclick="pickR(this,'qty-g','qty')">3000</button>
      <button class="rb" onclick="pickR(this,'qty-g','qty')">5000</button>
      <button class="rb" id="qty-other" onclick="showQtyInput()">其他</button>
    </div>
    <input class="tinput" id="qty-in" placeholder="輸入數量" style="display:none;margin-top:8px;" type="number" oninput="D.qty=this.value;chk3()">
  </div>
  <div class="card">
    <div class="card-title">不良比例</div>
    <div class="ratio-wrap">
      <div class="ratio-col">
        <div class="ratio-label">抽驗數</div>
        <div class="g4" id="samp-g">
          <button class="rb" onclick="pickR(this,'samp-g','samp')">3</button>
          <button class="rb" onclick="pickR(this,'samp-g','samp')">5</button>
          <button class="rb" onclick="pickR(this,'samp-g','samp')">10</button>
          <button class="rb" onclick="pickR(this,'samp-g','samp')">20</button>
          <button class="rb" onclick="pickR(this,'samp-g','samp')">50</button>
          <button class="rb" onclick="pickR(this,'samp-g','samp')">100</button>
          <button class="rb" onclick="pickR(this,'samp-g','samp')">200</button>
          <button class="rb" onclick="pickR(this,'samp-g','samp')">全</button>
        </div>
      </div>
      <div class="ratio-div">/</div>
      <div class="ratio-col">
        <div class="ratio-label">不良數</div>
        <div class="g4" id="bad-g">
          <button class="rb" onclick="pickR(this,'bad-g','bad')">1</button>
          <button class="rb" onclick="pickR(this,'bad-g','bad')">2</button>
          <button class="rb" onclick="pickR(this,'bad-g','bad')">3</button>
          <button class="rb" onclick="pickR(this,'bad-g','bad')">5</button>
          <button class="rb" onclick="pickR(this,'bad-g','bad')">10</button>
          <button class="rb" onclick="pickR(this,'bad-g','bad')">20</button>
          <button class="rb" onclick="pickR(this,'bad-g','bad')">50</button>
          <button class="rb" onclick="pickR(this,'bad-g','bad')">全</button>
        </div>
      </div>
    </div>
    <div class="ratio-result" id="ratio-res"></div>
  </div>
  <div class="nav">
    <button class="btn-back" onclick="go(2)">← 返回</button>
    <button class="btn-next" id="n3" onclick="go(4)" disabled>下一步 →</button>
  </div>
</div>

<!-- Step 4 -->
<div class="step" id="s4">
  <div class="card">
    <div class="card-title">異常照片</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:8px;">
      <button class="sb" style="padding:16px 6px;" onclick="document.getElementById('pf-camera').click()">📷<br><br>拍照</button>
      <button class="sb" style="padding:16px 6px;" onclick="document.getElementById('pf-gallery').click()">🖼️<br><br>從圖庫選</button>
    </div>
    <input type="file" id="pf-camera" accept="image/*" capture="environment" style="display:none" onchange="handlePic(this)">
    <input type="file" id="pf-gallery" accept="image/*" style="display:none" onchange="handlePic(this)">
    <img id="pp" style="width:100%;border-radius:8px;margin-top:10px;display:none;max-height:300px;object-fit:cover;">
  </div>
  <div class="nav">
    <button class="btn-back" onclick="go(3)">← 返回</button>
    <button class="btn-next" id="n4" onclick="go(5)" disabled>下一步 →</button>
  </div>
  <div class="skip" onclick="D.photo=false;D.photoData=null;document.getElementById('n4').disabled=false;go(5)">沒有照片，先跳過</div>
</div>

<!-- Step 5 -->
<div class="step" id="s5">
  <div class="card">
    <div class="card-title">品質判定</div>
    <div class="g3" id="judge-g">
      <button class="sb" style="padding:20px 6px;font-size:15px;" onclick="judgeBtn(this,'驗退X','on-red')">❌<br><br>驗退 X</button>
      <button class="sb" style="padding:20px 6px;font-size:15px;" onclick="judgeBtn(this,'特採△','on-amber')">⚠️<br><br>特採 △</button>
      <button class="sb" style="padding:20px 6px;font-size:15px;" onclick="judgeBtn(this,'加工○','on')">🔧<br><br>加工 ○</button>
    </div>
  </div>
  <div class="nav">
    <button class="btn-back" onclick="go(4)">← 返回</button>
    <button class="btn-next" id="n5" onclick="go(6)" disabled>下一步 →</button>
  </div>
</div>

<!-- Step 6 -->
<div class="step" id="s6">
  <div class="card">
    <div class="card-title">確認後送出</div>
    <table class="stbl" id="stbl"></table>
  </div>
  <div class="nav">
    <button class="btn-back" onclick="go(5)">← 修改</button>
    <button class="btn-next" id="n6" onclick="submitForm()">送出異常單 ✓</button>
  </div>
</div>

<!-- Done -->
<div class="step" id="s7">
  <div class="done">
    <div class="di">✅</div>
    <div class="dt">異常單建立完成！</div>
    <div class="dn" id="dnum"></div>
    <div class="ds">已自動通知品保與主管</div>
    <div class="nbox" id="nbox"></div>
    <button class="btn-next" style="margin-top:24px;max-width:200px;" onclick="closeLiff()">關閉</button>
  </div>
</div>

<script>
// ========== 設定區 ==========
// 部署後把這個換成你的 Render API 網址
var API_URL = 'https://YOUR-APP.onrender.com';
var LIFF_ID = 'YOUR_LIFF_ID';
// ============================

var D = {unit:'',resp:'',anom:[],prod:'',qty:'',samp:'',bad:'',photo:false,photoData:null,judge:''};
var cur = 1;
var userId = '';
var recentProds = JSON.parse(localStorage.getItem('recentProds')||'["WC4-795B-CR","WC4-800A","CB26-001"]');

liff.init({liffId: LIFF_ID}).then(function(){
  if(liff.isLoggedIn()){
    liff.getProfile().then(function(p){ userId = p.userId; });
  }
}).catch(function(){ console.log('LIFF init failed, running in browser mode'); });

renderRecent();

function renderRecent(){
  var el = document.getElementById('recent-list');
  el.innerHTML = '';
  recentProds.slice(0,5).forEach(function(p){
    var b = document.createElement('button');
    b.className = 'rbtn';
    b.textContent = p;
    b.onclick = function(){ setProd(p); };
    el.appendChild(b);
  });
}

function go(n){
  document.getElementById('s'+cur).classList.remove('active');
  cur = n;
  document.getElementById('s'+cur).classList.add('active');
  document.getElementById('hstep').textContent = n<=6 ? '步驟 '+n+' / 6' : '完成！';
  for(var i=1;i<=6;i++){
    var el = document.getElementById('ps'+i);
    el.className = 'ps'+(i===n?' active':i<n?' done':'');
  }
  if(n===6) buildSummary();
  window.scrollTo(0,0);
}

function pick(btn, gridId, key){
  document.getElementById(gridId).querySelectorAll('.sb').forEach(function(b){ b.className='sb'; });
  btn.classList.add('on');
  D[key] = btn.textContent.trim().replace(/\n.*/,'').replace(/[🏭🔍📦🚚⚙️✏️]\s*/,'');
  // 隱藏對應「其他」輸入框
  var inId = gridId.replace('-g','-other-in');
  var inp = document.getElementById(inId);
  if(inp){ inp.style.display='none'; inp.value=''; }
  if(key==='unit'||key==='resp') chk1();
}

function pickOther(btn, gridId, key, inputId){
  document.getElementById(gridId).querySelectorAll('.sb').forEach(function(b){ b.className='sb'; });
  btn.classList.add('on');
  D[key] = '';
  var inp = document.getElementById(inputId);
  inp.style.display = 'block';
  inp.focus();
  if(key==='unit'||key==='resp') chk1();
}

function chk1(){ document.getElementById('n1').disabled = !(D.unit && D.resp); }

function multi(btn){
  btn.classList.toggle('on');
  var val = btn.textContent.trim().replace(/\n.*/,'');
  var idx = D.anom.indexOf(val);
  if(btn.classList.contains('on')){ if(idx===-1) D.anom.push(val); }
  else { if(idx>-1) D.anom.splice(idx,1); }
  chk2();
}

function multiOther(btn){
  var isOn = btn.classList.contains('on');
  btn.classList.toggle('on', !isOn);
  var inp = document.getElementById('anom-other-in');
  if(!isOn){
    inp.style.display = 'block';
    inp.focus();
    // 清掉舊值
    if(D._anomOther){ var i=D.anom.indexOf(D._anomOther); if(i>-1)D.anom.splice(i,1); }
    D._anomOther = '';
  } else {
    inp.style.display = 'none';
    inp.value = '';
    if(D._anomOther){ var i=D.anom.indexOf(D._anomOther); if(i>-1)D.anom.splice(i,1); }
    D._anomOther = '';
    chk2();
  }
}

function updateAnomOther(val){
  // 移除舊的自訂值，換新的
  if(D._anomOther){ var i=D.anom.indexOf(D._anomOther); if(i>-1)D.anom.splice(i,1); }
  D._anomOther = val.trim();
  if(D._anomOther) D.anom.push(D._anomOther);
  chk2();
}

function setProd(v){
  document.getElementById('prod-in').value = v;
  D.prod = v;
  chk2();
}

function chk2(){
  D.prod = document.getElementById('prod-in').value.trim();
  document.getElementById('n2').disabled = !(D.anom.length > 0);
}

function pickR(btn, gridId, key){
  document.getElementById(gridId).querySelectorAll('.rb').forEach(function(b){ b.classList.remove('on'); });
  btn.classList.add('on');
  D[key] = btn.textContent.trim();
  updateRatio();
}

function showQtyInput(){
  document.getElementById('qty-in').style.display = 'block';
  document.getElementById('qty-in').focus();
}

function updateRatio(){
  var r = document.getElementById('ratio-res');
  if(D.samp && D.bad){
    r.style.display = 'block';
    r.textContent = '不良比例：' + D.samp + ' / ' + D.bad;
  }
  chk3();
}

function chk3(){ document.getElementById('n3').disabled = !(D.qty && D.samp && D.bad); }

function handlePic(inp){
  if(inp.files && inp.files[0]){
    var reader = new FileReader();
    reader.onload = function(e){
      var pp = document.getElementById('pp');
      pp.src = e.target.result;
      pp.style.display = 'block';
      D.photo = true;
      D.photoData = e.target.result;
      document.getElementById('n4').disabled = false;
    };
    reader.readAsDataURL(inp.files[0]);
  }
}

function judgeBtn(btn, val, cls){
  document.getElementById('judge-g').querySelectorAll('.sb').forEach(function(b){ b.className='sb'; b.style.padding='20px 6px'; b.style.fontSize='15px'; });
  btn.classList.add(cls);
  D.judge = val;
  document.getElementById('n5').disabled = false;
}

function genNum(){
  var t = new Date();
  return 'CB26-'+(t.getMonth()+1).toString().padStart(2,'0')+t.getDate().toString().padStart(2,'0')+'-'+(Math.floor(Math.random()*900)+100);
}

function buildSummary(){
  var t = new Date();
  var d = t.getFullYear()+'/'+(t.getMonth()+1).toString().padStart(2,'0')+'/'+t.getDate().toString().padStart(2,'0');
  D._num = genNum(); D._date = d;
  var jhtml = D.judge==='驗退'?'<span class="tag tg-red">驗退</span>':D.judge==='特採'?'<span class="tag tg-amber">特採</span>':'<span class="tag tg-green">加工</span>';
  var phtml = D.photo?'<span style="color:#006600">已附照片 ✓</span>':'<span style="color:#aaa">未附照片</span>';
  document.getElementById('stbl').innerHTML =
    row('單號',D._num)+row('日期',d)+row('發生單位',D.unit)+row('責任單位',D.resp)+
    row('品名',D.prod)+row('異常狀況',D.anom.join('、'))+
    row('訂單數量',D.qty)+row('不良比例',D.samp+' / '+D.bad)+
    row('品質判定',jhtml)+row('照片',phtml);
}

function row(l,v){ return '<tr><td class="lb">'+l+'</td><td class="vl">'+v+'</td></tr>'; }

function showErr(msg){
  var e = document.getElementById('errmsg');
  e.textContent = msg;
  e.style.display = 'block';
  setTimeout(function(){ e.style.display='none'; }, 5000);
}

function submitForm(){
  var btn = document.getElementById('n6');
  btn.disabled = true;
  btn.textContent = '送出中...';

  // 儲存最近品名
  var idx = recentProds.indexOf(D.prod);
  if(idx > -1) recentProds.splice(idx,1);
  recentProds.unshift(D.prod);
  recentProds = recentProds.slice(0,8);
  localStorage.setItem('recentProds', JSON.stringify(recentProds));

  var payload = {
    number: D._num,
    date: D._date,
    unit: D.unit,
    resp: D.resp,
    product: D.prod,
    anomaly: D.anom.join('、'),
    qty: D.qty,
    ratio: D.samp + ' / ' + D.bad,
    judge: D.judge,
    photoData: D.photoData,
    userId: userId
  };

  fetch(API_URL + '/api/anomaly', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(payload)
  }).then(function(r){ return r.json(); })
  .then(function(res){
    if(res.success){
      go(7);
      document.getElementById('dnum').textContent = res.number || D._num;
      document.getElementById('nbox').innerHTML =
        '<b>回報人：</b>' + (res.reporter || '(未知)') + '<br>' +
        '<b>已通知：</b>品保主管、' + D.unit + ' 主管<br>' +
        '<b>品名：</b>' + (D.prod || '(未填)') + '<br>' +
        '<b>異常：</b>' + D.anom.join('、') + '<br>' +
        '<b>比例：</b>' + D.samp + ' / ' + D.bad + '<br>' +
        '<b>判定：</b>' + D.judge;
    } else {
      showErr('送出失敗：' + (res.error||'請重試'));
      btn.disabled = false;
      btn.textContent = '送出異常單 ✓';
    }
  }).catch(function(e){
    showErr('網路錯誤，請確認網路連線後重試');
    btn.disabled = false;
    btn.textContent = '送出異常單 ✓';
  });
}

function closeLiff(){
  try{ liff.closeWindow(); } catch(e){ window.close(); }
}
</script>
</body>
</html>
