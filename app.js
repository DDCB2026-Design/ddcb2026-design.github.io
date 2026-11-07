
(function(){
"use strict";
var FONT_LINK='<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&family=Nunito+Sans:wght@400;600;700;800&display=swap">';
var XLSX_URL='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
var RESET=':root{color-scheme:light;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0;font:14px system-ui,sans-serif}img{max-width:100%}[hidden]{display:none!important}';
var SC='<'+'/script>';

function $(s,r){return (r||document).querySelector(s)}
function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]})}
function r2(n){return Math.round((Number(n)||0)*100)/100}
function money(n){return 'S/ '+(Number(n)||0).toFixed(2)}
function slug(s){return String(s||'').normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'').slice(0,60)||'x'}
function num(v){if(typeof v==='number')return v;var s=String(v||'').replace(/S\/\.?|s\/\.?|\$|\s/g,'');if(/,\d{1,2}$/.test(s)&&s.indexOf('.')<0)s=s.replace(',','.');s=s.replace(/,/g,'').replace(/^\./,'');var n=parseFloat(s);return isNaN(n)?null:n}
function norm(s){return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim()}
function gLabel(g){g=Number(g);return g>=1000?(g/1000).toString().replace('.',',')+' kg':g+' g'}
function sentence(s){s=String(s||'').trim().toLowerCase().replace(/\s+/g,' ');return s.charAt(0).toUpperCase()+s.slice(1)}

var DATA;
try{DATA=JSON.parse($('#store-data').textContent)}catch(e){DATA={settings:{},products:[]}}
if(window.ABAS_CONFIG){DATA.settings=Object.assign({},DATA.settings||{},window.ABAS_CONFIG.settings||{})}
var S=DATA.settings=Object.assign({storeName:'Tienda Frutos Secos',tagline:'',whatsapp:'',yapeNumber:'',yapeName:'',yapeQR:'',delivery:{mode:'flat',flat:15,districts:[]},pickup:false,pickupAddress:'',minOrder:0,heroImg:'',logo:'',catImgs:{},catOrder:[]},DATA.settings||{});
var PRODUCTS=(DATA.products||[]);
var byId={};PRODUCTS.forEach(function(p){byId[p.id]=p});

var cart={};
try{cart=JSON.parse(localStorage.getItem('fs-cart2')||'{}')||{}}catch(e){cart={}}
Object.keys(cart).forEach(function(k){if(!variantOf(k))delete cart[k]});
function saveCart(){try{localStorage.setItem('fs-cart2',JSON.stringify(cart))}catch(e){}}
function variantOf(key){var i=key.lastIndexOf('@'),p=byId[key.slice(0,i)],g=Number(key.slice(i+1));if(!p)return null;var v=(p.v||[]).find(function(x){return x[0]===g});return v?{p:p,g:g,price:v[1]}:null}

var st={view:'store',cat:'Todas',q:'',sel:{},drawer:false,step:'cart',isAdmin:false,tab:'productos',order:null,errors:{},
  form:{name:'',phone:'',method:'delivery',district:'',address:'',ref:'',when:'',notes:'',op:''}};
try{var f=JSON.parse(localStorage.getItem('fs-form')||'null');if(f)Object.assign(st.form,f)}catch(e){}
st.form.method='delivery';

var app=$('#app');

/* ---------- pricing ---------- */
var PRICING={mode:'markup',round:'0.5',defaultMargin:15,pres:{},presDefault:[250,500,1000],pack:{}};
function presOf(c){if(c.pres&&c.pres.length)return c.pres;return PRICING.pres[c.cat]||PRICING.presDefault||[250,500,1000]}
function roundP(raw){
  if(PRICING.round==='90')return r2(Math.ceil(raw-0.9-1e-9)+0.9);
  var step=Number(PRICING.round)||0.5;return r2(Math.ceil(raw/step-1e-9)*step);
}
function priceFor(costKg,g,margin){
  var m=(margin==null||margin===''?Number(PRICING.defaultMargin):Number(margin))/100;
  var base=(Number(costKg)||0)*g/1000+(Number((PRICING.pack||{})[g])||0);
  var raw=PRICING.mode==='margin'?(m>=0.95?base*20:base/(1-m)):base*(1+m);
  return roundP(raw);
}
function effMargin(row){if(row.margin!=null&&row.margin!=='')return Number(row.margin);var s=SUP[slug(row.supplier)];if(s&&s.margin!=null&&s.margin!=='')return Number(s.margin);return Number(PRICING.defaultMargin)}

/* ---------- delivery ---------- */
function deliveryCost(){
  if(st.form.method==='pickup')return 0;
  var d=S.delivery||{};
  if(d.mode==='district'){var x=(d.districts||[]).find(function(z){return z.name===st.form.district});return x?Number(x.cost):null}
  return Number(d.flat)||0;
}
function cartLines(){return Object.keys(cart).map(function(k){var v=variantOf(k);if(!v||!(cart[k]>0))return null;return {key:k,id:v.p.id,name:v.p.name,img:v.p.img,g:v.g,qty:cart[k],price:v.price,sub:r2(v.price*cart[k])}}).filter(Boolean)}
function cartCount(){return cartLines().reduce(function(a,l){return a+l.qty},0)}
function subtotal(){return r2(cartLines().reduce(function(a,l){return a+l.sub},0))}

/* ---------- render ---------- */
var PHS=['--ph1','--ph2','--ph3','--ph4'];
function hashIdx(s,n){var h=0;for(var i=0;i<s.length;i++)h=(h*31+s.charCodeAt(i))>>>0;return h%n}
function mono(name){var w=String(name).split(/\s+/).filter(function(x){return x.length>2});return ((w[0]||name)[0]+((w[1]||'')[0]||'')).toUpperCase()}
function cats(){var o=(S.catOrder||[]).slice();PRODUCTS.forEach(function(p){if(p.cat&&o.indexOf(p.cat)<0)o.push(p.cat)});return o}
function waLink(t){return S.whatsapp?'https://wa.me/'+String(S.whatsapp).replace(/\D/g,'')+(t?'?text='+encodeURIComponent(t):''):''}
var WA_SVG='<svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M20 11.5a8.5 8.5 0 0 1-12.6 7.4L3 20l1.2-4.2A8.5 8.5 0 1 1 20 11.5z"/><path d="M9.2 8.4c.3 2.7 2.5 5 5.2 5.3l1-1.3 1.9.9-.4 1.7c-3.9.3-8.2-3.7-8.1-7.8l1.6-.7.8 2z" fill="#fff" stroke-width="0.6"/></svg>';

function render(){
  var y=window.scrollY;
  app.innerHTML=header()+(st.view==='admin'&&st.isAdmin?adminView():storeView())+(st.view==='store'&&!st.drawer&&S.whatsapp?'<a class="wa'+(cartCount()?' up':'')+'" href="'+esc(waLink('Hola ABAS-TCT, quiero hacer una consulta.'))+'" target="_blank" rel="noopener" aria-label="Escríbenos por WhatsApp">'+WA_SVG+'<span>WhatsApp</span></a>':'')+(st.view==='store'&&!st.drawer&&cartCount()?cartBar():'')+(st.drawer?drawer():'')+(toastMsg?'<div class="toast" role="status">'+esc(toastMsg)+'</div>':'');
  if(st.view==='admin')afterAdminRender();
  window.scrollTo(0,y);
}
var PAL=['1','2','3','4','6'];
function catVars(c){var i=cats().indexOf(c);var k=PAL[i<0?4:Math.min(i,4)];return '--cc:var(--c'+k+');--ct:var(--t'+k+')'}
function header(){
  var n=cartCount();
  return '<header class="top"><div class="wrap">'+
   '<div class="logo">'+(S.logo?'<img src="'+esc(S.logo)+'" alt="'+esc(S.storeName)+'">':esc(S.storeName))+(st.view==='admin'?'<small>Panel de administración</small>':'')+'</div>'+
   (st.isAdmin?'<button class="btn sm ghost" data-act="toggleAdmin">'+(st.view==='admin'?'Ver tienda':'Administrar')+'</button>':'')+
   (st.view==='store'?'<button class="btn ghost sm" data-act="track">Mi pedido</button>':'')+(st.view==='store'?'<button class="btn navy" data-act="openCart" aria-label="Abrir carrito"><span class="lbl">Carrito</span><span class="badge num">'+n+'</span></button>':'')+
  '</div><div class="stripe" aria-hidden="true"><i></i><i></i><i></i><i></i></div></header>';
}
var IC={
 pay:'<span class="paychips" aria-hidden="true"><b class="py">Yape</b><b class="pl">Plin</b></span>',
 moto:'<svg class="ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="5.5" cy="17" r="3"/><circle cx="18.5" cy="17" r="3"/><path d="M8.5 17h5l3-6h-3.5"/><path d="M5.5 17l3-6h4l1.5 3"/><path d="M14 7h2.5l2 10"/><path d="M8 8.5h3"/></svg>',
 bag:'<svg class="ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 8h14l-1 12H6L5 8z"/><path d="M9 8V6a3 3 0 0 1 6 0v2"/><path d="M9.5 14.5l1.8 1.8 3.4-3.6"/></svg>'
};
function storeView(){
  var d=S.delivery||{};
  var delTxt=d.mode==='district'&&(d.districts||[]).length?'Delivery desde '+money(Math.min.apply(null,d.districts.map(function(z){return Number(z.cost)||0})))+' a toda Lima Metropolitana':'Delivery '+money(d.flat)+' a toda Lima Metropolitana';
  var cs=cats();
  var tiles='<div class="wrap tiles" role="group" aria-label="Categorías">'+['Todas'].concat(cs).map(function(c){
    var img=c==='Todas'?S.heroImg:(S.catImgs||{})[c],n=c==='Todas'?PRODUCTS.length:PRODUCTS.filter(function(p){return p.cat===c}).length;
    return '<button class="tile" style="'+(c==='Todas'?'--cc:var(--c5)':catVars(c))+'" data-act="cat" data-v="'+esc(c)+'" aria-pressed="'+(st.cat===c)+'"><span class="im">'+(img?'<img src="'+esc(img)+'" alt="" loading="lazy">':'')+'</span><span class="l">'+esc(c==='Todas'?'Todo el catálogo':c)+'<br><small style="font-family:var(--body);font-weight:600;font-size:12px;text-transform:none;opacity:.9">'+(n?n+' productos':'Próximamente')+'</small></span></button>'}).join('')+'</div>';
  return '<div class="searchbar"><div class="wrap"><input id="q" type="search" placeholder="Buscar pecanas, almendras, pasas…" value="'+esc(st.q)+'" aria-label="Buscar producto"></div></div>'+
   '<section class="wrap"><div class="hero"><div class="copy"><span class="eyebrow">Frutos secos · Deshidratados · Semillas · Especias</span>'+
    '<h2>'+esc(S.tagline||'Frutos secos frescos, directo a tu casa')+'</h2>'+
    '<p>Elige tus productos, paga con Yape o Plin y recibe tu pedido en Lima.</p>'+
    '<div class="perks"><span class="perk">'+IC.pay+'Paga con Yape o Plin</span><span class="perk">'+IC.moto+esc(delTxt)+'</span>'+(S.pickup?'<span class="perk">Recojo sin costo</span>':'')+(Number(S.minOrder)>0?'<span class="perk">'+IC.bag+'Pedido mínimo '+money(S.minOrder)+' (incluye delivery)</span>':'')+'</div></div>'+
    '<div class="pic">'+(S.heroImg?'<img src="'+esc(S.heroImg)+'" alt="Variedad de frutos secos">':'')+'</div></div>'+
    '<div class="strip"><div style="background:var(--c1)"><b>1</b>Elige</div><div style="background:var(--c2)"><b>2</b><span class="full">Paga con Yape o Plin</span><span class="short">Yape / Plin</span></div><div style="background:var(--c4)"><b>3</b><span class="full">Envía por WhatsApp</span><span class="short">WhatsApp</span></div><div style="background:var(--navy)"><b>4</b><span class="full">Recibe en casa</span><span class="short">Recibe</span></div></div></section>'+tiles+
   '<main class="wrap" id="list">'+listHtml()+'</main>'+
   '<footer class="foot"><div class="wrap"><b>'+esc(S.storeName)+'</b><span>Catálogo actualizado'+(DATA.publishedAt?' el '+esc(fmtDate(DATA.publishedAt)):'')+'. Precios en soles.</span>'+(S.whatsapp?'<span>Pedidos y consultas por WhatsApp: <a style="color:#fff;font-weight:800" href="'+esc(waLink('Hola ABAS-TCT'))+'" target="_blank" rel="noopener">'+esc(String(S.whatsapp).replace(/^51/,'').replace(/(\d{3})(\d{3})(\d{3})/,'$1 $2 $3'))+'</a></span>':'')+'</div></footer>';
}
function listHtml(){
  if(!PRODUCTS.length)return '<p class="empty">El catálogo se está preparando. Vuelve pronto.</p>';
  var q=norm(st.q);
  var match=function(p){return !q||norm(p.name+' '+p.cat+' '+(p.desc||'')).indexOf(q)>=0};
  var cs=st.cat==='Todas'?cats():[st.cat];
  var html=cs.map(function(c){
    var list=PRODUCTS.filter(function(p){return p.cat===c&&match(p)});
    if(!list.length&&!q&&!PRODUCTS.some(function(p){return p.cat===c}))return '<section class="sec" style="'+catVars(c)+'"><div class="sec-h"><h2>'+esc(c)+'</h2><span>Próximamente</span></div><div class="soon">'+((S.catImgs||{})[c]?'<img src="'+esc(S.catImgs[c])+'" alt="" loading="lazy">':'<span></span>')+'<div><h3>Muy pronto en la tienda</h3><p class="note" style="margin:0;font-size:15px">Estamos preparando nuestra línea de '+esc(c.toLowerCase())+'. Escríbenos para consultar disponibilidad y precios.</p>'+(S.whatsapp?'<a class="btn" href="'+esc(waLink('Hola ABAS-TCT, quiero consultar por '+c.toLowerCase()+'.'))+'" target="_blank" rel="noopener">Consultar por WhatsApp</a>':'')+'</div></div></section>';
    if(!list.length)return '';
    return '<section class="sec" style="'+catVars(c)+'"><div class="sec-h"><h2>'+esc(c)+'</h2><span class="num">'+list.length+' producto'+(list.length===1?'':'s')+'</span></div><div class="grid">'+list.map(card).join('')+'</div></section>';
  }).join('');
  return html||'<p class="empty">No encontramos productos con esa búsqueda.</p>';
}
function card(p){
  var vs=p.v||[];if(!vs.length)return '';
  var g=st.sel[p.id]||vs[0][0],v=vs.find(function(x){return x[0]===g})||vs[0];g=v[0];
  var key=p.id+'@'+g,q=cart[key]||0;
  var kgRef=vs.find(function(x){return x[0]===1000});
  var perkg=kgRef?money(kgRef[1])+' el kg':money(v[1]/g*1000)+' el kg';
  return '<article class="card"><div class="ph">'+
    (p.img?'<img src="'+esc(p.img)+'" alt="'+esc(p.name)+'" loading="lazy">':'<span class="mono" aria-hidden="true">'+esc(mono(p.name))+'</span>')+
    '<span class="kg num">'+perkg+'</span></div><div class="info"><h3>'+esc(p.name)+'</h3>'+(p.desc?'<div class="desc">'+esc(p.desc)+'</div>':'')+
    (vs.length>1?'<div class="pres" role="group" aria-label="Presentación">'+vs.map(function(x){return '<button data-act="sel" data-id="'+esc(p.id)+'" data-g="'+x[0]+'" aria-pressed="'+(x[0]===g)+'">'+gLabel(x[0])+'</button>'}).join('')+'</div>':'<div class="note">'+gLabel(g)+'</div>')+
    '<div class="buy"><span class="price num">'+money(v[1])+'</span>'+
    (q?'<span class="qty"><button data-act="dec" data-k="'+esc(key)+'" aria-label="Quitar uno">−</button><span class="num">'+q+'</span><button data-act="inc" data-k="'+esc(key)+'" aria-label="Agregar uno">+</button></span>'
      :'<button class="add" data-act="inc" data-k="'+esc(key)+'" aria-label="Agregar '+esc(p.name+' '+gLabel(g))+'">Agregar</button>')+
    '</div></div></article>';
}
function cartBar(){return '<div class="cartbar"><button data-act="openCart"><span>Ver mi pedido · <span class="num">'+cartCount()+'</span> producto'+(cartCount()===1?'':'s')+'</span><b class="num">'+money(subtotal())+'</b></button></div>'}
function fmtDate(iso){try{return new Date(iso).toLocaleDateString('es-PE',{day:'numeric',month:'long',year:'numeric'})}catch(e){return ''}}

/* ---------- drawer ---------- */
var CFG=window.ABAS_CONFIG||{};
var FRANJAS=CFG.franjas||['Mañana (9 a 12)','Tarde (12 a 4)','Noche (4 a 8)'];
var DIAS=['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
function limaNow(){var d=new Date();return new Date(d.getTime()+d.getTimezoneOffset()*60000-5*3600000)}
var RULES={anticipacion:1,sinReparto:'domingo'};
function deliveryDate(){var d=limaNow();d.setHours(12,0,0,0);d.setDate(d.getDate()+(Number(RULES.anticipacion)||0));var off=norm(RULES.sinReparto||'').split(' ');for(var i=0;i<7&&off.indexOf(norm(DIAS[d.getDay()]))>=0;i++)d.setDate(d.getDate()+1);return d}
function fmtDay(d){return DIAS[d.getDay()]+' '+String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0')}
function minDel(){var d=S.delivery||{};if(d.mode==='district'&&(d.districts||[]).length)return Math.min.apply(null,d.districts.map(function(z){return Number(z.cost)||0}));return Number(d.flat)||0}
function minMissing(){var m=Number(S.minOrder)||0;if(!m)return 0;var dc=deliveryCost();return r2(Math.max(0,m-(subtotal()+(dc!=null?dc:minDel()))))}
function drawer(){
  var title={cart:'Tu pedido',form:'Datos de entrega',yape:'Paga con Yape o Plin',track:'Consulta tu pedido'}[st.step]||'Tu pedido';
  var body='',foot='';
  var lines=cartLines(),sub=subtotal(),del=deliveryCost(),total=r2(sub+(del||0));
  if(st.step==='track'){
    var T=st.track||{};
    body='<p class="note" style="margin:0">Escribe el código de tu pedido (ej. FS-1106-4821) y los 4 últimos dígitos de tu celular.</p>'+
      '<div class="field"><label for="t-code">Código del pedido</label><input id="t-code" data-t="code" value="'+esc(T.code||'')+'" autocapitalize="characters" placeholder="FS-…"></div>'+
      '<div class="field"><label for="t-tel">Últimos 4 dígitos de tu celular</label><input id="t-tel" data-t="tel" inputmode="numeric" maxlength="4" value="'+esc(T.tel||'')+'"></div>'+
      (T.busy?'<p class="note">Buscando…</p>':'')+
      (T.err?'<p class="field"><span class="err">'+esc(T.err)+'</span></p>':'')+
      (T.res?'<div class="yape" style="border-color:var(--c3);text-align:left"><span class="eyebrow">Pedido '+esc(T.res.code)+'</span><div style="font-family:var(--display);font-size:22px;font-weight:800;color:var(--c3)">'+esc(T.res.estado)+'</div>'+
        '<div>Total: <b class="num">'+money(T.res.total)+'</b></div>'+
        (T.res.entrega?'<div>Entrega '+(T.res.confirmada?'confirmada':'estimada')+': <b>'+esc(T.res.entrega)+'</b>'+(T.res.franja?' · '+esc(T.res.franja):'')+'</div>':'')+
        (T.res.nota?'<div class="note">'+esc(T.res.nota)+'</div>':'')+'</div>':'');
    foot='<button class="btn primary" data-act="doTrack" '+(T.busy?'disabled':'')+'>Consultar</button>';
  } else if(st.step==='cart'){
    body=lines.length?lines.map(function(l){return '<div class="line"><div class="th">'+(l.img?'<img src="'+esc(l.img)+'" alt="">':'')+'</div><div><div class="n">'+esc(l.name)+'</div><div class="m">'+gLabel(l.g)+' · '+money(l.price)+'</div></div>'+
      '<div style="text-align:right"><span class="step"><button data-act="dec" data-k="'+esc(l.key)+'" aria-label="Quitar uno">−</button><span class="num">'+l.qty+'</span><button data-act="inc" data-k="'+esc(l.key)+'" aria-label="Agregar uno">+</button></span><div class="num" style="font-weight:800;margin-top:4px">'+money(l.sub)+'</div></div></div>'}).join('')
      :'<p class="empty">Tu carrito está vacío.</p>';
    var miss=minMissing();
    foot=lines.length?totals(sub,del,total)+(miss>0?'<p class="note" style="margin:0;color:var(--c2);font-weight:700">Pedido mínimo '+money(S.minOrder)+' (incluye delivery). Te faltan '+money(miss)+'.</p>':'')+'<button class="btn primary" data-act="toForm" '+(miss>0?'disabled':'')+'>Continuar</button>':'';
  } else if(st.step==='form'){
    var F=st.form,E=st.errors,d=S.delivery||{};
    body='<div class="form">'+
      fld('name','Nombre y apellido','text','name',F.name,E.name)+
      fld('phone','Celular (WhatsApp)','tel','tel',F.phone,E.phone)+
      (d.mode==='district'?'<div class="field"><label for="f-district">Distrito</label><select id="f-district" data-f="district"><option value="">Elige tu distrito</option>'+(d.districts||[]).map(function(z){return '<option '+(z.name===F.district?'selected':'')+' value="'+esc(z.name)+'">'+esc(z.name)+'</option>'}).join('')+'</select>'+(E.district?'<span class="err">'+esc(E.district)+'</span>':'')+'<span class="note">'+(deliveryCost()!=null?'Delivery a tu distrito: <b>'+money(deliveryCost())+'</b>':'El costo de delivery depende de tu distrito.')+' ¿No está tu distrito? Escríbenos por WhatsApp.</span></div>'
          :fld('district','Distrito','text','address-level2',F.district,E.district))+
      fld('address','Dirección','text','street-address',F.address,E.address)+
      fld('ref','Referencia (opcional)','text','off',F.ref)+
      '<div class="field"><label for="f-when">Franja preferida</label><select id="f-when" data-f="when"><option value="">Cualquier hora</option>'+FRANJAS.map(function(o){return '<option '+(o===F.when?'selected':'')+' value="'+esc(o)+'">'+esc(o)+'</option>'}).join('')+'</select></div>'+
      '<div class="yape" style="border-color:var(--c3);padding:10px 12px;text-align:left;gap:4px"><span>Entrega estimada: <b>'+fmtDay(deliveryDate())+'</b></span><span class="note">Confirmamos día y franja por WhatsApp al validar tu pago.</span></div>'+
      '<div class="field"><label for="f-notes">Comentario (opcional)</label><textarea id="f-notes" data-f="notes" rows="2">'+esc(F.notes)+'</textarea></div>'+
    '</div>';
    foot=totals(sub,del,total)+'<div class="toolbar"><button class="btn ghost" data-act="toCart">Atrás</button><button class="btn primary" style="flex:1" data-act="confirm">Ir a pagar '+money(total)+'</button></div>';
  } else {
    var o=st.order,sent=st.sent||{};
    var wa=S.whatsapp?waLink(buildMsg(o)):'';
    body='<div class="yape"><span class="eyebrow">Pedido '+esc(o.code)+'</span><div>Paga exactamente</div><div class="amt num">'+money(o.total)+'</div>'+
      (S.yapeQR?'<img src="'+esc(S.yapeQR)+'" alt="Código QR para pagar con Plin"><div class="note" style="margin-top:-4px">Escanea con Plin o la app de tu banco</div>':'')+
      (S.yapeNumber?'<div class="copyrow"><code class="num">'+esc(S.yapeNumber)+'</code><button class="btn sm" data-act="copy" data-v="'+esc(S.yapeNumber)+'">Copiar número</button></div>':'<p class="note">El número para pagar aún no está configurado.</p>')+
      '<div class="note">Desde Yape o Plin, paga a este número'+(S.yapeName?' · a nombre de <b>'+esc(S.yapeName)+'</b>':'')+'</div>'+'</div>';
    if(sent.ok){
      body+='<div class="yape" style="border-color:var(--c3);text-align:left"><span class="eyebrow" style="color:var(--c3)">Pedido registrado ✓</span><div>Tu pedido <b>'+esc(sent.code)+'</b> quedó registrado con entrega estimada el <b>'+esc(sent.entrega||fmtDay(deliveryDate()))+'</b>.</div><div class="note">Último paso: envíanos la captura de tu Yape por WhatsApp. Te confirmaremos el pago, el día y la franja de entrega por ese medio. Guarda tu código para consultar el estado en "Mi pedido".</div></div>';
      foot=(wa?'<a class="btn primary" style="background:var(--wa);border-color:var(--wa)" href="'+esc(wa)+'" target="_blank" rel="noopener">Enviar captura por WhatsApp</a>':'')+'<button class="btn ghost" data-act="newOrder">Terminar</button>';
    } else {
      body+='<ol class="steps"><li>Paga '+money(o.total)+' por Yape o Plin al número de arriba.</li><li>Escribe el número de operación de tu constancia.</li><li>Pulsa "Ya pagué" y envíanos la captura por WhatsApp.</li></ol>'+
        '<div class="field"><label for="f-op">N° de operación (Yape o Plin)</label><input id="f-op" data-f="op" inputmode="numeric" value="'+esc(st.form.op)+'" placeholder="Ej. 12345678">'+(st.errors.op?'<span class="err">'+esc(st.errors.op)+'</span>':'')+'</div>'+
        '<input type="text" id="f-web" tabindex="-1" autocomplete="off" aria-hidden="true" style="position:absolute;left:-9999px" value="">'+
        (sent.err?'<p class="field"><span class="err">'+esc(sent.err)+'</span></p>':'');
      foot=CFG.apiUrl&&!sent.err?'<button class="btn primary" data-act="submitOrder" '+(sent.busy?'disabled':'')+'>'+(sent.busy?'Registrando…':'Ya pagué · Registrar pedido')+'</button>':
        (wa?'<a class="btn primary" style="background:var(--wa);border-color:var(--wa)" href="'+esc(wa)+'" target="_blank" rel="noopener">Ya pagué · Enviar por WhatsApp</a>':'<p class="note">Falta configurar el WhatsApp de la tienda.</p>');
      foot+='<button class="btn ghost" data-act="toForm">Atrás</button>';
    }
  }
  return '<div class="scrim" data-act="closeCart"></div><aside class="drawer" role="dialog" aria-modal="true" aria-label="'+esc(title)+'"><header><h2>'+esc(title)+'</h2><button class="btn ghost sm" data-act="closeCart" aria-label="Cerrar">Cerrar</button></header><div class="scroll">'+body+'</div>'+(foot?'<footer>'+foot+'</footer>':'')+'</aside>';
}
function fld(k,label,type,ac,v,err){return '<div class="field"><label for="f-'+k+'">'+esc(label)+'</label><input id="f-'+k+'" type="'+type+'" autocomplete="'+ac+'" data-f="'+k+'" value="'+esc(v)+'">'+(err?'<span class="err">'+esc(err)+'</span>':'')+'</div>'}
function totals(sub,del,total){return '<div class="tot num"><span>Productos</span><span>'+money(sub)+'</span><span>Delivery</span><span>'+(del==null?'desde '+money(minDel()):money(del))+'</span><span class="big">Total'+(del==null?' desde':'')+'</span><span class="big">'+money(del==null?sub+minDel():total)+'</span></div>'}

function makeCode(){var d=limaNow();return 'FS-'+String(d.getMonth()+1).padStart(2,'0')+String(d.getDate()).padStart(2,'0')+'-'+Math.floor(1000+Math.random()*9000)}
function buildMsg(o){
  var code=(st.sent&&st.sent.code)||o.code;
  var L=['Hola, envío mi pedido.','Ref: '+code,''];
  o.items.forEach(function(i){L.push('• '+i.qty+' × '+i.name+' '+gLabel(i.g)+' — '+money(i.sub))});
  L.push('','Productos: '+money(o.subtotal),'Delivery: '+money(o.delivery),'Total: '+money(o.total),'Pago: Yape/Plin'+(st.form.op?' · Op. '+st.form.op:''),'',
    'Cliente: '+o.customer,'Celular: '+o.phone,'Distrito: '+o.district,'Dirección: '+o.address);
  if(o.ref)L.push('Referencia: '+o.ref);
  L.push('Entrega estimada: '+o.entrega+(o.when?' · '+o.when:''));
  if(o.notes)L.push('Comentario: '+o.notes);
  L.push('','Adjunto la captura del pago.');
  return L.join('\n');
}
function confirmOrder(){
  var F=st.form,E={};
  if(F.name.trim().length<3)E.name='Escribe tu nombre y apellido.';
  if(String(F.phone).replace(/\D/g,'').length<9)E.phone='Escribe un celular de 9 dígitos.';
  if(!F.district.trim())E.district='Indica tu distrito.';
  if(F.address.trim().length<5)E.address='Escribe tu dirección.';
  st.errors=E;if(Object.keys(E).length){render();return}
  var lines=cartLines(),sub=subtotal(),del=deliveryCost()||0;
  st.order={code:makeCode(),customer:F.name.trim(),phone:F.phone.trim(),district:F.district.trim(),address:F.address.trim(),ref:F.ref.trim(),when:F.when,notes:F.notes.trim(),entrega:fmtDay(deliveryDate()),
    items:lines.map(function(l){return {name:l.name,g:l.g,qty:l.qty,price:l.price,sub:l.sub}}),subtotal:sub,delivery:del,total:r2(sub+del)};
  st.sent={};st.step='yape';render();
}
function submitOrder(){
  var op=String(st.form.op||'').replace(/\s/g,'');
  if(op.length<4){st.errors={op:'Escribe el número de operación que aparece en tu constancia de Yape o Plin.'};render();return}
  st.errors={};st.sent={busy:true};render();
  var o=st.order,hp=($('#f-web')||{}).value||'';
  var ctl=window.AbortController?new AbortController():null,timer=setTimeout(function(){ctl&&ctl.abort()},20000);
  fetch(CFG.apiUrl,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(Object.assign({},o,{op:op,website:hp})),signal:ctl?ctl.signal:undefined})
    .then(function(r){return r.json()}).then(function(res){
      clearTimeout(timer);
      if(!res.ok)throw new Error(res.error||'error');
      st.sent={ok:true,code:res.code,entrega:res.entrega};o.code=res.code;
      if(res.total!=null)o.total=res.total;
      render();
    }).catch(function(){
      clearTimeout(timer);
      st.sent={err:'No pudimos registrar tu pedido en línea. No te preocupes: envíanoslo por WhatsApp con el botón de abajo.'};render();
    });
}
function doTrack(){
  var T=st.track=st.track||{};
  T.code=String(T.code||'').trim().toUpperCase();T.tel=String(T.tel||'').replace(/\D/g,'');
  if(!T.code||T.tel.length!==4){T.err='Completa el código y los 4 dígitos.';T.res=null;render();return}
  if(!CFG.apiUrl){T.err='La consulta en línea aún no está disponible. Escríbenos por WhatsApp.';render();return}
  T.busy=true;T.err='';T.res=null;render();
  fetch(CFG.apiUrl+'?action=status&code='+encodeURIComponent(T.code)+'&tel='+T.tel).then(function(r){return r.json()}).then(function(res){
    T.busy=false;if(res.ok)T.res=res;else T.err=res.error||'No encontramos ese pedido.';render();
  }).catch(function(){T.busy=false;T.err='No pudimos consultar ahora. Intenta en un momento o escríbenos por WhatsApp.';render()});
}

/* ---------- events ---------- */
var toastMsg='',toastT;
function toast(m){toastMsg=m;clearTimeout(toastT);render();toastT=setTimeout(function(){toastMsg='';var t=$('.toast');if(t)t.remove()},2600)}
function copyText(t){try{navigator.clipboard.writeText(t).then(function(){toast('Copiado')},function(){toast('Selecciona el texto y cópialo manualmente')})}catch(e){toast('Selecciona el texto y cópialo manualmente')}}

app.addEventListener('click',function(e){
  var el=e.target.closest('[data-act]');if(!el)return;
  var a=el.dataset.act,k=el.dataset.k,v=el.dataset.v;
  if(a==='inc'){cart[k]=(cart[k]||0)+1;saveCart();render();return}
  if(a==='dec'){cart[k]=(cart[k]||0)-1;if(cart[k]<=0)delete cart[k];saveCart();render();return}
  if(a==='sel'){st.sel[el.dataset.id]=Number(el.dataset.g);render();return}
  if(a==='cat'){st.cat=v;render();var l=$('#list');if(l)window.scrollTo({top:Math.max(0,l.getBoundingClientRect().top+window.scrollY-130),behavior:'smooth'});return}
  if(a==='openCart'){st.drawer=true;if(st.step!=='yape')st.step='cart';render();return}
  if(a==='closeCart'){st.drawer=false;render();return}
  if(a==='toForm'){st.step='form';render();return}
  if(a==='toCart'){st.step='cart';render();return}
  if(a==='method'){st.form.method=v;saveForm();render();return}
  if(a==='confirm'){confirmOrder();return}
  if(a==='submitOrder'){submitOrder();return}
  if(a==='doTrack'){doTrack();return}
  if(a==='track'){st.drawer=true;st.step='track';st.track=st.track||{};render();return}
  if(a==='copy'){copyText(v==='__msg'?buildMsg(st.order):v);return}
  if(a==='newOrder'){if(st.sent&&st.sent.ok){cart={};saveCart()}st.order=null;st.sent={};st.form.op='';st.step='cart';st.drawer=false;render();return}
  if(a==='toggleAdmin'){st.view=st.view==='admin'?'store':'admin';st.drawer=false;render();window.scrollTo(0,0);return}
  adminClick(a,el,e);
});
function saveForm(){try{localStorage.setItem('fs-form',JSON.stringify(Object.assign({},st.form,{op:''})))}catch(e){}}
app.addEventListener('input',function(e){
  var t=e.target;
  if(t.id==='q'){st.q=t.value;$('#list').innerHTML=listHtml();return}
  if(t.dataset.f){st.form[t.dataset.f]=t.value;saveForm();
    if(t.dataset.f==='op'&&st.order){var a=$('#walink');if(a&&S.whatsapp)a.href='https://wa.me/'+String(S.whatsapp).replace(/\D/g,'')+'?text='+encodeURIComponent(buildMsg(st.order));var m=$('#msgtxt');if(m)m.textContent=buildMsg(st.order)}
    return}
  if(t.dataset.t){st.track=st.track||{};st.track[t.dataset.t]=t.value;return}
  if(t.id==='pasteMsg')AD.paste=t.value;
});
app.addEventListener('change',function(e){var t=e.target;if(t.dataset.f){st.form[t.dataset.f]=t.value;saveForm();if(t.tagName==='SELECT')render();return}adminChange(t)});
document.addEventListener('keydown',function(e){if(e.key==='Escape'&&st.drawer){st.drawer=false;render()}});

var AD={paste:''};function adminClick(){}function adminChange(){}
/* ---------- catálogo desde Google Sheets ---------- */

function parseCSV(t){var rows=[],row=[],f='',q=false;for(var i=0;i<t.length;i++){var c=t[i];if(q){if(c==='"'){if(t[i+1]==='"'){f+='"';i++}else q=false}else f+=c}else if(c==='"')q=true;else if(c===','){row.push(f);f=''}else if(c==='\n'||c==='\r'){if(c==='\r'&&t[i+1]==='\n')i++;row.push(f);rows.push(row);row=[];f=''}else f+=c}if(f!==''||row.length){row.push(f);rows.push(row)}return rows}
function fromCSV(text){
  var rows=parseCSV(text);if(!rows.length)return [];
  var h=rows[0].map(function(x){return norm(x).replace(/ /g,'')});
  function col(r,k){var i=h.indexOf(k);return i<0?'':String(r[i]||'').trim()}
  return rows.slice(1).filter(function(r){return col(r,'producto')&&!/^no$/i.test(col(r,'activo'))}).map(function(r){
    var v=[];[1,2,3].forEach(function(n){var g=num(col(r,'g'+n)),p=num(col(r,'precio'+n));if(g&&p)v.push([Math.round(g),r2(p)])});
    var foto=col(r,'foto');
    return {id:slug(col(r,'producto')),name:col(r,'producto'),cat:col(r,'categoria'),desc:col(r,'descripcion'),img:foto?(/^https?:/.test(foto)?foto:'img/'+foto):'',v:v};
  }).filter(function(p){return p.v.length});
}
function loadSheet(){
  var url=(window.ABAS_CONFIG||{}).sheetCsvUrl;if(!url)return;
  fetch(url+(url.indexOf('?')<0?'?':'&')+'t='+Date.now()).then(function(r){if(!r.ok)throw 0;return r.text()}).then(function(t){
    var ps=fromCSV(t);if(!ps.length)return;
    PRODUCTS=ps;byId={};PRODUCTS.forEach(function(p){byId[p.id]=p});
    Object.keys(cart).forEach(function(k){if(!variantOf(k))delete cart[k]});saveCart();
    DATA.publishedAt=new Date().toISOString();render();
  }).catch(function(){});
}
function loadConfig(){
  if(!CFG.apiUrl)return;
  fetch(CFG.apiUrl+'?action=config').then(function(r){return r.json()}).then(function(c){
    if(!c||!c.ok||c.delivery==null)return;
    S.delivery=Object.assign({},S.delivery||{},{mode:'flat',flat:Number(c.delivery)});
    S.minOrder=Number(c.minimo)||0;
    if(c.whatsapp)S.whatsapp=String(c.whatsapp);
    if(c.pagoNumero)S.yapeNumber=String(c.pagoNumero);
    if(c.pagoTitular)S.yapeName=String(c.pagoTitular);
    if(c.franjas&&c.franjas.length)FRANJAS=c.franjas;
    if(c.districts&&c.districts.length)S.delivery={mode:'district',flat:Number(c.delivery),districts:c.districts.map(function(z){return {name:z.name,cost:Number(z.cost)}})};
    RULES={anticipacion:c.anticipacion,sinReparto:c.sinReparto};
    render();
  }).catch(function(){});
}
render();
loadSheet();
loadConfig();
})();
