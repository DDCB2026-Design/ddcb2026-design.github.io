
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
   (st.view==='store'?'<button class="btn navy" data-act="openCart" aria-label="Abrir carrito"><span class="lbl">Carrito</span><span class="badge num">'+n+'</span></button>':'')+
  '</div><div class="stripe" aria-hidden="true"><i></i><i></i><i></i><i></i></div></header>';
}
function storeView(){
  var d=S.delivery||{};
  var delTxt=d.mode==='district'?'Delivery según distrito':'Delivery '+money(d.flat)+' a todo Lima';
  var cs=cats();
  var tiles='<div class="wrap tiles" role="group" aria-label="Categorías">'+['Todas'].concat(cs).map(function(c){
    var img=c==='Todas'?S.heroImg:(S.catImgs||{})[c],n=c==='Todas'?PRODUCTS.length:PRODUCTS.filter(function(p){return p.cat===c}).length;
    return '<button class="tile" style="'+(c==='Todas'?'--cc:var(--c5)':catVars(c))+'" data-act="cat" data-v="'+esc(c)+'" aria-pressed="'+(st.cat===c)+'"><span class="im">'+(img?'<img src="'+esc(img)+'" alt="" loading="lazy">':'')+'</span><span class="l">'+esc(c==='Todas'?'Todo el catálogo':c)+'<br><small style="font-family:var(--body);font-weight:600;font-size:12px;text-transform:none;opacity:.9">'+(n?n+' productos':'Próximamente')+'</small></span></button>'}).join('')+'</div>';
  return '<div class="searchbar"><div class="wrap"><input id="q" type="search" placeholder="Buscar pecanas, almendras, pasas…" value="'+esc(st.q)+'" aria-label="Buscar producto"></div></div>'+
   '<section class="wrap"><div class="hero"><div class="copy"><span class="eyebrow">Frutos secos · Deshidratados · Semillas · Especias</span>'+
    '<h2>'+esc(S.tagline||'Frutos secos frescos, directo a tu casa')+'</h2>'+
    '<p>Elige tus productos, paga con Yape y recibe tu pedido en Lima.</p>'+
    '<div class="perks"><span class="perk"><i class="y"></i>Paga con Yape</span><span class="perk"><i></i>'+esc(delTxt)+'</span>'+(S.pickup?'<span class="perk"><i></i>Recojo sin costo</span>':'')+(Number(S.minOrder)>0?'<span class="perk"><i></i>Pedido mínimo '+money(S.minOrder)+'</span>':'')+'</div></div>'+
    '<div class="pic">'+(S.heroImg?'<img src="'+esc(S.heroImg)+'" alt="Variedad de frutos secos">':'')+'</div></div>'+
    '<div class="strip"><div style="background:var(--c1)"><b>1</b>Elige</div><div style="background:var(--c2)"><b>2</b>Paga con Yape</div><div style="background:var(--c4)"><b>3</b>Envía por WhatsApp</div><div style="background:var(--navy)"><b>4</b>Recibe en casa</div></div></section>'+tiles+
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
function drawer(){
  var title=st.step==='cart'?'Tu pedido':st.step==='form'?'Datos de entrega':'Paga con Yape';
  var body='',foot='';
  var lines=cartLines(),sub=subtotal(),del=deliveryCost(),total=r2(sub+(del||0));
  if(st.step==='cart'){
    body=lines.length?lines.map(function(l){return '<div class="line"><div class="th">'+(l.img?'<img src="'+esc(l.img)+'" alt="">':'')+'</div><div><div class="n">'+esc(l.name)+'</div><div class="m">'+gLabel(l.g)+' · '+money(l.price)+'</div></div>'+
      '<div style="text-align:right"><span class="step"><button data-act="dec" data-k="'+esc(l.key)+'" aria-label="Quitar uno">−</button><span class="num">'+l.qty+'</span><button data-act="inc" data-k="'+esc(l.key)+'" aria-label="Agregar uno">+</button></span><div class="num" style="font-weight:800;margin-top:4px">'+money(l.sub)+'</div></div></div>'}).join('')
      :'<p class="empty">Tu carrito está vacío.</p>';
    var minOk=!(Number(S.minOrder)>0&&sub<Number(S.minOrder));
    foot=lines.length?totals(sub,del,total)+(minOk?'':'<p class="note">El pedido mínimo es '+money(S.minOrder)+'.</p>')+'<button class="btn primary" data-act="toForm" '+(minOk?'':'disabled')+'>Continuar</button>':'';
  } else if(st.step==='form'){
    var F=st.form,E=st.errors,d=S.delivery||{};
    body='<div class="form">'+
      (S.pickup?'<div class="seg" role="group" aria-label="Tipo de entrega"><button data-act="method" data-v="delivery" aria-pressed="'+(F.method==='delivery')+'">Delivery</button><button data-act="method" data-v="pickup" aria-pressed="'+(F.method==='pickup')+'">Recojo</button></div>':'')+
      fld('name','Nombre y apellido','text','name',F.name,E.name)+
      fld('phone','Celular (WhatsApp)','tel','tel',F.phone,E.phone)+
      (F.method==='pickup'?'<p class="note">Recojo en: '+esc(S.pickupAddress||'se coordina por WhatsApp')+'</p>':
        (d.mode==='district'?'<div class="field"><label for="f-district">Distrito</label><select id="f-district" data-f="district"><option value="">Elige tu distrito</option>'+(d.districts||[]).map(function(z){return '<option '+(z.name===F.district?'selected':'')+' value="'+esc(z.name)+'">'+esc(z.name)+' · '+money(z.cost)+'</option>'}).join('')+'</select>'+(E.district?'<span class="err">'+esc(E.district)+'</span>':'')+'</div>'
          :fld('district','Distrito','text','address-level2',F.district,E.district))+
        fld('address','Dirección','text','street-address',F.address,E.address)+
        fld('ref','Referencia (opcional)','text','off',F.ref)+
        '<div class="field"><label for="f-when">Horario preferido (opcional)</label><select id="f-when" data-f="when">'+['','Mañana (9 a 12)','Tarde (12 a 4)','Noche (4 a 8)'].map(function(o){return '<option '+(o===F.when?'selected':'')+' value="'+esc(o)+'">'+(o||'Cualquier hora')+'</option>'}).join('')+'</select></div>')+
      '<div class="field"><label for="f-notes">Comentario (opcional)</label><textarea id="f-notes" data-f="notes" rows="2">'+esc(F.notes)+'</textarea></div>'+
    '</div>';
    foot=totals(sub,del,total)+'<div class="toolbar"><button class="btn ghost" data-act="toCart">Atrás</button><button class="btn primary" style="flex:1" data-act="confirm">Ir a pagar '+money(total)+'</button></div>';
  } else {
    var o=st.order;
    var wa=S.whatsapp?'https://wa.me/'+String(S.whatsapp).replace(/\D/g,'')+'?text='+encodeURIComponent(buildMsg(o)):'';
    body='<div class="yape"><span class="eyebrow">Pedido '+esc(o.code)+'</span><div>Yapea exactamente</div><div class="amt num">'+money(o.total)+'</div>'+
      (S.yapeQR?'<img src="'+esc(S.yapeQR)+'" alt="Código QR de Yape">':'')+
      (S.yapeNumber?'<div class="copyrow"><code class="num">'+esc(S.yapeNumber)+'</code><button class="btn sm" data-act="copy" data-v="'+esc(S.yapeNumber)+'">Copiar número</button></div>':'<p class="note">El número de Yape aún no está configurado.</p>')+
      (S.yapeName?'<div class="note">A nombre de <b>'+esc(S.yapeName)+'</b></div>':'')+'</div>'+
      '<ol class="steps"><li>Paga '+money(o.total)+' por Yape al número o QR de arriba.</li><li>Escribe el número de operación de tu constancia.</li><li>Envía el pedido por WhatsApp y adjunta la captura. Validamos el pago y despachamos.</li></ol>'+
      '<div class="field"><label for="f-op">N° de operación Yape</label><input id="f-op" data-f="op" inputmode="numeric" value="'+esc(st.form.op)+'" placeholder="Ej. 12345678"></div>'+
      '<details><summary class="note">Ver el mensaje del pedido</summary><div class="msg" id="msgtxt">'+esc(buildMsg(o))+'</div><button class="btn sm" data-act="copy" data-v="__msg" style="margin-top:6px">Copiar mensaje</button></details>';
    foot=(wa?'<a class="btn primary" id="walink" href="'+esc(wa)+'" target="_blank" rel="noopener">Ya pagué · Enviar por WhatsApp</a><p class="note" style="margin:0">Si WhatsApp no se abre, copia el mensaje y envíalo al '+esc(S.whatsapp)+'.</p>':'<p class="note">Falta configurar el WhatsApp de la tienda.</p>')+
      '<button class="btn ghost" data-act="newOrder">Terminar y vaciar carrito</button>';
  }
  return '<div class="scrim" data-act="closeCart"></div><aside class="drawer" role="dialog" aria-modal="true" aria-label="'+esc(title)+'"><header><h2>'+esc(title)+'</h2><button class="btn ghost sm" data-act="closeCart" aria-label="Cerrar">Cerrar</button></header><div class="scroll">'+body+'</div>'+(foot?'<footer>'+foot+'</footer>':'')+'</aside>';
}
function fld(k,label,type,ac,v,err){return '<div class="field"><label for="f-'+k+'">'+esc(label)+'</label><input id="f-'+k+'" type="'+type+'" autocomplete="'+ac+'" data-f="'+k+'" value="'+esc(v)+'">'+(err?'<span class="err">'+esc(err)+'</span>':'')+'</div>'}
function totals(sub,del,total){return '<div class="tot num"><span>Productos</span><span>'+money(sub)+'</span><span>'+(st.form.method==='pickup'?'Recojo':'Delivery')+'</span><span>'+(del==null?'Elige distrito':money(del))+'</span><span class="big">Total</span><span class="big">'+money(total)+'</span></div>'}

function makeCode(){var d=new Date();return 'FS-'+String(d.getMonth()+1).padStart(2,'0')+String(d.getDate()).padStart(2,'0')+'-'+Math.floor(1000+Math.random()*9000)}
function buildMsg(o){
  var L=['Hola, quiero hacer este pedido.','Ref: '+o.code,''];
  o.items.forEach(function(i){L.push('• '+i.qty+' × '+i.name+' '+gLabel(i.g)+' — '+money(i.sub))});
  L.push('','Productos: '+money(o.subtotal),(o.method==='pickup'?'Recojo: ':'Delivery: ')+money(o.delivery),'Total: '+money(o.total),'Pago: Yape'+(st.form.op?' · Op. '+st.form.op:''),'',
    'Cliente: '+o.customer,'Celular: '+o.phone,'Entrega: '+(o.method==='pickup'?'Recojo':'Delivery'));
  if(o.method!=='pickup'){L.push('Distrito: '+o.district,'Dirección: '+o.address);if(o.ref)L.push('Referencia: '+o.ref);if(o.when)L.push('Horario: '+o.when)}
  if(o.notes)L.push('Comentario: '+o.notes);
  L.push('','Adjunto la captura del Yape.');
  return L.join('\n');
}
function confirmOrder(){
  var F=st.form,E={};
  if(F.name.trim().length<3)E.name='Escribe tu nombre y apellido.';
  if(String(F.phone).replace(/\D/g,'').length<9)E.phone='Escribe un celular de 9 dígitos.';
  if(F.method!=='pickup'){if(!F.district.trim())E.district='Indica tu distrito.';if(F.address.trim().length<5)E.address='Escribe tu dirección.'}
  st.errors=E;if(Object.keys(E).length){render();return}
  var lines=cartLines(),sub=subtotal(),del=deliveryCost()||0;
  var o={code:makeCode(),createdAt:new Date().toISOString(),customer:F.name.trim(),phone:F.phone.trim(),method:F.method,district:F.method==='pickup'?'':F.district.trim(),address:F.method==='pickup'?'':F.address.trim(),ref:F.ref.trim(),when:F.when,notes:F.notes.trim(),
    items:lines.map(function(l){return {id:l.id,name:l.name,g:l.g,qty:l.qty,price:l.price,sub:l.sub}}),subtotal:sub,delivery:del,total:r2(sub+del),status:'Pago pendiente',source:'web'};
  st.order=o;st.step='yape';render();
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
  if(a==='copy'){copyText(v==='__msg'?buildMsg(st.order):v);return}
  if(a==='newOrder'){cart={};saveCart();st.order=null;st.form.op='';st.step='cart';st.drawer=false;render();return}
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
render();
loadSheet();
})();
