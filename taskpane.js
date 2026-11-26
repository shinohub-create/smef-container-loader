/* SMEF Container Loader – Excel task pane.
   Reads Input / Containers / Loading Plan from the open workbook, packs with the
   same rules as the workbook's Calc sheet, and draws the load in 3D (three.js r128). */
(function(){
"use strict";

/* ---------- defaults ---------- */
const DEFAULT_CONTAINERS=[
  {type:"20ft",L:589,W:235,H:239,payload:28200},
  {type:"40ft",L:1203,W:235,H:239,payload:26700},
  {type:"40ft HC",L:1203,W:235,H:269,payload:26500}];
const COLORS=["#1f4fd1","#e8761c","#1a9a4a","#d81b7a","#7b2fd0","#f2c12e","#0f8f9a","#c62828","#3d3d9e","#8bc34a","#ff5fa2","#00a3e0","#a0522d","#6d7b8d","#b39ddb"];

let containers=DEFAULT_CONTAINERS.map(c=>({...c}));
let cargo=[];            // {row, name, qty, l, w, h, kg, stack}
let ctype="40ft", cnum=0, view="iso", result=null;
let inExcel=false, hasLoadingPlan=false;

const $=id=>document.getElementById(id);
const esc=s=>String(s==null?"":s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const fmt=(n,d=0)=>Number(n).toLocaleString("en-GB",{minimumFractionDigits:d,maximumFractionDigits:d});
function setStatus(t,err){const s=$("status");s.textContent=t;s.classList.toggle("err",!!err)}

/* ---------- workbook parsing (shared by Excel and file import) ---------- */
function parseCargo(rows){
  const norm=v=>String(v??"").toLowerCase().replace(/\(.*?\)|\bcm\b|[^a-z ]/g," ").replace(/\s+/g," ").trim();
  const isW=h=>h==="w"||/\bwidth\b/.test(h),isD=h=>h==="d"||/depth|length/.test(h),isH=h=>h==="h"||/height/.test(h);
  const hi=rows.findIndex(r=>{const n=r.map(norm);return n.some(isW)&&n.some(isD)&&n.some(isH)});
  if(hi<0)return{error:"Couldn't find the cargo header row (Model, Quantity, W, D, H…) on the Input sheet."};
  const H=rows[hi].map(norm);
  const col=t=>H.findIndex(typeof t==="function"?t:h=>t.test(h));
  const ci={name:col(/model|desc|item/),qty:col(/qty|quantity|pcs/),l:col(isD),w:col(isW),h:col(isH),kg:col(/weight|kg/),stack:col(/stack/)};
  if(ci.l<0||ci.w<0||ci.h<0||ci.qty<0)return{error:"The Input header row needs Quantity, W, D and H columns."};
  const out=[];
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i],nm=ci.name>=0?String(r[ci.name]??"").trim():"";
    if(/^total$/i.test(nm))break;
    const num=k=>{const v=ci[k]>=0?r[ci[k]]:"";const n=typeof v==="number"?v:parseFloat(String(v??"").replace(",","."));return isFinite(n)?n:0};
    const it={row:i,name:nm,qty:num("qty"),l:num("l"),w:num("w"),h:num("h"),kg:num("kg"),stack:num("stack")};
    if(it.qty>0&&it.l>0&&it.w>0&&it.h>0)out.push(it);
  }
  return{cargo:out};
}
function parseContainers(rows){
  const found=[];
  (rows||[]).forEach(r=>{const t=String(r[0]??"").trim();if(/^\d+\s*ft/i.test(t)&&+r[1]>0&&+r[2]>0&&+r[3]>0)found.push({type:t,L:+r[1],W:+r[2],H:+r[3],payload:+r[5]||0})});
  return found.length?found:DEFAULT_CONTAINERS.map(c=>({...c}));
}
function applyData(inputRows,containerRows,sel,source){
  const pc=parseCargo(inputRows||[]);
  if(pc.error){cargo=[];setStatus(pc.error,true)}
  else{cargo=pc.cargo;setStatus(`${cargo.length} model${cargo.length===1?"":"s"} from ${source} · updated ${new Date().toLocaleTimeString([], {hour:"2-digit",minute:"2-digit",second:"2-digit"})}`)}
  containers=parseContainers(containerRows);
  if(sel&&containers.some(c=>c.type===sel))ctype=sel;
  else if(!containers.some(c=>c.type===ctype))ctype=containers[0].type;
  matCache.clear();recompute();
}

/* ---------- packing (mirrors the workbook's Calc sheet) ---------- */
function planItem(it,c){
  const {qty:q,l,w,h,stack:st}=it;
  const layers=Math.min(Math.floor(c.H/h),st>0?st:99);
  const acA=Math.floor(c.W/w),acB=Math.floor(c.W/l);
  const lenA=(acA*layers>0&&l<=c.L)?Math.ceil(q/(acA*layers))*l:Infinity;
  const lenB=(acB*layers>0&&w<=c.L)?Math.ceil(q/(acB*layers))*w:Infinity;
  if(!isFinite(Math.min(lenA,lenB)))return{fit:false,why:layers<1?"too tall":"too wide or long"};
  const A=lenA<=lenB,across=A?acA:acB,along=A?l:w,acrossDim=A?w:l,perRow=across*layers,rows=Math.ceil(q/perRow);
  return{fit:true,orient:A?"Lengthwise":"Turned 90°",layers,across,along,acrossDim,perRow,rows,floorLen:rows*along,h};
}
function pack(items,c){
  const plans=items.map(it=>planItem(it,c));
  const list=[];let cur=null;
  const fresh=()=>{cur={units:[],x:0,kg:0,vol:0,rows:0,items:new Map()};list.push(cur)};
  fresh();
  items.forEach((it,idx)=>{
    const p=plans[idx];if(!p.fit)return;
    let remaining=it.qty,unitNo=0;
    for(let r=0;r<p.rows;r++){
      if(cur.x+p.along>c.L+1e-6)fresh();
      const n=Math.min(p.perRow,remaining);cur.rows++;
      for(let i=0;i<n;i++){
        const layer=Math.floor(i/p.across),col=i%p.across;
        cur.units.push({item:idx,unitNo:++unitNo,qty:it.qty,x:cur.x,y:layer*p.h,z:col*p.acrossDim,sx:p.along,sy:p.h,sz:p.acrossDim,row:cur.rows,layer:layer+1,orient:p.orient});
        cur.kg+=it.kg||0;cur.vol+=p.along*p.h*p.acrossDim;
        cur.items.set(idx,(cur.items.get(idx)||0)+1);
      }
      remaining-=n;cur.x+=p.along;
    }
  });
  if(list.length>1&&!list[list.length-1].units.length)list.pop();
  const totalKg=items.reduce((s,it)=>s+(it.kg||0)*it.qty,0);
  const noFit=plans.filter(p=>!p.fit).length;
  const byWeight=c.payload>0?Math.ceil(totalKg/c.payload):0;
  const hasUnits=list.some(k=>k.units.length);
  return{plans,list,noFit,totalKg,needed:hasUnits?Math.max(list.length,byWeight):0,byWeight};
}
const getC=t=>containers.find(c=>c.type===t)||containers[0];

/* ---------- three.js scene ---------- */
const stage=$("stage"),tip=$("tip");
const css=n=>getComputedStyle(document.documentElement).getPropertyValue(n).trim();
let renderer,scene,camera,controls,world,cargoGroup,hovered=null,unitMeshes=[];
const raycaster=new THREE.Raycaster(),pointer=new THREE.Vector2(),geoCache=new Map(),matCache=new Map();

function initScene(){
  renderer=new THREE.WebGLRenderer({antialias:true});
  renderer.setPixelRatio(Math.min(window.devicePixelRatio||1,2));
  stage.prepend(renderer.domElement);
  scene=new THREE.Scene();
  camera=new THREE.PerspectiveCamera(40,1,5,40000);
  controls=new THREE.OrbitControls(camera,renderer.domElement);
  controls.enableDamping=true;controls.dampingFactor=.12;controls.maxPolarAngle=Math.PI*.495;controls.screenSpacePanning=true;
  scene.add(new THREE.HemisphereLight(0xffffff,0x8a94a0,.75));
  const d=new THREE.DirectionalLight(0xffffff,.55);d.position.set(.6,1.2,.9);scene.add(d);
  const d2=new THREE.DirectionalLight(0xffffff,.2);d2.position.set(-.8,.5,-.6);scene.add(d2);
  let lastW=0;
  new ResizeObserver(()=>{resize();const w=stage.clientWidth;if(Math.abs(w-lastW)>40){lastW=w;setView(view,true)}}).observe(stage);resize();
  renderer.domElement.addEventListener("pointermove",onHover);
  renderer.domElement.addEventListener("pointerleave",()=>setHover(null));
  (function loop(){requestAnimationFrame(loop);if(tween)stepTween();controls.update();renderer.render(scene,camera)})();
}
function resize(){const w=stage.clientWidth,h=stage.clientHeight;if(!w||!h)return;renderer.setSize(w,h,false);camera.aspect=w/h;camera.updateProjectionMatrix()}

function textCanvas(lines,{w,h,bg,fg,border}){
  const cv=document.createElement("canvas");cv.width=w;cv.height=h;const g=cv.getContext("2d");
  g.fillStyle=bg;g.fillRect(0,0,w,h);
  g.strokeStyle=border;g.lineWidth=Math.max(2,Math.min(w,h)*.02);g.strokeRect(0,0,w,h);
  g.fillStyle=fg;g.textBaseline="top";
  const px=w*.07,avail=w-px*2,base=Math.min(h*.16,w*.13);
  let y=h*.08;
  lines.forEach((t,i)=>{
    const wt=i===0?600:400;let s=i===0?base:base*.78;
    g.font=`${wt} ${s}px "Segoe UI",Arial,sans-serif`;const m=g.measureText(t).width;if(m>avail)s=s*avail/m;s=Math.max(s,6);
    g.font=`${wt} ${s}px "Segoe UI",Arial,sans-serif`;if(y+s<h*.98)g.fillText(t,px,y);y+=s*1.25;
  });
  return cv;
}
function lum(hex){const n=parseInt(hex.slice(1),16);return(.2126*(n>>16&255)+.7152*(n>>8&255)+.0722*(n&255))/255}
function shade(hex,k){const n=parseInt(hex.slice(1),16);const f=v=>Math.max(0,Math.min(255,Math.round(v*k)));return"#"+[f(n>>16&255),f(n>>8&255),f(n&255)].map(v=>v.toString(16).padStart(2,"0")).join("")}
function faceMat(idx,fw,fh,labels,ghost){
  const it=cargo[idx],col=COLORS[idx%COLORS.length];
  const key=[idx,Math.round(fw),Math.round(fh),labels,ghost].join("|");
  if(matCache.has(key))return matCache.get(key);
  const opts={transparent:ghost,opacity:ghost?.35:1,depthWrite:!ghost};
  let m;
  if(labels){
    const scale=Math.min(4,900/Math.max(fw,fh));
    const cv=textCanvas([`#${idx+1} ${it.name||"Model "+(idx+1)}`,`W${it.w} × D${it.l} × H${it.h} cm`,it.kg?`${it.kg} kg`:""].filter(Boolean),
      {w:Math.max(32,Math.round(fw*scale)),h:Math.max(32,Math.round(fh*scale)),bg:col,fg:lum(col)>.55?"#111":"#fff",border:shade(col,.6)});
    const tex=new THREE.CanvasTexture(cv);tex.anisotropy=4;
    m=new THREE.MeshLambertMaterial({map:tex,...opts});
  }else m=new THREE.MeshLambertMaterial({color:col,...opts});
  matCache.set(key,m);return m;
}
function sprite(text,{size=18,fg,bg,bold=false}){
  const f=`${bold?700:500} 40px Consolas,"Cascadia Mono",monospace`;
  const c0=document.createElement("canvas").getContext("2d");c0.font=f;const tw=c0.measureText(text).width;
  const cv=document.createElement("canvas");cv.width=Math.ceil(tw+24);cv.height=56;const g=cv.getContext("2d");
  if(bg){g.fillStyle=bg;g.fillRect(0,0,cv.width,cv.height)}
  g.font=f;g.fillStyle=fg;g.textBaseline="middle";g.fillText(text,12,cv.height/2+2);
  const s=new THREE.Sprite(new THREE.SpriteMaterial({map:new THREE.CanvasTexture(cv),depthTest:false,transparent:true}));
  s.scale.set(size*cv.width/cv.height,size,1);s.renderOrder=5;return s;
}
function line(pts,color,opacity=1){const g=new THREE.BufferGeometry().setFromPoints(pts.map(p=>new THREE.Vector3(...p)));return new THREE.LineSegments(g,new THREE.LineBasicMaterial({color,transparent:opacity<1,opacity}))}

function buildScene(){
  if(!scene)return;
  if(world){scene.remove(world);world.traverse(o=>{if(o.geometry&&!o.userData.sharedGeo)o.geometry.dispose()})}
  world=new THREE.Group();scene.add(world);
  const c=getC(ctype),L=c.L,W=c.W,H=c.H;
  scene.background=new THREE.Color(css("--scene"));
  const wire=css("--wire"),tick=css("--tick"),floorCol=css("--floor"),muted=css("--muted");
  world.position.set(-L/2,0,-W/2);
  const ground=new THREE.Mesh(new THREE.PlaneGeometry(L+700,W+700),new THREE.MeshBasicMaterial({color:floorCol}));
  ground.rotation.x=-Math.PI/2;ground.position.set(L/2,-.6,W/2);world.add(ground);
  const cfloor=new THREE.Mesh(new THREE.PlaneGeometry(L,W),new THREE.MeshBasicMaterial({color:new THREE.Color(floorCol).lerp(new THREE.Color(wire),.12)}));
  cfloor.rotation.x=-Math.PI/2;cfloor.position.set(L/2,-.3,W/2);world.add(cfloor);
  const box=new THREE.LineSegments(new THREE.EdgesGeometry(new THREE.BoxGeometry(L,H,W)),new THREE.LineBasicMaterial({color:wire}));
  box.position.set(L/2,H/2,W/2);world.add(box);
  const ribs=[];for(let x=30;x<L;x+=30)ribs.push([x,0,0],[x,H,0]);
  world.add(line(ribs,wire,.12));
  world.add(line([[L,0,0],[L,H,0],[L,H,0],[L,H,W],[L,H,W],[L,0,W],[L,0,W/2],[L,H,W/2]],tick,.6));
  if($("optScale").checked){
    const t=[];
    for(let x=0;x<=L;x+=50){const big=x%100===0;t.push([x,0,W],[x,0,W+(big?22:12)])}
    t.push([0,0,W],[L,0,W]);
    for(let z=0;z<=W;z+=50)t.push([0,0,z],[-14,0,z]);
    for(let y=0;y<=H;y+=50)t.push([L,y,W],[L+12,y,W]);
    world.add(line(t,tick,.85));
    for(let x=0;x<L;x+=100){const s=sprite(String(x),{size:21,fg:"#fff",bg:tick});s.position.set(x,0,W+40);world.add(s)}
    const sl=sprite(`${L} cm`,{size:21,fg:"#fff",bg:tick,bold:true});sl.position.set(L,0,W+40);world.add(sl);
    const zs=[];for(let z=0;z<=W-25;z+=50)zs.push(z);zs.push(W);
    zs.forEach(v=>{const s=sprite(String(v),{size:18,fg:"#fff",bg:tick});s.position.set(-40,0,v);world.add(s)});
    for(let y=50;y<=H-25;y+=50){const s=sprite(String(y),{size:18,fg:"#fff",bg:tick});s.position.set(L+34,y,W);world.add(s)}
    const sh=sprite(`${H}`,{size:18,fg:"#fff",bg:tick,bold:true});sh.position.set(L+34,H,W);world.add(sh);
    const fr=sprite("FRONT WALL",{size:22,fg:muted});fr.position.set(-20,H+22,W/2);world.add(fr);
    const dr=sprite("DOOR",{size:22,fg:muted});dr.position.set(L+20,H+22,W/2);world.add(dr);
  }
  cargoGroup=new THREE.Group();world.add(cargoGroup);
  buildCargo();
}
function buildCargo(){
  if(!cargoGroup)return;
  setHover(null);cargoGroup.clear();unitMeshes=[];
  const k=result&&result.list[cnum],empty=$("emptyMsg");
  if(!k||!k.units.length){
    empty.hidden=false;
    empty.textContent=cargo.length?"Nothing fits in this container type. See the rows listed under Cargo.":"No cargo yet. Fill in quantity and W × D × H on the Input sheet.";
    updateSeq(0,true);return;
  }
  empty.hidden=true;
  const labels=$("optLabels").checked,ghost=$("optGhost").checked;
  const edgeMat=new THREE.LineBasicMaterial({color:0x111111,transparent:true,opacity:ghost?.5:.75});
  const gap=.8;
  k.units.forEach(u=>{
    const gk=[u.sx,u.sy,u.sz].join("x");
    let g=geoCache.get(gk);if(!g){g={box:new THREE.BoxGeometry(u.sx-gap,u.sy-gap,u.sz-gap)};g.edges=new THREE.EdgesGeometry(g.box);geoCache.set(gk,g)}
    const mats=[faceMat(u.item,u.sz,u.sy,labels,ghost),faceMat(u.item,u.sz,u.sy,labels,ghost),faceMat(u.item,u.sx,u.sz,labels,ghost),faceMat(u.item,u.sx,u.sz,false,ghost),faceMat(u.item,u.sx,u.sy,labels,ghost),faceMat(u.item,u.sx,u.sy,labels,ghost)];
    const m=new THREE.Mesh(g.box,mats);m.userData={u,sharedGeo:true};
    m.position.set(u.x+u.sx/2,u.y+u.sy/2,u.z+u.sz/2);
    const e=new THREE.LineSegments(g.edges,edgeMat);e.userData.sharedGeo=true;m.add(e);
    cargoGroup.add(m);unitMeshes.push(m);
  });
  updateSeq(unitMeshes.length,true);
}

/* ---------- camera ---------- */
let tween=null;
function setView(v,instant){
  view=v;document.querySelectorAll("#viewSeg button").forEach(b=>b.setAttribute("aria-pressed",String(b.dataset.view===v)));
  if(!camera)return;
  const c=getC(ctype),L=c.L,W=c.W,H=c.H,span=Math.max(L,W*2.2);
  const tan=Math.tan(THREE.MathUtils.degToRad(camera.fov/2)),asp=Math.max(camera.aspect,.5);
  const fit=span*1.12/(2*tan*Math.min(asp,1.8));
  let pos,tgt=new THREE.Vector3(0,H*.2,0);
  if(v==="top"){pos=new THREE.Vector3(0,fit*1.15,1);tgt=new THREE.Vector3(0,0,0)}
  else if(v==="side"){pos=new THREE.Vector3(0,H*.5,fit*1.08);tgt=new THREE.Vector3(0,H*.5,0)}
  else if(v==="door"){const d=Math.max(W,H)/(2*tan*Math.min(asp,1))*1.5;pos=new THREE.Vector3(L/2+d,H*.55,0);tgt=new THREE.Vector3(0,H*.45,0)}
  else if(camera.aspect<1.15){const d=Math.max(L,W*2)*1.25/Math.sqrt(asp);pos=new THREE.Vector3(d*.88,d*.5,d*.34);tgt=new THREE.Vector3(L*.06,H*.1,0)}  // narrow pane: look down the length from the door corner
  else pos=new THREE.Vector3(L*.28,fit*.6,fit*.92);
  if(instant){tween=null;camera.position.copy(pos);controls.target.copy(tgt);controls.update();return}
  tween={t0:performance.now(),dur:matchMedia("(prefers-reduced-motion: reduce)").matches?1:600,p0:camera.position.clone(),p1:pos,q0:controls.target.clone(),q1:tgt};
}
function stepTween(){const k=Math.min(1,(performance.now()-tween.t0)/tween.dur),e=k<.5?2*k*k:1-Math.pow(-2*k+2,2)/2;camera.position.lerpVectors(tween.p0,tween.p1,e);controls.target.lerpVectors(tween.q0,tween.q1,e);if(k>=1)tween=null}

/* ---------- hover ---------- */
function onHover(ev){
  const r=renderer.domElement.getBoundingClientRect();
  pointer.x=((ev.clientX-r.left)/r.width)*2-1;pointer.y=-((ev.clientY-r.top)/r.height)*2+1;
  raycaster.setFromCamera(pointer,camera);
  const hit=raycaster.intersectObjects(unitMeshes.filter(m=>m.visible),false)[0];
  setHover(hit?hit.object:null,ev.clientX-r.left,ev.clientY-r.top);
}
function setHover(m,x,y){
  if(hovered&&hovered!==m){hovered.children[0].material=hovered.userData.prevEdge;hovered.scale.set(1,1,1)}
  if(!m){hovered=null;tip.hidden=true;return}
  if(hovered!==m){m.userData.prevEdge=m.children[0].material;m.children[0].material=new THREE.LineBasicMaterial({color:css("--accent")});m.scale.set(1.012,1.012,1.012)}
  hovered=m;const u=m.userData.u,it=cargo[u.item];
  tip.innerHTML=`<b>#${u.item+1} ${esc(it.name||"Model")}</b>Unit ${u.unitNo} of ${u.qty} · ${u.orient}<div class="mono">Row ${u.row}, layer ${u.layer}<br>${Math.round(u.x)}–${Math.round(u.x+u.sx)} cm from front wall<br>W${it.w} × D${it.l} × H${it.h} cm${it.kg?` · ${it.kg} kg`:""}</div>`;
  tip.hidden=false;
  tip.style.left=Math.max(4,Math.min(x+12,stage.clientWidth-tip.offsetWidth-6))+"px";
  tip.style.top=Math.max(6,y-tip.offsetHeight-10)+"px";
}

/* ---------- loading sequence ---------- */
const seq=$("seq"),seqOut=$("seqOut"),playBtn=$("playBtn");let playing=null;
function updateSeq(n,resetMax){if(resetMax)seq.max=unitMeshes.length;seq.value=n;applySeq()}
function applySeq(){const n=+seq.value;unitMeshes.forEach((m,i)=>m.visible=i<n);seqOut.textContent=`${n} / ${unitMeshes.length}`}
seq.addEventListener("input",()=>{stopPlay();applySeq()});
function stopPlay(){if(playing){cancelAnimationFrame(playing);playing=null}playBtn.textContent="▶ Play"}
playBtn.addEventListener("click",()=>{
  if(playing){stopPlay();return}
  const total=unitMeshes.length;if(!total)return;
  if(+seq.value>=total)seq.value=0;
  const start=performance.now(),from=+seq.value,dur=Math.min(9000,Math.max(3000,total*45));
  playBtn.textContent="❚❚ Pause";
  const step=t=>{const n=Math.min(total,Math.round(from+(total-from)*(t-start)/dur));seq.value=n;applySeq();if(n<total)playing=requestAnimationFrame(step);else stopPlay()};
  playing=requestAnimationFrame(step);
});

/* ---------- side panels ---------- */
function recompute(){
  result=pack(cargo,getC(ctype));
  if(cnum>=result.list.length)cnum=0;
  renderControls();renderSide();buildScene();
}
function chooseType(t){
  ctype=t;cnum=0;matCache.clear();recompute();setView(view);
  if(inExcel&&hasLoadingPlan)writeSelection(t);
}
function renderControls(){
  const seg=$("ctypeSeg");seg.innerHTML="";
  containers.forEach(c=>{const b=document.createElement("button");b.type="button";b.setAttribute("aria-pressed",String(c.type===ctype));b.textContent=c.type;b.onclick=()=>chooseType(c.type);seg.appendChild(b)});
  const cs=$("cnumSeg"),n=result.list.length;cs.innerHTML="";$("cnumGrp").hidden=n<2;
  result.list.forEach((k,i)=>{const b=document.createElement("button");b.type="button";b.setAttribute("aria-pressed",String(i===cnum));b.innerHTML=`#${i+1}<small>${k.units.length}</small>`;b.onclick=()=>{cnum=i;renderControls();renderSide();buildCargo()};cs.appendChild(b)});
  const c=getC(ctype);
  $("bayTitle").innerHTML=`${esc(c.type)} · Container ${cnum+1}${n>1?` of ${n}`:""}<span class="dims">${c.L} × ${c.W} × ${c.H} cm inside</span>`;
}
function meter(label,val,max,unit,d=0){const p=max>0?val/max:0;return`<div class="meter"><div class="r"><span>${label}</span><span class="v">${fmt(val,d)} / ${fmt(max,d)} ${unit}</span></div><div class="bar"><i class="${p>1?"over":""}" style="width:${Math.min(100,p*100).toFixed(1)}%"></i></div></div>`}
function renderSide(){
  const c=getC(ctype),k=result.list[cnum]||{units:[],x:0,kg:0,vol:0,items:new Map()};
  let head,sub;
  if(!result.needed){head="No cargo loaded";sub=result.noFit?`${result.noFit} model(s) do not fit a ${c.type}.`:"Enter cargo on the Input sheet."}
  else{head=`${result.needed} × ${c.type}`;sub=result.needed===1?"Everything fits in one container.":`Floor length needs ${result.list.length}${result.byWeight>result.list.length?`, weight needs ${result.byWeight}`:""}.`;if(result.noFit)sub+=` ${result.noFit} model(s) do not fit at all.`}
  $("verdict").innerHTML=`<span class="label">Containers needed</span><strong>${head}</strong><span>${sub}</span>`;
  $("meters").innerHTML=meter("Floor length used",k.x,c.L,"cm")+meter("Cargo volume",k.vol/1e6,c.L*c.W*c.H/1e6,"CBM",1)+meter("Weight",k.kg,c.payload,"kg")+
    `<div class="meter"><div class="r"><span>Units in this container</span><span class="v">${fmt(k.units.length)}</span></div></div>`+
    (k.kg>c.payload?`<div style="margin-top:8px"><span class="pill bad">Over max payload</span></div>`:"");
  const cmp=$("cmp");cmp.innerHTML="";
  containers.forEach(ct=>{
    const r=ct.type===ctype?result:pack(cargo,ct);
    const used=r.list.reduce((s,q)=>s+q.vol,0)/1e6,cap=ct.L*ct.W*ct.H/1e6*Math.max(1,r.needed);
    const status=r.noFit?`<span class="pill bad">${r.noFit} won't fit</span>`:r.needed===1?`<span class="pill ok">Fits in one</span>`:r.needed>1?`<span class="pill warn">Needs ${r.needed}</span>`:"";
    const b=document.createElement("button");b.type="button";b.setAttribute("aria-pressed",String(ct.type===ctype));
    b.innerHTML=`<span class="t">${esc(ct.type)}</span><span class="n">${r.needed||"–"}</span><span class="s">${status} ${r.needed?`${fmt(used/cap*100)}% of volume used`:""}</span>`;
    b.onclick=()=>chooseType(ct.type);cmp.appendChild(b);
  });
  const lg=$("legend");lg.innerHTML="";
  if(!k.items.size)lg.innerHTML=`<li style="grid-template-columns:1fr" class="muted">No units in this container.</li>`;
  k.items.forEach((n,idx)=>{const li=document.createElement("li");li.innerHTML=`<span class="sw" style="background:${COLORS[idx%COLORS.length]}"></span><span class="name">#${idx+1} ${esc(cargo[idx].name||"Model")}</span><span class="q">${n} pcs</span>`;lg.appendChild(li)});
  $("nofit").innerHTML=result.plans.map((p,i)=>p.fit?"":`<div><span class="pill bad">Does not fit</span> #${i+1} ${esc(cargo[i].name)} is ${p.why} for a ${esc(c.type)}</div>`).join("");
}
$("viewSeg").addEventListener("click",e=>{const b=e.target.closest("button");if(b)setView(b.dataset.view)});
["optLabels","optGhost","optScale"].forEach(id=>$(id).addEventListener("change",()=>{if(id==="optScale")buildScene();else buildCargo()}));

/* ---------- Excel integration ---------- */
const findSheet=(ws,re)=>ws.items.find(s=>re.test(s.name));
async function readWorkbook(){
  try{
    await Excel.run(async ctx=>{
      const sheets=ctx.workbook.worksheets;sheets.load("items/name");await ctx.sync();
      const inp=findSheet(sheets,/^input$/i)||findSheet(sheets,/input/i)||ctx.workbook.worksheets.getActiveWorksheet();
      const con=findSheet(sheets,/^containers?$/i);
      const lp=findSheet(sheets,/loading plan/i);
      // anchor every read at A1 so column positions match the sheet
      const fromA1=ws=>ws.getRange("A1").getBoundingRect(ws.getUsedRange());
      const r1=fromA1(inp);r1.load("values");
      const r2=con?fromA1(con):null;if(r2)r2.load("values");
      const r3=lp?lp.getRange("C3"):null;if(r3)r3.load("values");
      await ctx.sync();
      hasLoadingPlan=!!lp;
      applyData(r1.values,r2?r2.values:null,r3?String(r3.values[0][0]).trim():null,"this workbook");
    });
  }catch(err){setStatus("Couldn't read the workbook: "+(err&&err.message||err),true)}
}
async function writeSelection(t){
  try{await Excel.run(async ctx=>{const sheets=ctx.workbook.worksheets;sheets.load("items/name");await ctx.sync();const lp=findSheet(sheets,/loading plan/i);if(lp){lp.getRange("C3").values=[[t]];await ctx.sync()}})}
  catch(err){/* leave the workbook as it is */}
}
let refreshT=null;
const scheduleRefresh=()=>{clearTimeout(refreshT);refreshT=setTimeout(readWorkbook,450)};
async function watchWorkbook(){
  try{
    await Excel.run(async ctx=>{
      const sheets=ctx.workbook.worksheets;sheets.load("items/name");await ctx.sync();
      sheets.items.filter(s=>/input|container|loading plan/i.test(s.name)).forEach(s=>s.onChanged.add(async()=>{scheduleRefresh()}));
      await ctx.sync();
    });
  }catch(err){/* events need ExcelApi 1.7; the Refresh button still works */}
}
function applyOfficeTheme(){
  try{
    const th=Office.context.officeTheme;if(!th)return;
    const hex=String(th.bodyBackgroundColor||"#ffffff").replace("#","");
    const n=parseInt(hex,16);const l=(.2126*(n>>16&255)+.7152*(n>>8&255)+.0722*(n&255))/255;
    document.documentElement.setAttribute("data-theme",l<.4?"dark":"light");
  }catch(e){}
}
$("refreshBtn").addEventListener("click",()=>{if(inExcel)readWorkbook();else $("file").click()});

/* ---------- browser fallback: import a file ---------- */
$("importBtn").addEventListener("click",()=>$("file").click());
$("file").addEventListener("change",e=>{
  const f=e.target.files[0];if(!f)return;e.target.value="";
  const r=new FileReader();
  r.onload=()=>{try{
    const wb=XLSX.read(new Uint8Array(r.result),{type:"array"});
    const get=re=>{const n=wb.SheetNames.find(s=>re.test(s));return n?XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true,defval:""}):null};
    const lp=wb.SheetNames.find(s=>/loading plan/i.test(s));
    const sel=lp&&wb.Sheets[lp].C3?String(wb.Sheets[lp].C3.v).trim():null;
    applyData(get(/^input$/i)||get(/./),get(/^containers?$/i),sel,f.name);setView(view);
  }catch(err){setStatus("Couldn't read that file. Save it as .xlsx and try again.",true)}};
  r.readAsArrayBuffer(f);
});

/* ---------- boot ---------- */
function boot(host){
  initScene();
  inExcel=host==="Excel";
  if(inExcel){applyOfficeTheme();recompute();setView("iso",true);readWorkbook().then(()=>setView("iso",true));watchWorkbook()}
  else{
    $("browserCard").hidden=false;$("refreshBtn").textContent="Import";
    if(matchMedia("(prefers-color-scheme: dark)").matches)document.documentElement.setAttribute("data-theme","dark");
    setStatus("Not connected to Excel. Import the workbook file to preview.");
    recompute();setView("iso",true);
  }
}
if(window.Office&&Office.onReady)Office.onReady(info=>boot(info&&info.host)).catch?.(()=>boot(null));
else boot(null);
})();
