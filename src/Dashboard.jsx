import React, { useState, useMemo, useRef, useCallback, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Upload, RefreshCw, FileSpreadsheet, X, MapPin, Trash2, Database } from 'lucide-react';

const LOGO = "data:image/png;base64,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";
const LS_KEY = 'tkab_dashboard_v1';
const C = { ink:'#16243A', navy:'#1F3864', navy2:'#2E5496', bp:'#F3F6FB', paper:'#FFFFFF',
  line:'#CBD5E1', line2:'#E2E8F0', green:'#00B050', amber:'#E8A317', red:'#E5484D',
  muted:'#64748B', grid:'#D7E0EC' };
const EST = {
  'FINAL':{c:'#00B050',t:'#ffffff',label:'Aprobado final (PP)'},
  'INTERNO':{c:'#A9D08E',t:'#16243A',label:'Aprobado interno (PC)'},
  'C/OBS':{c:'#FFD966',t:'#16243A',label:'Aprobado c/obs'},
  'PROCESO':{c:'#F4B183',t:'#16243A',label:'En proceso'},
  'RECHAZADO':{c:'#FF5050',t:'#ffffff',label:'Rechazado (NCR)'},
  'NO INICIADO':{c:'#E7E6E6',t:'#64748B',label:'No iniciado'},
};
const EST_ORDER=['FINAL','INTERNO','C/OBS','PROCESO','RECHAZADO','NO INICIADO'];
const RANK={'NO INICIADO':0,'PROCESO':1,'C/OBS':2,'INTERNO':3,'FINAL':4};

const clean=v=>(v==null?'':String(v).trim());
function numOrNull(v){ if(v==null||v==='')return null; let n=typeof v==='number'?v:parseFloat(String(v).replace('%','').replace(',','.')); if(isNaN(n))return null; if(n>1.5)n=n/100; return Math.max(0,Math.min(1,n)); }
function serialToStr(v){ if(v==null||v==='')return ''; if(typeof v==='number'){ const d=new Date(Date.UTC(1899,11,30)+v*86400000); const p=x=>String(x).padStart(2,'0'); return p(d.getUTCDate())+'-'+p(d.getUTCMonth()+1)+'-'+String(d.getUTCFullYear()).slice(2);} return String(v); }
function protoMeta(n){ if(/Picado/i.test(n))return{o:0,short:'Picado'}; if(/Enfierr/i.test(n))return{o:1,short:'Enfierr.'}; if(/Hormig/i.test(n))return{o:2,short:'Hormigón'}; if(/FRP/i.test(n))return{o:3,short:'FRP'}; return{o:9,short:n}; }
function estadoOf(r){ if(!r)return'NO INICIADO'; if(r.resInt==='Rechazado'||r.resFin==='Rechazado')return'RECHAZADO'; if(r.resFin==='Aprobado')return'FINAL'; if(r.resFin==='Aprob. c/obs'||r.resInt==='Aprob. c/obs')return'C/OBS'; if(r.resInt==='Aprobado')return'INTERNO'; if((r.avance||0)>0)return'PROCESO'; return'NO INICIADO'; }
function heat(p){ if(p==null)return{bg:'#F8FAFC',fg:'#94A3B8'}; p=Math.max(0,Math.min(1,p)); const lerp=(a,b,t)=>a.map((x,i)=>Math.round(x+(b[i]-x)*t)); let c=p<=0.5?lerp([255,255,255],[198,239,206],p/0.5):lerp([198,239,206],[0,176,80],(p-0.5)/0.5); return{bg:'rgb('+c[0]+','+c[1]+','+c[2]+')',fg:p>0.66?'#0B3D1E':'#0F2A1A'}; }
const pctTxt=p=>p==null?'':Math.round(p*100)+'%';

function parse(buf){
  const wb=XLSX.read(buf,{type:'array'});
  const sh=wb.Sheets['REGISTRO'];
  if(!sh)throw new Error("El archivo no contiene la hoja 'REGISTRO'. Usa la planilla del mapa de control TK-AB.");
  const aoa=XLSX.utils.sheet_to_json(sh,{header:1,raw:true,defval:null});
  let hr=-1;
  for(let i=0;i<Math.min(aoa.length,12);i++){ const row=(aoa[i]||[]).map(x=>x==null?'':String(x).trim()); if(row.some(c=>/^Elemento$/i.test(c))&&row.some(c=>/Avance/i.test(c))){hr=i;break;} }
  if(hr<0)throw new Error("No se encontró el encabezado del REGISTRO.");
  const H=aoa[hr].map(x=>x==null?'':String(x).trim());
  const f=p=>H.findIndex(p);
  const ix={ elemento:f(h=>/^Elemento$/i.test(h)), pano:f(h=>/Pa.?[ñn]o/i.test(h)), nivel:f(h=>/Nivel/i.test(h)||/Fila/i.test(h)),
    proto:H.findIndex(h=>/Protocolo/i.test(h)&&!/C[óo]d/i.test(h)), avance:f(h=>/Avance/i.test(h)), ejecuto:f(h=>/Ejecut/i.test(h)),
    ito:f(h=>/ITO/i.test(h)), fInt:f(h=>/Fecha.*Interno/i.test(h)), resInt:f(h=>/Result.*Interno/i.test(h)),
    fFin:f(h=>/Fecha.*Final/i.test(h)), resFin:f(h=>/Result.*Final/i.test(h)), ncr:f(h=>/NCR/i.test(h)), obs:f(h=>/Observ/i.test(h)) };
  const recs=[];
  for(let i=hr+1;i<aoa.length;i++){ const row=aoa[i]; if(!row)continue; const elem=clean(row[ix.elemento]); const proto=clean(row[ix.proto]); if(!elem||!proto)continue;
    recs.push({ elemento:elem, pano:clean(row[ix.pano]), nivel:clean(row[ix.nivel]), protocolo:proto,
      avance:numOrNull(row[ix.avance]), ejecuto:clean(row[ix.ejecuto]), ito:clean(row[ix.ito]), fInt:serialToStr(row[ix.fInt]),
      resInt:clean(row[ix.resInt]), fFin:serialToStr(row[ix.fFin]), resFin:clean(row[ix.resFin]), ncr:clean(row[ix.ncr]), obs:clean(row[ix.obs]) }); }
  if(!recs.length)throw new Error("El REGISTRO no tiene filas de datos.");
  return recs;
}

function Ring({v,size=58,stroke=7,color=C.green}){
  const r=(size-stroke)/2, cx=size/2, circ=2*Math.PI*r, off=circ*(1-(v||0));
  return (<svg width={size} height={size}><circle cx={cx} cy={cx} r={r} fill="none" stroke={C.line2} strokeWidth={stroke}/>
    <circle cx={cx} cy={cx} r={r} fill="none" stroke={color} strokeWidth={stroke} strokeDasharray={circ} strokeDashoffset={off} strokeLinecap="round" transform={'rotate(-90 '+cx+' '+cx+')'}/>
    <text x="50%" y="50%" dominantBaseline="central" textAnchor="middle" style={{fontSize:size*0.26,fontWeight:700,fill:C.ink}}>{Math.round((v||0)*100)+'%'}</text></svg>);
}

function Seg({val,set,opts}){
  return (<div className="inline-flex rounded-lg overflow-hidden" style={{border:'1px solid '+C.line}}>
    {opts.map((o,i)=>(<button key={o.v} onClick={()=>set(o.v)} className="text-xs font-semibold px-3 py-1.5 transition-colors"
      style={{background:val===o.v?C.navy:C.paper, color:val===o.v?'#ffffff':C.muted, borderLeft:i?'1px solid '+C.line:'none'}}>{o.l}</button>))}</div>);
}

function WallGrid({title,avg,accent,panos,niveles,getCell,isSel,onSelect}){
  return (<div style={{display:'inline-block'}}>
    <div className="text-white text-[11px] font-bold uppercase tracking-wide px-2 py-1 flex justify-between gap-3 rounded-t" style={{background:accent}}>
      <span>{title}</span><span style={{opacity:0.9}}>{pctTxt(avg)}</span></div>
    <table style={{borderCollapse:'collapse'}}><tbody>
      <tr><td style={{background:C.navy,border:'1px solid '+C.grid,width:26}}></td>
        {panos.map(p=>(<td key={p} className="text-white text-[8px] font-bold text-center" style={{background:C.navy,border:'1px solid '+C.grid,minWidth:30,padding:'2px'}}>{p}</td>))}</tr>
      {niveles.map(n=>(<tr key={n}>
        <td className="text-[8px] font-bold text-center" style={{background:'#E2E8F0',border:'1px solid '+C.grid,color:C.ink}}>{n}</td>
        {panos.map(p=>{ const s=getCell(p,n); const sel=isSel(p,n);
          return (<td key={p} onClick={()=>onSelect(p,n)} title={title+' · '+p+' · '+n} className="text-[9px] text-center cursor-pointer"
            style={{background:s.bg,color:s.fg,border:sel?'2px solid '+C.ink:'1px solid '+C.grid,height:26,minWidth:30,fontWeight:600,fontVariantNumeric:'tabular-nums'}}>{s.txt}</td>); })}
      </tr>))}
    </tbody></table></div>);
}

function Matrix({elements,protos,matrix,wallAvg}){
  const fmt=v=>Math.round(v*100)+'%';
  return (<div className="overflow-auto"><table className="w-full text-xs" style={{borderCollapse:'collapse'}}><thead>
    <tr><th className="text-left p-1.5" style={{color:C.muted}}>Elemento</th>
      {protos.map(p=>(<th key={p} className="p-1.5 text-center" style={{color:C.muted}}>{protoMeta(p).short}</th>))}
      <th className="p-1.5 text-center" style={{color:C.muted}}>Prom.</th></tr></thead><tbody>
    {elements.map(e=>(<tr key={e}><td className="p-1.5 font-semibold" style={{color:C.ink}}>{e}</td>
      {protos.map(p=>{ const v=matrix[e][p]; const h=heat(v); return (<td key={p} className="p-1.5 text-center font-semibold" style={{background:h.bg,color:h.fg,border:'1px solid '+C.line2,fontVariantNumeric:'tabular-nums'}}>{fmt(v)}</td>); })}
      <td className="p-1.5 text-center font-bold" style={{color:C.ink,fontVariantNumeric:'tabular-nums'}}>{fmt(wallAvg[e])}</td></tr>))}
  </tbody></table></div>);
}

function EstadoBar({estCount,total}){
  const tot=total||1;
  return (<div><div className="flex w-full h-5 rounded-md overflow-hidden" style={{border:'1px solid '+C.line2}}>
    {EST_ORDER.map(k=>{ const w=estCount[k]/tot*100; if(w<=0)return null; return (<div key={k} title={EST[k].label+': '+estCount[k]} style={{width:w+'%',background:EST[k].c}}/>); })}</div>
    <div className="flex flex-wrap gap-x-3 gap-y-1 mt-2">{EST_ORDER.map(k=>(<span key={k} className="inline-flex items-center gap-1 text-[11px]" style={{color:C.muted}}>
      <span className="inline-block w-2.5 h-2.5 rounded-sm" style={{background:EST[k].c}}/>{EST[k].label} <b style={{color:C.ink}}>{estCount[k]}</b></span>))}</div></div>);
}

function ProcBars({protos,procAvg}){
  return (<div className="space-y-2">{protos.map(p=>{ const v=procAvg[p]; return (
    <div key={p} className="flex items-center gap-2"><div className="w-20 text-[11px] font-semibold" style={{color:C.ink}}>{protoMeta(p).short}</div>
      <div className="flex-1 h-3 rounded-full overflow-hidden" style={{background:C.line2}}><div className="h-full rounded-full" style={{width:(v*100)+'%',background:C.green}}/></div>
      <div className="w-10 text-right text-[11px] font-bold" style={{color:C.ink,fontVariantNumeric:'tabular-nums'}}>{Math.round(v*100)+'%'}</div></div>); })}</div>);
}

function Inspector({sel,protos,byKey,onClose}){
  if(!sel)return (<div className="h-full flex flex-col items-center justify-center text-center" style={{color:C.muted}}>
    <MapPin size={22}/><div className="text-sm font-semibold mt-2" style={{color:C.ink}}>Inspeccionar sector</div>
    <div className="text-xs mt-1">Haz clic en cualquier celda del plano para ver el avance y los controles de sus 4 protocolos.</div></div>);
  const list=protos.map(p=>({proto:p, r:byKey.get(sel.elem+'|'+sel.pano+'|'+sel.nivel+'|'+p)||null}));
  return (<div>
    <div className="flex items-center justify-between"><div className="text-sm font-bold" style={{color:C.ink}}>{sel.elem}</div>
      <button onClick={onClose} style={{color:C.muted}}><X size={16}/></button></div>
    <div className="text-[11px] uppercase tracking-wide font-semibold mb-3" style={{color:C.navy2}}>{sel.pano} · {sel.nivel}</div>
    <div className="space-y-2">
      {list.map(({proto,r})=>{ const es=estadoOf(r); const E=EST[es];
        return (<div key={proto} className="rounded-lg p-2" style={{background:'#F8FAFC',border:'1px solid '+C.line2}}>
          <div className="flex items-center justify-between gap-2">
            <span className="text-xs font-semibold" style={{color:C.ink}}>{protoMeta(proto).short}</span>
            <span className="text-xs font-bold" style={{color:C.ink,fontVariantNumeric:'tabular-nums'}}>{r&&r.avance!=null?pctTxt(r.avance):'—'}</span></div>
          <div className="mt-1"><span className="text-[10px] font-bold px-1.5 py-0.5 rounded" style={{background:E.c,color:E.t}}>{E.label}</span></div>
          {r&&(r.resInt||r.fInt)&&<div className="text-[10px] mt-1" style={{color:C.muted}}>PC: {r.resInt||'—'}{r.ejecuto?' · '+r.ejecuto:''}{r.fInt?' · '+r.fInt:''}</div>}
          {r&&(r.resFin||r.fFin)&&<div className="text-[10px]" style={{color:C.muted}}>PP: {r.resFin||'—'}{r.ito?' · ITO '+r.ito:''}{r.fFin?' · '+r.fFin:''}</div>}
          {r&&r.ncr&&<div className="text-[10px] font-semibold" style={{color:C.red}}>NCR: {r.ncr}</div>}
        </div>); })}
    </div></div>);
}

function Header({global,fileName,savedLabel,onReload,onClear}){
  return (<div className="sticky top-0 z-10" style={{background:C.navy,boxShadow:'0 2px 8px rgba(16,36,58,.18)'}}>
    <div className="max-w-[1200px] mx-auto px-4 sm:px-6 py-3 flex items-center gap-4">
      <div className="bg-white rounded-md px-2 py-1 flex items-center" style={{height:38}}><img src={LOGO} alt="SAENS" style={{height:24}}/></div>
      <div className="flex-1 min-w-0"><div className="text-white font-bold leading-tight">Control de Protocolos QA/QC</div>
        <div className="text-[11px] truncate" style={{color:'#A9C0E0'}}>Seguimiento y Control de Avance · Gestión de Protocolos de Calidad</div></div>
      {fileName&&<div className="hidden sm:flex flex-col items-end mr-1">
        <div className="flex items-center gap-1.5 text-[11px] px-2 py-1 rounded-md" style={{background:'rgba(255,255,255,.12)',color:'#ffffff'}}><FileSpreadsheet size={13}/><span className="max-w-[160px] truncate">{fileName}</span></div>
        {savedLabel&&<div className="flex items-center gap-1 text-[10px] mt-0.5" style={{color:'#A9C0E0'}}><Database size={11}/>{savedLabel}</div>}</div>}
      {onReload&&<button onClick={onReload} className="flex items-center gap-1.5 text-xs font-semibold text-white px-3 py-1.5 rounded-md" style={{background:'rgba(255,255,255,.14)'}}><RefreshCw size={14}/>Otra planilla</button>}
      {onClear&&<button onClick={onClear} title="Borrar datos guardados" className="flex items-center gap-1.5 text-xs font-semibold text-white px-2 py-1.5 rounded-md" style={{background:'rgba(229,72,77,.85)'}}><Trash2 size={14}/></button>}
      {global!=null&&<div className="hidden md:block"><Ring v={global} size={46} stroke={6} color="#7CD992"/></div>}
    </div></div>);
}

export default function Dashboard(){
  const [recs,setRecs]=useState(null);
  const [fileName,setFileName]=useState('');
  const [savedAt,setSavedAt]=useState('');
  const [err,setErr]=useState('');
  const [drag,setDrag]=useState(false);
  const [proc,setProc]=useState('Todos');
  const [mode,setMode]=useState('avance');
  const [sel,setSel]=useState(null);
  const inputRef=useRef(null);

  // cargar datos persistidos al iniciar
  useEffect(()=>{
    // 1. localStorage para render instantáneo
    try{ const raw=localStorage.getItem(LS_KEY); if(raw){ const o=JSON.parse(raw); if(o&&Array.isArray(o.recs)&&o.recs.length){ setRecs(o.recs); setFileName(o.fileName||''); setSavedAt(o.savedAt||''); } } }catch(e){}
    // 2. Vercel Blob como fuente compartida (silencioso en dev local)
    fetch('/api/dashboard').then(r=>r.ok?r.json():null).then(data=>{
      if(data&&Array.isArray(data.recs)&&data.recs.length){
        setRecs(data.recs); setFileName(data.fileName||''); setSavedAt(data.savedAt||'');
        try{ localStorage.setItem(LS_KEY, JSON.stringify(data)); }catch(e){}
      }
    }).catch(()=>{});
  },[]);

  const onFile=useCallback(async(file)=>{ try{ setErr(''); const buf=await file.arrayBuffer(); const r=parse(buf); const when=new Date().toISOString();
    setRecs(r); setFileName(file.name); setSavedAt(when); setSel(null); setProc('Todos');
    const payload={recs:r, fileName:file.name, savedAt:when};
    try{ localStorage.setItem(LS_KEY, JSON.stringify(payload)); }catch(e){}
    fetch('/api/dashboard',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}).catch(()=>{});
  }catch(e){ setErr(e.message||String(e)); } },[]);
  const onDrop=useCallback(e=>{ e.preventDefault(); setDrag(false); const file=e.dataTransfer.files&&e.dataTransfer.files[0]; if(file)onFile(file); },[onFile]);
  const clearData=useCallback(()=>{ try{ localStorage.removeItem(LS_KEY); }catch(e){} fetch('/api/dashboard',{method:'DELETE'}).catch(()=>{}); setRecs(null); setFileName(''); setSavedAt(''); setSel(null); setErr(''); },[]);

  const savedLabel = savedAt ? ('Guardado · '+new Date(savedAt).toLocaleString('es-CL',{day:'2-digit',month:'2-digit',year:'2-digit',hour:'2-digit',minute:'2-digit'})) : '';

  const m=useMemo(()=>{
    if(!recs)return null;
    const byKey=new Map(); recs.forEach(r=>byKey.set(r.elemento+'|'+r.pano+'|'+r.nivel+'|'+r.protocolo,r));
    const protos=[...new Set(recs.map(r=>r.protocolo))].sort((a,b)=>protoMeta(a).o-protoMeta(b).o);
    const elements=[...new Set(recs.map(r=>r.elemento))];
    const dims={}; elements.forEach(e=>{ const rs=recs.filter(r=>r.elemento===e); const panos=[...new Set(rs.map(r=>r.pano))].sort(); let niv=[...new Set(rs.map(r=>r.nivel))].sort(); if(niv[0]&&/^N/i.test(niv[0]))niv=niv.reverse(); dims[e]={panos,niveles:niv}; });
    const estCount={}; EST_ORDER.forEach(k=>estCount[k]=0); recs.forEach(r=>{estCount[estadoOf(r)]++;});
    const avgOf=arr=>{ const v=arr.map(r=>r.avance).filter(x=>x!=null); return v.length?v.reduce((a,b)=>a+b,0)/v.length:0; };
    const wallAvg={}; elements.forEach(e=>wallAvg[e]=avgOf(recs.filter(r=>r.elemento===e)));
    const matrix={}; elements.forEach(e=>{ matrix[e]={}; protos.forEach(p=>matrix[e][p]=avgOf(recs.filter(r=>r.elemento===e&&r.protocolo===p))); });
    const procAvg={}; protos.forEach(p=>procAvg[p]=avgOf(recs.filter(r=>r.protocolo===p)));
    return { byKey, protos, elements, dims, estCount, total:recs.length, global:avgOf(recs), wallAvg, matrix, procAvg, pending:estCount['INTERNO'] };
  },[recs]);

  const getVal=(e,p,n)=>{ if(!m)return null; if(proc==='Todos'){ const vs=m.protos.map(pr=>m.byKey.get(e+'|'+p+'|'+n+'|'+pr)).map(r=>r?r.avance:null).filter(x=>x!=null); return vs.length?vs.reduce((a,b)=>a+b,0)/vs.length:null; } const r=m.byKey.get(e+'|'+p+'|'+n+'|'+proc); return r?(r.avance==null?0:r.avance):null; };
  const getEst=(e,p,n)=>{ if(!m)return null; if(proc==='Todos'){ const es=m.protos.map(pr=>m.byKey.get(e+'|'+p+'|'+n+'|'+pr)).filter(Boolean).map(estadoOf); if(!es.length)return null; if(es.includes('RECHAZADO'))return'RECHAZADO'; let min=null,mr=99; es.forEach(x=>{const rk=RANK[x]==null?99:RANK[x]; if(rk<mr){mr=rk;min=x;}}); return min; } const r=m.byKey.get(e+'|'+p+'|'+n+'|'+proc); return r?estadoOf(r):null; };
  const cellStyle=(e,p,n)=>{ if(mode==='estado'){ const es=getEst(e,p,n); if(!es)return{bg:'#F8FAFC',fg:'#94A3B8',txt:''}; return{bg:EST[es].c,fg:EST[es].t,txt:pctTxt(getVal(e,p,n))}; } const v=getVal(e,p,n); const h=heat(v); return{bg:h.bg,fg:h.fg,txt:pctTxt(v)}; };
  const wallProps=(elem,accent)=>({ title:elem, avg:m.wallAvg[elem], accent, panos:m.dims[elem].panos, niveles:m.dims[elem].niveles,
    getCell:(p,n)=>cellStyle(elem,p,n), isSel:(p,n)=>!!(sel&&sel.elem===elem&&sel.pano===p&&sel.nivel===n), onSelect:(p,n)=>setSel({elem,pano:p,nivel:n}) });

  const fileInput=(<input ref={inputRef} type="file" accept=".xlsx,.xlsm" className="hidden" onChange={e=>{const f=e.target.files&&e.target.files[0]; if(f)onFile(f); e.target.value='';}}/>);

  if(!recs){
    return (<div className="min-h-screen w-full flex flex-col" style={{background:C.bp,color:C.ink,fontFamily:'ui-sans-serif,system-ui,Segoe UI,Roboto,Arial'}}>
      <Header global={null}/>
      <div className="flex-1 flex items-center justify-center p-6">
        <div onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={onDrop} onClick={()=>inputRef.current&&inputRef.current.click()}
          className="w-full max-w-xl rounded-2xl p-10 text-center cursor-pointer transition-all" style={{background:C.paper, border:'2px dashed '+(drag?C.navy2:C.line), boxShadow:'0 1px 2px rgba(16,36,58,.06)'}}>
          <div className="mx-auto mb-4 flex items-center justify-center rounded-xl" style={{width:56,height:56,background:'#EAF1FB'}}><Upload size={26} color={C.navy2}/></div>
          <div className="text-lg font-bold" style={{color:C.ink}}>Carga tu planilla de control</div>
          <div className="text-sm mt-1" style={{color:C.muted}}>Arrastra aquí el archivo <b>Mapa_Control_Protocolos_TK-AB.xlsx</b> o haz clic para seleccionarlo.</div>
          <div className="text-xs mt-4" style={{color:C.muted}}>El dashboard lee la hoja REGISTRO y se dibuja solo. Los datos quedan guardados en este navegador para la próxima vez.</div>
          {fileInput}
        </div>
      </div>
      {err&&<div className="px-6 pb-6"><div className="max-w-xl mx-auto text-sm rounded-lg p-3" style={{background:'#FDECEC',color:C.red,border:'1px solid #F5C2C2'}}>{err}</div></div>}
    </div>);
  }

  const kpis=[
    {l:'Avance global', node:<Ring v={m.global}/>},
    {l:'Cerrados (PP)', v:m.estCount['FINAL'], c:EST['FINAL'].c},
    {l:'Pendiente PP', v:m.pending, c:EST['INTERNO'].c},
    {l:'Rechazados (NCR)', v:m.estCount['RECHAZADO'], c:EST['RECHAZADO'].c},
    {l:'No iniciados', v:m.estCount['NO INICIADO'], c:'#94A3B8'},
  ];
  const find=re=>m.elements.find(e=>re.test(e));
  const sur=find(/Sur/i),norte=find(/Norte/i),este=find(/Este/i),oeste=find(/Oeste/i),floor=find(/Piso/i);
  const procOpts=[{v:'Todos',l:'Todos'}].concat(m.protos.map(p=>({v:p,l:protoMeta(p).short})));

  return (<div className="min-h-screen w-full" style={{background:C.bp,color:C.ink,fontFamily:'ui-sans-serif,system-ui,Segoe UI,Roboto,Arial'}}>
    <Header global={m.global} fileName={fileName} savedLabel={savedLabel} onReload={()=>inputRef.current&&inputRef.current.click()} onClear={clearData}/>
    {fileInput}
    <div className="max-w-[1200px] mx-auto px-4 sm:px-6 py-5 space-y-5">

      <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
        {kpis.map((k,i)=>(<div key={i} className="rounded-xl p-3 flex items-center gap-3" style={{background:C.paper,border:'1px solid '+C.line2,boxShadow:'0 1px 2px rgba(16,36,58,.05)'}}>
          {k.node?k.node:<div className="text-2xl font-bold" style={{minWidth:42,color:C.ink,fontVariantNumeric:'tabular-nums'}}>{k.v}</div>}
          <div><div className="text-[11px] uppercase tracking-wide font-semibold" style={{color:C.muted}}>{k.l}</div>
            {k.c&&<div className="mt-1 h-1.5 w-12 rounded-full" style={{background:k.c}}/>}</div>
        </div>))}
      </div>

      <div className="rounded-2xl overflow-hidden" style={{background:C.paper,border:'1px solid '+C.line2,boxShadow:'0 1px 3px rgba(16,36,58,.06)'}}>
        <div className="flex flex-wrap items-center justify-between gap-3 px-4 py-3" style={{borderBottom:'1px solid '+C.line2}}>
          <div><div className="text-sm font-bold uppercase tracking-wide" style={{color:C.ink}}>Plano del estanque · avance por proceso</div>
            <div className="text-[11px]" style={{color:C.muted}}>Muros en paño × nivel, piso en columna × fila. Haz clic en un sector para ver el detalle.</div></div>
          <div className="flex items-center gap-2 flex-wrap">
            <Seg val={mode} set={setMode} opts={[{v:'avance',l:'Avance %'},{v:'estado',l:'Estado'}]}/>
            <Seg val={proc} set={setProc} opts={procOpts}/>
          </div>
        </div>
        <div className="flex flex-col lg:flex-row">
          <div className="flex-1 p-4 overflow-auto">
            <div className="inline-block rounded-xl p-4" style={{background:C.bp,backgroundImage:'linear-gradient('+C.grid+' 1px,transparent 1px),linear-gradient(90deg,'+C.grid+' 1px,transparent 1px)',backgroundSize:'22px 22px',border:'1px solid '+C.grid}}>
              <div className="flex justify-center mb-3">{sur&&<WallGrid {...wallProps(sur,C.navy2)}/>}</div>
              <div className="flex justify-center items-start gap-3">
                {este&&<WallGrid {...wallProps(este,C.navy2)}/>}
                {floor&&<WallGrid {...wallProps(floor,C.navy)}/>}
                {oeste&&<WallGrid {...wallProps(oeste,C.navy2)}/>}
              </div>
              <div className="flex justify-center mt-3">{norte&&<WallGrid {...wallProps(norte,C.navy2)}/>}</div>
            </div>
            <div className="mt-3">
              {mode==='avance'
                ? (<div className="flex items-center gap-2 text-[11px]" style={{color:C.muted}}><span className="font-semibold">Color = % avance</span>
                    <div className="h-3 w-40 rounded" style={{background:'linear-gradient(90deg,#ffffff,#C6EFCE,#00B050)',border:'1px solid '+C.line}}/><span>0%</span><span>100%</span></div>)
                : (<div className="flex flex-wrap items-center gap-2">{EST_ORDER.map(k=>(<span key={k} className="text-[10px] font-semibold px-2 py-0.5 rounded" style={{background:EST[k].c,color:EST[k].t}}>{EST[k].label}</span>))}</div>)}
            </div>
          </div>
          <div className="lg:w-80 p-4 lg:border-l" style={{borderColor:C.line2}}>
            <Inspector sel={sel} protos={m.protos} byKey={m.byKey} onClose={()=>setSel(null)}/>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
        <div className="rounded-2xl p-4" style={{background:C.paper,border:'1px solid '+C.line2,boxShadow:'0 1px 3px rgba(16,36,58,.06)'}}>
          <div className="text-sm font-bold uppercase tracking-wide mb-3" style={{color:C.ink}}>Avance promedio por elemento y proceso</div>
          <Matrix elements={m.elements} protos={m.protos} matrix={m.matrix} wallAvg={m.wallAvg}/>
        </div>
        <div className="rounded-2xl p-4" style={{background:C.paper,border:'1px solid '+C.line2,boxShadow:'0 1px 3px rgba(16,36,58,.06)'}}>
          <div className="text-sm font-bold uppercase tracking-wide mb-3" style={{color:C.ink}}>Estado de los controles</div>
          <EstadoBar estCount={m.estCount} total={m.total}/>
          <div className="text-sm font-bold uppercase tracking-wide mt-5 mb-3" style={{color:C.ink}}>Avance global por proceso</div>
          <ProcBars protos={m.protos} procAvg={m.procAvg}/>
        </div>
      </div>

      <div className="text-[11px] text-center pb-2 leading-5" style={{color:C.muted}}>
        <div>Dashboard de Control Protocolos · QA/QC · Diego Cartes Solorza</div>
        <div>Datos {fileName||'—'} · SAENS Polímeros y Revestimientos Ltda.</div>
      </div>
    </div>
  </div>);
}
