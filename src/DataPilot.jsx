import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend,
} from "recharts";

/* ── Design Tokens ─────────────────────────────────────────────────────────── */
const C = {
  bg:       "#000008",
  s1:       "#07070f",
  s2:       "#0d0d1c",
  s3:       "#141428",
  border:   "rgba(255,255,255,0.07)",
  borderHi: "rgba(255,255,255,0.14)",
  accent:   "#6366f1",
  accentHi: "#818cf8",
  emerald:  "#10b981",
  sky:      "#38bdf8",
  amber:    "#f59e0b",
  rose:     "#f43f5e",
  violet:   "#a78bfa",
  text1:    "#ffffff",
  text2:    "#a1a1c2",
  text3:    "#4a4a72",
};

const SERIES = ["#6366f1","#10b981","#38bdf8","#f59e0b","#f43f5e","#a78bfa","#34d399","#7dd3fc"];

/* ── Global CSS ────────────────────────────────────────────────────────────── */
const GLOBAL_CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700;800&family=DM+Sans:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500;600&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  html { scroll-behavior: smooth; }
  body { background: ${C.bg}; color: ${C.text2}; font-family: 'DM Sans', sans-serif; min-height: 100vh; overflow-x: hidden; }
  ::-webkit-scrollbar { width: 4px; }
  ::-webkit-scrollbar-track { background: ${C.bg}; }
  ::-webkit-scrollbar-thumb { background: ${C.s3}; border-radius: 4px; }
  @keyframes fadeUp   { from { opacity:0; transform:translateY(24px); } to { opacity:1; transform:translateY(0); } }
  @keyframes fadeIn   { from { opacity:0; } to { opacity:1; } }
  @keyframes float    { 0%,100% { transform:translateY(0) scale(1); } 50% { transform:translateY(-28px) scale(1.04); } }
  @keyframes float2   { 0%,100% { transform:translateY(0) scale(1); } 50% { transform:translateY(22px) scale(0.97); } }
  @keyframes pulse    { 0%,100% { opacity:.6; } 50% { opacity:1; } }
  @keyframes spin     { to { transform:rotate(360deg); } }
  @keyframes shimmer  { 0% { background-position:-600px 0; } 100% { background-position:600px 0; } }
  @keyframes gradMove { 0%,100% { background-position:0% 50%; } 50% { background-position:100% 50%; } }
  @keyframes borderPulse { 0%,100% { border-color:rgba(99,102,241,0.3); } 50% { border-color:rgba(99,102,241,0.7); } }
  .fade-up    { animation: fadeUp .55s cubic-bezier(.16,1,.3,1) both; }
  .fade-up-2  { animation: fadeUp .55s cubic-bezier(.16,1,.3,1) .12s both; }
  .fade-up-3  { animation: fadeUp .55s cubic-bezier(.16,1,.3,1) .24s both; }
  .fade-up-4  { animation: fadeUp .55s cubic-bezier(.16,1,.3,1) .36s both; }
  .btn        { display:inline-flex; align-items:center; gap:8px; border:none; cursor:pointer; font-family:'DM Sans',sans-serif; font-weight:600; transition:all .18s ease; white-space:nowrap; text-decoration:none; }
  .btn:disabled { opacity:.35; cursor:not-allowed; }
  .inp        { width:100%; background:${C.s1}; border:1.5px solid ${C.border}; border-radius:10px; padding:10px 14px; color:${C.text1}; font-family:'DM Sans',sans-serif; font-size:13px; outline:none; transition:border-color .15s; }
  .inp:focus  { border-color:${C.accent}; }
  .inp::placeholder { color:${C.text3}; }
  .card       { background:${C.s2}; border:1px solid ${C.border}; border-radius:16px; }
  .tag        { display:inline-flex; align-items:center; gap:4px; font-size:10px; font-weight:700; letter-spacing:.7px; padding:3px 9px; border-radius:6px; font-family:'JetBrains Mono',monospace; }
  .feat-card:hover { border-color:rgba(99,102,241,0.35) !important; transform:translateY(-3px); }
  .nav-link:hover { color:${C.text1} !important; }
  .pill-btn:hover { background:rgba(99,102,241,0.12) !important; color:${C.accentHi} !important; }
`;

/* ── File Parsers ──────────────────────────────────────────────────────────── */
function parseDelimited(text, delim=",") {
  const [header,...lines] = text.trim().split("\n");
  const cols = header.split(delim).map(h=>h.trim().replace(/^"|"$/g,""));
  return lines.filter(Boolean).map(line=>{
    const vals=line.split(delim).map(v=>v.trim().replace(/^"|"$/g,""));
    const o={}; cols.forEach((c,i)=>{ const n=parseFloat(vals[i]); o[c]=isNaN(n)?vals[i]:n; }); return o;
  });
}
function parseJSON(text) {
  const p=JSON.parse(text);
  if (Array.isArray(p)) return p;
  const arr=Object.values(p).find(v=>Array.isArray(v));
  if (arr) return arr; throw new Error("No array in JSON");
}
function parseExcel(buf) {
  const wb=XLSX.read(buf,{type:"array"}); const ws=wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws,{defval:""});
}
function detectDelim(text) {
  const first=text.split("\n")[0];
  if ((first.match(/\t/g)||[]).length>(first.match(/,/g)||[]).length) return "\t";
  if ((first.match(/;/g)||[]).length>(first.match(/,/g)||[]).length) return ";";
  return ",";
}
async function parseFile(file) {
  const ext=file.name.split(".").pop().toLowerCase();
  return new Promise((res,rej)=>{
    if (["xlsx","xls","ods"].includes(ext)) {
      const r=new FileReader(); r.onload=e=>{ try{res(parseExcel(new Uint8Array(e.target.result)));}catch(err){rej(err);} }; r.onerror=rej; r.readAsArrayBuffer(file);
    } else {
      const r=new FileReader(); r.onload=e=>{ try{ const t=e.target.result; if(ext==="json") res(parseJSON(t)); else if(ext==="tsv") res(parseDelimited(t,"\t")); else res(parseDelimited(t,detectDelim(t))); }catch(err){rej(err);} }; r.onerror=rej; r.readAsText(file);
    }
  });
}
const FILE_ICON=ext=>({ xlsx:"📗",xls:"📗",ods:"📗",json:"📋",tsv:"📑" }[ext]||"📄");
function exportCSV(rows,name="results.csv") {
  if (!rows?.length) return;
  const cols=Object.keys(rows[0]);
  const csv=[cols.join(","),...rows.map(r=>cols.map(c=>{ const v=r[c]; return typeof v==="string"&&v.includes(",") ?`"${v}"`:v; }).join(","))].join("\n");
  Object.assign(document.createElement("a"),{href:URL.createObjectURL(new Blob([csv],{type:"text/csv"})),download:name}).click();
}
function fmt(v){ return typeof v==="number"?v.toLocaleString():String(v??"—"); }

/* ── SQL Formatter + Highlighter ───────────────────────────────────────────── */
function formatSQL(sql) {
  if (!sql) return sql;
  let s=sql.replace(/\s+/g," ").trim();
  ["SELECT","FROM","WHERE","GROUP BY","ORDER BY","HAVING","LIMIT","LEFT JOIN","RIGHT JOIN","INNER JOIN","JOIN","UNION"].forEach(c=>{ s=s.replace(new RegExp(`\\b(${c})\\b`,"g"),`\n$1`); });
  s=s.replace(/\(SELECT/gi,"\n  (SELECT").replace(/\b(AND|OR)\b/g,"\n  $1").trim().replace(/;*$/,";");
  return s;
}
function highlightSQL(sql) {
  return sql
    .replace(/'([^']*)'/g,`<em style="color:#34d399;font-style:normal">'$1'</em>`)
    .replace(/(?<![:#"'\`a-zA-Z-])\b(\d+)\b(?![a-zA-Z"%])/g,`<em style="color:#f59e0b;font-style:normal">$1</em>`)
    .replace(/\b(SELECT|FROM|WHERE|GROUP BY|ORDER BY|HAVING|LIMIT|JOIN|LEFT JOIN|RIGHT JOIN|INNER JOIN|ON|AS|AND|OR|NOT|IN|LIKE|BETWEEN|IS|NULL|SUM|AVG|COUNT|MIN|MAX|DISTINCT|DESC|ASC|BY|ROUND|WITH|CASE|WHEN|THEN|ELSE|END)\b/g,
      `<strong style="color:#a78bfa;font-weight:600">$1</strong>`);
}

/* ── Badge ─────────────────────────────────────────────────────────────────── */
function Badge({children,color=C.accent,bg}){
  return <span className="tag" style={{color,background:bg||color+"18",border:`1px solid ${color}30`}}>{children}</span>;
}

/* ── Chart ─────────────────────────────────────────────────────────────────── */
const TTP={contentStyle:{background:C.s3,border:`1px solid ${C.border}`,borderRadius:10,color:C.text1,fontSize:12,fontFamily:"DM Sans"},cursor:{fill:C.s3}};
const AX={tick:{fill:C.text3,fontSize:11},axisLine:{stroke:C.border},tickLine:false};
function Chart({data:cd}){
  if(!cd?.data?.length) return null;
  const {type,data,xKey,yKey,yKeys,title}=cd;
  const keys=yKeys?.length?yKeys:yKey?[yKey]:[];
  if(!keys.length||!xKey) return null;
  return(
    <div style={{marginTop:24}}>
      <p style={{fontSize:11,fontWeight:600,color:C.text3,marginBottom:12,letterSpacing:.5,textTransform:"uppercase"}}>{title}</p>
      <ResponsiveContainer width="100%" height={220}>
        {type==="pie"?(
          <PieChart><Pie data={data} dataKey={keys[0]} nameKey={xKey} cx="50%" cy="50%" outerRadius={85} paddingAngle={3}>
            {data.map((_,i)=><Cell key={i} fill={SERIES[i%SERIES.length]}/>)}
          </Pie><Tooltip {...TTP}/><Legend/></PieChart>
        ):type==="bar"||type==="grouped-bar"?(
          <BarChart data={data}><CartesianGrid strokeDasharray="2 4" stroke={C.border} vertical={false}/>
            <XAxis dataKey={xKey} {...AX}/><YAxis {...AX}/>
            <Tooltip {...TTP}/>{keys.length>1&&<Legend/>}
            {keys.map((k,i)=><Bar key={k} dataKey={k} fill={SERIES[i%SERIES.length]} radius={[6,6,0,0]} maxBarSize={48}/>)}
          </BarChart>
        ):(
          <LineChart data={data}><CartesianGrid strokeDasharray="2 4" stroke={C.border} vertical={false}/>
            <XAxis dataKey={xKey} {...AX}/><YAxis {...AX}/>
            <Tooltip {...TTP}/>{keys.length>1&&<Legend/>}
            {keys.map((k,i)=><Line key={k} type="monotone" dataKey={k} stroke={SERIES[i%SERIES.length]} strokeWidth={2.5} dot={false} activeDot={{r:5,strokeWidth:0}}/>)}
          </LineChart>
        )}
      </ResponsiveContainer>
    </div>
  );
}

/* ── SQL View ──────────────────────────────────────────────────────────────── */
function SQLView({sql_mysql,sql_postgresql}){
  const [tab,setTab]=useState("mysql");
  const [cp,setCp]=useState(false);
  const active=tab==="mysql"?sql_mysql:sql_postgresql;
  const copy=()=>{ navigator.clipboard.writeText(active); setCp(true); setTimeout(()=>setCp(false),2000); };
  const TABS=[{id:"mysql",label:"🐬 MySQL",color:"#f59e0b"},{id:"postgresql",label:"🐘 PostgreSQL",color:"#818cf8"}];
  return(
    <div style={{marginBottom:20}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
        <div style={{display:"flex",gap:3,background:C.s1,border:`1px solid ${C.border}`,borderRadius:10,padding:3}}>
          {TABS.map(d=>(
            <button key={d.id} className="btn" onClick={()=>{setTab(d.id);setCp(false);}}
              style={{borderRadius:7,padding:"5px 14px",background:tab===d.id?C.s3:"transparent",border:tab===d.id?`1px solid ${C.border}`:"1px solid transparent",color:tab===d.id?d.color:C.text3,fontSize:12,fontWeight:tab===d.id?600:400}}>
              {d.label}
            </button>
          ))}
        </div>
        <button className="btn" onClick={copy} style={{background:cp?"rgba(16,185,129,.12)":C.s3,border:`1px solid ${cp?C.emerald:C.border}`,borderRadius:8,padding:"5px 12px",color:cp?C.emerald:C.text2,fontSize:11}}>
          {cp?"✓ Copied":"⧉ Copy"}
        </button>
      </div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:8}}>
        <Badge color={tab==="mysql"?"#f59e0b":"#818cf8"}>{tab==="mysql"?"BACKTICK QUOTING · MySQL 5.7+":"DOUBLE-QUOTE QUOTING · PostgreSQL 12+"}</Badge>
        <div style={{display:"flex",gap:10}}>
          {[{color:"#a78bfa",label:"keyword"},{color:"#34d399",label:"'string'"},{color:"#f59e0b",label:"number"}].map(l=>(
            <span key={l.label} style={{display:"flex",alignItems:"center",gap:4,fontSize:10,color:C.text3}}>
              <span style={{width:7,height:7,borderRadius:2,background:l.color,display:"inline-block"}}/>
              {l.label}
            </span>
          ))}
        </div>
      </div>
      <div style={{background:C.s1,border:`1px solid ${C.border}`,borderRadius:12,padding:"16px 18px",fontFamily:"'JetBrains Mono',monospace",fontSize:12.5,lineHeight:1.9,color:"#c8d0e0",overflowX:"auto",whiteSpace:"pre-wrap"}}
        dangerouslySetInnerHTML={{__html:highlightSQL(active||"")}}/>
    </div>
  );
}

/* ── Python View ───────────────────────────────────────────────────────────── */
function PyView({code,explanation,output,packages}){
  const [cp,setCp]=useState(false);
  const [cpO,setCpO]=useState(false);
  const copy=()=>{ navigator.clipboard.writeText(code); setCp(true); setTimeout(()=>setCp(false),2000); };
  const copyOut=()=>{ navigator.clipboard.writeText(output||""); setCpO(true); setTimeout(()=>setCpO(false),2000); };
  const openColab=()=>{ window.open(`https://colab.research.google.com/#create=true`,"_blank"); };
  const h=raw=>raw
    .replace(/\b(import|from|as|def|return|for|in|if|elif|else|with|lambda|class|try|except|and|or|not|is|None|True|False|print|len|range|sorted|sum|min|max|round)\b/g,`<span style="color:#a78bfa;font-weight:500">$1</span>`)
    .replace(/\b(pd|np|df)\b/g,`<span style="color:#38bdf8;font-weight:600">$1</span>`)
    .replace(/(#[^\n]*)/g,`<span style="color:#4a5568;font-style:italic">$1</span>`)
    .replace(/(".*?"|'.*?')/g,`<span style="color:#34d399">$1</span>`)
    .replace(/(?<![a-zA-Z])\b(\d+\.?\d*)\b/g,`<span style="color:#f59e0b">$1</span>`)
    .replace(/\b(groupby|agg|merge|sort_values|reset_index|fillna|dropna|describe|value_counts|pivot_table|mean|std|median|corr|cumsum|pct_change|apply|head|tail|copy|astype)\b/g,`<span style="color:#f87171">$1</span>`);
  return(
    <div style={{marginBottom:20}}>
      {packages?.length>0&&(
        <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:12,flexWrap:"wrap"}}>
          <span style={{fontSize:10,color:C.text3,fontWeight:700,letterSpacing:.5}}>REQUIRES</span>
          {packages.map(p=><Badge key={p} color={C.sky}>{p}</Badge>)}
        </div>
      )}
      <div style={{marginBottom:14}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
          <Badge color={C.amber}>PYTHON · PANDAS</Badge>
          <div style={{display:"flex",gap:6}}>
            <button className="btn" onClick={openColab} style={{borderRadius:8,padding:"5px 12px",background:"rgba(255,111,0,0.08)",border:"1px solid rgba(255,111,0,0.25)",color:"#ff6f00",fontSize:11}}>▶ Open in Colab</button>
            <button className="btn" onClick={copy} style={{background:cp?"rgba(16,185,129,.12)":C.s3,border:`1px solid ${cp?C.emerald:C.border}`,borderRadius:8,padding:"5px 12px",color:cp?C.emerald:C.text2,fontSize:11}}>
              {cp?"✓ Copied":"⧉ Copy"}
            </button>
          </div>
        </div>
        <div style={{background:C.s1,border:`1px solid ${C.border}`,borderRadius:12,overflow:"hidden"}}>
          <div style={{display:"flex"}}>
            <div style={{background:C.bg,borderRight:`1px solid ${C.border}`,padding:"16px 12px",userSelect:"none",minWidth:40,textAlign:"right"}}>
              {(code||"").split("\n").map((_,i)=><div key={i} style={{fontFamily:"'JetBrains Mono',monospace",fontSize:11,lineHeight:1.9,color:C.text3}}>{i+1}</div>)}
            </div>
            <div style={{padding:"16px 18px",fontFamily:"'JetBrains Mono',monospace",fontSize:12.5,lineHeight:1.9,color:"#c8d0e0",overflowX:"auto",flex:1,whiteSpace:"pre-wrap"}}
              dangerouslySetInnerHTML={{__html:h(code||"")}}/>
          </div>
        </div>
      </div>
      {explanation&&<div style={{background:C.s1,borderLeft:`3px solid ${C.amber}`,borderRadius:"0 10px 10px 0",padding:"12px 16px",marginBottom:14,fontSize:13,color:C.text2,lineHeight:1.7}}>{explanation}</div>}
      {output&&(
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
            <Badge color={C.emerald}>OUTPUT</Badge>
            <button className="btn" onClick={copyOut} style={{background:cpO?"rgba(16,185,129,.12)":C.s3,border:`1px solid ${cpO?C.emerald:C.border}`,borderRadius:8,padding:"5px 12px",color:cpO?C.emerald:C.text2,fontSize:11}}>
              {cpO?"✓ Copied":"⧉ Copy"}
            </button>
          </div>
          <div style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:12,padding:"14px 18px",fontFamily:"'JetBrains Mono',monospace",fontSize:12,lineHeight:1.8,color:C.emerald,whiteSpace:"pre-wrap",overflowX:"auto"}}>{output}</div>
        </div>
      )}
    </div>
  );
}

/* ── Results Table ─────────────────────────────────────────────────────────── */
function Table({rows,filename}){
  if(!rows?.length) return null;
  const cols=Object.keys(rows[0]);
  return(
    <div style={{marginBottom:20}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
        <div style={{display:"flex",alignItems:"center",gap:8}}><Badge color={C.sky}>RESULTS</Badge><span style={{fontSize:12,color:C.text3}}>{rows.length} rows</span></div>
        <button className="btn" onClick={()=>exportCSV(rows,filename||"results.csv")} style={{background:C.s3,border:`1px solid ${C.border}`,borderRadius:8,padding:"5px 12px",color:C.emerald,fontSize:11}}>↓ Export CSV</button>
      </div>
      <div style={{borderRadius:12,border:`1px solid ${C.border}`,overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead>
              <tr style={{background:C.s1}}>
                {cols.map(c=><th key={c} style={{padding:"10px 16px",textAlign:"left",color:C.text3,fontSize:10,fontWeight:700,letterSpacing:.8,borderBottom:`1px solid ${C.border}`,fontFamily:"'JetBrains Mono',monospace"}}>{c.toUpperCase()}</th>)}
              </tr>
            </thead>
            <tbody>
              {rows.map((row,i)=>(
                <tr key={i} style={{background:i%2===0?C.s2:"transparent",transition:"background .1s"}}
                  onMouseOver={e=>e.currentTarget.style.background=C.s3} onMouseOut={e=>e.currentTarget.style.background=i%2===0?C.s2:"transparent"}>
                  {cols.map(c=><td key={c} style={{padding:"10px 16px",color:typeof row[c]==="number"?C.text1:C.text2,fontSize:13,borderBottom:`1px solid ${C.border}30`,fontFamily:typeof row[c]==="number"?"'JetBrains Mono',monospace":"inherit",whiteSpace:"nowrap"}}>{fmt(row[c])}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ── Skeleton ──────────────────────────────────────────────────────────────── */
function Skeleton(){
  const b=(w,h=14)=><div style={{width:w,height:h,borderRadius:6,background:`linear-gradient(90deg,${C.s3} 25%,${C.border} 50%,${C.s3} 75%)`,backgroundSize:"600px 100%",animation:"shimmer 1.4s infinite"}}/>;
  return(
    <div className="card fade-up" style={{padding:28,display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"flex",gap:10,alignItems:"center"}}>{b("40%",20)}{b("20%",12)}</div>
      {b("100%",90)}{b("60%")}{b("80%")}{b("70%")}
    </div>
  );
}

/* ── DB Modal ──────────────────────────────────────────────────────────────── */
function DBModal({onData,onConn,onClose}){
  const [type,setType]=useState("postgresql");
  const [f,setF]=useState({host:"localhost",port:"",db:"",user:"",pass:"",table:"",proxy:"http://localhost:3001"});
  const [step,setStep]=useState("form");
  const [tables,setTables]=useState([]);
  const [loading,setLoading]=useState(false);
  const [err,setErr]=useState(null);
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const acc=type==="mysql"?"#f59e0b":C.accentHi;
  const connect=async()=>{
    setLoading(true);setErr(null);
    try{
      const r=await fetch(`${f.proxy}/connect`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({dbType:type,host:f.host,port:f.port||(type==="mysql"?"3306":"5432"),database:f.db,user:f.user,password:f.pass})});
      const d=await r.json(); if(!r.ok) throw new Error(d.error||"Connection failed");
      setTables(d.tables||[]); setStep("tables");
    }catch(e){setErr(e.message);}finally{setLoading(false);}
  };
  const load=async()=>{
    if(!f.table) return; setLoading(true);setErr(null);
    try{
      const r=await fetch(`${f.proxy}/table`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({dbType:type,host:f.host,port:f.port,database:f.db,user:f.user,password:f.pass,table:f.table})});
      const d=await r.json(); if(!r.ok) throw new Error(d.error||"Failed");
      onData(d.rows,`${type}:${f.db}.${f.table}`); onConn({dbType:type,...f}); onClose();
    }catch(e){setErr(e.message);}finally{setLoading(false);}
  };
  const Inp=({label,k,ph,t="text"})=>(
    <div style={{display:"flex",flexDirection:"column",gap:5}}>
      <label style={{fontSize:11,fontWeight:600,color:C.text3,letterSpacing:.5}}>{label}</label>
      <input className="inp" value={f[k]} onChange={e=>set(k,e.target.value)} placeholder={ph} type={t}/>
    </div>
  );
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,8,.85)",backdropFilter:"blur(8px)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:200,padding:16}}>
      <div className="card fade-up" style={{width:"100%",maxWidth:460,padding:28,position:"relative",maxHeight:"90vh",overflowY:"auto"}}>
        <button className="btn" onClick={onClose} style={{position:"absolute",top:14,right:16,background:"transparent",border:"none",color:C.text3,fontSize:18,padding:4}}>✕</button>
        <div style={{marginBottom:22}}>
          <div style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:18,color:C.text1,marginBottom:4}}>Connect Database</div>
          <div style={{fontSize:13,color:C.text3}}>Pull live data directly into DataPilot</div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:22}}>
          {[{id:"postgresql",icon:"🐘",label:"PostgreSQL"},{id:"mysql",icon:"🐬",label:"MySQL"}].map(d=>(
            <button key={d.id} className="btn" onClick={()=>{setType(d.id);setErr(null);}}
              style={{justifyContent:"center",flexDirection:"column",gap:3,padding:"12px",borderRadius:12,background:type===d.id?acc+"15":"transparent",border:`1.5px solid ${type===d.id?acc:C.border}`,color:type===d.id?C.text1:C.text3,fontSize:13,fontWeight:type===d.id?700:400}}>
              <span style={{fontSize:22}}>{d.icon}</span>{d.label}
            </button>
          ))}
        </div>
        {step==="form"&&(
          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            <div style={{background:C.s1,border:`1px solid ${C.border}`,borderRadius:10,padding:"10px 14px",fontSize:12,color:C.text3}}>
              ℹ Run <code style={{background:C.s3,padding:"1px 6px",borderRadius:4,fontFamily:"'JetBrains Mono',monospace",fontSize:11,color:C.text2}}>node server.js</code> first, then connect.
            </div>
            <Inp label="PROXY URL" k="proxy" ph="http://localhost:3001"/>
            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:10}}><Inp label="HOST" k="host" ph="localhost"/><Inp label="PORT" k="port" ph={type==="mysql"?"3306":"5432"}/></div>
            <Inp label="DATABASE" k="db" ph="mydb"/>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}><Inp label="USERNAME" k="user" ph={type==="mysql"?"root":"postgres"}/><Inp label="PASSWORD" k="pass" ph="••••••••" t="password"/></div>
            {err&&<div style={{background:"rgba(244,63,94,.08)",border:"1px solid rgba(244,63,94,.25)",borderRadius:10,padding:"10px 14px",fontSize:12,color:"#f87171"}}>⚠ {err}</div>}
            <button className="btn" onClick={connect} disabled={loading||!f.db||!f.user}
              style={{marginTop:4,justifyContent:"center",background:acc,borderRadius:10,padding:"12px",color:"#fff",fontSize:14,fontWeight:600}}>
              {loading?"Connecting…":`Connect to ${type==="mysql"?"MySQL":"PostgreSQL"}`}
            </button>
          </div>
        )}
        {step==="tables"&&(
          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            <div style={{background:"rgba(16,185,129,.08)",border:"1px solid rgba(16,185,129,.25)",borderRadius:10,padding:"10px 14px",fontSize:12,color:C.emerald,fontWeight:500}}>✓ Connected to {f.db}</div>
            {tables.length>0&&(
              <div>
                <label style={{fontSize:11,fontWeight:600,color:C.text3,letterSpacing:.5,display:"block",marginBottom:8}}>SELECT TABLE</label>
                <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                  {tables.slice(0,16).map(t=>(
                    <button key={t} className="btn" onClick={()=>set("table",t)}
                      style={{borderRadius:8,padding:"5px 12px",background:f.table===t?acc+"18":"transparent",border:`1px solid ${f.table===t?acc:C.border}`,color:f.table===t?C.text1:C.text2,fontSize:12}}>
                      {t}
                    </button>
                  ))}
                </div>
              </div>
            )}
            <Inp label="OR TYPE TABLE NAME" k="table" ph="table_name"/>
            {err&&<div style={{background:"rgba(244,63,94,.08)",border:"1px solid rgba(244,63,94,.25)",borderRadius:10,padding:"10px 14px",fontSize:12,color:"#f87171"}}>{err}</div>}
            <div style={{display:"grid",gridTemplateColumns:"auto 1fr",gap:8}}>
              <button className="btn" onClick={()=>setStep("form")} style={{padding:"11px 16px",borderRadius:10,background:"transparent",border:`1px solid ${C.border}`,color:C.text2,fontSize:13}}>←</button>
              <button className="btn" onClick={load} disabled={loading||!f.table}
                style={{justifyContent:"center",background:acc,borderRadius:10,padding:"11px",color:"#fff",fontSize:13,fontWeight:600}}>
                {loading?"Loading…":"Load Table →"}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/* LANDING PAGE COMPONENTS                                                     */
/* ═══════════════════════════════════════════════════════════════════════════ */

function Nav({onLaunch}){
  const [scrolled,setScrolled]=useState(false);
  useEffect(()=>{
    const h=()=>setScrolled(window.scrollY>30);
    window.addEventListener("scroll",h); return()=>window.removeEventListener("scroll",h);
  },[]);
  return(
    <nav style={{position:"fixed",top:0,left:0,right:0,zIndex:100,padding:"0 32px",height:64,display:"flex",alignItems:"center",justifyContent:"space-between",background:scrolled?"rgba(0,0,8,.85)":"transparent",backdropFilter:scrolled?"blur(20px)":"none",borderBottom:scrolled?`1px solid ${C.border}`:"none",transition:"all .3s ease"}}>
      <div style={{display:"flex",alignItems:"center",gap:10}}>
        <div style={{width:30,height:30,borderRadius:8,background:`linear-gradient(135deg,${C.accent},${C.sky})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>⚡</div>
        <span style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:16,color:C.text1,letterSpacing:-.3}}>DataPilot</span>
        <span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:9,color:C.accent,background:C.accent+"18",border:`1px solid ${C.accent}30`,padding:"2px 7px",borderRadius:5,letterSpacing:.5}}>v4</span>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:4}}>
        {["Features","SQL","Python","Database"].map(l=>(
          <a key={l} className="nav-link" href="#features" style={{padding:"6px 14px",fontSize:13,color:C.text3,textDecoration:"none",borderRadius:8,transition:"color .15s",fontWeight:500}}>{l}</a>
        ))}
      </div>
      <button className="btn" onClick={onLaunch}
        style={{background:`linear-gradient(135deg,${C.accent},${C.accentHi})`,borderRadius:10,padding:"9px 20px",color:"#fff",fontSize:13,fontWeight:600,boxShadow:`0 0 24px ${C.accent}40`}}>
        Launch App →
      </button>
    </nav>
  );
}

function Hero({onLaunch}){
  const QUERIES=["Show revenue by month where returns are above average","Calculate month-over-month growth rate in Pandas","Which platform had the highest net earnings in Q4?","Compare orders vs customers across all months"];
  const [qi,setQi]=useState(0);
  const [displayed,setDisplayed]=useState("");
  const [typing,setTyping]=useState(true);
  useEffect(()=>{
    const target=QUERIES[qi]; let i=0; setDisplayed(""); setTyping(true);
    const iv=setInterval(()=>{
      if(i<target.length){ setDisplayed(target.slice(0,i+1)); i++; }
      else { setTyping(false); clearInterval(iv); setTimeout(()=>setQi(p=>(p+1)%QUERIES.length),2000); }
    },38);
    return()=>clearInterval(iv);
  },[qi]);

  return(
    <section style={{minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"120px 24px 80px",position:"relative",overflow:"hidden",textAlign:"center"}}>
      {/* Orbs */}
      <div style={{position:"absolute",top:"15%",left:"10%",width:500,height:500,borderRadius:"50%",background:`radial-gradient(circle,${C.accent}18 0%,transparent 70%)`,animation:"float 8s ease-in-out infinite",pointerEvents:"none"}}/>
      <div style={{position:"absolute",bottom:"10%",right:"8%",width:420,height:420,borderRadius:"50%",background:`radial-gradient(circle,${C.sky}14 0%,transparent 70%)`,animation:"float2 10s ease-in-out infinite",pointerEvents:"none"}}/>
      <div style={{position:"absolute",top:"50%",left:"50%",width:600,height:300,transform:"translate(-50%,-50%)",background:`radial-gradient(ellipse,${C.violet}08 0%,transparent 70%)`,pointerEvents:"none"}}/>
      {/* Grid */}
      <div style={{position:"absolute",inset:0,backgroundImage:`linear-gradient(${C.border} 1px,transparent 1px),linear-gradient(90deg,${C.border} 1px,transparent 1px)`,backgroundSize:"48px 48px",pointerEvents:"none",opacity:.4}}/>

      {/* Badge */}
      <div className="fade-up" style={{display:"inline-flex",alignItems:"center",gap:8,background:"rgba(99,102,241,.1)",border:`1px solid ${C.accent}35`,borderRadius:100,padding:"6px 18px",marginBottom:32}}>
        <span style={{width:6,height:6,borderRadius:"50%",background:C.accent,animation:"pulse 2s infinite"}}/>
        <span style={{fontSize:11,fontWeight:600,color:C.accentHi,letterSpacing:2,fontFamily:"'JetBrains Mono',monospace"}}>AI ANALYTICS ENGINE</span>
      </div>

      {/* Headline */}
      <h1 className="fade-up-2" style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:"clamp(40px,7vw,76px)",lineHeight:1.08,color:C.text1,letterSpacing:-2,marginBottom:8,maxWidth:820}}>
        Turn Data Into
      </h1>
      <h1 className="fade-up-2" style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:"clamp(40px,7vw,76px)",lineHeight:1.08,letterSpacing:-2,marginBottom:24,background:`linear-gradient(135deg,${C.accent},${C.sky},${C.violet})`,backgroundSize:"200% 200%",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",animation:"gradMove 4s ease infinite",backgroundClip:"text"}}>
        Decisions. Instantly.
      </h1>

      {/* Subheadline */}
      <p className="fade-up-3" style={{fontSize:"clamp(15px,2vw,19px)",color:C.text2,lineHeight:1.7,maxWidth:580,marginBottom:40,fontWeight:300}}>
        The AI analytics engine that speaks <span style={{color:C.text1,fontWeight:500}}>SQL</span>, <span style={{color:C.text1,fontWeight:500}}>Python</span>, and plain English. Connect your database, upload your files, get answers in seconds.
      </p>

      {/* CTAs */}
      <div className="fade-up-4" style={{display:"flex",gap:12,flexWrap:"wrap",justifyContent:"center",marginBottom:56}}>
        <button className="btn" onClick={onLaunch}
          style={{background:`linear-gradient(135deg,${C.accent},${C.accentHi})`,borderRadius:12,padding:"14px 28px",color:"#fff",fontSize:15,fontWeight:700,boxShadow:`0 0 40px ${C.accent}45`,fontFamily:"'Sora',sans-serif"}}>
          Start Analyzing →
        </button>
        <button className="btn" onClick={onLaunch}
          style={{background:"transparent",border:`1.5px solid ${C.border}`,borderRadius:12,padding:"14px 24px",color:C.text2,fontSize:15,fontWeight:500}}>
          View Demo ↓
        </button>
      </div>

      {/* Live query typewriter */}
      <div className="fade-up-4" style={{width:"100%",maxWidth:640,background:C.s1,border:`1px solid ${C.border}`,borderRadius:14,padding:"16px 20px",textAlign:"left",animation:"borderPulse 3s ease-in-out infinite"}}>
        <div style={{fontSize:10,color:C.text3,fontWeight:600,letterSpacing:1.5,marginBottom:10,fontFamily:"'JetBrains Mono',monospace"}}>LIVE QUERY</div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{color:C.accent,fontFamily:"'JetBrains Mono',monospace",fontSize:14}}>⚡</span>
          <span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:13,color:C.text2,lineHeight:1.5}}>
            {displayed}<span style={{display:"inline-block",width:2,height:14,background:C.accent,marginLeft:2,animation:typing?"pulse .6s infinite":"none",verticalAlign:"middle"}}/>
          </span>
        </div>
      </div>
    </section>
  );
}

function StatsBar(){
  const stats=[
    {value:"3",label:"Query Modes"},
    {value:"5+",label:"File Formats"},
    {value:"2",label:"Databases"},
    {value:"∞",label:"Questions"},
  ];
  return(
    <div style={{borderTop:`1px solid ${C.border}`,borderBottom:`1px solid ${C.border}`,padding:"28px 32px",display:"flex",justifyContent:"center",gap:0,flexWrap:"wrap",background:C.s1}}>
      {stats.map((s,i)=>(
        <div key={s.label} style={{flex:"1 1 140px",textAlign:"center",padding:"8px 24px",borderRight:i<stats.length-1?`1px solid ${C.border}`:"none"}}>
          <div style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:32,color:C.text1,letterSpacing:-1}}>{s.value}</div>
          <div style={{fontSize:12,color:C.text3,fontWeight:500,marginTop:2,letterSpacing:.5}}>{s.label}</div>
        </div>
      ))}
    </div>
  );
}

function Features({onLaunch}){
  const feats=[
    {icon:"✦",color:C.accent,title:"AI Insights",desc:"Ask any business question in plain English and get a headline, chart, and actionable recommendation instantly.",tag:"MOST POPULAR"},
    {icon:"{ }",color:C.sky,title:"SQL Query",desc:"Plain English becomes production-ready MySQL and PostgreSQL queries with proper syntax, formatting, and results table.",tag:"SQL"},
    {icon:"⬡",color:C.amber,title:"Python Code",desc:"Generate runnable Pandas code with line numbers, syntax highlighting, output preview, and one-click Google Colab integration.",tag:"PYTHON"},
    {icon:"🗄",color:C.emerald,title:"Live Database",desc:"Connect to MySQL or PostgreSQL directly via the backend proxy. Run queries on real data, not samples.",tag:"DATABASE"},
    {icon:"📁",color:C.violet,title:"Any File Format",desc:"CSV, Excel (.xlsx/.xls), JSON, TSV, and TXT — drag, drop, or upload multiple files and switch between them instantly.",tag:"FILES"},
    {icon:"↓",color:C.rose,title:"Export Everything",desc:"Every query result, chart dataset, and analysis can be exported as a clean CSV with one click.",tag:"EXPORT"},
  ];
  return(
    <section id="features" style={{padding:"96px 32px",maxWidth:1100,margin:"0 auto"}}>
      <div style={{textAlign:"center",marginBottom:64}}>
        <div style={{fontFamily:"'JetBrains Mono',monospace",fontSize:10,fontWeight:700,color:C.accent,letterSpacing:3,marginBottom:16}}>FEATURES</div>
        <h2 style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:"clamp(28px,4vw,44px)",color:C.text1,letterSpacing:-1.5,marginBottom:16}}>Everything you need to analyze data</h2>
        <p style={{fontSize:16,color:C.text3,maxWidth:480,margin:"0 auto",lineHeight:1.7}}>Three modes, five file formats, two databases. One clean interface.</p>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(300px,1fr))",gap:16}}>
        {feats.map((f,i)=>(
          <div key={i} className="feat-card" style={{background:C.s2,border:`1px solid ${C.border}`,borderRadius:16,padding:"28px",cursor:"default",transition:"all .2s ease",position:"relative",overflow:"hidden"}}>
            <div style={{position:"absolute",top:0,right:0,width:120,height:120,background:`radial-gradient(circle,${f.color}10 0%,transparent 70%)`}}/>
            <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:18}}>
              <div style={{width:44,height:44,borderRadius:12,background:f.color+"15",border:`1px solid ${f.color}30`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,color:f.color,fontFamily:"'JetBrains Mono',monospace"}}>
                {f.icon}
              </div>
              <span style={{fontSize:9,fontWeight:700,color:f.color,background:f.color+"15",border:`1px solid ${f.color}25`,padding:"3px 8px",borderRadius:5,letterSpacing:1,fontFamily:"'JetBrains Mono',monospace"}}>{f.tag}</span>
            </div>
            <div style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:17,color:C.text1,marginBottom:10,letterSpacing:-.3}}>{f.title}</div>
            <div style={{fontSize:13,color:C.text3,lineHeight:1.7}}>{f.desc}</div>
          </div>
        ))}
      </div>
      <div style={{textAlign:"center",marginTop:56}}>
        <button className="btn" onClick={onLaunch}
          style={{background:`linear-gradient(135deg,${C.accent},${C.accentHi})`,borderRadius:12,padding:"14px 32px",color:"#fff",fontSize:15,fontWeight:700,boxShadow:`0 0 40px ${C.accent}40`,fontFamily:"'Sora',sans-serif"}}>
          Try DataPilot Free →
        </button>
      </div>
    </section>
  );
}

function Footer(){
  return(
    <footer style={{borderTop:`1px solid ${C.border}`,padding:"40px 32px",background:C.s1}}>
      <div style={{maxWidth:1100,margin:"0 auto",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:16}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:26,height:26,borderRadius:7,background:`linear-gradient(135deg,${C.accent},${C.sky})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12}}>⚡</div>
          <span style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:14,color:C.text1}}>DataPilot</span>
          <span style={{fontSize:12,color:C.text3}}>— AI Analytics Engine</span>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:24}}>
          <a href="https://github.com/Ajaypatel06" target="_blank" rel="noreferrer" style={{fontSize:12,color:C.text3,textDecoration:"none",fontWeight:500,transition:"color .15s"}} onMouseOver={e=>e.target.style.color=C.text1} onMouseOut={e=>e.target.style.color=C.text3}>GitHub</a>
          <a href="https://linkedin.com/in/ajay-patel" target="_blank" rel="noreferrer" style={{fontSize:12,color:C.text3,textDecoration:"none",fontWeight:500,transition:"color .15s"}} onMouseOver={e=>e.target.style.color=C.text1} onMouseOut={e=>e.target.style.color=C.text3}>LinkedIn</a>
        </div>
        <div style={{fontSize:11,color:C.text3,fontFamily:"'JetBrains Mono',monospace",letterSpacing:.5}}>Built by Ajay Patel · Powered by Claude Sonnet</div>
      </div>
    </footer>
  );
}

/* ═══════════════════════════════════════════════════════════════════════════ */
/* ANALYTICS APP                                                               */
/* ═══════════════════════════════════════════════════════════════════════════ */

const SAMPLE=[
  {month:"Jan",revenue:42000,orders:320,customers:210,returns:18},
  {month:"Feb",revenue:38000,orders:290,customers:195,returns:24},
  {month:"Mar",revenue:51000,orders:410,customers:280,returns:12},
  {month:"Apr",revenue:47000,orders:370,customers:255,returns:20},
  {month:"May",revenue:63000,orders:490,customers:340,returns:9},
  {month:"Jun",revenue:58000,orders:450,customers:310,returns:15},
  {month:"Jul",revenue:71000,orders:530,customers:390,returns:7},
  {month:"Aug",revenue:65000,orders:500,customers:360,returns:11},
];

function AppPage({onBack}){
  const [datasets,setDatasets]=useState([{name:"sample",label:"Sample Data",icon:"📊",rows:SAMPLE}]);
  const [activeIdx,setActiveIdx]=useState(0);
  const [dbCfg,setDbCfg]=useState(null);
  const [q,setQ]=useState("");
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const [history,setHistory]=useState([]);
  const [err,setErr]=useState(null);
  const [mode,setMode]=useState("insight");
  const [showDB,setShowDB]=useState(false);
  const [fileLoading,setFileLoading]=useState(false);
  const fileRef=useRef();

  const active=datasets[activeIdx]||datasets[0];
  const data=active.rows;
  const src=active.name;
  const cols=Object.keys(data[0]||{});

  const handleFiles=useCallback(async e=>{
    const files=Array.from(e.target.files); if(!files.length) return;
    setFileLoading(true); setErr(null);
    const loaded=[];
    for(const file of files){
      try{ const rows=await parseFile(file); if(!rows.length) throw new Error("Empty"); const ext=file.name.split(".").pop().toLowerCase(); loaded.push({name:file.name,label:file.name.replace(/\.[^.]+$/,""),icon:FILE_ICON(ext),rows}); }
      catch(ex){ setErr(`Could not read "${file.name}": ${ex.message}`); }
    }
    if(loaded.length){ setDatasets(prev=>{ const base=prev.filter(d=>d.name==="sample"); const next=[...base,...loaded]; setActiveIdx(next.length-1); return next; }); setResult(null); setHistory([]); setDbCfg(null); }
    setFileLoading(false); e.target.value="";
  },[]);

  const removeDataset=useCallback(idx=>{ setDatasets(prev=>{ const next=prev.filter((_,i)=>i!==idx); if(!next.length) return [{name:"sample",label:"Sample Data",icon:"📊",rows:SAMPLE}]; return next; }); setActiveIdx(prev=>Math.max(0,prev>=idx?prev-1:prev)); setResult(null); },[]);

  const runOnDB=async sql=>{ if(!dbCfg) return null; const r=await fetch(`${dbCfg.proxy}/query`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({dbType:dbCfg.dbType,host:dbCfg.host,port:dbCfg.port,database:dbCfg.db,user:dbCfg.user,password:dbCfg.pass,sql})}); const d=await r.json(); if(!r.ok) throw new Error(d.error); return d.rows; };

  const analyze=useCallback(async()=>{
    if(!q.trim()) return; setLoading(true); setErr(null); setResult(null);
    const preview=JSON.stringify(data.slice(0,50)), colStr=cols.join(", ");
    const isSQL=mode==="sql", isPY=mode==="python";
    const tbl=src.includes(":")?src.split(".").pop():"data";
    const sys=isSQL
      ?`You are a SQL expert. Generate correct production-ready SQL for both MySQL and PostgreSQL.
Table: ${tbl} | Columns: ${colStr}
QUOTING: MySQL→backticks, PostgreSQL→double quotes, strings→single quotes, numbers→no quotes
SIMPLICITY: Use only SELECT FROM WHERE GROUP BY ORDER BY HAVING LIMIT and basic aggregates. No FIELD(), no CASE WHEN for sorting, no window functions, no CTEs. Max 8 lines. End with semicolon.
Respond ONLY with valid JSON:
{"sql_mysql":"SELECT \`col\`\\nFROM \`${tbl}\`\\nORDER BY \`col\` DESC;","sql_postgresql":"SELECT \\"col\\"\\nFROM \\"${tbl}\\"\\nORDER BY \\"col\\" DESC;","explanation":"one sentence","results":[computed rows],"insight":"1-2 sentence finding","chart":{"type":"bar","title":"...","xKey":"...","yKeys":["col1"],"data":[max 10 rows]}}`
      :isPY
      ?`You are DataPilot Python Engine. Generate clean Pandas code.
df is already loaded with columns: ${colStr}
Respond using EXACTLY these delimiters:
<packages>pandas,numpy</packages>
<explanation>one sentence</explanation>
<code>
import pandas as pd
# code here using df
</code>
<output>
what code prints
</output>
Rules: never use read_csv(), always print(), max 20 lines`
      :`You are DataPilot AI Analyst.
Columns: ${colStr}
Respond ONLY with valid JSON:
{"headline":"punchy finding","insight":"2-3 sentence insight","recommendation":"one action","chart":{"type":"bar","title":"...","xKey":"...","yKeys":["col1","col2"],"data":[max 10 rows]}}`;

    try{
      // Routed through the backend proxy so the Anthropic key never reaches
      // the browser bundle. Reuses the same proxy URL as the DB connection
      // (defaults to http://localhost:3001), overridable via REACT_APP_BACKEND_URL.
      const backendUrl=(dbCfg&&dbCfg.proxy)||process.env.REACT_APP_BACKEND_URL||"http://localhost:3001";
      const res=await fetch(`${backendUrl}/analyze`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({system:sys,question:`Columns:${colStr}\nData:${preview}\nRows:${data.length}\nQuestion:${q}`,model:"claude-sonnet-4-20250514",max_tokens:2500})});
      const api=await res.json(); if(!res.ok) throw new Error(api.error||"Analysis request failed");
      const txt=api.content?.find(b=>b.type==="text")?.text||""; if(!txt) throw new Error("Empty response");
      let parsed;
      if(isPY){
        const ex=tag=>{ const m=txt.match(new RegExp(`<${tag}>([\\s\\S]*?)<\\/${tag}>`)); return m?m[1].trim():""; };
        parsed={code:ex("code"),explanation:ex("explanation"),output:ex("output"),packages:ex("packages").split(",").map(p=>p.trim()).filter(Boolean)||["pandas"]};
        if(!parsed.code) throw new Error("No code returned");
      }else{
        parsed=JSON.parse(txt.replace(/```json|```/g,"").trim());
        if(parsed.chart?.yKey&&!parsed.chart?.yKeys) parsed.chart.yKeys=[parsed.chart.yKey];
        if(parsed.chart?.yKeys) parsed.chart.yKey=parsed.chart.yKeys[0];
        const fixMQ=sql=>sql?sql.replace(/"([^"]+)"/g,(m,inner)=>/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(inner)?`\`${inner}\``:m):sql;
        const fixPQ=sql=>sql?sql.replace(/`([^`]+)`/g,(m,inner)=>/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(inner)?`"${inner}"`:m):sql;
        if(parsed.sql_mysql) parsed.sql_mysql=formatSQL(fixMQ(parsed.sql_mysql));
        if(parsed.sql_postgresql) parsed.sql_postgresql=formatSQL(fixPQ(parsed.sql_postgresql));
        if(isSQL&&dbCfg&&(parsed.sql_mysql||parsed.sql_postgresql)){
          const liveSql=dbCfg.dbType==="mysql"?parsed.sql_mysql:parsed.sql_postgresql;
          try{ const live=await runOnDB(liveSql); if(live?.length){parsed.results=live;if(parsed.chart)parsed.chart.data=live.slice(0,10);} }catch(e){parsed.dbWarn=e.message;}
        }
      }
      setResult(parsed); setHistory(p=>[{q,result:parsed,mode},...p.slice(0,6)]); setQ("");
    }catch(e){ setErr("Analysis failed: "+(e.message||"Please try again.")); }
    finally{ setLoading(false); }
  },[q,data,mode,dbCfg,cols,src]);

  const sugg=mode==="sql"
    ?["Show total revenue by month","Find months where returns > average","Compare orders and customers","Revenue per order by month"]
    :mode==="python"
    ?["Calculate month-over-month growth rate","Find correlation between all columns","Show top 3 months by revenue","Calculate revenue per customer"]
    :["Which month had highest revenue?","Compare all metrics over time","Where should we focus next?"];

  const srcLabel=dbCfg?`${dbCfg.dbType==="mysql"?"🐬":"🐘"} ${dbCfg.db}.${dbCfg.table||"—"}`:active.label;
  const MODES=[{id:"insight",icon:"✦",label:"AI Insights",color:C.accent},{id:"sql",icon:"{}",label:"SQL Query",color:C.sky},{id:"python",icon:"⬡",label:"Python Code",color:C.amber}];

  return(
    <div style={{minHeight:"100vh",background:C.bg,color:C.text2,fontFamily:"'DM Sans',sans-serif"}}>
      {showDB&&<DBModal onData={(rows,label)=>{ const ds={name:label,label:label.split(":")[1]||label,icon:"🗄",rows}; setDatasets(prev=>{const next=[...prev,ds];setActiveIdx(next.length-1);return next;}); setResult(null);setHistory([]); }} onConn={setDbCfg} onClose={()=>setShowDB(false)}/>}

      {/* App Header */}
      <div style={{borderBottom:`1px solid ${C.border}`,padding:"0 28px",display:"flex",alignItems:"center",height:60,gap:20,position:"sticky",top:0,zIndex:50,background:"rgba(0,0,8,.9)",backdropFilter:"blur(20px)"}}>
        <button className="btn" onClick={onBack} style={{background:"transparent",border:`1px solid ${C.border}`,borderRadius:8,padding:"6px 12px",color:C.text3,fontSize:12}}>← Home</button>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{width:26,height:26,borderRadius:7,background:`linear-gradient(135deg,${C.accent},${C.sky})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12}}>⚡</div>
          <span style={{fontFamily:"'Sora',sans-serif",fontWeight:800,fontSize:15,color:C.text1}}>DataPilot</span>
        </div>
        <div style={{flex:1}}/>
        <div style={{fontSize:10,color:C.emerald,fontFamily:"'JetBrains Mono',monospace",background:"rgba(16,185,129,.08)",padding:"4px 10px",borderRadius:20,border:"1px solid rgba(16,185,129,.2)"}}>
          ● {data.length} rows · {cols.length} cols
        </div>
        <button className="btn" onClick={()=>setShowDB(true)}
          style={{borderRadius:9,padding:"7px 14px",background:dbCfg?"rgba(16,185,129,.1)":C.s2,border:`1px solid ${dbCfg?"rgba(16,185,129,.4)":C.border}`,color:dbCfg?C.emerald:C.text2,fontSize:12}}>
          🗄 {dbCfg?"Connected":"Connect DB"}
        </button>
        <button className="btn" onClick={()=>fileRef.current.click()}
          style={{borderRadius:9,padding:"7px 14px",background:C.s2,border:`1px solid ${C.border}`,color:fileLoading?C.accent:C.text2,fontSize:12}}>
          {fileLoading?"Loading…":"↑ Upload Files"}
        </button>
        <input ref={fileRef} type="file" multiple accept=".csv,.tsv,.txt,.json,.xlsx,.xls,.ods" onChange={handleFiles} style={{display:"none"}}/>
      </div>

      <div style={{maxWidth:860,margin:"0 auto",padding:"28px 20px 80px"}}>

        {/* Dataset tabs */}
        {datasets.length>1&&(
          <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:20}}>
            {datasets.map((ds,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",background:activeIdx===i?C.s3:C.s2,border:`1px solid ${activeIdx===i?C.borderHi:C.border}`,borderRadius:9,overflow:"hidden",transition:"all .15s"}}>
                <button className="btn" onClick={()=>{setActiveIdx(i);setResult(null);}}
                  style={{padding:"5px 12px",background:"transparent",border:"none",color:activeIdx===i?C.text1:C.text3,fontSize:12,gap:5,fontWeight:activeIdx===i?600:400}}>
                  {ds.icon} {ds.label} <span style={{fontSize:10,color:C.text3}}>({ds.rows.length})</span>
                </button>
                {ds.name!=="sample"&&<button className="btn" onClick={()=>removeDataset(i)} style={{padding:"5px 8px",background:"transparent",border:"none",borderLeft:`1px solid ${C.border}`,color:C.text3,fontSize:11}}>✕</button>}
              </div>
            ))}
          </div>
        )}

        {/* Schema strip */}
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:24,padding:"10px 14px",background:C.s2,border:`1px solid ${C.border}`,borderRadius:10}}>
          <span style={{fontSize:11,color:C.text3,fontWeight:600,letterSpacing:.5,fontFamily:"'JetBrains Mono',monospace"}}>TABLE</span>
          <span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:11,color:C.text2,background:C.s3,padding:"2px 9px",borderRadius:5,border:`1px solid ${C.border}`}}>
            {src.includes(":")?src.split(":")[1]:src==="sample"?"sample_data":active.label}
          </span>
          <span style={{color:C.border}}>·</span>
          {cols.map(c=><span key={c} style={{fontFamily:"'JetBrains Mono',monospace",fontSize:10,color:C.text3,background:C.s2,padding:"2px 7px",borderRadius:4,border:`1px solid ${C.border}`}}>{c}</span>)}
        </div>

        {/* Mode tabs */}
        <div style={{display:"flex",gap:3,marginBottom:20,background:C.s2,border:`1px solid ${C.border}`,borderRadius:13,padding:4,width:"fit-content"}}>
          {MODES.map(m=>(
            <button key={m.id} className="btn" onClick={()=>{setMode(m.id);setResult(null);setQ("");}}
              style={{borderRadius:10,padding:"9px 22px",background:mode===m.id?C.s3:"transparent",border:mode===m.id?`1px solid ${C.border}`:"1px solid transparent",color:mode===m.id?C.text1:C.text3,fontSize:13,fontWeight:mode===m.id?600:400}}>
              <span style={{color:mode===m.id?m.color:"inherit",fontFamily:m.id==="sql"?"'JetBrains Mono',monospace":"inherit"}}>{m.icon}</span> {m.label}
            </button>
          ))}
        </div>

        {/* Input card */}
        <div className="card" style={{padding:22,marginBottom:20}}>
          <div style={{fontSize:10,fontWeight:700,letterSpacing:.8,color:C.text3,marginBottom:12,fontFamily:"'JetBrains Mono',monospace"}}>
            {mode==="sql"?"PLAIN ENGLISH → MYSQL + POSTGRESQL + RESULTS":mode==="python"?"DESCRIBE ANALYSIS → GET RUNNABLE PANDAS CODE":"ASK YOUR DATA ANYTHING"}
          </div>
          <div style={{display:"flex",gap:10}}>
            <input className="inp" value={q} onChange={e=>setQ(e.target.value)} onKeyDown={e=>e.key==="Enter"&&!loading&&analyze()}
              placeholder={mode==="sql"?"e.g. Show revenue by month where returns are above average":mode==="python"?"e.g. Calculate month-over-month growth rate":"e.g. Which month had the best performance?"}
              style={{flex:1}}/>
            <button className="btn" onClick={analyze} disabled={loading||!q.trim()}
              style={{borderRadius:10,padding:"10px 24px",background:loading||!q.trim()?"transparent":mode==="python"?`linear-gradient(135deg,${C.amber},#f97316)`:mode==="sql"?`linear-gradient(135deg,${C.sky},${C.accent})`:`linear-gradient(135deg,${C.accent},${C.accentHi})`,border:`1px solid ${loading||!q.trim()?C.border:"transparent"}`,color:loading||!q.trim()?C.text3:"#fff",fontSize:13,fontWeight:700,minWidth:110,fontFamily:"'Sora',sans-serif",boxShadow:loading||!q.trim()?"none":`0 0 20px ${C.accent}35`}}>
              {loading?"···":mode==="sql"?"Run Query":mode==="python"?"Generate":"Analyze"}
            </button>
          </div>
          <div style={{display:"flex",gap:6,marginTop:12,flexWrap:"wrap"}}>
            {sugg.map((s,i)=>(
              <button key={i} className="btn pill-btn" onClick={()=>setQ(s)}
                style={{borderRadius:20,padding:"4px 12px",background:"transparent",border:`1px solid ${C.border}`,color:C.text3,fontSize:11,fontWeight:400,transition:"all .15s"}}>
                {s}
              </button>
            ))}
          </div>
        </div>

        {loading&&<Skeleton/>}
        {err&&<div style={{background:"rgba(244,63,94,.08)",border:"1px solid rgba(244,63,94,.25)",borderRadius:12,padding:"14px 18px",color:"#f87171",fontSize:13,marginBottom:16}}>⚠ {err}</div>}

        {/* Result */}
        {result&&!loading&&(
          <div className="card fade-up" style={{padding:28,marginBottom:24}}>
            {result.dbWarn&&<div style={{background:`${C.amber}12`,border:`1px solid ${C.amber}30`,borderRadius:10,padding:"10px 14px",fontSize:12,color:C.amber,marginBottom:16}}>⚠ {result.dbWarn}</div>}
            {/* Python */}
            {result.code&&<PyView code={result.code} explanation={result.explanation} output={result.output} packages={result.packages}/>}
            {/* SQL */}
            {(result.sql_mysql||result.sql_postgresql)&&<>
              <SQLView sql_mysql={result.sql_mysql} sql_postgresql={result.sql_postgresql}/>
              <div style={{background:C.s1,borderLeft:`3px solid ${C.accent}`,borderRadius:"0 10px 10px 0",padding:"12px 16px",marginBottom:20,fontSize:13,color:C.text2,lineHeight:1.7}}>{result.explanation}</div>
              <Table rows={result.results} filename="query_results.csv"/>
              <Chart data={result.chart}/>
              {result.insight&&<div style={{marginTop:20,background:"rgba(16,185,129,.06)",border:"1px solid rgba(16,185,129,.2)",borderRadius:12,padding:"14px 18px",fontSize:13,color:C.text2,lineHeight:1.7,display:"flex",gap:10}}><span style={{color:C.emerald,fontWeight:700,flexShrink:0}}>Insight</span>{result.insight}</div>}
            </>}
            {/* AI Insight */}
            {result.headline&&!result.sql_mysql&&!result.code&&<>
              <div style={{fontFamily:"'Sora',sans-serif",fontSize:20,fontWeight:800,color:C.text1,marginBottom:14,lineHeight:1.3,letterSpacing:-.5}}>{result.headline}</div>
              <div style={{background:C.s1,borderLeft:`3px solid ${C.accent}`,borderRadius:"0 10px 10px 0",padding:"14px 16px",marginBottom:20,fontSize:14,color:C.text2,lineHeight:1.8}}>{result.insight}</div>
              <Chart data={result.chart}/>
              {result.chart?.data?.length>0&&<div style={{marginTop:10,textAlign:"right"}}><button className="btn" onClick={()=>exportCSV(result.chart.data,"chart_data.csv")} style={{background:"transparent",border:`1px solid ${C.border}`,borderRadius:8,padding:"5px 12px",color:C.emerald,fontSize:11}}>↓ Export Chart Data</button></div>}
              {result.recommendation&&<div style={{marginTop:20,padding:"16px 18px",background:C.accentHi+"08",border:`1px solid ${C.accent}25`,borderRadius:12,display:"flex",gap:10}}>
                <span style={{color:C.accent,fontWeight:700,fontFamily:"'Sora',sans-serif"}}>→</span>
                <div>
                  <div style={{fontSize:9,color:C.accent,fontFamily:"'JetBrains Mono',monospace",letterSpacing:1,marginBottom:6}}>RECOMMENDED ACTION</div>
                  <div style={{fontSize:13,color:C.text2,lineHeight:1.7}}>{result.recommendation}</div>
                </div>
              </div>}
            </>}
          </div>
        )}

        {/* History */}
        {history.length>1&&(
          <div>
            <div style={{fontSize:10,color:C.text3,fontFamily:"'JetBrains Mono',monospace",letterSpacing:1,marginBottom:10}}>RECENT QUERIES</div>
            <div style={{display:"flex",flexDirection:"column",gap:6}}>
              {history.slice(1).map((h,i)=>(
                <button key={i} className="btn" onClick={()=>setResult(h.result)}
                  style={{width:"100%",justifyContent:"space-between",borderRadius:10,padding:"11px 16px",background:C.s2,border:`1px solid ${C.border}`,color:C.text2,fontSize:13,fontWeight:400,transition:"all .15s"}}
                  onMouseOver={e=>e.currentTarget.style.borderColor=C.borderHi} onMouseOut={e=>e.currentTarget.style.borderColor=C.border}>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <Badge color={h.mode==="sql"?C.sky:h.mode==="python"?C.amber:C.accent}>{h.mode==="sql"?"SQL":h.mode==="python"?"PY":"AI"}</Badge>
                    <span style={{color:C.text2}}>{h.q}</span>
                  </div>
                  <span style={{color:C.text3,fontSize:11}}>view →</span>
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Empty state */}
        {!result&&!loading&&!err&&history.length===0&&(
          <div style={{textAlign:"center",padding:"56px 24px"}}>
            <div style={{width:60,height:60,borderRadius:18,background:`linear-gradient(135deg,${C.accent}20,${C.sky}20)`,border:`1px solid ${C.accent}30`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:26,margin:"0 auto 20px"}}>⚡</div>
            <div style={{fontFamily:"'Sora',sans-serif",fontSize:18,fontWeight:800,color:C.text1,marginBottom:8}}>Ready to analyze</div>
            <div style={{fontSize:13,color:C.text3,marginBottom:28}}>Ask a question, run a SQL query, or generate Python code</div>
            <div style={{display:"flex",justifyContent:"center",gap:8,flexWrap:"wrap"}}>
              {["🐬 MySQL","🐘 PostgreSQL","📗 Excel","📄 CSV","📋 JSON","✦ AI Insights","{ } SQL Query","⬡ Python","↓ Export"].map(f=>(
                <span key={f} style={{fontSize:11,color:C.text3,fontFamily:"'JetBrains Mono',monospace",border:`1px solid ${C.border}`,padding:"4px 10px",borderRadius:20}}>{f}</span>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ── Root ──────────────────────────────────────────────────────────────────── */
export default function DataPilot(){
  const [view,setView]=useState("landing");
  return(
    <>
      <style>{GLOBAL_CSS}</style>
      {view==="landing"?(
        <div>
          <Nav onLaunch={()=>setView("app")}/>
          <Hero onLaunch={()=>setView("app")}/>
          <StatsBar/>
          <Features onLaunch={()=>setView("app")}/>
          <Footer/>
        </div>
      ):(
        <AppPage onBack={()=>setView("landing")}/>
      )}
    </>
  );
}
