import {useEffect,useMemo,useState} from 'react';
import {Document,Packer,Paragraph,HeadingLevel} from 'docx';
import {jsPDF} from 'jspdf';
import './styles.css';

type Hoja={name:string;cells:Record<string,string|boolean>;merged?:string[];validations?:{range:string;formula:string}[];maxRow:number;maxCol:number};
type Unidad={id:number;title:string;originalWorkbook?:string;sheets:Hoja[]};
type Archivo={plantillas?:Unidad[];app?:string;edits?:Record<string,string>};
const STORE='sap-ef-2027-v2-ediciones';
const abc=(n:number)=>{let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s};
const pos=(s:string):[number,number]=>{const m=s.replaceAll('$','').match(/^([A-Z]+)(\d+)$/);return m?[m[1].split('').reduce((a,c)=>a*26+c.charCodeAt(0)-64,0),Number(m[2])]:[0,0]};
const between=(ref:string,range:string)=>{const [a,b=a]=range.replaceAll('$','').split(':');const [x,y]=pos(ref),[x1,y1]=pos(a),[x2,y2]=pos(b);return x>=x1&&x<=x2&&y>=y1&&y<=y2};
const save=(blob:Blob,name:string)=>{const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000)};
const safe=(s:string)=>s.normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-zA-Z0-9_-]/g,'_').slice(0,60);
const options=(sheet:Hoja,ref:string)=>{
 const rule=(sheet.validations||[]).find(v=>v.range.split(/\s+/).some(r=>between(ref,r)));
 if(!rule)return [];
 const f=rule.formula.trim();if(f.startsWith('"'))return [...new Set(f.replace(/^"|"$/g,'').split(',').map(x=>x.trim()).filter(Boolean))];
 const m=f.match(/(?:^|!)(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)$/);if(!m)return [];
 const [a,b]=m[1].split(':'),[x1,y1]=pos(a),[x2,y2]=pos(b),out:string[]=[];
 for(let r=y1;r<=Math.min(y2,y1+300);r++)for(let c=x1;c<=x2;c++){const value=sheet.cells[abc(c)+r];if(value!==undefined&&String(value).trim())out.push(String(value).trim())}
 return [...new Set(out)];
};
const validate=(obj:unknown):Unidad[]=>{
 const units=Array.isArray(obj)?obj:(obj&&typeof obj==='object'?(obj as Archivo).plantillas:undefined);
 if(!Array.isArray(units)||units.length!==12||!units.every(x=>x&&typeof x.title==='string'&&Array.isArray(x.sheets)&&x.sheets.length>0&&x.sheets.every((s:Hoja)=>s&&s.cells&&typeof s.cells==='object'&&Number.isInteger(s.maxRow))))throw Error('Se necesita el JSON completo de las 12 SAP originales');
 return units as Unidad[];
};
export default function App(){
 const [data,setData]=useState<Unidad[]>([]),[idx,setIdx]=useState(0),[tab,setTab]=useState(0),[mode,setMode]=useState<'ficha'|'todo'>('ficha'),[filter,setFilter]=useState(''),[showEmpty,setShowEmpty]=useState(false),[msg,setMsg]=useState('');
 const [edits,setEdits]=useState<Record<string,string>>(()=>{try{return JSON.parse(localStorage.getItem(STORE)||'{}') as Record<string,string>}catch{return {}}});
 useEffect(()=>{try{localStorage.setItem(STORE,JSON.stringify(edits))}catch{setMsg('No se han podido guardar los cambios en el navegador. Descarga una copia JSON.')}},[edits]);
 const u=data[idx],sheet=u?.sheets[tab];
 const key=(i:number,t:number,c:string)=>`${i}/${t}/${c}`;
 const value=(i:number,t:number,c:string)=>edits[key(i,t,c)]??String(data[i]?.sheets[t]?.cells[c]??'');
 const setCell=(c:string,v:string)=>setEdits(prev=>({...prev,[key(idx,tab,c)]:v}));
 const load=async(f?:File)=>{if(!f)return;try{const obj:unknown=JSON.parse(await f.text()),units=validate(obj);setData(units);setIdx(0);setTab(0);setFilter('');if(obj&&typeof obj==='object'&&!Array.isArray(obj)&&(obj as Archivo).edits)setEdits((obj as Archivo).edits||{});setMsg('12 SAP cargadas. Los desplegables originales están disponibles en sus celdas.')}catch(e){setMsg(e instanceof Error?e.message:'Archivo JSON no compatible')}};
 const exportData=()=>{const plantillas=data.map((unit,i)=>({...unit,sheets:unit.sheets.map((s,t)=>({...s,cells:{...s.cells,...Object.fromEntries(Object.entries(edits).filter(([k])=>k.startsWith(`${i}/${t}/`)).map(([k,v])=>[k.split('/')[2],v]))}}))}));save(new Blob([JSON.stringify({plantillas},null,2)],{type:'application/json'}),'MIS_12_SAP_2027_EDITADAS.json')};
 const exportBackup=()=>save(new Blob([JSON.stringify({app:'SAP-EF-2027',version:2,plantillas:data,edits},null,2)],{type:'application/json'}),'COPIA_COMPLETA_SAP_EF_2027.json');
 const entries=()=>{if(!sheet)return [];const result:{c:string;v:string}[]=[];for(let r=1;r<=sheet.maxRow;r++)for(let c=1;c<=Math.min(sheet.maxCol,mode==='ficha'?4:sheet.maxCol);c++){const ref=abc(c)+r,v=value(idx,tab,ref);if(v.trim())result.push({c:ref,v})}return result};
 const word=async()=>{if(!u)return;const doc=new Document({sections:[{children:[new Paragraph({text:`SAP ${u.id} · ${u.title}`,heading:HeadingLevel.HEADING_1}),new Paragraph(`Hoja: ${sheet.name}`),...entries().map(({c,v})=>new Paragraph({text:`${c}: ${v}`}))]}]});save(await Packer.toBlob(doc),`SAP_${u.id}_${safe(u.title)}.docx`)};
 const pdf=()=>{if(!u)return;const d=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});let y=17;d.setFontSize(14);d.text(`SAP ${u.id} - ${safe(u.title).replaceAll('_',' ')}`,13,y);y+=10;d.setFontSize(9);for(const {c,v} of entries()){const lines=d.splitTextToSize(`${c}: ${v}`,266);if(y+lines.length*4.5>190){d.addPage();y=17}d.text(lines,13,y);y+=lines.length*4.5+2}d.save(`SAP_${u.id}_${safe(u.title)}.pdf`)};
 const rows=useMemo(()=>{if(!sheet)return [];const cols=mode==='ficha'?['B','C','D']:Array.from({length:sheet.maxCol},(_,j)=>abc(j+1));return Array.from({length:sheet.maxRow},(_,j)=>j+1).map(r=>({r,refs:cols.map(c=>c+r).filter(ref=>showEmpty||String(sheet.cells[ref]??'').trim()||(sheet.validations||[]).some(v=>v.range.split(/\s+/).some(range=>between(ref,range))))})).filter(({r,refs})=>refs.length&&(!filter||refs.some(ref=>(value(idx,tab,ref)+' '+ref+' '+r).toLowerCase().includes(filter.toLowerCase()))))},[sheet,mode,showEmpty,filter,edits,idx,tab]);
 return <main className="app"><header className="cabecera"><div><p className="eyebrow">OPOSICIONES · LA RIOJA · 2027</p><h1>APP SAP · Editor de 12 situaciones de aprendizaje</h1><p>Desplegables originales · Campos independientes · Exportación Word y PDF</p></div><div className="tag">Datos cargados por ti</div></header>
 <section className="panel" style={{marginTop:18}}><h2>1. Cargar las plantillas privadas</h2><p className="help">Esta versión pública no incorpora los materiales del preparador. Selecciona el archivo <strong>plantillas.json</strong> de la versión privada, o una copia completa creada con esta APP. La lectura se realiza en tu navegador.</p><input type="file" accept=".json,application/json" onChange={e=>load(e.target.files?.[0])}/>{msg&&<p className="notice" role="status">{msg}</p>}
 {!u?<p>Esperando el JSON original de las 12 SAP. El archivo se carga una vez por sesión; tus cambios se conservan en este navegador, pero conviene descargar una copia completa.</p>:<><div className="toolbar"><div><label htmlFor="sap">Situación de aprendizaje</label><select id="sap" value={idx} onChange={e=>{setIdx(Number(e.target.value));setTab(0);setFilter('')}}>{data.map((unit,i)=><option key={i} value={i}>SAP {unit.id} · {unit.title}</option>)}</select></div><div><label htmlFor="hoja">Hoja de la plantilla</label><select id="hoja" value={tab} onChange={e=>{setTab(Number(e.target.value));setFilter('')}}>{u.sheets.map((s,i)=><option key={i} value={i}>{s.name}</option>)}</select></div></div>
 <div className="panel-head"><div><h2>2. {u.title}</h2><p>Hoja «{sheet.name}». Los desplegables proceden de las validaciones y tablas auxiliares del Excel.</p></div><div className="switch"><button className={mode==='ficha'?'active':''} onClick={()=>setMode('ficha')}>Ficha principal</button><button className={mode==='todo'?'active':''} onClick={()=>setMode('todo')}>Todas las columnas</button></div></div>
 <div className="controls"><input placeholder="Buscar texto o celda..." aria-label="Buscar" value={filter} onChange={e=>setFilter(e.target.value)}/><label className="check"><input type="checkbox" checked={showEmpty} onChange={e=>setShowEmpty(e.target.checked)}/> Mostrar celdas vacías</label><button onClick={()=>{if(window.confirm('¿Restaurar los cambios de esta hoja?'))setEdits(o=>Object.fromEntries(Object.entries(o).filter(([k])=>!k.startsWith(`${idx}/${tab}/`))))}}>Restaurar hoja</button></div>
 <p className="help">Cada celda mantiene su propia selección: por ejemplo, B24, B25 y B26 permiten tres metodologías independientes. Puedes escribir una opción nueva seleccionando «Otro valor…».</p>
 <div className="sheet"><table><thead><tr><th>Fila</th><th>Campos editables y desplegables</th></tr></thead><tbody>{rows.map(({r,refs})=><tr key={r}><th className="num">{r}</th><td><div className="cellrow">{refs.map(ref=>{const opts=options(sheet,ref),v=value(idx,tab,ref),other=opts.length>0&&!opts.includes(v)&&v!=='';return <div className={'field '+(v.endsWith(':')?'heading':'')} key={ref}><label htmlFor={`f-${ref}`}>{ref}{Object.prototype.hasOwnProperty.call(edits,key(idx,tab,ref))?' · modificado':''}</label>{opts.length?<><select id={`f-${ref}`} value={other?'__custom__':v} onChange={e=>setCell(ref,e.target.value==='__custom__'?(other?v:''):e.target.value)}><option value="">— Seleccionar —</option>{opts.map((o,i)=><option value={o} key={i}>{o}</option>)}<option value="__custom__">Otro valor…</option></select>{(other||v==='')&&<input aria-label={`Valor libre ${ref}`} placeholder="Escribir opción personalizada" value={other?v:''} onChange={e=>setCell(ref,e.target.value)}/>}</>:<textarea id={`f-${ref}`} rows={v.length>140?4:v.length>70?3:1} value={v} onChange={e=>setCell(ref,e.target.value)}/>}</div>})}</div></td></tr>)}</tbody></table></div>
 <div className="export"><button className="primary" onClick={exportBackup}>Guardar copia completa JSON</button><button onClick={exportData}>Exportar 12 SAP editadas (JSON)</button><button onClick={word}>Exportar hoja a Word</button><button onClick={pdf}>Exportar hoja a PDF</button></div><p className="help">Word y PDF incluyen los campos no vacíos de la hoja seleccionada; todavía no reproducen exactamente la maquetación del Excel. La copia JSON completa conserva las 12 SAP, las hojas auxiliares y los desplegables.</p></>}</section><footer>APP SAP EF 2027 · Versión 2 · Los materiales privados no están incluidos en el repositorio público.</footer></main>
}
