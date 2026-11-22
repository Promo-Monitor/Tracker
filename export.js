/* Promo Tracker \u2013 export Excel (.xlsx cu filtre) \u0219i PDF (vizualizarea curent\u0103). F\u0103r\u0103 biblioteci externe. */
(function(){
const TE=new TextEncoder();
let DLNS;
const dlns=()=>DLNS||(DLNS=(Promise.resolve(null)).catch(()=>null));
dlns();
async function dl(blob,name){
  const ns=await dlns();
  if(ns){
    try{await ns.save({filename:name,data:blob});return "saved"}
    catch(e){const c=e&&e.code;if(c==="declined")return "declined";if(c==="rate_limited")throw new Error("A\u0219teapt\u0103 c\u00e2teva secunde \u0219i \u00eencearc\u0103 din nou");if(c&&!["unavailable","not_granted","capability_disabled","capability_removed"].includes(c))throw new Error(e.message||c)}
  }
  const u=URL.createObjectURL(blob),a=document.createElement("a");
  a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(u),60000);return "saved";
}
const stamp=()=>new Date(Date.now()+2*3600e3).toISOString().slice(0,10);

/* ---------- ZIP (stored) ---------- */
const CRC=(()=>{const t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0}return t})();
function crc32(b){let c=0xFFFFFFFF;for(let i=0;i<b.length;i++)c=CRC[(c^b[i])&255]^(c>>>8);return (c^0xFFFFFFFF)>>>0}
function zip(files){
  const parts=[],cent=[];let off=0;
  for(const [name,str] of files){
    const nb=TE.encode(name),db=TE.encode(str),crc=crc32(db);
    const h=new DataView(new ArrayBuffer(30));
    h.setUint32(0,0x04034b50,true);h.setUint16(4,20,true);h.setUint16(6,0x0800,true);h.setUint16(8,0,true);
    h.setUint16(10,0,true);h.setUint16(12,0x21,true);h.setUint32(14,crc,true);h.setUint32(18,db.length,true);h.setUint32(22,db.length,true);
    h.setUint16(26,nb.length,true);h.setUint16(28,0,true);
    parts.push(new Uint8Array(h.buffer),nb,db);
    const c=new DataView(new ArrayBuffer(46));
    c.setUint32(0,0x02014b50,true);c.setUint16(4,20,true);c.setUint16(6,20,true);c.setUint16(8,0x0800,true);c.setUint16(10,0,true);
    c.setUint16(12,0,true);c.setUint16(14,0x21,true);c.setUint32(16,crc,true);c.setUint32(20,db.length,true);c.setUint32(24,db.length,true);
    c.setUint16(28,nb.length,true);c.setUint32(42,off,true);
    cent.push(new Uint8Array(c.buffer),nb);
    off+=30+nb.length+db.length;
  }
  const cs=cent.reduce((s,x)=>s+x.length,0),e=new DataView(new ArrayBuffer(22));
  e.setUint32(0,0x06054b50,true);e.setUint16(8,files.length,true);e.setUint16(10,files.length,true);e.setUint32(12,cs,true);e.setUint32(16,off,true);
  return new Blob([...parts,...cent,new Uint8Array(e.buffer)],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}

/* ---------- XLSX ---------- */
const xe=s=>String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,"");
const colL=i=>{let s="";i++;while(i>0){const m=(i-1)%26;s=String.fromCharCode(65+m)+s;i=Math.floor((i-1)/26)}return s};
const serial=d=>{if(!d||!/^\d{4}-\d\d-\d\d/.test(d))return null;return Math.round((Date.UTC(+d.slice(0,4),+d.slice(5,7)-1,+d.slice(8,10))-Date.UTC(1899,11,30))/864e5)};
/* styles: 0 normal, 1 header, 2 date, 3 money, 4 link, 5 int */
const STYLES=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="dd.mm.yyyy"/><numFmt numFmtId="165" formatCode="0.00"/></numFmts>
<fonts count="3"><font><sz val="10"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font><font><u/><sz val="10"/><color rgb="FF1D4ED8"/><name val="Calibri"/><family val="2"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1B1F24"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
/* cols: [header, width, type] type: s text, d date, m money, i int, l link */
function sheetXml(cols,rows,links){
  let x=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>`;
  cols.forEach((c,i)=>x+=`<col min="${i+1}" max="${i+1}" width="${c[1]}" customWidth="1"/>`);
  x+=`</cols><sheetData><row r="1">`;
  cols.forEach((c,i)=>x+=`<c r="${colL(i)}1" t="inlineStr" s="1"><is><t>${xe(c[0])}</t></is></c>`);
  x+=`</row>`;
  rows.forEach((r,ri)=>{const rn=ri+2;x+=`<row r="${rn}">`;
    r.forEach((v,ci)=>{const ref=colL(ci)+rn,t=cols[ci][2];
      if(v==null||v==="")return;
      if(t==="d"){const s=serial(v);if(s!=null){x+=`<c r="${ref}" s="2"><v>${s}</v></c>`;return}}
      if((t==="m"||t==="i")&&typeof v==="number"&&isFinite(v)){x+=`<c r="${ref}" s="${t==="m"?3:5}"><v>${v}</v></c>`;return}
      if(t==="l"){links.push([ref,String(v)]);x+=`<c r="${ref}" t="inlineStr" s="4"><is><t>${xe(cols[ci][3]||v)}</t></is></c>`;return}
      x+=`<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xe(v)}</t></is></c>`});
    x+=`</row>`});
  x+=`</sheetData><autoFilter ref="A1:${colL(cols.length-1)}${rows.length+1}"/>`;
  if(links.length)x+=`<hyperlinks>`+links.map((l,i)=>`<hyperlink ref="${l[0]}" r:id="rId${i+1}"/>`).join("")+`</hyperlinks>`;
  return x+`</worksheet>`;
}
const relsXml=links=>`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+links.map((l,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xe(l[1])}" TargetMode="External"/>`).join("")+`</Relationships>`;
function xlsx(sheets){
  const files=[],ct=[],wbs=[],wbr=[];
  sheets.forEach((s,i)=>{const n=i+1,links=[];
    files.push([`xl/worksheets/sheet${n}.xml`,sheetXml(s.cols,s.rows,links)]);
    if(links.length)files.push([`xl/worksheets/_rels/sheet${n}.xml.rels`,relsXml(links)]);
    ct.push(`<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    wbs.push(`<sheet name="${xe(s.name)}" sheetId="${n}" r:id="rId${n}"/>`);
    wbr.push(`<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/>`);
  });
  const k=sheets.length+1;
  const dn=sheets.map((s,i)=>`<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${xe(s.name)}'!$A$1:$${colL(s.cols.length-1)}$${s.rows.length+1}</definedName>`).join("");
  return zip([
    ["[Content_Types].xml",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${ct.join("")}</Types>`],
    ["_rels/.rels",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ["xl/workbook.xml",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${wbs.join("")}</sheets><definedNames>${dn}</definedNames></workbook>`],
    ["xl/_rels/workbook.xml.rels",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbr.join("")}<Relationship Id="rId${k}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
    ["xl/styles.xml",STYLES],
    ...files]);
}
window.exportXLSX=async function(D,today){
  const st=p=>p.valid_to&&p.valid_to<today?"Expirat":p.valid_from>today?"Viitor":"Activ";
  const abs=u=>{try{return new URL(u,location.href).href}catch(e){return u}};
  const cols=[["Status",9,"s"],["Retailer",11,"s"],["Catalog",30,"s"],["Valabil de la",12,"d"],["Valabil p\u00e2n\u0103 la",12,"d"],["Pagina",8,"i"],["Categorie",10,"s"],["Companie",11,"s"],["Brand",22,"s"],["Produs",45,"s"],["Ambalaj",10,"s"],["Pack size (L)",11,"m"],["Buc.",6,"i"],["Pre\u021b promo (lei)",13,"m"],["Pre\u021b vechi (lei)",13,"m"],["Discount %",10,"i"],["Pre\u021b/L (lei)",11,"m"],["Mecanic\u0103",34,"s"],["Condi\u021bie",24,"s"],["Not\u0103",28,"s"],["Catalog online",14,"l","Vezi \u00een catalog"],["Foto",8,"l","Foto"],["Detectat la",11,"d"]];
  const rows=D.promos.slice().sort((a,b)=>(b.valid_from||"").localeCompare(a.valid_from||"")||a.retailer.localeCompare(b.retailer)||a.page-b.page)
    .map(p=>[st(p),p.retailer,p.catalog_title,p.valid_from,p.valid_to,p.page,p.category,p.company,p.brand,p.product,p.pack_type,p.pack_size_l,p.pieces,p.promo_price,p.regular_price,p.discount_pct,p.price_per_l,p.mechanic,p.condition,p.note,p.viewer_url,p.img?abs(p.img):"",p.detected_on]);
  const ccols=[["Retailer",11,"s"],["Catalog",38,"s"],["Valabil de la",12,"d"],["Valabil p\u00e2n\u0103 la",12,"d"],["Pagini",8,"i"],["Pagini cu branduri",20,"s"],["Procesat la",11,"d"],["Not\u0103",30,"s"]];
  const crows=(D.catalogs||[]).map(c=>[c.retailer,c.title,c.valid_from,c.valid_to,c.pages,(c.hit_pages||[]).join(", "),c.processed_on,c.note||""]);
  const r=await dl(xlsx([{name:"Promo\u021bii",cols,rows},{name:"Cataloage",cols:ccols,rows:crows}]),`Promo_Tracker_${stamp()}.xlsx`);
  return r==="declined"?-1:rows.length;
};

/* ---------- PDF (Helvetica, JPEG nativ) ---------- */
const HW=[278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,333,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,334,260,334,584];
const MAP={"\u0103":136,"\u0102":137,"\u0219":138,"\u0218":139,"\u021b":140,"\u021a":141,"\u015f":138,"\u015e":139,"\u0163":140,"\u0162":141,"\u2013":150,"\u2014":151,"\u00b7":183,"\u201e":132,"\u201d":148,"\u201c":147,"\u2019":146,"\u2026":133,"\u00d7":215,"\u00e2":226,"\u00c2":194,"\u00ee":238,"\u00ce":206,"\u00e9":233,"\u00fc":252,"\u00f6":246};
const DIFF="/Differences [136 /abreve /Abreve /scedilla /Scedilla /tcedilla /Tcedilla]";
function enc(s){let o="";for(const ch of String(s??"")){const c=ch.charCodeAt(0);let b=c<128?c:(MAP[ch]??(c<256?c:63));if(b===40||b===41||b===92)o+="\\";o+=String.fromCharCode(b)}return o}
function tw(s,size,bold){let w=0;for(const ch of String(s)){const c=ch.charCodeAt(0);w+=(c>=32&&c<=126?HW[c-32]:556)}return w*size/1000*(bold?1.06:1)}
function wrap(s,size,maxw,maxl,bold){const words=String(s??"").split(/\s+/).filter(Boolean),L=[];let cur="";
  for(const w of words){const t=cur?cur+" "+w:w;if(tw(t,size,bold)<=maxw)cur=t;else{if(cur)L.push(cur);cur=w}if(L.length>=maxl)break}
  if(cur&&L.length<maxl)L.push(cur);
  if(L.length===maxl&&words.join(" ").length>L.join(" ").length){let l=L[maxl-1];while(l&&tw(l+"\u2026",size,bold)>maxw)l=l.slice(0,-1);L[maxl-1]=l+"\u2026"}
  return L.map(l=>{while(tw(l,size,bold)>maxw&&l.length>1)l=l.slice(0,-2)+"\u2026";return l})}
function jpgSize(b){let i=2;while(i<b.length){if(b[i]!==0xFF){i++;continue}const m=b[i+1],len=(b[i+2]<<8)|b[i+3];
  if(m>=0xC0&&m<=0xCF&&m!==0xC4&&m!==0xC8&&m!==0xCC)return{h:(b[i+5]<<8)|b[i+6],w:(b[i+7]<<8)|b[i+8],c:b[i+9]};i+=2+len}return null}
const RGB={TCCC:"0.82 0.07 0.12",PepsiCo:"0.12 0.31 0.66",Competitor:"0.36 0.40 0.45"};
const n2=v=>v==null?"":v.toFixed(2).replace(".",",");
const dm=s=>s?s.slice(8,10)+"."+s.slice(5,7):"";
window.exportPDF=async function(list,subtitle,catro){
  const W=595.28,H=841.89,M=26,COLS=3,ROWS=3,G=9,HEAD=44,cw=(W-2*M-G*(COLS-1))/COLS,ch=(H-2*M-HEAD-G*(ROWS-1))/ROWS,IH=128;
  const objs=[];const add=v=>{objs.push(v);return objs.length};
  const cat=add(null),pages=add(null),f1=add(`<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding ${DIFF} >> >>`),
        f2=add(`<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding ${DIFF} >> >>`);
  const imgs={};
  await Promise.all(list.map(async p=>{if(!p.img)return;try{const b=new Uint8Array(await (await fetch(p.img)).arrayBuffer()),s=jpgSize(b);if(!s)return;imgs[p.id]={b,...s}}catch(e){}}));
  const imgObj={};
  for(const [id,im] of Object.entries(imgs)){
    const head=`<< /Type /XObject /Subtype /Image /Width ${im.w} /Height ${im.h} /ColorSpace /${im.c===1?"DeviceGray":im.c===4?"DeviceCMYK":"DeviceRGB"} /BitsPerComponent 8 /Filter /DCTDecode${im.c===4?" /Decode [1 0 1 0 1 0 1 0]":""} /Length ${im.b.length} >>\nstream\n`;
    imgObj[id]=add([head,im.b,"\nendstream"]);
  }
  const per=COLS*ROWS,np=Math.max(1,Math.ceil(list.length/per)),kids=[];
  const T=(x,y,s,f,txt,col)=>`BT ${col||"0.11 0.12 0.14"} rg /${f} ${s} Tf ${x.toFixed(2)} ${y.toFixed(2)} Td (${enc(txt)}) Tj ET\n`;
  for(let pg=0;pg<np;pg++){
    let c="";const annots=[];
    c+=T(M,H-M-14,15,"F2","Promo Tracker \u2013 cataloage retail");
    c+=T(M,H-M-30,8,"F1",subtitle,"0.36 0.40 0.45");
    c+=T(W-M-tw(`Pagina ${pg+1}/${np}`,8),H-M-30,8,"F1",`Pagina ${pg+1}/${np}`,"0.36 0.40 0.45");
    list.slice(pg*per,pg*per+per).forEach((p,i)=>{
      const col=i%COLS,row=Math.floor(i/COLS),x=M+col*(cw+G),yTop=H-M-HEAD-row*(ch+G),y0=yTop-ch,pad=7,tx=x+pad,iw=cw-2*pad;
      c+=`0.89 0.90 0.92 RG 0.6 w ${x.toFixed(2)} ${y0.toFixed(2)} ${cw.toFixed(2)} ${ch.toFixed(2)} re S\n`;
      c+=`0.98 0.98 0.98 rg ${(x+0.3).toFixed(2)} ${(yTop-IH-0.3).toFixed(2)} ${(cw-0.6).toFixed(2)} ${IH} re f\n`;
      const im=imgs[p.id];
      if(im){const s=Math.min((cw-8)/im.w,(IH-8)/im.h),w=im.w*s,h=im.h*s;c+=`q ${w.toFixed(2)} 0 0 ${h.toFixed(2)} ${(x+(cw-w)/2).toFixed(2)} ${(yTop-IH+(IH-h)/2).toFixed(2)} cm /Im${p.id} Do Q\n`}
      if(p.discount_pct!=null){const t=`-${p.discount_pct}%`,bw=tw(t,8,1)+10;c+=`0.11 0.12 0.14 rg ${(x+5).toFixed(2)} ${(yTop-19).toFixed(2)} ${bw.toFixed(2)} 14 re f\n`+T(x+10,yTop-15,8,"F2",t,"1 1 1")}
      const cl=(catro[p.category]||p.category||""),clw=tw(cl,6.5)+8;
      c+=`1 1 1 rg 0.89 0.90 0.92 RG ${(x+cw-5-clw).toFixed(2)} ${(yTop-17).toFixed(2)} ${clw.toFixed(2)} 11 re B\n`+T(x+cw-1-clw,yTop-14,6.5,"F1",cl,"0.36 0.40 0.45");
      let y=yTop-IH-11;
      c+=T(tx,y,7,"F2",p.retailer)+T(tx+tw(p.retailer,7,1)+2,y,7,"F1",`\u00b7 ${dm(p.valid_from)}\u2013${dm(p.valid_to)} \u00b7 p.${p.page}`,"0.36 0.40 0.45");y-=12;
      wrap(p.brand,10,iw,1,1).forEach(l=>{c+=T(tx,y,10,"F2",l,RGB[p.company]||RGB.Competitor);y-=10});
      wrap(p.product,7.5,iw,2).forEach(l=>{c+=T(tx,y,7.5,"F1",l,"0.2 0.2 0.2");y-=9});
      const pk=[p.pack_type,p.pack_size_l?String(p.pack_size_l).replace(".",",")+" L":"",p.pieces>1?p.pieces+" buc.":""].filter(Boolean).join(" \u00b7 ")+(p.price_per_l?` \u00b7 ${n2(p.price_per_l)} lei/L`:"");
      c+=T(tx,y,7,"F1",wrap(pk,7,iw,1)[0]||"","0.36 0.40 0.45");y-=15;
      const pr=p.promo_price!=null?n2(p.promo_price)+" lei":"\u2014";c+=T(tx,y,13,"F2",pr);
      if(p.regular_price!=null){const ox=tx+tw(pr,13,1)+5,ot=n2(p.regular_price);c+=T(ox,y,8,"F1",ot,"0.36 0.40 0.45")+`0.36 0.40 0.45 RG 0.5 w ${ox.toFixed(2)} ${(y+2.8).toFixed(2)} m ${(ox+tw(ot,8)).toFixed(2)} ${(y+2.8).toFixed(2)} l S\n`}
      y-=11;
      wrap((p.mechanic||"")+(p.condition?" \u00b7 "+p.condition:""),6.8,iw,2).forEach(l=>{c+=T(tx,y,6.8,"F1",l,"0.36 0.40 0.45");y-=8.5});
      if(p.viewer_url){c+=T(tx,y0+6,7,"F1","Vezi \u00een catalog >","0.11 0.31 0.85");annots.push(add(`<< /Type /Annot /Subtype /Link /Rect [${x} ${y0} ${x+cw} ${yTop}] /Border [0 0 0] /A << /S /URI /URI (${enc(p.viewer_url)}) >> >>`))}
    });
    if(!list.length)c+=T(M,H/2,11,"F1","Nicio apari\u021bie pentru filtrele alese.");
    const cs=add(`<< /Length ${c.length} >>\nstream\n${c}endstream`);
    const xo=list.slice(pg*per,pg*per+per).filter(p=>imgObj[p.id]).map(p=>`/Im${p.id} ${imgObj[p.id]} 0 R`).join(" ");
    kids.push(add(`<< /Type /Page /Parent ${pages} 0 R /MediaBox [0 0 ${W} ${H}] /Resources << /Font << /F1 ${f1} 0 R /F2 ${f2} 0 R >> /XObject << ${xo} >> >> /Contents ${cs} 0 R${annots.length?` /Annots [${annots.map(a=>a+" 0 R").join(" ")}]`:""} >>`));
  }
  objs[cat-1]=`<< /Type /Catalog /Pages ${pages} 0 R >>`;
  objs[pages-1]=`<< /Type /Pages /Kids [${kids.map(k=>k+" 0 R").join(" ")}] /Count ${kids.length} >>`;
  const L1=s=>{const u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i)&255;return u};
  const out=[L1("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n")];let off=out[0].length;const xref=[];
  objs.forEach((o,i)=>{xref.push(off);const parts=[`${i+1} 0 obj\n`,...(Array.isArray(o)?o:[o]),"\nendobj\n"].map(x=>typeof x==="string"?L1(x):x);parts.forEach(u=>{out.push(u);off+=u.length})});
  let t=`xref\n0 ${objs.length+1}\n0000000000 65535 f \n`+xref.map(o=>String(o).padStart(10,"0")+" 00000 n \n").join("")+`trailer\n<< /Size ${objs.length+1} /Root ${cat} 0 R >>\nstartxref\n${off}\n%%EOF\n`;
  out.push(L1(t));
  const r=await dl(new Blob(out,{type:"application/pdf"}),`Promo_Tracker_${stamp()}.pdf`);
  return r==="declined"?-1:list.length;
};
})();
