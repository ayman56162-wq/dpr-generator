const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign
} = require('docx');
const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');
const os   = require('os');
const { IncomingForm } = require('formidable');
const { buildPdf } = require('./pdf-generator');

// Logo embedded as base64 (no filesystem dependency)
const LOGO_B64 = "/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCAB+ApgDASIAAhEBAxEB/8QAHQABAAICAwEBAAAAAAAAAAAAAAYHBQgDBAkCAf/EAFsQAAEDAwIDBAQFCxEECAcAAAEAAgMEBREGIQcSMRMiQVEIFHGRFTJSYYEWFyNWc5ShsbLR0gkYMzQ1NjdCU1VicnSSlcHTJDiToiVDVIKks+HwY3aEo7TCw//EABsBAQACAwEBAAAAAAAAAAAAAAABBAIDBgUH/8QANREAAgECAwUGBAUFAQAAAAAAAAECAxEEITEFEhNBUWFxgZGh8BSxwdEGIkJS4RUjMjPxcv/aAAwDAQACEQMRAD8A3LREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAX4XNBwXAe0qOav1ZR2KIxsxPVn4sY6D2lUPr9l91LTSSwaguNFXjLo3w1UkTCfItacY+jZUKu08NRrKlOWur6d5Zhg6tSm5xX8mzXOz5bfenOz5bfevOy83zXdorn0dfqO/RSsPjcJcH2HmXS+rDV3203z/ABCX9JdLDZfEipRmmmeTLGbrtKOZ6Qc7Plt96c7Plt9685qLU2s6uXs49UXz5ybhLgD+8u7edSapo4IWM1TfCTnmd8IS7/8AMtq2LNre3irPbFKNWNK2bPQznZ8tvvTnZ8tvvXm/9WGrvtpvn+IS/pJ9WGrvtpvn+IS/pLD+kS/cWPjl0PSHnZ8tvvTmb8oe9aG8PdT6nqGziXUN4mdzYAdWyOPh862L4U6K1HXGG8aju9zZS454qY1knM/yLsHYePVY1tmKjT35z9Dy47eqVca8JSottau+S78i6EQDAwEXknRhERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAWD1tXXK32N89sg7SXmw53ixuD3se73rOIRkYK11YOcHFO1+ZlCSjJNq5rpUSyzzPmnkfJI85c55JJPzkr4Voa70S2pD7jaGYnyTLCG/H+cfP+PKxOiNETVkgrbvG+GnaRyQubh0ntz0H41xdTZWJVfhWvfnyOhjjaTp79/Ar7UPD+LVekrpcquBscVBSSzMqC3v8zGFwaD7cZ38VQFBpZlU/ImcGDqcLcDjFrq1WWzVelqGJlRU1FM+CRrHgNga8Fu+x3xk428Frp9jp4PBjGhfV/wAMbPlhMJuTu7vK/wBOw+QfjD8Qt4pQw0vzJWdvebMXTWGmpYeSN/KB1PL1WFvVvp6p7WMldyszvjqspcK91RmNmWx/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";
const LOGO_BUFFER = Buffer.from(LOGO_B64, 'base64');

const GOLD="C9A84C", DARK_BLUE="1F3864", LIGHT_BLUE="D6E4F0";
const WHITE="FFFFFF", LGRAY="F2F2F2", TEXT_DARK="1A1A1A";
const PAGE_W=11906, PAGE_H=16838, MARGIN=576, CW=PAGE_W-(MARGIN*2);
const BD={style:BorderStyle.SINGLE,size:1,color:"CCCCCC"};
const BDS={top:BD,bottom:BD,left:BD,right:BD};
const NB={style:BorderStyle.NONE,size:0,color:"FFFFFF"};
const NBDS={top:NB,bottom:NB,left:NB,right:NB};

function mc(text,o={}){
  const{bold=false,color=TEXT_DARK,bg=WHITE,align=AlignmentType.LEFT,
    size=15,colspan=1,rowspan=1,valign=VerticalAlign.CENTER,w=null,bds=BDS,italic=false}=o;
  return new TableCell({columnSpan:colspan,rowSpan:rowspan,verticalAlign:valign,
    width:w?{size:w,type:WidthType.DXA}:undefined,
    shading:{fill:bg,type:ShadingType.CLEAR},
    margins:{top:50,bottom:50,left:80,right:80},borders:bds,
    children:[new Paragraph({alignment:align,spacing:{before:0,after:0},
      children:[new TextRun({text:String(text??''),bold,color,size,font:"Arial",italics:italic})]})]});
}
function hc(text,o={}){return mc(text,{bold:true,color:WHITE,bg:DARK_BLUE,align:AlignmentType.CENTER,size:15,...o});}
function banner(label){
  return new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[CW],
    rows:[new TableRow({children:[new TableCell({
      shading:{fill:GOLD,type:ShadingType.CLEAR},borders:BDS,
      margins:{top:80,bottom:80,left:140,right:140},
      children:[new Paragraph({alignment:AlignmentType.CENTER,
        children:[new TextRun({text:label,bold:true,color:WHITE,size:18,font:"Arial"})]})]})]})]});
}
const sp=(pts=70)=>new Paragraph({spacing:{before:pts,after:0},children:[]});
function groupActs(acts){const out=[];let cur=null;for(const r of acts){if(!cur||r.system!==cur.system){cur={system:r.system,rows:[]};out.push(cur);}cur.rows.push(r);}return out;}
const AC=[1500,1400,700,1800,3754,700,900];
function buildActTable(rows){
  const groups=groupActs(rows);
  const tRows=[new TableRow({tableHeader:true,children:[
    hc("SYSTEM",{w:AC[0]}),hc("ZONE",{w:AC[1]}),hc("LEVEL",{w:AC[2]}),
    hc("DISCIPLINE",{w:AC[3]}),hc("TYPE OF WORK",{w:AC[4]}),
    hc("QTY",{w:AC[5],align:AlignmentType.CENTER}),hc("UNIT",{w:AC[6],align:AlignmentType.CENTER}),
  ]})];
  for(const grp of groups){
    const cnt=grp.rows.length;
    grp.rows.forEach((row,i)=>{
      const cells=[];
      if(i===0){cells.push(new TableCell({rowSpan:cnt,verticalAlign:VerticalAlign.CENTER,
        shading:{fill:LIGHT_BLUE,type:ShadingType.CLEAR},borders:BDS,
        width:{size:AC[0],type:WidthType.DXA},margins:{top:50,bottom:50,left:80,right:80},
        children:[new Paragraph({alignment:AlignmentType.CENTER,
          children:[new TextRun({text:grp.system,bold:true,size:15,font:"Arial",color:DARK_BLUE})]})]}));}
      cells.push(mc(row.zone,{size:15,w:AC[1]}));
      cells.push(mc(row.level,{size:15,align:AlignmentType.CENTER,w:AC[2]}));
      cells.push(mc(row.discipline,{size:15,w:AC[3]}));
      cells.push(mc(row.work,{size:15,w:AC[4]}));
      cells.push(mc(row.qty||'',{size:15,align:AlignmentType.CENTER,w:AC[5]}));
      cells.push(mc(row.unit||'',{size:15,align:AlignmentType.CENTER,w:AC[6]}));
      tRows.push(new TableRow({children:cells}));
    });
  }
  return new Table({width:{size:CW,type:WidthType.DXA},columnWidths:AC,rows:tRows});
}

function parseExcel(filePath){
  const wb=XLSX.readFile(filePath);
  const ws=wb.Sheets[wb.SheetNames[0]];
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
  const activities=[],nextDay=[];
  let section=null,headerRow=null;
  for(let i=0;i<raw.length;i++){
    const row=raw[i];
    const c0=String(row[0]||'').trim();
    if(c0==="Today's Activities"){section='today';headerRow=null;continue;}
    if(c0==="Next Day Activities"){section='next';headerRow=null;continue;}
    if(!section)continue;
    const rowStr=row.map(v=>String(v).toLowerCase()).join('|');
    if(!headerRow&&(rowStr.includes('zone')||rowStr.includes('level'))){
      headerRow=row.map(v=>String(v).trim().toLowerCase());continue;}
    if(!headerRow)continue;
    const num=String(row[0]).trim();
    if(!num||isNaN(parseFloat(num)))continue;
    const get=(key)=>{const idx=headerRow.findIndex(h=>h.includes(key));return idx>=0?String(row[idx]||'').trim():'';};
    const zone=get('zone'),level=get('level');
    const discipline=get('discipline')||get('disc');
    const system=get('system');
    let work=get('type')||get('work');
    const qty=get('qty')||get('quantity');
    const unit=get('unit');
    const remark=get('remark')||get('remarks');
    if(!work&&remark)work=remark;
    if(!work)continue;
    const entry={zone,level,discipline,system:system||'General',work,qty,unit};
    if(section==='today')activities.push(entry);
    if(section==='next')nextDay.push(entry);
  }
  return{activities,nextDay};
}

async function buildDocx({date,reportNo,supervision,labors,pkg,activities,nextDay,photos}){
  const ch=[];
  const d=new Date(date);
  const dateStr=`${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;

  ch.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:130},
    children:[new ImageRun({data:LOGO_BUFFER,type:"png",transformation:{width:370,height:70}})]}));

  const gi=[2100,3900,1754,3000];
  ch.push(new Table({width:{size:CW,type:WidthType.DXA},columnWidths:gi,rows:[
    new TableRow({children:[mc("Contractor / SubCont.",{bold:true,bg:LGRAY,size:14,w:gi[0]}),mc("MBL / First Fix",{size:14,w:gi[1]}),mc("Package",{bold:true,bg:LGRAY,size:14,w:gi[2]}),mc(pkg,{size:14,w:gi[3]})]}),
    new TableRow({children:[mc("Client",{bold:true,bg:LGRAY,size:14,w:gi[0]}),mc("Jeddah Central Development Co.",{size:14,w:gi[1]}),mc("Subcontract Ref.",{bold:true,bg:LGRAY,size:14,w:gi[2]}),mc("P1020096-OCN-014",{size:14,w:gi[3]})]})
  ]}));
  ch.push(sp(90));

  const leftW=4800,titleW=CW-4800;
  ch.push(new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[leftW,titleW],
    rows:[new TableRow({children:[
      new TableCell({verticalAlign:VerticalAlign.CENTER,shading:{fill:WHITE,type:ShadingType.CLEAR},borders:BDS,
        width:{size:leftW,type:WidthType.DXA},margins:{top:0,bottom:0,left:0,right:0},
        children:[new Table({width:{size:leftW,type:WidthType.DXA},columnWidths:[1700,3100],rows:[
          new TableRow({children:[mc("Date",{bold:true,bg:LGRAY,size:16,w:1700}),mc(dateStr,{size:16,w:3100})]}),
          new TableRow({children:[mc("Report No",{bold:true,bg:LGRAY,size:16,w:1700}),mc(reportNo,{size:16,w:3100})]}),
          new TableRow({children:[mc("Manpower",{bold:true,bg:LGRAY,align:AlignmentType.CENTER,size:16,colspan:2,w:4800})]}),
          new TableRow({children:[mc("Supervision",{bold:true,bg:LGRAY,align:AlignmentType.CENTER,size:15,w:1700}),mc(String(supervision),{align:AlignmentType.CENTER,size:16,bold:true,w:3100})]}),
          new TableRow({children:[mc("Labors",{bold:true,bg:LGRAY,align:AlignmentType.CENTER,size:15,w:1700}),mc(String(labors),{align:AlignmentType.CENTER,size:16,bold:true,w:3100})]}),
        ]})]
      }),
      new TableCell({verticalAlign:VerticalAlign.CENTER,shading:{fill:WHITE,type:ShadingType.CLEAR},borders:BDS,
        width:{size:titleW,type:WidthType.DXA},margins:{top:70,bottom:70,left:100,right:100},
        children:[new Paragraph({alignment:AlignmentType.CENTER,
          children:[new TextRun({text:"Daily Site Work Report",bold:true,size:28,color:DARK_BLUE,font:"Arial"})]})]}),
    ]})]
  }));

  ch.push(sp(100));
  ch.push(banner("DAILY SITE ACTIVITIES"));
  ch.push(sp(50));
  ch.push(new Paragraph({spacing:{before:30,after:50},children:[new TextRun({text:"MECHANICAL PROGRESS",bold:true,size:16,color:DARK_BLUE,font:"Arial"})]}));
  ch.push(buildActTable(activities));
  ch.push(sp(100));
  ch.push(banner("next day activities"));
  ch.push(sp(50));
  if(nextDay.length>0){ch.push(buildActTable(nextDay));}
  else{ch.push(new Paragraph({spacing:{before:40,after:40},children:[new TextRun({text:"To be updated",size:15,color:"999999",italics:true,font:"Arial"})]}));}
  ch.push(sp(100));
  ch.push(banner("PHOTOS"));
  ch.push(sp(70));
  const photoW=Math.floor(CW/2);
  for(let i=0;i<photos.length;i+=2){
    const cells=[];
    for(let j=i;j<Math.min(i+2,photos.length);j++){
      const{buffer,ext}=photos[j];
      cells.push(new TableCell({borders:NBDS,width:{size:photoW,type:WidthType.DXA},
        margins:{top:50,bottom:50,left:50,right:50},
        children:[new Paragraph({alignment:AlignmentType.CENTER,
          children:[new ImageRun({data:buffer,type:ext==='jpg'||ext==='jpeg'?'jpeg':'png',transformation:{width:250,height:188}})]})]
      }));
    }
    if(cells.length===1)cells.push(new TableCell({borders:NBDS,width:{size:photoW,type:WidthType.DXA},children:[new Paragraph({children:[]})]}));
    ch.push(new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[photoW,photoW],rows:[new TableRow({children:cells})]}));
    ch.push(sp(40));
  }

  const doc=new Document({
    styles:{default:{document:{run:{font:"Arial",size:15,color:TEXT_DARK}}}},
    sections:[{properties:{page:{size:{width:PAGE_W,height:PAGE_H},margin:{top:MARGIN,right:MARGIN,bottom:MARGIN,left:MARGIN}}},children:ch}]
  });
  return Packer.toBuffer(doc);
}

module.exports=async(req,res)=>{
  if(req.method!=='POST')return res.status(405).json({error:'Method not allowed'});
  const tmpDir=fs.mkdtempSync(path.join(os.tmpdir(),'dpr-'));
  try{
    const form=new IncomingForm({uploadDir:tmpDir,keepExtensions:true,multiples:true});
    const{fields,files}=await new Promise((ok,err)=>form.parse(req,(e,f,fi)=>e?err(e):ok({fields:f,files:fi})));
    const g=(f,k)=>Array.isArray(f[k])?f[k][0]:f[k];
    const date=g(fields,'date'),reportNo=g(fields,'reportNo');
    const supervision=g(fields,'supervision'),labors=g(fields,'labors');
    const pkg=g(fields,'package')||'MEP';
    const excelArr=Array.isArray(files.excel)?files.excel:(files.excel?[files.excel]:[]);
    const photoArr=Array.isArray(files.photos)?files.photos:(files.photos?[files.photos]:[]);
    let activities=[],nextDay=[];
    for(const f of excelArr){const fp=f.filepath||f.path;const p=parseExcel(fp);activities=activities.concat(p.activities);nextDay=nextDay.concat(p.nextDay);}

    const photos=photoArr.map(f=>{
      const fp=f.filepath||f.path;
      const ext=path.extname(f.originalFilename||fp).toLowerCase().replace('.','');
      return{buffer:fs.readFileSync(fp),ext:ext==='jpg'?'jpeg':ext};
    });

    const d=new Date(date);
    const dateStr=`${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;

    const [docxBuffer, pdfBuffer] = await Promise.all([
      buildDocx({date,reportNo,supervision,labors,pkg,activities,nextDay,photos}),
      buildPdf({date:dateStr,reportNo,supervision,labors,pkg,activities,nextDay,photos,logoBuffer:LOGO_BUFFER})
    ]);

    res.setHeader('Content-Type','application/json');
    res.status(200).json({
      docx: Buffer.from(docxBuffer).toString('base64'),
      pdf: Buffer.from(pdfBuffer).toString('base64'),
      filename: `DPR_${reportNo}`
    });
  }catch(e){console.error(e);res.status(500).json({error:e.message});}
  finally{try{fs.rmSync(tmpDir,{recursive:true});}catch{}}
};
