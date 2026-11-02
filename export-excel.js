/* Browser-side XLSX export. Native cells, formulas and pictures, no CSV masquerading as Excel. */
(() => {
  'use strict';
  const R = window.KimReport, X = R.xml;
  const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const relns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const relationship = (id, type, target) => `<Relationship Id="${id}" Type="${relns}/${type}" Target="${X(target)}"/>`;
  const rels = content => `${declaration}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
  const cell = (ref, value, style = 0, formula) => {
    if (formula) return `<c r="${ref}" s="${style}"${typeof value === 'string' ? ' t="str"' : ''}><f>${X(formula)}</f><v>${X(value)}</v></c>`;
    if (typeof value === 'number') return `<c r="${ref}" s="${style}"><v>${value}</v></c>`;
    return `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${X(value)}</t></is></c>`;
  };
  const row = (n, cells, height = 20) => `<row r="${n}" ht="${height}" customHeight="1">${cells.join('')}</row>`;
  const mergeXml = merges => merges.length ? `<mergeCells count="${merges.length}">${merges.map(ref => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>` : '';
  function sheet(rows, merges, drawing, primary = false) {
    return `${declaration}<worksheet xmlns="${ns}" xmlns:r="${relns}"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="${primary ? 12 : 1}" topLeftCell="A${primary ? 13 : 2}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>${[12,12,18,18,8,18,18,18].map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>${rows.join('')}</sheetData>${mergeXml(merges)}${primary ? `<conditionalFormatting sqref="C6:D6">${['C6&gt;=50','AND(C6&gt;=25,C6&lt;50)','AND(C6&gt;=10,C6&lt;25)','C6&lt;10'].map((f,i)=>`<cfRule type="expression" dxfId="${3-i}" priority="${i+1}" stopIfTrue="1"><formula>${f}</formula></cfRule>`).join('')}</conditionalFormatting>` : ''}<pageMargins left="0.25" right="0.25" top="0.3" bottom="0.3" header="0.1" footer="0.1"/><pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="${primary ? 1 : 0}"/><headerFooter><oddFooter>&amp;L${X('KIM 1')}&amp;R&amp;P</oddFooter></headerFooter>${drawing ? '<drawing r:id="rIdDrawing"/>' : ''}</worksheet>`;
  }
  function styles() {
    const font = (size, bold = false, color = '000000', name = R.FONT) => `<font><sz val="${size}"/><name val="${name}"/><family val="2"/>${bold ? '<b/>' : ''}<color rgb="FF${color}"/></font>`;
    const fill = color => `<fill><patternFill patternType="solid"><fgColor rgb="FF${color}"/><bgColor rgb="FF${color}"/></patternFill></fill>`;
    const xf = (fontId, fillId = 0, borderId = 0, align = 'left') => `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" applyAlignment="1"><alignment horizontal="${align}" vertical="top" wrapText="1"/></xf>`;
    return `${declaration}<styleSheet xmlns="${ns}"><fonts count="5">${font(11)}${font(11,true)}${font(26,false,'000000',R.HEADING)}${font(11,true,'FFFFFF')}${font(16,true)}</fonts><fills count="9"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fill('2C6267')}${fill('F3F3F2')}${fill('E7F6F5')}${['E9F0D8','FFF2BD','F8E3CD','F6DEDB'].map(fill).join('')}</fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border>${['left','right','top','bottom'].map(side=>`<${side} style="thin"><color rgb="FFC6C6C6"/></${side}>`).join('')}<diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="12">${xf(0)}${xf(1)}${xf(2)}${xf(3,2,1)}${xf(0,0,1)}${xf(4,4,1,'center')}${xf(4,3,1)}${xf(0,5)}${xf(0,6)}${xf(0,7)}${xf(0,8)}${xf(1,3)}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="4">${['E9F0D8','FFF2BD','F8E3CD','F6DEDB'].map(c=>`<dxf>${fill(c)}</dxf>`).join('')}</dxfs></styleSheet>`;
  }
  async function build(report, evidence, assets = {}) {
    const m = R.model(report, evidence), zip = new JSZip(), images = [];
    const r = [], merges = [];
    const merged = (ref, value, style, formula) => { merges.push(ref); return cell(ref.split(':')[0], value, style, formula); };
    r.push(row(1,[merged('A1:F1','ERGONOMIANALYS KIM 1',2)],38));
    r.push(row(3,[merged('A3:B3','Fabrik',1),cell('C3',report.metadata[0],0),cell('D3','Process',1),merged('E3:H3',report.metadata[1],0)],25));
    r.push(row(4,[merged('A4:B4','Underprocess',1),merged('C4:H4',report.metadata[2],0)],25));
    const classFormula = 'IF(C6>=50,4,IF(C6>=25,3,IF(C6>=10,2,1)))';
    const labelFormula = 'IF(C6>=50,"Hög belastning",IF(C6>=25,"Kraftigt ökad belastning",IF(C6>=10,"Ökad belastning","Låg belastning")))';
    const actionFormula = 'IF(C6>=50,"Arbetsplatsens utformning måste ändras.",IF(C6>=25,"Ny utformning av arbetsplatsen rekommenderas.",IF(C6>=10,"Anpassning av arbetsplatsen är fördelaktig för personer med lägre fysisk kapacitet.","Fysisk överbelastning är osannolik.")))';
    r.push(row(6,[merged('A6:B6','RISKPOÄNG',1),merged('C6:D6',m.risk,2,'SUM(E14:E16)*E13'),merged('E6:H6',`Riskklass ${m.riskInfo.level}: ${m.riskInfo.label}`,6,`"Riskklass "&${classFormula}&": "&${labelFormula}`)],42));
    r.push(row(7,[merged('A7:H8',m.riskInfo.action,0,actionFormula)],24));
    r.push(row(9,[merged('A9:H9','Riskpoäng = (last + arbetsförhållanden + arbetsställning) × tidspoäng',0)],22));
    r.push(row(12,[merged('A12:B12','Faktor',3),merged('C12:D12','Vald nivå',3),cell('E12','Poäng',3),merged('F12:H12','Motivering',3)],23));
    m.factors.forEach((f,i)=>{
      const n=13+i;
      r.push(row(n,[merged(`A${n}:B${n}`,f.label,4),merged(`C${n}:D${n}`,f.choice,4),cell(`E${n}`,R.number(f.score),5),merged(`F${n}:H${n}`,f.commentPreview,4)],54));
    });
    r.push(row(18,[merged('A18:B18','Summa belastning',1),merged('C18:D18',m.scores.slice(0,3).reduce((sum,n)=>sum+R.number(n),0),6,'SUM(E14:E16)'),merged('E18:F18','Tidspoäng',1),merged('G18:H18',R.number(m.scores[4]),6,'E13')],25));
    r.push(row(20,[merged('A20:D20','PROCESS OCH AVGRÄNSNING',1),merged('E20:H20','ÖVRIGA OBSERVATIONER',1)],22));
    r.push(row(21,[merged('A21:D23',m.intro,0),merged('E21:H23',m.observation,0)],18));
    r.push(row(22,[],18),row(23,[],18));
    r.push(row(25,[merged('A25:B25','<10  Låg',7),merged('C25:D25','10–<25  Ökad',8),merged('E25:F25','25–<50  Kraftigt ökad',9),merged('G25:H25','≥50  Hög',10)],22));
    r.push(row(27,[merged('A27:H27','Poängen i kolumn E är redigerbara. Riskpoäng och riskklass räknas om i Excel.',0)],24));
    const sheets = [{ name:'Analys', rows:r, merges, primary:true, pictures:[] }];
    if(assets.logoData) sheets[0].pictures.push({data:assets.logoData,col:6,row:0,w:1.6,h:1.6*269/2023,name:'Dometic'});
    if(m.notes.length || m.photos.length || m.attachments.length){
      const details={name:'Underlag',rows:[],merges:[],pictures:[]};sheets.push(details);
      let cursor=1;
      const addMerged=(value,style=0,height=22)=>{const n=cursor++;details.merges.push(`A${n}:H${n}`);details.rows.push(row(n,[cell(`A${n}`,value,style)],height));};
      addMerged('UNDERLAG KIM 1',2,38);
      for(const note of m.notes){
        addMerged(note.title,11,25);
        const chars=Array.from(note.text);
        for(let i=0;i<chars.length;i+=500)addMerged(chars.slice(i,i+500).join(''),0,Math.max(24,Math.ceil(Math.min(500,chars.length-i)/115)*16));
        cursor++;
      }
      for(let at=0;at<m.photos.length;at+=2){
        const group=m.photos.slice(at,at+2), start=cursor;
        group.forEach((p,i)=>{
          const a=i===0?'A':'E',b=i===0?'D':'H';
          details.merges.push(`${a}${start}:${b}${start}`,`${a}${start+1}:${b}${start+1}`);
          let iw=4.6,ih=iw*p.height/p.width;if(ih>2.75){ih=2.75;iw=ih*p.width/p.height;}
          details.pictures.push({data:p.data,col:i*4,row:start+1,w:iw,h:ih,name:`Fråga ${p.number}: ${p.caption||p.name}`});
        });
        details.rows.push(row(start,group.map((p,i)=>cell(`${i?'E':'A'}${start}`,`${p.number}. ${p.label}`,1)),24));
        details.rows.push(row(start+1,group.map((p,i)=>cell(`${i?'E':'A'}${start+1}`,p.captionPreview,0)),32));
        for(let i=2;i<18;i++)details.rows.push(row(start+i,[],15));
        cursor+=20;
      }
      if(m.attachments.length){
        addMerged('DOKUMENTBILAGOR',11,26);
        for(const file of m.attachments){addMerged(file.name,1,24);if(file.description)addMerged(file.description,0,Math.max(24,Math.min(160,Math.ceil(file.description.length/115)*16)));}
      }
    }
    const drawingTypes=[];
    for(let index=0;index<sheets.length;index++){
      const s=sheets[index], hasPictures=s.pictures.length>0;
      zip.file(`xl/worksheets/sheet${index+1}.xml`,sheet(s.rows,s.merges,hasPictures,s.primary));
      if(!hasPictures)continue;
      const anchors=[],links=[];
      for(const picture of s.pictures){
        const match=/^data:image\/(png|jpeg);base64,(.+)$/.exec(picture.data);
        if(!match)throw new Error('Bildformatet kunde inte exporteras till Excel.');
        const id=images.length+1, extension=match[1], media=`image${id}.${extension}`;
        images.push({extension});zip.file(`xl/media/${media}`,match[2],{base64:true});
        links.push(relationship(`rIdImage${id}`,'image',`../media/${media}`));
        anchors.push(`<xdr:oneCellAnchor><xdr:from><xdr:col>${picture.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${picture.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${Math.round(picture.w*914400)}" cy="${Math.round(picture.h*914400)}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${X(picture.name)}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage${id}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`);
      }
      const drawing=`drawing${index+1}.xml`;
      zip.file(`xl/drawings/${drawing}`,`${declaration}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${relns}">${anchors.join('')}</xdr:wsDr>`);
      zip.file(`xl/drawings/_rels/${drawing}.rels`,rels(links.join('')));
      zip.file(`xl/worksheets/_rels/sheet${index+1}.xml.rels`,rels(relationship('rIdDrawing','drawing',`../drawings/${drawing}`)));
      drawingTypes.push(`<Override PartName="/xl/drawings/${drawing}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
    }
    zip.file('xl/styles.xml',styles());
    zip.file('xl/workbook.xml',`${declaration}<workbook xmlns="${ns}" xmlns:r="${relns}"><bookViews><workbookView/></bookViews><sheets>${sheets.map((s,i)=>`<sheet name="${s.name}" sheetId="${i+1}" r:id="rIdSheet${i+1}"/>`).join('')}</sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Analys'!$A$1:$H$27</definedName></definedNames><calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels',rels(sheets.map((_,i)=>relationship(`rIdSheet${i+1}`,'worksheet',`worksheets/sheet${i+1}.xml`)).join('')+relationship('rIdStyles','styles','styles.xml')));
    zip.file('_rels/.rels',rels(relationship('rIdWorkbook','officeDocument','xl/workbook.xml')));
    zip.file('[Content_Types].xml',`${declaration}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${[...new Set(images.map(i=>i.extension))].map(ext=>`<Default Extension="${ext}" ContentType="image/${ext}"/>`).join('')}<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}${drawingTypes.join('')}</Types>`);
    return zip.generateAsync({type:'arraybuffer',compression:'DEFLATE'});
  }
  async function exportReport(){
    if(window.KimExportBusy)return;
    window.KimExportBusy=true;
    const buttons=['pptxBtn','excelBtn','printBtn'].map(id=>document.getElementById(id));buttons.forEach(b=>b.disabled=true);
    const button=document.getElementById('excelBtn'),old=button.textContent;button.textContent='Skapar Excel…';
    try{
      await window.KimEvidence.flush();
      const report=window.KimAnalysis.report(),evidence=structuredClone(window.KimEvidence.getState());
      const buffer=await build(report,evidence,await R.assets());
      window.KimEvidence.download(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),R.safeName(['KIM1',...report.metadata.filter(Boolean)].join('_'))+'.xlsx');
      window.KimEvidence.message('Excel är klar. Poäng och kommentarer kan redigeras.');
    }catch(error){console.error('Excel export failed',error);window.KimEvidence.message('Excel-exporten misslyckades: '+error.message,true);}
    finally{window.KimExportBusy=false;buttons.forEach(b=>b.disabled=false);button.textContent=old;}
  }
  document.getElementById('excelBtn').onclick=exportReport;
  window.KimExcel={build};
})();
