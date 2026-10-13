/* v1.1.20 — 單一工作表的看望報表；本機產生 OOXML，不上傳個人資料。 */
(function (root) {
  'use strict';
  const encoder = new TextEncoder();
  const xml = value => String(value ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,'');
  const utf = s => encoder.encode(s);
  const order = ['近期未聚會','零星聚會','偶聚會','常聚會','週週聚會'];
  const groups = ['年長','中壯','青壯','青職','大學','大專','青少年','國小','學齡前'];
  const groupName = s => ['國中','高中','中學'].includes(String(s||'').trim()) ? '青少年' : (String(s||'').trim()||'未分類');
  const compareText = (a,b) => String(a).localeCompare(String(b),'zh-Hant-TW');
  const dateTime = p => p.lastAttendanceDate instanceof Date && Number.isFinite(p.lastAttendanceDate.getTime()) ? p.lastAttendanceDate.getTime() : null;
  const printableDate = d => d instanceof Date && Number.isFinite(d.getTime()) ? `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}` : '無紀錄';
  const pad = n => String(n).padStart(2,'0');
  const today = () => {const d = new Date();return `${d.getFullYear()}/${pad(d.getMonth()+1)}/${pad(d.getDate())}`;};
  const cell = (ref, value, style=0) => `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${xml(value)}</t></is></c>`;
  const row = (n, vals, style, height, customStyles) => `<row r="${n}" ht="${height}" customHeight="1">${vals.map((v,i)=>cell(`${'ABCDEFG'[i]}${n}`,v,customStyles?.[i]??style)).join('')}</row>`;
  const crcTable = (() => {const t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0;}return t;})();
  function crc32(bytes) {let c=-1;for(const b of bytes)c=crcTable[(c^b)&255]^(c>>>8);return (c^-1)>>>0;}
  function zip(entries) {
    const files=[], directory=[];let offset=0;
    const number=(dv,pos,val,width)=>width===2?dv.setUint16(pos,val,true):dv.setUint32(pos,val>>>0,true);
    for(const [name,source] of entries){
      const nameBytes=utf(name), bytes=utf(source), crc=crc32(bytes);
      const local=new Uint8Array(30+nameBytes.length),ld=new DataView(local.buffer);
      number(ld,0,0x04034b50,4);number(ld,4,20,2);number(ld,6,2048,2);number(ld,8,0,2);
      number(ld,14,crc,4);number(ld,18,bytes.length,4);number(ld,22,bytes.length,4);number(ld,26,nameBytes.length,2);
      local.set(nameBytes,30);
      const central=new Uint8Array(46+nameBytes.length),cd=new DataView(central.buffer);
      number(cd,0,0x02014b50,4);number(cd,4,20,2);number(cd,6,20,2);number(cd,8,2048,2);number(cd,10,0,2);
      number(cd,16,crc,4);number(cd,20,bytes.length,4);number(cd,24,bytes.length,4);number(cd,28,nameBytes.length,2);number(cd,42,offset,4);
      central.set(nameBytes,46);files.push(local,bytes);directory.push(central);offset+=local.length+bytes.length;
    }
    const dirSize=directory.reduce((n,b)=>n+b.length,0);
    const end=new Uint8Array(22),ed=new DataView(end.buffer);
    number(ed,0,0x06054b50,4);number(ed,8,directory.length,2);number(ed,10,directory.length,2);number(ed,12,dirSize,4);number(ed,16,offset,4);
    return new Blob([...files,...directory,end],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }
  function create({people,dateColumns,rows,indexes,attendance,testDate,filterLabel}) {
    const weeks12=dateColumns.slice(-12), weeks52=dateColumns.slice(-52);
    const normalized=people.map(p=>{
      const raw=rows[p.rowNumber-1] || [];
      const count=cols=>cols.reduce((sum,col)=>sum+(attendance(raw[col.col])?1:0),0);
      return {...p,reportGroup:groupName(p.group),count12:count(weeks12),count52:count(weeks52)};
    });
    const lastDate=dateColumns.length ? printableDate(dateColumns[dateColumns.length-1].date) : '無資料';
    const blocks=[];
    for(const status of order){
      const matches=normalized.filter(p=>p.status===status);
      if(!matches.length)continue;
      blocks.push({kind:'status',value:`${status}（${matches.length} 人）`});
      // A rank name is only meaningful within its small district.
      const areas=new Map();
      for(const person of matches){
        const small=String(person.smallDistrict||'').trim()||'未分小區';
        const rank=String(person.rank||'').trim()||'未分排';
        const key=JSON.stringify([person.district||'',small,rank]);
        if(!areas.has(key))areas.set(key,{small,rank,district:person.district||'',people:[]});
        areas.get(key).people.push(person);
      }
      const compareNatural=(a,b)=>compareText(a,b);
      const sortedAreas=[...areas.values()].sort((a,b)=>
        compareNatural(a.district,b.district) ||
        a.small.localeCompare(b.small,'zh-Hant-TW',{numeric:true}) ||
        (a.rank==='未分排'?1:0)-(b.rank==='未分排'?1:0) ||
        a.rank.localeCompare(b.rank,'zh-Hant-TW',{numeric:true})
      );
      for(const area of sortedAreas){
        const inArea=area.people;
        blocks.push({kind:'rank',value:`${area.district ? area.district+'／' : ''}${area.small}／${area.rank}（${inArea.length} 人）`});
        const groupsFound=[...new Set(inArea.map(p=>p.reportGroup))].sort((a,b)=>{
          const ai=groups.indexOf(a),bi=groups.indexOf(b);
          return (ai<0?999:ai)-(bi<0?999:bi)||compareText(a,b);
        });
        for(const group of groupsFound){
          const peopleInGroup=inArea.filter(p=>p.reportGroup===group).sort((a,b)=>{
            const at=dateTime(a),bt=dateTime(b);
            if(at===null&&bt!==null)return -1;
            if(bt===null&&at!==null)return 1;
            return (at??0)-(bt??0)||compareText(a.name,b.name)||a.rowNumber-b.rowNumber;
          });
          blocks.push({kind:'group',value:`${group}（${peopleInGroup.length} 人）`});
          for(const person of peopleInGroup)blocks.push({kind:'person',person});
        }
      }
    }
    // v1.1.20: 真正逐頁編排。每個人工頁面都寫入自己的標題，
    // 不依賴 Excel 固定的 Print_Titles 來顯示動態聚會情況。
    const r=[];const merges=[];const breaks=[];let n=0;
    const add=(vals,style,height,merge=false,special)=>{n++;r.push(row(n,vals,style,height,special));if(merge)merges.push(`<mergeCell ref="A${n}:G${n}"/>`);return n;};
    const header=['姓名','小區','排組','最近出席','近12週','近1年','看望紀錄／備註'];
    const pageHeight=700; // 可用列高估算；由 Excel A4 直式實際列印預覽再校對。
    const commonHeaderHeight=30+24+24+34+26;
    let pageHeightUsed=0, pageNo=0, currentStatus=null;
    let currentRank=null,currentGroup=null;
    const heights={status:34,rank:21,group:19,person:24};
    const startPage=(status,continued=false)=>{
      if(pageNo>0)breaks.push(n);
      pageNo++;
      add(['召會生活看望名單'],1,30,true);
      add([`篩選：${filterLabel || '人員明細目前條件'}　｜　人數：${normalized.length} 人`],2,24,true);
      add([`資料截止週次：${lastDate}　｜　匯出日期：${testDate || today()}　｜　近12週：${weeks12.length} 週／近1年：${weeks52.length} 週`],2,24,true);
      add([`${status.value}${continued ? '（續）' : ''}`],9,34,true);
      add(header,3,26);
      pageHeightUsed=commonHeaderHeight;
    };
    const contentCapacity=pageHeight-commonHeaderHeight;
    const writeRank=(block,continued=false)=>{
      add([continued?block.value.replace(/（\d+ 人）$/,'（續）'):block.value],8,21,true);
      pageHeightUsed+=21;
    };
    const writeGroup=(block,continued=false)=>{
      add([continued?block.value.replace(/（\d+ 人）$/,'（續）'):block.value],8,19,true);
      pageHeightUsed+=19;
    };
    const writePerson=block=>{
      const person=block.person;
      add([person.name,person.smallDistrict,String(person.rank || ''),printableDate(person.lastAttendanceDate),`${person.count12}/${weeks12.length}`,`${person.count52}/${weeks52.length}`,''],4,24,false,[4,4,4,4,5,5,6]);
      pageHeightUsed+=24;
    };
    let i=0;
    while(i<blocks.length){
      const block=blocks[i];
      if(block.kind==='status'){
        currentStatus=block;currentRank=null;currentGroup=null;
        startPage(currentStatus,false);
        i++;continue;
      }
      if(block.kind==='rank'){
        currentRank=block;currentGroup=null;
        // 如果整排能裝進一頁，優先避免拆排。
        let total=0;
        for(let j=i;j<blocks.length&&blocks[j].kind!=='rank'&&blocks[j].kind!=='status';j++)total+=heights[blocks[j].kind];
        if(pageHeightUsed+total>pageHeight && total<=contentCapacity){
          startPage(currentStatus,true);
        } else if(pageHeightUsed+21+19+24>pageHeight){
          startPage(currentStatus,true);
        }
        writeRank(block);i++;continue;
      }
      if(block.kind==='group'){
        currentGroup=block;
        // 不讓羣組標題落單在頁尾。
        if(pageHeightUsed+19+24>pageHeight){
          startPage(currentStatus,true);
          if(currentRank)writeRank(currentRank,true);
        }
        writeGroup(block);i++;continue;
      }
      if(block.kind==='person'){
        if(pageHeightUsed+24>pageHeight){
          startPage(currentStatus,true);
          if(currentRank)writeRank(currentRank,true);
          if(currentGroup)writeGroup(currentGroup,true);
        }
        writePerson(block);i++;continue;
      }
      i++;
    }
    // 無資料時也產生可辨識的空白報表。
    if(pageNo===0){
      startPage({value:'無符合條件人員（0 人）'},false);
    }
    const sheet=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:G${n}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="24"/><cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="2" width="9" customWidth="1"/><col min="3" max="3" width="9" customWidth="1"/><col min="4" max="4" width="15" customWidth="1"/><col min="5" max="5" width="9" customWidth="1"/><col min="6" max="6" width="9" customWidth="1"/><col min="7" max="7" width="29" customWidth="1"/></cols><sheetData>${r.join('')}</sheetData><mergeCells count="${merges.length}">${merges.join('')}</mergeCells><printOptions horizontalCentered="1"/><pageMargins left="0.25" right="0.25" top="0.35" bottom="0.35" header="0.15" footer="0.15"/><pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0"/><rowBreaks count="${breaks.length}" manualBreakCount="${breaks.length}">${breaks.map(x=>`<brk id="${x}" man="1" max="16383"/>`).join('')}</rowBreaks></worksheet>`;
    const styles=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="4"><font><sz val="10"/><name val="Microsoft JhengHei"/></font><font><b/><sz val="16"/><name val="Microsoft JhengHei"/></font><font><b/><sz val="10"/><name val="Microsoft JhengHei"/></font><font><b/><sz val="15"/><name val="Microsoft JhengHei"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDCE9EE"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF1F4F6"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD4DCE0"/></left><right style="thin"><color rgb="FFD4DCE0"/></right><top style="thin"><color rgb="FFD4DCE0"/></top><bottom style="thin"><color rgb="FFD4DCE0"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="10"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    const wb=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="看望名單" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'看望名單'!$A$1:$G$${n}</definedName></definedNames></workbook>`;
    const entries=[
      ['[Content_Types].xml','<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'],
      ['_rels/.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
      ['xl/workbook.xml',wb],
      ['xl/_rels/workbook.xml.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
      ['xl/styles.xml',styles],['xl/worksheets/sheet1.xml',sheet]
    ];
    return {blob:zip(entries),count:normalized.length,lastDate};
  }
  root.createVisitationWorkbook=create;
})(window);
