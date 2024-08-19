const borderLine = require('./borderLine');
const {convertMsToDate, convertISO8601StrToDate} = require('./util');

handlePayload = function(newPayload, myPayload){
  let exportPayload = newPayload;
  if (myPayload.from) exportPayload["from"] = myPayload.from;
  if (myPayload.to) exportPayload["to"] = myPayload.to;
  if (myPayload.numFmts) exportPayload['numFmts'] = myPayload.numFmts;
  if (myPayload.align) exportPayload['align'] = myPayload.align;
  // key-value  pairTable[i] = {"key": <key label>, "value": <value>}
  exportPayload["pairTable"] = (myPayload.pairTable || []);
  return exportPayload;
}

renderXlsx = function(sheet, payload){
  let reportTableIndex = "B9";
  /* Heading Table  */
  const headingTableFont = {
    name: 'Calibri',
    color: { argb: 'FFFFFFFF' },
    family: 2,
    size: 11,
    bold: true
  };
  sheet.getColumn('A').width = 2;
  sheet.getColumn('E').width = 20;
  sheet.getColumn('F').width = 20;

  // from-to format
  sheet.getCell('E3').value = 'Thời gian thu thập';
  sheet.getCell('E4').value = 'Từ';
  sheet.getCell('F4').value = 'Đến';
  sheet.mergeCells('E3:F3');
  sheet.getCell('E4').fill = sheet.getCell('F4').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4189B3' },
  };
  sheet.getCell('E5').alignment = sheet.getCell('F5').alignment = sheet.getCell('E3').alignment = sheet.getCell('E4').alignment = sheet.getCell('F4').alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getCell('E3').font = sheet.getCell('E4').font = sheet.getCell('F4').font = headingTableFont;
  sheet.getCell('E5').style.numFmt = sheet.getCell('F5').style.numFmt = 'hh:mm:ss dd/mm/yyyy';
  sheet.getCell('E3').fill = sheet.getCell('F3').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF316886' },
  };
  // Render from/to time
  sheet.getCell('F5').style.font = sheet.getCell('E5').style.font = {
      name: 'Calibri',
      size: 10
  };
  if (payload.from != undefined && payload.to != undefined) {
      if (typeof payload.from === 'number') sheet.getCell('E5').value = convertMsToDate(payload.from);
      else sheet.getCell('E5').value = convertISO8601StrToDate(payload.from, payload.addingHourToDate);
      if (typeof payload.to === 'number') sheet.getCell('F5').value = convertMsToDate(payload.to);
      else sheet.getCell('F5').value = convertISO8601StrToDate(payload.to, payload.addingHourToDate);
  }
  // render border
  sheet.getCell('E3').border = { top: borderLine.outline, left: borderLine.outline, right: borderLine.outline, bottom: borderLine.inline };
  sheet.getCell('E4').border = { top: borderLine.inline, left: borderLine.outline, right: borderLine.inline, bottom: borderLine.inline };
  sheet.getCell('F4').border = { top: borderLine.inline, left: borderLine.inline, right: borderLine.outline, bottom: borderLine.inline };
  sheet.getCell('E5').border = { top: borderLine.inline, left: borderLine.outline, right: borderLine.inline, bottom: borderLine.outline };
  sheet.getCell('F5').border = { top: borderLine.inline, left: borderLine.inline, right: borderLine.outline, bottom: borderLine.outline };

  // key-value
  sheet.getColumn('C').width = 22;

  for (let i = 0; i < payload.pairTable.length; i++) {
    sheet.getCell(`B${3+i}`).value = (payload.pairTable[i].key+':');
    sheet.getCell(`C${3+i}`).value = payload.pairTable[i].value;
    sheet.getCell(`B${3+i}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4189B3' },
    };
    sheet.getCell(`B${3+i}`).font = headingTableFont;
    sheet.getCell(`C${3+i}`).alignment = { vertical: 'middle', horizontal: 'right' };
    sheet.getCell(`B${3+i}`).border = { top: borderLine.outline, left: borderLine.outline, right: borderLine.inline, bottom: borderLine.outline };
    sheet.getCell(`C${3+i}`).border = { top: borderLine.outline, left: borderLine.inline, right: borderLine.outline, bottom: borderLine.outline };
  }

  let addRow = payload.pairTable.length-3;
  if (addRow < 0) addRow = 0;
  let reportTableRowIndex = parseInt(reportTableIndex[1]);
  reportTableAboveIndex = `${reportTableIndex[0]}${(reportTableRowIndex+addRow).toString()}`;
  reportTableIndex = `${reportTableIndex[0]}${(reportTableRowIndex+addRow+1).toString()}`;
  
  
  // Handle merge header
  let reportTableIndexCharAtZero = reportTableIndex.charCodeAt(0);
  let c_iter = reportTableIndexCharAtZero;
  let rowIdxAsStr = reportTableIndex.substring(1);
  let aboveRowIdxAsStr = reportTableAboveIndex.substring(1);
  let isHasAnyMerge = payload.merge_mark.some((e)=>{return e!=undefined;});
  payload.merge_mark.forEach((m) => {
    let rCell = `${String.fromCharCode(c_iter)}${aboveRowIdxAsStr}`;
    if (m) {
      sheet.mergeCells(`${rCell}:${String.fromCharCode(c_iter + m.len - 1)}${aboveRowIdxAsStr}`);
      sheet.getCell(rCell).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'F00287B5' },
      };
      sheet.getCell(rCell).value = m.text;
      sheet.getCell(rCell).font = {
        name: 'Calibri',
        color: { argb: 'FFFFFFFF' },
        family: 2,
        size: 11,
        bold: true,
      };
      sheet.getCell(rCell).alignment = { vertical: 'middle', horizontal: 'center' };
      sheet.getCell(rCell).border = { top: borderLine.blueoutline, left: borderLine.blueoutline, right: borderLine.blueoutline, bottom: borderLine.blueoutline };
      c_iter += m.len;
    }
    else {
      if (isHasAnyMerge){
        let bCell = `${String.fromCharCode(c_iter)}${rowIdxAsStr}`;
        sheet.mergeCells(`${rCell}:${bCell}`);
      }
      c_iter++;
    }
  });
  //
  c_iter = reportTableIndexCharAtZero;
  for (var i = 0; i < payload.header.length; i++){
    let rCell = `${String.fromCharCode(c_iter)}${reportTableIndex.substring(1)}`;
    sheet.getCell(rCell).border = { left: borderLine.blueoutline, right: borderLine.blueoutline, top: borderLine.blueoutline, bottom: borderLine.blueoutline };
    sheet.getCell(rCell).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'F00287B5' },
    };
    sheet.getCell(rCell).font = headingTableFont;
    sheet.getCell(rCell).value = payload.header[i];
    sheet.getCell(rCell).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    c_iter++;
  }
  //
  c_iter = reportTableIndexCharAtZero;
  payload.merge_mark.forEach((m) => {
    if (m) {
      let ll = m.len;
      for (let i = 0; i < ll; i++) {
        let rCell = `${String.fromCharCode(c_iter)}${rowIdxAsStr}`;
        if (i < ll-1) sheet.getCell(rCell).border.right = borderLine.darkblueinline;
        if (i > 0) sheet.getCell(rCell).border.left = borderLine.darkblueinline;
        c_iter++;
      }
    } else {
      c_iter++;
    }
  });

  // render cell
  let rLen = payload.data.length;
  let cLen = payload.header.length;
  let rowIterIndex = reportTableRowIndex+addRow+2;
  let colIter;
  for (var r = 0; r < rLen; r++) {
    let mergeRowMark = {};
    colIter = reportTableIndexCharAtZero;
    for (var c = 0; c < cLen; c++) {
      let cellIdx = `${String.fromCharCode(colIter)}${rowIterIndex}`;
      let cell = sheet.getCell(cellIdx);
      //
      if (c == 0) cell.border = { left: borderLine.blueoutline, right: borderLine.blueinline, bottom: borderLine.blueinline };
      else if (c == cLen-1) cell.border = { left: borderLine.blueinline, right: borderLine.blueoutline, bottom: borderLine.blueinline };
      else cell.border = { left: borderLine.blueinline, right: borderLine.blueinline, bottom: borderLine.blueinline };
      if (r == rLen-1) cell.border.bottom = borderLine.blueoutline;
      //
      if (r % 2 == 1) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'E0C0EDFC' },
        };
      }
      //
      if (payload.align!=undefined && payload.align[c] != undefined && payload.align[c] != ""){
        cell.alignment = { vertical: 'middle', horizontal: payload.align[c] };
      }
      //
      let numFmtType = 0;// 1 => date
      if (payload.numFmts!=undefined && payload.numFmts[c] != undefined && payload.numFmts[c] != "") {
        cell.style.numFmt = payload.numFmts[c];
        if (cell.style.numFmt == "hh:mm:ss dd/mm/yyyy") numFmtType = 1;
      }
      //
      if (payload.data[r][c]!="undefined"){
        if (numFmtType == 1) cell.value = convertISO8601StrToDate(payload.data[r][c], payload.addingHourToDate);
        else cell.value = payload.data[r][c];
      }
      else
      {
        if (c > 0){
          let prevCellIdx = `${String.fromCharCode(colIter-1)}${rowIterIndex}`;
          mergeRowMark[prevCellIdx] = cellIdx;
        }
        cell.value = "";
      }
      //
      colIter++;
    }
    //
    let isHasAnyMergeRow = false;
    while (true) {
      let keys = Object.keys(mergeRowMark);
      if (keys.length === 0) break;
      isHasAnyMergeRow = true;
      let firstCellIdx = keys[0];
      let secondCellIdx = firstCellIdx;
      while (mergeRowMark.hasOwnProperty(secondCellIdx)) {
        let tmp = mergeRowMark[secondCellIdx];
        delete mergeRowMark[secondCellIdx];
        secondCellIdx = tmp;
      }
      sheet.mergeCells(`${firstCellIdx}:${secondCellIdx}`);
      sheet.getCell(firstCellIdx).alignment = { vertical: 'middle', horizontal: 'center' };
    }
    if (isHasAnyMergeRow){
      sheet.getCell(`${String.fromCharCode(reportTableIndexCharAtZero)}${rowIterIndex}`).border.left = borderLine.blueoutline;
      sheet.getCell(`${String.fromCharCode(reportTableIndexCharAtZero+cLen-1)}${rowIterIndex}`).border.right = borderLine.blueoutline;
    }
    //
    rowIterIndex++;
  }
}

renderPdf = function(doc, payload) {

};

const projectId = "chuan";

module.exports = {handlePayload, renderXlsx, renderPdf, projectId};