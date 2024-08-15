const borderLine = require('./borderLine');
const {convertMsToDate} = require('./util');

handlePayload = function(newPayload, myPayload){
  let exportPayload = newPayload;
  exportPayload["title"] = (myPayload.title || '');
  exportPayload["unit"] = (myPayload.unit || '');
  return exportPayload;
}

renderXlsx = function(sheet, payload){
  function int2ColStr(i) {
    const beginI = 'A'.charCodeAt(0);
    let t;
    let n = i;
    let res = '';
    do {
      t = (n % 26);
      n = Math.floor(n/26)-1;
      res = String.fromCharCode(beginI + t) + res;
    } while (n >= 0);
    return res;
  };

  let reportTableIndex = "C9";
  /* Heading Table  */
  const headingTableFont = {
    name: 'Calibri',
    color: { argb: 'FFFFFFFF' },
    family: 2,
    size: 11,
    bold: true
  };
  sheet.getColumn('A').width = 2;

  // key-value table format
  sheet.getColumn('C').width = 22;
  sheet.getColumn('C').style.numFmt = 'hh:mm:ss dd/mm/yyyy';
  sheet.getCell('B3').value = 'Đối tượng';
  sheet.getCell('B4').value = 'Đ/v dữ liệu';
  sheet.getCell('B5').value = 'Ghi chú';
  sheet.getCell('B3').fill = sheet.getCell('B4').fill = sheet.getCell('B5').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4189B3' },
  };
  sheet.getCell('C3').alignment = sheet.getCell('C4').alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getCell('B3').font = sheet.getCell('B4').font = sheet.getCell('B5').font = headingTableFont;
  sheet.getCell('C3').value = payload.title;
  sheet.getCell('C4').value = payload.unit;

  // from-to format
  sheet.getCell('D3').value = 'Thời gian thu thập';
  sheet.getCell('D4').value = 'Từ';
  sheet.getCell('E4').value = 'Đến';
  sheet.mergeCells('D3:E3');
  sheet.getCell('D4').fill = sheet.getCell('E4').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4189B3' },
  };
  sheet.getCell('D3').alignment = sheet.getCell('D4').alignment = sheet.getCell('E4').alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getCell('D3').font = sheet.getCell('D4').font = sheet.getCell('E4').font = headingTableFont;
  sheet.getCell('D5').style.numFmt = sheet.getCell('E5').style.numFmt = 'hh:mm dd/mm/yyyy';
  sheet.getCell('D3').fill = sheet.getCell('E3').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF316886' },
  };
  // Render from/to time
  sheet.getCell('E5').style.font = sheet.getCell('D5').style.font = {
      name: 'Calibri',
      size: 10
  };
  if (payload.data != undefined && payload.data.length > 0) {
      let fromTs = payload.data[0][0];
      let toTs = payload.data[payload.data.length - 1][0];
      sheet.getCell('D5').value = convertMsToDate(fromTs);
      sheet.getCell('E5').value = convertMsToDate(toTs);
  }

  // Border (medium  double   thin)
  for (var r = 3; r <= 5; r++) {
      for (var c = 'B'.charCodeAt(0); c <= 'E'.charCodeAt(0); c++) {
          sheet.getCell(`${String.fromCharCode(c)}${r}`).border = { top: borderLine.inline, left: borderLine.inline, bottom: borderLine.inline, right: borderLine.inline };
      }
  }
  for (var c = 'B'.charCodeAt(0); c <= 'E'.charCodeAt(0); c++) {
      sheet.getCell(`${String.fromCharCode(c)}3`).border.top = borderLine.outline;
      sheet.getCell(`${String.fromCharCode(c)}5`).border.bottom = borderLine.outline;
  }
  for (var r = 3; r <= 5; r++) {
      sheet.getCell(`B${r}`).border.left = borderLine.outline;
      sheet.getCell(`E${r}`).border.right = borderLine.outline;
  }

  /* ========== Heading Table ==============  */
  // add a table to a sheet
  let spaceID = 0;
  let myColumns = payload.header.map((ele) => {
      let spaceArr = new Array(spaceID++).fill(' ');
      let spaceIDStr = spaceArr.join('');
      return { name: (spaceIDStr + ele + spaceIDStr), totalsRowFunction: 'average', filterButton: false }
  });
  if (myColumns.length > 0) {
      myColumns[0].totalsRowFunction = 'none';
      myColumns[0].filterButton = true;
      myColumns[0].totalsRowLabel = 'Trung bình:';
  }
  let myRows = payload.data.map((ele) => {
      return [convertMsToDate(ele[0]), ...ele.slice(1)];
  });

  if (myColumns.length > 0 && myRows.length > 0) {
      sheet.addTable({
          name: 'DataTable',
          ref: reportTableIndex,
          headerRow: true,
          totalsRow: true,
          style: payload.table_style || {
              theme: 'TableStyleMedium2',
              showColumnStripes: false
          },
          columns: myColumns,
          rows: myRows,
      });
  }

  sheet.getRow(String.fromCharCode(reportTableIndex.charCodeAt(1)-1)).alignment = sheet.getRow(String.fromCharCode(reportTableIndex.charCodeAt(1))).alignment = { vertical: 'middle', horizontal: 'center' };

  // Handle merge header
  let firstCol = reportTableIndex.charCodeAt(0);
  let cIter = firstCol-('A'.charCodeAt(0));
  let aboveRowIdxAsStr = String.fromCharCode(reportTableIndex.charCodeAt(1)-1);
  payload.merge_mark.forEach((m) => {
      if (m) {
          let rCell = `${int2ColStr(cIter)}${aboveRowIdxAsStr}`;
        //   console.log("Merge: ", `${rCell}:${int2ColStr(cIter + m.len - 1)}${aboveRowIdxAsStr}`);
          sheet.mergeCells(`${rCell}:${int2ColStr(cIter + m.len - 1)}${aboveRowIdxAsStr}`);
        //   console.log("=========== merged ==========");
          sheet.getCell(rCell).fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FF4189B3' },
          };
          sheet.getCell(rCell).value = m.text;
          sheet.getCell(rCell).font = {
              color: { argb: 'FF16365C' },
              alignment: { vertical: 'middle', horizontal: 'center' },
              bold: true
          };
          sheet.getCell(rCell).border = { top: borderLine.outline, left: borderLine.outline, right: borderLine.outline, bottom: borderLine.outline };
          cIter += m.len;
      }
      else {
        cIter++;
      }
  });
  //
  cIter = firstCol-('A'.charCodeAt(0));
  for (var i = 0; i < payload.header.length; i++){
    let rCell = `${int2ColStr(cIter)}${reportTableIndex[1]}`;
    sheet.getCell(rCell).border = { left: borderLine.outline, right: borderLine.outline, top: borderLine.outline, bottom: borderLine.outline };
    cIter++;
  }
}

const projectId = "thingsboard";

module.exports = {handlePayload, renderXlsx, projectId};