const pdf = require('pdfjs');
const fs = require('fs');
const path = require('path');
const moment = require('moment');

const borderLine = require('./borderLine');
const {convertMsToDate} = require('./util');

const TotalFunctionsLabel = {
    average: 'Trung bình',
    sum: 'Tổng',
};
const calDataTotal = (data, dataColAmount, totalFunc) => {
    const dataLen = data.length;
    let dataTotal = new Array(dataColAmount).fill(0);

    if (totalFunc === 'average' || totalFunc === 'sum') {
        // sum data
        for (let i = 0; i < data.length; i++) {
            for (let j = 1; j < dataColAmount + 1; j++) {
                dataTotal[j-1] += data[i][j];
            }
        }
    }
    if (totalFunc === 'average') {
        // average data
        for (let i = 0; i < dataColAmount; i++) {
            dataTotal[i] = Math.round(dataTotal[i] / dataLen * 100) / 100;
        }
    }
};
///////////////////////////////////////////////
const PDF_CellConfig = {
    alignment: 'center',
    textAlign: 'center',
};
const PDF_PrimaryColor = 0x0388fc;
const PDF_PrimaryContentColor = 0xffffff;
const PDF_SecondaryColor = 0x005bab;
const PDF_SecondaryContentColor = 0xffffff;
const PDF_MediumColor = 0x616161;
const PDF_LightColor = 0xd6d6d6;
const PDF_ContentColor = 0x000000;
const PDF_ContentBgColor = 0xffffff;
//
const PDF_FontSize = 12;
const PDF_Font = new pdf.Font(fs.readFileSync(path.join(__dirname, './fonts/Times.otf')));
//
const PDF_ReportTitleFontSize = 24;
//
const PDF_DateTimeFormat = 'HH:mm:ss DD/MM/YYYY';
///////////////////////////////////////////////

handlePayload = function(newPayload, myPayload){
  let exportPayload = newPayload;
  exportPayload["title"] = (myPayload.title || '');
  exportPayload["unit"] = (myPayload.unit || '');
  exportPayload["fromTs"] = (myPayload.fromTs);
  exportPayload["toTs"] = (myPayload.toTs);
  exportPayload["totalFunc"] = (myPayload.totalFunc || 'average');
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
      let fromTs = payload.fromTs || payload.data[0]?.[0] || 0;
      let toTs = payload.toTs || payload.data[payload.data.length - 1]?.[0] || 0;
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
      return { name: (spaceIDStr + ele + spaceIDStr), totalsRowFunction: payload.totalFunc || "average", filterButton: false }
  });
  if (myColumns.length > 0) {
      myColumns[0].totalsRowFunction = 'none';
      myColumns[0].filterButton = true;
      myColumns[0].totalsRowLabel = TotalFunctionsLabel[payload.totalFunc || "average"]+':';
  }
  let myRows = payload.data.map((ele) => {
      return [...ele];
  });

  if (myColumns.length > 0 && myRows.length > 0) {
      sheet.addTable({
          name: 'DataTable',
          ref: reportTableIndex,
          headerRow: true,
          totalsRow: true,
          style: payload.table_style || {
              theme: 'TableStyleMedium2',
              showColumnStripes: false,
          },
          columns: myColumns,
          rows: myRows,
      });
  }

  sheet.getRow(String.fromCharCode(reportTableIndex.charCodeAt(1)-1)).alignment = sheet.getRow(String.fromCharCode(reportTableIndex.charCodeAt(1))).alignment = { vertical: 'middle', horizontal: 'center' };

  // Handle merge header
  if ((payload.merge_mark || []).some((m) => m)) {
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
}

// return doc
renderPdf = function(payload) {
    const { heading='', title='', unit='', data=[], merge_mark:mergeMark=[], header=[], note='', table_style: style={}, fromTs: tableFromTs, toTs: tableToTs, totalFunc='average' } = payload;
    const { dataCell={} } = style;
    const { width: dataCellWidth=50 } = dataCell;

    const fromTs = tableFromTs || data[0]?.[0] || 0;
    const toTs = tableToTs || data[payload.data.length - 1]?.[0] || 0;
    const fromStr = moment(fromTs).format(PDF_DateTimeFormat);
    const toStr = moment(toTs).format(PDF_DateTimeFormat);

    const dataColAmount = header.length - 1;

    // calculate data total
    const dataTotal = calDataTotal(data, dataColAmount, totalFunc);

    const doc = new pdf.Document({
        font: PDF_Font,//require('pdfjs/font/Times'),
        fontSize: PDF_FontSize,
        padding: 10,
        width: (10+10) + Math.max((120+dataCellWidth*dataColAmount), (120+200+120+120)),
        properties: {
            title: 'Huynh Duc Nham',
            author: 'Huynh Duc Nham',
        }
        });
    
    // title align center having primary color
    doc.text(heading, {textAlign: 'center', color: PDF_PrimaryColor, fontSize: PDF_ReportTitleFontSize});

    // gap between title and info table
    doc.cell('', {minHeight: 20});

    // info table
    const infoTable = doc.table({
        widths: [120, 200, 120, 120],
        borderWidth: 1,
	    padding: 5,
        borderColor: PDF_MediumColor,
    });
    const infoRow1 = infoTable.row();
    infoRow1.cell("Đối tượng", {paddingLeft: 6, color: PDF_PrimaryContentColor, backgroundColor: PDF_PrimaryColor});
    infoRow1.cell(title, {color: PDF_ContentColor, backgroundColor: PDF_ContentBgColor});
    infoRow1.cell("Thời gian thu thập", {colspan: 2, textAlign: 'center', alignment: 'center', backgroundColor: PDF_SecondaryColor, color: PDF_SecondaryContentColor});
    //
    const infoRow2 = infoTable.row();
    infoRow2.cell("Đ/v dữ liệu", {paddingLeft: 6, color: PDF_PrimaryContentColor, backgroundColor: PDF_PrimaryColor})
    infoRow2.cell(unit, {color: PDF_ContentColor, backgroundColor: PDF_ContentBgColor})
    infoRow2.cell("Từ", {textAlign: 'center', alignment: 'center', backgroundColor: PDF_PrimaryColor, color: PDF_PrimaryContentColor})
    infoRow2.cell("Đến", {textAlign: 'center', alignment: 'center', backgroundColor: PDF_PrimaryColor, color: PDF_PrimaryContentColor});
    //
    const infoRow3 = infoTable.row();
    infoRow3.cell("Ghi chú", {paddingLeft: 6, color: PDF_PrimaryContentColor, backgroundColor: PDF_PrimaryColor})
    infoRow3.cell(note, {color: PDF_ContentColor, backgroundColor: PDF_ContentBgColor})
    infoRow3.cell(fromStr, {textAlign: 'center', alignment: 'center', color: PDF_ContentColor, backgroundColor: PDF_ContentBgColor})
    infoRow3.cell(toStr, {textAlign: 'center', alignment: 'center', color: PDF_ContentColor, backgroundColor: PDF_ContentBgColor});

    // gap between info table and content
    doc.cell('', {minHeight: 24});

    // render something onto the document
    const table = doc.table({
        widths: [120, ...(new Array(dataColAmount).fill(dataCellWidth))],
        borderWidth: 1,
	    padding: 5,
        borderColor: PDF_MediumColor,
    })

    if (mergeMark.some((m) => m)) {
        const aboveHeader = table.header({

        });
        for (let mergeHeader of mergeMark) {
            const text = mergeHeader?.text || '';
            const span = mergeHeader?.len || 1;
            aboveHeader.cell(text, {
                alignment: 'center',
                textAlign: 'center',
                colspan: span,
                ...(mergeHeader ? {backgroundColor: PDF_PrimaryColor} : {}),
            });
        }
    }

    const belowHeader = table.header({

    });
    for (let headerItem of header) {
        belowHeader.cell(headerItem, {
            ...PDF_CellConfig,
        });
    };

    for (let dataItem of data) {
        const [index, ...dataCols] = dataItem;
        //
        const row = table.row({
            ...PDF_CellConfig,
        });
        // first col is index
        row.cell(index.toString(), {
            ...PDF_CellConfig,
        });
        // other cols are data
        for (let dataCol of dataCols) {
            row.cell(dataCol.toString(), {
                ...PDF_CellConfig,
            });
        }
    }

    const avgRow = table.row({
        ...PDF_CellConfig,
        backgroundColor: PDF_LightColor,
    });
    avgRow.cell(TotalFunctionsLabel[totalFunc]+':', {
        ...PDF_CellConfig,
    });
    // avg row
    for (let avgData of dataTotal) {
        avgRow.cell(avgData.toString(), {
            ...PDF_CellConfig,
        });      
    }

    return doc;
};

const projectId = "general";

module.exports = {handlePayload, renderXlsx, renderPdf, projectId};
