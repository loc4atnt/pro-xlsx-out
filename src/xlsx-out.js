const ExcelJS = require('exceljs');
const {handlePayload, renderXlsx} = require('./render');

// Extract columns and rows
// Filter data & header
function handleRawPayload(myPayload, projectId = undefined) {
    myPayload.header.forEach((e, i, arr) => { if (Array.isArray(e) && e.length == 1) arr[i] = e[0]; });
    myPayload.data.forEach((d) => { d.forEach((e, i, arr) => { if (Array.isArray(e) && e.length == 1) arr[i] = e[0]; }); });

    const tableHeaders = myPayload.header.reduce((res, h) => {
        if (Array.isArray(h)) {
            for (let i = 1; i < h.length; i++) res.push(h[i]);
        } else { res.push(h); }
        return res;
    }, []);

    const tableData = myPayload.data.map((r) => {
        return r.reduce((res, h) => {
            if (Array.isArray(h)) {
                res = res.concat(h);
            } else { res.push(h); }
            return res;
        }, [])
    });

    const mergeMark = myPayload.header.reduce((res, h) => {
        if (Array.isArray(h)) {
            res.push({ "text": h[0], "len": (h.length - 1) });
        } else res.push(undefined);
        return res;
    }, []);

    var newPayload = {
        "header": tableHeaders,
        "data": tableData,
        "merge_mark": mergeMark,
        "heading": (myPayload.heading || 'BÁO CÁO'),
    };
    if (myPayload.table_style) newPayload['table_style'] = myPayload.table_style;
    if (myPayload.addingHourToDate) newPayload['addingHourToDate'] = myPayload.addingHourToDate;
    if (projectId != undefined) return handlePayload(newPayload, myPayload, projectId);
    return newPayload;
}

function exportReportAsXlsx(payload, projectId, isExportFile=false) {
    const handledPayload = handleRawPayload(payload, projectId);

    const workbook = new ExcelJS.Workbook();

    workbook.title = handledPayload.heading;
    workbook.creator = 'Phenikaa MaaS';
    workbook.created = new Date();
    workbook.properties.date1904 = true;// Set workbook dates to 1904 date system

    const sheet = workbook.addWorksheet("Trang 1", {
        properties: {
            defaultColWidth: 15
        },
        views: [{ showGridLines: false }]
    });

    // Heading
    sheet.getRow('1').height = 36;
    sheet.getCell('B1').value = handledPayload.heading;
    sheet.getCell('B1').style.font = {
        name: 'Tahoma',
        color: { argb: 'FF316886' },
        family: 2,
        size: 22,
        bold: true
    };

    renderXlsx(sheet, handledPayload, projectId);

    if (isExportFile) workbook.xlsx.writeFile("hihi.xlsx");
    let buffer = workbook.xlsx.writeBuffer();// write to a new buffer
    return buffer;
}

module.exports = {
    exportReportAsXlsx,
}

/////////////////////////////////////////////////////////
// LƯU Ý:   CỘT 1 luôn là cột thứ tự
//          CỘT 2 luôn là cột Biển số xe
// const headers = {
//     "route_hop_chuan": {
//         "heading": "Báo cáo B.1. Hành trình xe chạy",
//         "header": ["TT", "Biển số xe", "Thời điểm", ["Tọa độ", "Kinh độ", "Vĩ độ"], "Địa điểm", "Ghi chú"],
//         "numFmts": ["", "", "hh:mm:ss dd/mm/yyyy", "", "", "", ""],
//         "align": ["center", "", "", "", "", "", ""],
//         "dataIndexes": ["stt", "deviceId", "fixTime", "longitude", "latitude", "address", "note"]
//     },
//     "trips_hop_chuan": {
//         "heading": "Báo cáo B.1.2. Lộ trình xe chạy",
//         "header": ["TT", "Biển số xe", "Họ tên lái xe", "Số Giấy phép lái xe", ["Thời điểm bắt đầu", "Thời điểm", "Kinh độ", "Vĩ độ", "Địa điểm"], ["Thời điểm kết thúc", "Thời điểm", "Kinh độ", "Vĩ độ", "Địa điểm"], "Khoảng cách (Km)", "Khoảng thời gian", "Tốc độ trung bình (Km/H)", "Tốc độ cao nhất (Km/H)", "Ghi chú"],
//         "numFmts": ["", "", "", "", "hh:mm:ss dd/mm/yyyy", "", "", "", "hh:mm:ss dd/mm/yyyy", "", "", "", "", "", "", "", ""],
//         "align": ["center", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
//         "dataIndexes": ["stt", "deviceId", "driverName", "driverUniqueId", "startTime", "startLon", "startLat", "startAddress", "endTime", "endLon", "endLat", "endAddress", "distance", "duration", "averageSpeed", "maxSpeed", "note"],
//     },
//     "speed_hop_chuan": {
//         "heading": "Báo cáo B.2.1. Tốc độ của xe",
//         "header": ["TT", "Biển số xe", "Thời điểm", "Các tốc độ", "Ghi chú"],
//         "numFmts": ["", "", "hh:mm:ss dd/mm/yyyy", "", ""],
//         "align": ["center", "", "", "", ""],
//         "dataIndexes": ["stt", "deviceId", "fixTime", "speeds", "note"]
//     },
//     "speed_limit_hop_chuan": {
//         "heading": "Báo cáo B.2.2. Quá tốc độ giới hạn",
//         "header": ["TT", "Biển số xe", "Họ tên lái xe", "Số Giấy phép lái xe", "Loại hình hoạt động", "Thời điểm", "Tốc độ trung bình khi quá tốc độ giới hạn (km/h)", "Tốc độ giới hạn (km/h)", ["Tọa độ quá tốc độ giới hạn", "Kinh độ", "Vĩ độ"], "Địa điểm quá tốc độ giới hạn", "Ghi chú"],
//         "numFmts": ["", "", "", "", "", "hh:mm:ss dd/mm/yyyy", "", "", "", "", "", ""],
//         "align": ["center", "", "", "", "", "", "", "", "", "", "", ""],
//         "dataIndexes": ["stt", "deviceId", "driverName", "driverUniqueId", "category", "startTime", "avgSpeed", "speedLimit", "longitude", "latitude", "address", "note"]
//     },
//     "trips_by_driver_hop_chuan": {
//         "heading": "Báo cáo B.3. Thời gian lái xe liên tục",
//         "header": ["TT", "Biển số xe", "Họ tên lái xe", "Số Giấy phép lái xe", "Loại hình hoạt động", ["Thời điểm bắt đầu", "Thời điểm", "Kinh độ", "Vĩ độ", "Địa điểm"], ["Thời điểm kết thúc", "Thời điểm", "Kinh độ", "Vĩ độ", "Địa điểm"], "Thời gian lái xe", "Ghi chú"],
//         "numFmts": ["", "", "", "", "", "hh:mm:ss dd/mm/yyyy", "", "", "", "hh:mm:ss dd/mm/yyyy", "", "", "", "", ""],
//         "align": ["center", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
//         "dataIndexes": ["stt", "deviceId", "driverName", "driverUniqueId", "category", "startTime", "startLon", "startLat", "startAddress", "endTime", "endLon", "endLat", "endAddress", "duration", "note"]
//     },
//     "stops_hop_chuan": {
//         "heading": "Báo cáo B.4. Dừng đỗ",
//         "header": ["TT", "Biển số xe", "Họ tên lái xe", "Số Giấy phép lái xe", "Loại hình hoạt động", "Thời điểm dừng đỗ", "Thời gian dừng đỗ", ["Tọa độ dừng đỗ", "Kinh độ", "Vĩ độ"], "Địa điểm dừng đỗ", "Ghi chú"],
//         "numFmts": ["", "", "", "", "", "hh:mm:ss dd/mm/yyyy", "", "", "", "", ""],
//         "align": ["center", "", "", "", "", "", "center", "", "", "", ""],
//         "dataIndexes": ["stt", "deviceId", "driverName", "driverUniqueId", "category", "startTime", "duration", "longitude", "latitude", "address", "note"]
//     },
//     "speed_limit_stops_hop_chuan": {
//         "heading": "Báo cáo B.5.1. Báo cáo tổng hợp theo xe",
//         "header": ["TT", "Biển số xe", "Loại hình hoạt động", "Tổng KM", ["Tỷ lệ km quá tốc độ giới hạn/tổng km (%)", "Tỷ lệ quá tốc độ từ 5 km/h đến dưới 10 km/h", "Tỷ lệ quá tốc độ từ 10 km/h đến dưới 20 km/h", "Tỷ lệ quá tốc độ từ 20 km/h đến 35 km/h", "Tỷ lệ quá tốc độ trên 35 km/h"], ["Tổng số lần quá tốc độ giới hạn (lần)", "Số lần quá tốc độ từ 5 km/h đến dưới 10 km/h", "Số lần quá tốc độ từ 10 km/h đến dưới 20 km/h", "Số lần quá tốc độ từ 20 km/h đến 35 km/h", "Số lần quá tốc độ trên 35 km/h"], "Tổng số lần dừng đỗ", "Ghi chú"],
//         "numFmts": ["", "", "", "", , "", "", "", "", "", "", "", "", "", ""],
//         "align": ["center", "", "", "", , "", "", "", "", "", "", "", "", "", ""],
//         "dataIndexes": ["stt", "deviceId", "category", "distance", "over_speed_ratio_0", "over_speed_ratio_1", "over_speed_ratio_2", "over_speed_ratio_3", "breakout_speed_ratio_0", "breakout_speed_ratio_1", "breakout_speed_ratio_2", "breakout_speed_ratio_3", "stop_times", "note"]
//     },
//     "speed_limit_stops_by_driver_hop_chuan": {
//         "heading": "Báo cáo B.5.2. Báo cáo tổng hợp theo lái xe",
//         "header": ["TT", "Họ tên lái xe", "Số Giấy phép lái xe", "Tổng KM", ["Tỷ lệ km quá tốc độ giới hạn/tổng km (%)", "Tỷ lệ quá tốc độ từ 5 km/h đến dưới 10 km/h", "Tỷ lệ quá tốc độ từ 10 km/h đến dưới 20 km/h", "Tỷ lệ quá tốc độ từ 20 km/h đến 35 km/h", "Tỷ lệ quá tốc độ trên 35 km/h"], ["Tổng số lần quá tốc độ giới hạn (lần)", "Số lần quá tốc độ từ 5 km/h đến dưới 10 km/h", "Số lần quá tốc độ từ 10 km/h đến dưới 20 km/h", "Số lần quá tốc độ từ 20 km/h đến 35 km/h", "Số lần quá tốc độ trên 35 km/h"], "Tổng số lần lái xe liên tục quá 04 giờ", "Ghi chú"],
//         "numFmts": ["", "", "", "", , "", "", "", "", "", "", "", "", "", ""],
//         "align": ["center", "", "", "", , "", "", "", "", "", "", "", "", "", ""],
//         "dataIndexes": ["stt", "driverName", "driverUniqueId", "distance", "over_speed_ratio_0", "over_speed_ratio_1", "over_speed_ratio_2", "over_speed_ratio_3", "breakout_speed_ratio_0", "breakout_speed_ratio_1", "breakout_speed_ratio_2", "breakout_speed_ratio_3", "_4hour_times", "note"]
//     }
// }
// function capitalizeFirstLetter(string) {
//     return string.charAt(0).toUpperCase() + string.slice(1);
//   }
// function getDeviceAndGroupInfo(payload, isMergeDeviceGroupId=true){
//     payload.data.devices = payload.data.devices.reduce((acc, e)=>{
//         if (payload.data.deviceId.some((i)=>{return e.id === i;})) {
//             acc[e.id.toString()] = e.name;
//             if (isMergeDeviceGroupId) payload.data.groupId.push(e.groupId);
//         }
//         return acc;
//     }, {});
//     payload.data.groups = payload.data.groups.reduce((acc, e)=>{
//         if (payload.data.groupId.some((i)=>{return e.id === i;})) {
//             acc[e.id.toString()] = e.name;
//         }
//         return acc;
//     }, {});
//     return payload;
// }
// function renderData(payload){
//     let newPayload = {};

//     newPayload.from = payload.data.params.from;
//     newPayload.to = payload.data.params.to;

//     newPayload.rpCode = payload.data.reportType;
//     const myHeader = headers[newPayload.rpCode];
//     if (myHeader != undefined){
//         newPayload.heading = myHeader.heading;
//         newPayload.header = myHeader.header;
//         newPayload.numFmts = myHeader.numFmts;
//         newPayload.align = myHeader.align;
//     }
    
//     let pairTable = [];
//     // don vi kinh doanh
//     let businessPairObj = {"key": "Đơn vị kinh doanh vận tải"};
//     businessPairObj.value = Object.keys(payload.data.groups).reduce((acc, kg, i, a)=>{let g = payload.data.groups[kg]; acc += g; if (i < a.length-1) acc += ", "; return acc;}, "");
//     pairTable.push(businessPairObj);
//     if (newPayload.rpCode != "speed_limit_stops_by_driver_hop_chuan"){
//         // bien so xe
//         let devicesPairObj = {"key": "Biển số xe"};
//         devicesPairObj.value = Object.keys(payload.data.devices).reduce((acc, kd, i, a)=>{let d = payload.data.devices[kd]; acc += d; if (i < a.length-1) acc += ", "; return acc;}, "");
//         pairTable.push(devicesPairObj);
//         // lai xe lien tuc 4h hoac chon tat ca
//         if (newPayload.rpCode == "trips_by_driver_hop_chuan") {
//             let _4hPairObj = {"key": "Phân loại"};
//             if (payload.data.endpoint.indexOf("driverOver4Hours=true")!=-1) _4hPairObj.value = "Lái xe liên tục quá 04h";
//             else _4hPairObj.value = "Tất cả";
//             pairTable.push(_4hPairObj);
//         }
//     } else {// B.5.2
//         //
//     }
//     //
//     newPayload.pairTable = pairTable;

//     if (myHeader != undefined){
//         const dataIndexes = myHeader.dataIndexes;
//         let data = payload.data.items.map((raw)=>{
//             let row = [];
//             for (let i = 0; i < dataIndexes.length; i++){
//                 if (dataIndexes[i] == "deviceId") row.push(payload.data.devices[raw[dataIndexes[i]]]);
//                 else if (dataIndexes[i] == "category") {
//                     let categoryKey = ("category"+capitalizeFirstLetter(raw[dataIndexes[i]]));
//                     row.push(payload.data.strings[categoryKey]);
//                 }
//                 else if (dataIndexes[i] == "duration") {
//                     let duration = raw[dataIndexes[i]]/1000;
//                     let durationAsMin = Math.floor(duration/60);
//                     let durationAsHour = Math.floor(durationAsMin/60);
//                     let minRemain = durationAsMin%60;
//                     row.push(`${durationAsHour}h ${minRemain}m`);
//                 }
//                 else if (dataIndexes[i] == "distance" || dataIndexes[i] == "maxSpeed" || dataIndexes[i] == "averageSpeed") {
//                     row.push(Math.floor(raw[dataIndexes[i]]*100)/100);
//                 }
//                 else {
//                     if (raw[dataIndexes[i]]!=undefined) row.push(raw[dataIndexes[i]]);
//                     else row.push("");
//                 }
//             }
//             return row;
//         });
//         //
//         if (newPayload.rpCode == "stops_hop_chuan"){
//             let totalDuration = payload.data.items.reduce((acc, e)=>{
//                 acc += e.duration/1000;
//                 return acc;
//             }, 0);
//             let totalDurationAsMin = Math.floor(totalDuration/60);
//             let totalDurationAsHour = Math.floor(totalDurationAsMin/60);
//             let totalMinRemain = totalDurationAsMin%60;
//             data.push(["Tổng", "undefined", "undefined", "undefined", "undefined", "undefined", `${totalDurationAsHour}h ${totalMinRemain}m`, "undefined", "undefined", "undefined", "undefined"]);
//         }
//         newPayload.data = data;
//     }

//     return newPayload;
// }
// function test(payload){
//     payload = getDeviceAndGroupInfo(payload, payload.data.reportType!="speed_limit_stops_by_driver_hop_chuan");
//     payload = renderData(payload);
//     return payload;
// }
/////////////////////////////////////////////////////////
// const ppll1 = {"title":"Bà Chiểu","unit":"°C","header":["Thời gian",["P. Lạnh (Điểm 1)","MIN","AVG","MAX"],["P. Lạnh (Điểm 2)","MIN","AVG","MAX"]],"data":[[1660795985291,-24.9,-23.98,-22.8,-25.6,-24.57,-23.1],[1660799585291,-26,-24.87,-21.9,-26.6,-25.37,-21.9],[1660803185291,-25.9,-23.76,-19.7,-26.6,-24.19,-17.8],[1660806785291,-26,-23.84,-19.9,-26.6,-24.19,-18.1],[1660810385291,-26,-23.5,-19.7,-26.7,-23.78,-17.7],[1660813985291,-26,-23.03,-14.9,-26.7,-22.64,-6.2],[1660817585291,-25.6,-12.36,11.4,-26.2,-8.04,23.1],[1660821185291,-26,-23.69,-19.9,-26.7,-23.99,-18.1],[1660824785291,-26,-22.91,-17.7,-26.7,-23.19,-16.9],[1660828385291,-26,-23.8,-20,-26.9,-24.28,-18.2],[1660831985291,-26,-23.7,-19.9,-26.9,-24.13,-18.6],[1660835585291,-26,-22.67,-10.4,-26.8,-21.55,1.8],[1660839185291,-25.3,-10.61,13,-26.1,-7.01,23.8],[1660842785291,-26,-23.7,-20,-26.8,-24.1,-18.3],[1660846385291,-26,-23.66,-20,-26.8,-24.07,-18.5],[1660849985291,-26,-23.42,-20,-26.8,-23.72,-18.3],[1660853585291,-26,-23.16,-20,-26.8,-23.44,-18.3],[1660857185291,-26,-22.71,-11.3,-26.8,-22.31,-1],[1660860785291,-25.9,-12.3,11.4,-26.7,-8.19,23],[1660864385291,-26,-21.52,-5.5,-26.8,-22.07,-11.3],[1660867985291,-26,-23.78,-19.9,-26.6,-24.07,-18.3],[1660871585291,-26,-23.56,-19.8,-26.6,-23.83,-18],[1660875185291,-26,-23.17,-19.9,-26.7,-23.32,-18]]};
// const ppll4 = {
//     "heading": "Báo cáo hành trình B.1",
//     "header": ["c1", ["phu", "c2", "c3", "c4"], "c5"],
//     "numFmts": ["","","","",""],
//     "align": ["center", "", "", "", ""],
//     "data": [
//         [2, 4, 2, 4, 5],
//         [4, 2, 1, 2, 9],
//         [4, 2, 1, 2, 9],
//         ["Tổng", "undefined", "undefined", 22, "undefined"],
//     ],
//     "from": "2022-08-18T11:13:05.291Z",
//     "to": "2022-08-18T21:13:05.291Z",
//     "pairTable": [
//         {"key": "Đơn vị kinh doanh vận tải", "value": "ABCDE"},
//         {"key": "Biển số xe", "value": "21344"},
//     ],
// };
// const rawppll3 = {};
// const ppll3 = test(rawppll3);
// console.log(ppll3);
// exportReportAsXlsx(ppll5, "chuan", true);