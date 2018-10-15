
//Tạo menu gọi lệnh///////////////////////////////////////
function onOpen() {
  SpreadsheetApp.flush();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name: "Bước 1 - Nhập thông tin đầu vào", functionName: "forcusSheet"},
                     {name: "Bước 2 - Tạo sheet lịch tuần mới", functionName: "taoSheet"},
                     {name: "Bước 3 - Cập nhật ngày", functionName: "capnhatNgaytuanmoi"},
                     {name: "Bước 4 - Xóa tên KTV", functionName: "xoaKTV"}
                    ];
  var d = new Date();
  var n = d.getDay();
  if (n == 4){
    ss.addMenu("SXBG - Tạo lịch ghi hình", menuEntries); 
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bảng điều khiển").showSheet();
  } else {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bảng điều khiển").hideSheet(); 
  }
};

function taoSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNguon = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bảng điều khiển").getRange(5, 2).getValue();
  var sheetDich = "-->>" + SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bảng điều khiển").getRange(6, 2).getValue();
  dupName(sheetNguon,sheetDich);
  ss.getSheetByName(sheetDich).setTabColor(makeColorHex());
  dichuyenSheet(sheetDich);
};

function capnhatNgaytuanmoi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNguon = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bảng điều khiển").getRange(5, 2).getValue();
  var sheetDich = "-->>" + SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bảng điều khiển").getRange(6, 2).getValue();
  setATvSh(sheetDich);
  capnhatNgay(sheetNguon,sheetDich);  
};

function dupName(sheetNguon,sheetDich) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var sheet = ss.getActiveSheet();
  var sheet = ss.getSheetByName(sheetNguon);
  //var name = Browser.inputBox('Enter new sheet name');
  //sheetDich = "-->>" + sheetDich;
  ss.insertSheet(sheetDich, {template: sheet});
}

function forcusSheet() {
  var nhapTT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bảng điều khiển");
  nhapTT.getRange("A2:B6").activate();
}

function setATvSh(sheetActive) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName(sheetActive));
}
