function doGet(e) {
  var x = HtmlService.createTemplateFromFile("index");
  var y = x.evaluate();
  var z = y.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return z;
}

function checkLogin(username, password) {  
  var ws=SpreadsheetApp.openById('1tnl2ZsuPkuNxMbOCXIYDgsrUMrt_k7xSA4uYmfORSiY');
  var ss = ws.getSheetByName("Data");
  var getLastRow =  ss.getLastRow();
  var found_record = '';
  for(var i = 2; i <= getLastRow; i++){
   if(ss.getRange(i, 1).getDisplayValue().toUpperCase() == username.toUpperCase() && 
     ss.getRange(i, 2).getDisplayValue().toUpperCase() == password.toUpperCase())
     if(ss.getRange(i,5).getValue()=="Approved"){
        found_record = 'TRUE';
      } else{
        found_record = 'FALSE';
        }   
  }
  if(found_record == ''){
    found_record = 'FALSE'; 
  }  
  return found_record;
  
}

function AddRecord(usernamee, passwordd, email, phone) {
  var ws=SpreadsheetApp.openById('1tnl2ZsuPkuNxMbOCXIYDgsrUMrt_k7xSA4uYmfORSiY');
  var ss = ws.getSheetByName("Data");  
  ss.appendRow([usernamee,passwordd,email,phone]);  
}