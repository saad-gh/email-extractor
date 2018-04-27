var CONFIG = {
  settingsSheet:{
    NAME:"Settings CSVE",
    HEADERS:[
      [
        "Entity","From","Subject","File Name","Is Active","Date of receiving",
        "Donot extract again","Ignore in recursive mode","Append data",
        "Add formulas","Columns","Formulas","Filter rows","Having value","in cell",
        "Remove rows on trigger","Sort before deleting?","Date column",
        "Number of days worth of data to be deleted","Email for status notification",
        "Daily trigger","At hour","Check again","After","Up to","Delete rows weekly","Delete rows daily"
      ]
    ],
  },
  VARS:{
    REX_SHEET_CSVE:/SHEET_CSVE/g
  }
}

var FORMULAS = {
  ADD:{ 
    NAME : "ADD",
    FUNC : function(arr,colsInput){
      var sum = 0;
      for(var i = 0;i < colsInput.length; i++)      
          sum += parseFloat(arr[colsInput[i]]);
      return sum;
    }
  },
  CONCATENATE : {
    NAME : "CONCATENATE",
    FUNC : function(arr,colsInput){
      var s = "";
      for(var i = 0;i < colsInput.length; i++)
        s += arr[colsInput[i]];
      return s;
    }
  },
  MULTIPLY : {
    NAME : "MULTIPLY",
    FUNC : function(arr,colsInput){
      var p = 1;
      for(var i = 0;i < colsInput.length; i++)
        p *= parseFloat(arr[colsInput[i]]);
      return p; 
    }
  }
}

var MODES = {
  running : {
    testing : 0,
    ui : 1,
    dailyTrigger : 2,
    recursive : 3
  },
  recTrigger:false,
  RowRemovalTrigger:{
    freq : {
      daily : 1,
      weekly: 2
    },
    mode : {
      unset : 0,
      set : 1
    }
  }
}

var DEFAULT_SETTINGS = {
  TYPE:{
    SHEETS : [
      { key : "isAlreadyUpdated", value : true },
      { key : "ignoreForRecTrg", value : true },
      { key : "exctractMode", value : "latest" }
    ]
  }
}

var MSGS = {
  LOG : {
    errors : 0,
    updated : 1
  },
  EMAIL : {
    errors : 0,
    updated : 1
  }
}

var REPORT_TYPES = {
  shtsUpdated : 1,
  errors : 2
}

var EXCEPTIONS = {
  LatestFileNotReceivedExp:function(){
    return {
      message:"Latest file not received",
      from:shtProps.from,
      subject:shtProps.subject,
      fileName:shtProps.fileName,      
      report:ss.getName(),
      sheetName:shtProps.name
    }
  },
  AttchExp : function(length){
    return {
      message : length + " files attached",
      from:shtProps.from,
      subject:shtProps.subject,
      fileName:shtProps.fileName, 
      report:ss.getName(),
      sheetName:shtProps.name
    }
  },
  InvalidFileExp : function(fn){
    return {
      message : fn + " is not in a recognized format",
      from:shtProps.from,
      subject:shtProps.subject,
      fileName:shtProps.fileName, 
      report:ss.getName(),
      sheetName:shtProps.name
    }
  },
  UnknownDelimiterExp : function(){
    return {
      message : "Unkown delimiter",
      from:shtProps.from,
      subject:shtProps.subject,
      fileName:shtProps.fileName, 
      report:ss.getName(),
      sheetName:shtProps.name
    }
  },
  LastStartedUndefinedExp : function(){
    return {
      message : "Property : \"lastStarted\" is undefined",
      from:shtProps.from,
      subject:shtProps.subject,
      fileName:shtProps.fileName, 
      report:ss.getName(),
      sheetName:shtProps.name
    }
  },
  InvalidQHSNameExp : function(name){
    if(name == "")
      name = "[Empty string]";
    return {
      message : name + "is an invalid Quesry Host Sheet name"
    }
  }
}

// Object variables
var ss;
var allShtsProps;
var sheetsForOp;
var shtProps;
var mode;
var msg;
var attch;
var rawData;
var processedData;
var header;
var shtOp;
var shtNames;
var ps;
var dt;
var arrErrors = [];
var arrShtsUpdated = [];
var reportErrors;
var reportUpdatedShts;
var rngProcessedData;

// Extensions
String.prototype.normalizeForId = function(){
  return this.split(" ").join("").replace(/'/g,"").replace(/\(/g,"").replace(/\)/g,"");
}

String.prototype.toDateObject = function(format){
  var sDt,dt;
  if(format == "yyyy-m-d"){
    sDt = this.split("-").map(function(item){ return parseInt(item); });
    dt = new Date(sDt[0],sDt[1] - 1,sDt[2]);
  } else if(format == "mm-dd-yy"){
    sDt = this.split("/").map(function(item){ return parseInt(item); });
    dt = new Date(sDt[2],sDt[0] - 1,sDt[1]); 
  }
  return dt;
}

Date.prototype.toFormattedString = function(){
 return this.getYear() + "-" + (this.getMonth() + 1) + "-" + (this.getDay() + 1); 
}

// Object functions
function setSpreadsheet(){
  ss = ss || SpreadsheetApp.getActive();
  return ss;
}

function setPropertyStore(){ // Start here
  this.ps = this.ps || PropertiesService.getDocumentProperties(); 
  return this.ps;
}

function setDefaultProps(){
  var sn = ss.getSheets().map(function(sht){ return sht.getName(); });                 
  return ps
  .setProperty("SheetNames_CSVE",
               JSON.stringify(sn)).getProperties();              
}

function setShtPropsObj(){
  setPropertyStore();
  allShtsProps = allShtsProps || (function(){
    var props = ps.getProperties();
    var shtId, shtProps,formulaProps,columns,values,filterProps,keys,values;
    var arr = [];
    var createColumnValueObjs = function(values,forFormulas){
      return function(column,index){
        var obj = {};
        obj["column"] = column;
        obj["value"] = values[index];
        if(forFormulas){
          if(values[index].search("=") != 0)
            obj["custom"] = true;
          else
            obj["custom"] = false;
        }
        return obj;
      }
    }
    
    for(var key in props){
      if(key.match(CONFIG.VARS.REX_SHEET_CSVE)){
        shtProps = JSON.parse(props[key]);
        shtProps["PropertyStoreKey"] = key;
        
        shtProps["formulaProps"] = {};
        formulaProps = shtProps["formulaProps"];
        if(shtProps["appendFormula"] === "true" || shtProps["appendFormula"]){
          formulaProps["isSet"] = true;
          formulaProps["props"] = [];
          columns = shtProps["formulaColumns"].split(",");
          values = shtProps["formulas"].split("\n").map(function(formula){
            if(formula.search(/^\=/) == 0)
              return formula; 
            else
              return formula.split(' ').join('').toUpperCase();
          });        
          formulaProps["props"] = columns.map(createColumnValueObjs(values,true));
        } else {
          formulaProps["isSet"] = false;
        }
        
        shtProps["filterProps"] = {};
        filterProps = shtProps["filterProps"];
        if(shtProps["filter"] === "true" || shtProps["filter"]){
          filterProps["isSet"] = true;
          filterProps["props"] = [];
          columns = shtProps["filterColumns"].split(",");
          values = shtProps["filterValues"].split("\n");
          filterProps["props"] = columns.map(createColumnValueObjs(values,false));
        } else {
          filterProps["isSet"] = false;
        }
        
        shtProps["lastUpdated"] = shtProps["lastUpdated"] || "1970-1-1";
        
        arr.push(shtProps);  
      }
    }
    return arr;
  })();
  return allShtsProps;
}

function daysDifference(d0, d1) {
  var diff = new Date(+d1).setHours(12) - new Date(+d0).setHours(12);
  return Math.round(diff/8.64e7);
}

function setSheetsForOp(){
  var rawSheetsForOp;
  setShtPropsObj();
  rawSheetsForOp = allShtsProps.map(function(shtProps){
    if(shtProps.active !== true) {
      return undefined;
    } else if(shtProps.isAlreadyUpdated && (daysDifference(shtProps.lastUpdated.toDateObject("yyyy-m-d"),Date.now()) == 0)){ // Create prototype compareDate date for Date object 
      return undefined;
    } else if(shtProps.ignoreForRecTrigger && mode == MODES.running.recursive) {  // $Set modes variable
      return undefined;
    } else {    
      return shtProps;
    }
  });
  sheetsForOp = rawSheetsForOp.filter(function(item){ return item != undefined; });
  return sheetsForOp;
}

function getMsg(threads){ 
  return threads[0].getMessages()[threads[0].getMessageCount() - 1];
}

function setAttachment(){
  attch = msg.getAttachments().filter(function(attch){
    return attch.getName().search(shtProps.fileName) !== -1;
  });
  if(attch.length == 0 || attch.length > 1){
    throw EXCEPTIONS.AttchExp(attch.length); // create exp object 
  } else {
    attch = attch[0]; 
  }
}

function parseCsv(sData){
  rawData = Utilities.parseCsv(sData);
  if(rawData[0].length == 1 && rawData[0][0].search('\t') != -1)
    rawData = Utilities.parseCsv(sData,'\t');
  else if(rawData[0].length == 1)
    throw EXCEPTIONS.UnknownDelimiterExp();
}

function extractRawData(){
  var blobs, csv,zip;
  csv = "csv";
  zip = "zip";
  var ext = attch.getName().split('.').pop();
  var sData;
  if(ext == csv){
    attch.setContentType("text/csv");
    sData = attch.getDataAsString();
    if (sData.search('\x00') != -1 || sData.search('\ufffd') != -1) {
      sData = sData.replace('\x00', "", 'g').replace('\ufffd', "", 'g');      
      parseCsv(sData);
    } else {
      parseCsv(sData);
    }
  } else if(ext == zip) {
    attch.setContentType("application/zip");
    blobs = Utilities.unzip(attch);
    if(blobs.length > 1 || blobs.length == 0){
      throw exp.AttchExp(blobs.length);
    } else {
      parseCsv(blobs[0].getDataAsString());
    }
  } else {
    throw EXCEPTIONS.InvalidFileExp(attch.getName()); 
  }
}

var filterFunc = function(filterProps,len){
  return function(arr,ix){
    var v1 = filterProps,v2,v3;    
    for(var iy = 0; iy < len;iy++){
      v2 = arr[parseInt(filterProps[iy].column) - 1];
      if(v2 == undefined)
        continue;
      else
        v2 = v2.toString();
      v3 = filterProps[iy].value.toString();
      if(v2 == v3)
        return 0;
    }
    return 1;
  }
}

function filterData(filterProps,data){ 
  processedData = data.filter(filterFunc(filterProps,filterProps.length));
}

//if invert set to 0 than func returns filter columns same as formula columns
var funcFilterProps = function(formulaCols,invert){
  return function(obj){
    return invert ?  formulaCols.indexOf(obj.column) !== -1 : formulaCols.indexOf(obj.column) === -1;
  }
}

var sortFunction = function (dateColIndex){
  
  return function (a, b) {
    return a[dateColIndex].toDateObject("mm/dd/yyyy") - b[dateColIndex].toDateObject("mm/dd/yyyy");
  }
}

function getFormulas(custom){
  return shtProps.formulaProps.props.filter((function(custom){
    return function(obj){    
      return obj.custom == custom ? true : false;
    }
  })(custom));
}

function applyFormula(arr,formula){
  var res;
  var colsInput = formula.match(/\d{1,3}/g);
  colsInput = colsInput.map(function(col){ return parseInt(col) - 1; });
  if(formula.search(FORMULAS.ADD.NAME) == 0){
    res = FORMULAS.ADD.FUNC(arr,colsInput);
  } else if(formula.search(FORMULAS.CONCATENATE.NAME) == 0){
    res = FORMULAS.CONCATENATE.FUNC(arr,colsInput);
  } else if(formula.search(FORMULAS.MULTIPLY.NAME) == 0){
    res = FORMULAS.MULTIPLY.FUNC(arr,colsInput);
  }
  
  return res;
}

function generateFormulaCol(data,formula,column){
  return data.map((function(formula,column,applyFormula){ 
    return function(arr){
    arr.splice(column,0,applyFormula(arr,formula));
    return arr;
    }
  })(formula,column,applyFormula));  
}

function formulaCase(lastRow,data){
  var filterProps, formulaCols_custom; 
  var formulaArray = getFormulas(true);

  if(formulaArray.length > 0){
    formulaCols_custom = formulaArray.map(function(obj){
      return obj.column;
    });
    for(var iFormulas = 0;iFormulas < formulaArray.length;iFormulas++){                
      var newCol = formulaArray[iFormulas].column;
      var formula = formulaArray[iFormulas].value;   
      data = generateFormulaCol(data,formula,parseInt(newCol) - 1);   
    }
    
    if(shtProps.filterProps.isSet){
      filterProps = shtProps.filterProps.props.filter(funcFilterProps(formulaCols_custom,0)); 
      if(filterProps.length > 0)
        filterData(filterProps,data);
    }
  }
  
  formulaArray = getFormulas(false);
  
  if(formulaArray.length > 0){
    data.sort(sortFunction(1));
    for(var iRow=0;iRow<data.length;iRow++){                  
      for(var iFormulaObj = 0;iFormulaObj<formulaArray.length;iFormulaObj++)
        data[iRow][parseInt(formulaArray[iFormulaObj].column) - 1] = formulaArray[iFormulaObj].value.replace(/#/g,iRow + lastRow + 1);                        
    }
  }
}

function putData(lastRow){  
  shtOp = shtOp || ss.getSheetByName(shtProps.name);
  if((!shtProps.append || lastRow === 1)){
    if(!shtProps.append)
      shtOp.clearContents();
    shtOp
    .getRange(1,1,1,header[0].length)
    .setValues(header);
    rngProcessedData = shtOp
    .getRange(2,1,processedData.length,processedData[0].length)
    .setValues(processedData);    
  }else if(shtProps.append){
    rngProcessedData = shtOp
    .getRange(lastRow + 1,1,processedData.length,processedData[0].length)
    .setValues(processedData);
  }
}

function filterCase(formula, lastRow){
  var filterProps,formulaCols;
  if(formula){
    formulaCols = shtProps.formulaProps.props.map(function(obj){
      return obj.column;
    });
    
    filterProps = shtProps.filterProps.props.filter(funcFilterProps(formulaCols,0));
    
    if(filterProps.length > 0)
      filterData(filterProps,processedData);
    
    formulaCase(lastRow,processedData);
    putData(lastRow);
    processedData = undefined;
    filterProps = shtProps.filterProps.props.filter(funcFilterProps(formulaCols,1));
    
    if(filterProps.length > 0){
      processedData = rngProcessedData.getValues();
      filterData(filterProps,processedData);
      rngProcessedData.clearContent();
    } 
  } else {
    filterData(shtProps.filterProps.props,processedData);
  }
}

function selectCase(){
  var lastRow;
  shtOp = ss.getSheetByName(shtProps.name);
  lastRow = shtOp.getLastRow();
  if(!lastRow || !shtProps.append)
    lastRow = 1
  
  header = rawData.splice(0, 1);
  processedData = rawData.slice(0);
  
  if(shtProps.filterProps.isSet){
    filterCase(shtProps.formulaProps.isSet,lastRow);      
  }else if(shtProps.formulaProps.isSet){  
    formulaCase(lastRow,processedData);      
  }
  
  if(processedData)
    putData(lastRow);
  
  shtProps.lastUpdated = dt.toFormattedString();
}

function startExtraction(){  
  setAttachment();
  extractRawData();
  selectCase();
}

function findData(){
  var query = "from: " + shtProps.from + ' subject:"' + shtProps.subject + '"';
  var threads = GmailApp.search(query);
  msg = getMsg(threads);  
  var msgDate;
  
  if(shtProps.extractMode == "latest"){
    msgDate = msg.getDate();
    if(daysDifference(msgDate,Date.now())){
      throw EXCEPTIONS.LatestFileNotReceivedExp();
    }
  }
}

function createReport(type){
  var arr;
  var keysShtsUpdated = ["name","from","fileName","subject","spreadsheet"];
  if(type == REPORT_TYPES.errors)
    arr = arrErrors;
  else if(type == REPORT_TYPES.shtsUpdated){
    arr = arrShtsUpdated;
    arr.push({"spreadsheet" : ss.getName()});
  }
  
  var data;
  data = "\n"
  
  for(var i = 0; i < arr.length; i++){
    for(var key in arr[i]){
      if(type == REPORT_TYPES.errors)
        data += key + " : " + arr[i][key] + "\t"; 
      else if(type == REPORT_TYPES.shtsUpdated && keysShtsUpdated.indexOf(key))
        data += key + " : " + arr[i][key] + "\t"; 
    }
    data += "\n";
  }
  
  return data;
}

function emailNotification(type){
  var gProps = JSON.parse(ps.getProperty("GLOBAL_CSVE"));
  var email = gProps["customEmailNotification"];
  var subject,body;
  
  if(type == REPORT_TYPES.errors){
    subject = "Error report CSVE: " + ss.getName();
    body = reportErrors;
  } else if(type == REPORT_TYPES.shtsUpdated){
    subject = "Updated sheets report CSVE: " + ss.getName();
    body = reportUpdatedShts;
  }
  
  if(email != "")
  MailApp.sendEmail({
    to:email,
    subject:subject,
    body:body
  });
}

function handleErrorsAndUpdates(){
  var msgForUi = {};
   if(arrErrors.length > 0){
     reportErrors = createReport(REPORT_TYPES.errors);
    if(mode == MODES.running.testing){
      Logger.log(reportErrors);
    } else if(mode == MODES.running.ui){
      throw new Error(reportErrors);      
    } else if(mode == MODES.running.recursive || mode == MODES.running.dailyTrigger){
      emailNotification(REPORT_TYPES.errors);
    } 
  }
  if(arrShtsUpdated.length > 0){
    reportUpdatedShts = createReport(REPORT_TYPES.shtsUpdated);
    if(mode == MODES.running.testing){
      Logger.log(reportUpdatedShts);
    }else if(mode == MODES.running.recursive || mode == MODES.running.dailyTrigger){
      emailNotification(REPORT_TYPES.shtsUpdated);
    }
  }
}

function updateSheets(m){
  mode = m;
  setSheetsForOp();
  dt = new Date();
  sheetsForOp.map((function(dt,that){ 
    var ps2,shtProps2; 
    return function(shtProps){    
      that.shtProps = shtProps; 
      that.shtProps["lastStarted"] = dt.toFormattedString();
      
      if(that.mode == that.MODES.running.ui){
        ps2 = ps2 || PropertiesService.getDocumentProperties();
        shtProps2 = JSON.parse(ps2.getProperty(that.shtProps["PropertyStoreKey"]));
        if(that.shtProps.isAlreadyUpdated)
          shtProps2["lastStarted"] = that.shtProps["lastStarted"];
      }
      
      try{
        that.findData();        
        that.startExtraction();  
        that.arrShtsUpdated.push(that.shtProps);
      } catch(e){
        that.arrErrors.push(e);
      }
      
      if(that.mode == that.MODES.running.ui){
        if(that.shtProps.isAlreadyUpdated)
          shtProps2["lastUpdated"] = that.shtProps["lastUpdated"];
        ps2.setProperty(shtProps2["PropertyStoreKey"],JSON.stringify(shtProps2));
      } else {
        that.ps.setProperty(that.shtProps["PropertyStoreKey"],JSON.stringify(that.shtProps));
      }
    }
  })(dt,this));
  
  handleErrorsAndUpdates();
}

function getProps(){ 
  return {
    allSheets:allSheets.map(function(sht){ 
      return {
        name:sht.getName(),
        value:sht.getName().normalizeForId()
      }; 
    }),
    shtProps:k
  }
}

// Weekly trigger module
function delTrigger(id){
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    if(id == allTriggers[i].getUniqueId())
    ScriptApp.deleteTrigger(allTriggers[i]);
  }  
}

function columnToLetter(column){
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
  
(function(host){
  
  function QueryControl(propsWeeklyTrigger){
    if(!(this instanceof QueryControl)){
      return new QueryControl(propsWeeklyTrigger); 
    }
    
    if (propsWeeklyTrigger != undefined){       
      this.propsWeeklyTrigger = propsWeeklyTrigger;
      this.sheetsMeta = this.propsWeeklyTrigger.sheetsMeta;
      this.ss = this.setSpreadsheet();
      if(propsWeeklyTrigger.queryHostSheet == "")
        throw new host.Error("[Empty string] is not a valid Query Host Sheet name");
      this.queryHostSheet = this.ss.getSheetByName(propsWeeklyTrigger.queryHostSheet) || 
        this.createQueryHostSheet(propsWeeklyTrigger.queryHostSheet);
      this.dataRangeQHS = this.queryHostSheet.getDataRange();
      this.dataQHS = this.dataRangeQHS.getValues();
      this.lastRowQHS = this.dataQHS[0][0] === "" ? 0 : this.dataQHS.length;
      this.sheets = this.dataQHS.map(function(arr){      
        if(arr[0] != "")
          return arr[0];
        else
          return arr[1];      
      });
      
    }
    
    this.formula_query = '=query(\'[sheet-name]\'![date-col]:[date-col],' +
      '"Select count([date-col]) where ' + 
        '[date-col] >= date \'"&[oldDt-ub][last-row-QHS]&"\' and ' + 
          '[date-col] <= date \'"&[oldDt-lb][last-row-QHS]&"\'")';
    
    this.formula_getOldestDate_upperBound = '=text(year([min-date-col][last-row-QHS]),"0000") & "-" & '+
      ' text(month([min-date-col][last-row-QHS]),"00") & "-" & ' + 
        'text(day([min-date-col][last-row-QHS]),"00")';
    
    this.formula_getOldestDate_lowerBound = '=text(year([min-date-col][last-row-QHS] + [days-worth]),"0000") & "-" & '+
      ' text(month([min-date-col][last-row-QHS] + [days-worth]),"00") & "-" & ' + 
        'text(day([min-date-col][last-row-QHS] + [days-worth]),"00")';
    
    this.formula_minDate = '=MIN(\'[sheet-name]\'![date-col]:[date-col])';   
    
    return this;    
  }
  
  QueryControl.prototype.host = host;
  QueryControl.prototype.setSpreadsheet = host.setSpreadsheet;
  QueryControl.prototype.delTrigger = host.delTrigger;
  QueryControl.prototype.setPropertyStore = host.setPropertyStore;
  QueryControl.prototype.columnToLetter = host.columnToLetter;
  QueryControl.prototype.shtNames = host.setSheetNames();
  QueryControl.prototype.createQueryHostSheet = host.createSheet;
  
  QueryControl.prototype.controlQueries = function(action){
    for(var i = 0;i < this.sheetsMeta.length;i++){
      if(!this.sheetsMeta[i].remove ){
        if(action == "insert" && this.sheets.indexOf(this.sheetsMeta[i].name) === -1){
          this.buildAndExecuteQuery(this.sheetsMeta[i],action);        
          this.lastRowQHS += 2;
        } else if (action == "update" && this.sheets.indexOf(this.sheetsMeta[i].name) !== -1){
          this.buildAndExecuteQuery(this.sheetsMeta[i],action); 
        }        
      }
    }
    return this;    
  }
  
  QueryControl.prototype.removeQueries = function () {
    for(var i = 0;i < this.sheetsMeta.length;i++){
      if(this.sheetsMeta[i].remove)
        this.removeQuery(this.sheetsMeta[i]);
    }
    return this;
  }
  
  QueryControl.prototype.removeQuery = function(sheetMeta){
    var row = this.sheets.indexOf(sheetMeta.name);
    if(row !== -1){
      this.queryHostSheet.deleteRows(row + 1, 2);
      this.sheets.splice(row,2);
    }
  }
  
  QueryControl.prototype.updateProperties = function(sheetMeta){
    this.setPropertyStore();
    var old_JSON_weeklyTrigger_Meta_CSVE = JSON.parse(this.ps.getProperty("weeklyTrigger_Meta_CSVE"));
    for(var i=0;this.sheetsMeta.length;i++){
      if(this.sheetsMeta[i].name == sheetMeta.name){
        this.sheetsMeta[i] = sheetMeta;  
        old_JSON_weeklyTrigger_Meta_CSVE.propsWeeklyTrigger.sheetsMeta = this.sheetsMeta;      
        this.ps.setProperty("weeklyTrigger_Meta_CSVE",JSON.stringify(old_JSON_weeklyTrigger_Meta_CSVE));
        return;
      }
    }
  }
  QueryControl.prototype.buildAndExecuteQuery = function (sheetMeta, action){
    
    var queryRow = [],f_q,f_dt_u,f_dt_l,f_md,dt_letter = this.columnToLetter(sheetMeta.dateCol);
    //      var oldDt_ub = this.columnToLetter(parseInt(sheetMeta.dateCol) + 1),oldDt_lb = this.columnToLetter(parseInt(sheetMeta.dateCol) + 2);
    var oldDt_ub = "C",oldDt_lb = "D", minDt = "E";
    var actionRow;
    
    if(action == "insert")
      actionRow = this.lastRowQHS % 2 ? this.lastRowQHS + 2 : this.lastRowQHS + 1;
    else if(action == "update"){
      this.updateProperties(sheetMeta); 
      actionRow = this.sheets.indexOf(sheetMeta.name) + 1;        
    }    
    
    queryRow[0] = [];
    
    f_q = this.formula_query;
    f_q = f_q
    .split("[sheet-name]").join(sheetMeta.name)
    .split("[date-col]").join(dt_letter)
    .split("[last-row-QHS]").join(actionRow)
    .split("[oldDt-ub]").join(oldDt_ub)
    .split("[oldDt-lb]").join(oldDt_lb);
    
    f_dt_u = this.formula_getOldestDate_upperBound;
    f_dt_u = f_dt_u
    .split("[min-date-col]").join(minDt)
    .split("[last-row-QHS]").join(actionRow);
    
    f_dt_l = this.formula_getOldestDate_lowerBound;
    f_dt_l = f_dt_l
    .split("[min-date-col]").join(minDt)
    .split("[last-row-QHS]").join(actionRow)
    .split("[days-worth]").join(sheetMeta.daysWorth - 1);
    
    f_md = this.formula_minDate;
    f_md = f_md
    .split("[sheet-name]").join(sheetMeta.name)
    .split("[date-col]").join(dt_letter);
    
    queryRow[0].push(sheetMeta.name);
    queryRow[0].push(f_q);
    queryRow[0].push(f_dt_u);
    queryRow[0].push(f_dt_l);
    queryRow[0].push(f_md);
    
    this.queryHostSheet    
    .getRange(actionRow,1,1,queryRow[0].length)
    .setValues(queryRow); 
    
  }
  
  QueryControl.prototype.deleteRows = function (){
    var sht,meta;    
    for(var i = 0; i < this.sheets.length - 1;i++){
      if(this.shtNames.indexOf(this.sheets[i]) !== -1 && !isNaN(this.sheets[i + 1]) && this.sheets[i+1] != ""){
        meta = this.sheetsMeta.filter((function (sht){ return function(meta){ return meta.name == sht ; } }(this.sheets[i])))[0];
        sht = ss.getSheetByName(this.sheets[i]);
        if(meta.sortDel)
          sht.getRange(2, 1, sht.getLastRow() - 1, sht.getLastColumn()).sort(parseInt(meta.dateCol));         
        
        sht.deleteRows(2,this.sheets[i + 1]);        
      }
    }
    return this;
  }
  
  QueryControl.prototype.setWeeklyTrigger = function (propsWeeklyTrigger){    
    var new_JSON_weeklyTrigger_Meta_CSVE = {};
    var old_JSON_weeklyTrigger_Meta_CSVE = JSON.parse(this.ps.getProperty("weeklyTrigger_Meta_CSVE"));
    
    if(old_JSON_weeklyTrigger_Meta_CSVE != undefined)
      this.delTrigger(old_JSON_weeklyTrigger_Meta_CSVE["weeklyTrigger_Id"]);      
    
    
    new_JSON_weeklyTrigger_Meta_CSVE["weeklyTrigger_Id"] = ScriptApp.newTrigger("weeklyTrigger")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .create()
    .getUniqueId();
    new_JSON_weeklyTrigger_Meta_CSVE["propsWeeklyTrigger"] = propsWeeklyTrigger || this.propsWeeklyTrigger;
    //    new_JSON_weeklyTrigger_Meta_CSVE.propsWeeklyTrigger.isSet = true;
    this.ps.setProperty("weeklyTrigger_Meta_CSVE",JSON.stringify(new_JSON_weeklyTrigger_Meta_CSVE));
    
  }
  
  QueryControl.prototype.setDailyTriggerDR = function (propsWeeklyTrigger){
    this.setPropertyStore();
    var new_JSON_weeklyTrigger_Meta_CSVE = {};
    var old_JSON_weeklyTrigger_Meta_CSVE = JSON.parse(this.ps.getProperty("weeklyTrigger_Meta_CSVE"));
    
    if(old_JSON_weeklyTrigger_Meta_CSVE != undefined)
      this.delTrigger(old_JSON_weeklyTrigger_Meta_CSVE["dailyTrigger_Id"]);      
    
    
    new_JSON_weeklyTrigger_Meta_CSVE["dailyTrigger_Id"] = ScriptApp.newTrigger("weeklyTrigger")
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create()
    .getUniqueId();
    new_JSON_weeklyTrigger_Meta_CSVE["propsWeeklyTrigger"] = propsWeeklyTrigger || this.propsWeeklyTrigger;
    //    new_JSON_weeklyTrigger_Meta_CSVE.propsWeeklyTrigger.isSet = true;
    this.ps.setProperty("weeklyTrigger_Meta_CSVE",JSON.stringify(new_JSON_weeklyTrigger_Meta_CSVE));
    
  }
  
  QueryControl.prototype.removeWeeklyTrigger = function (){
    this.setPropertyStore();
    var old_JSON_weeklyTrigger_Meta_CSVE = JSON.parse(this.ps.getProperty("weeklyTrigger_Meta_CSVE"));
    if(old_JSON_weeklyTrigger_Meta_CSVE != undefined){
      this.delTrigger(old_JSON_weeklyTrigger_Meta_CSVE.weeklyTrigger_Id);
      old_JSON_weeklyTrigger_Meta_CSVE.propsWeeklyTrigger.isSetWeekly = false;
      this.ps.setProperty("weeklyTrigger_Meta_CSVE",JSON.stringify(old_JSON_weeklyTrigger_Meta_CSVE));
    }
    
  }
  
  QueryControl.prototype.removeDailyTrigger = function (){
    this.setPropertyStore();
    var old_JSON_weeklyTrigger_Meta_CSVE = JSON.parse(this.ps.getProperty("weeklyTrigger_Meta_CSVE"));
    if(old_JSON_weeklyTrigger_Meta_CSVE != undefined){
      this.delTrigger(old_JSON_weeklyTrigger_Meta_CSVE.dailyTrigger_Id);
      old_JSON_weeklyTrigger_Meta_CSVE.propsWeeklyTrigger.isSetDaily = false;
      this.ps.setProperty("weeklyTrigger_Meta_CSVE",JSON.stringify(old_JSON_weeklyTrigger_Meta_CSVE));
    }
    
  }
  
  host.removeWeeklyTrigger = function () {
    QueryControl()
    .removeWeeklyTrigger();
  }
  host.removeDailyTrigger = function () {
    QueryControl()
    .removeDailyTrigger();
  }
  host.removeRowsWeekly = function(){
    this.setPropertyStore();
    QueryControl(JSON.parse(ps.getProperty("weeklyTrigger_Meta_CSVE")).propsWeeklyTrigger)
    .deleteRows()
    .controlQueries("update");
  }
  
  host.QueryControl = QueryControl;
  
  
}(this));

function getWeeklyTriggerProperties(){
  setPropertyStore();
  var propsWeeklyTrigger = ps.getProperty("weeklyTrigger_Meta_CSVE");
  if(propsWeeklyTrigger)
    return JSON.parse(propsWeeklyTrigger).propsWeeklyTrigger;
  else
    return {};
}

function updateQuery(sheetMeta){
  setPropertyStore();
  QueryControl(JSON.parse(ps.getProperty("weeklyTrigger_Meta_CSVE")).propsWeeklyTrigger)
  .buildAndExecuteQuery(sheetMeta,"update");
}

function updateRowRemovalTrigger(props,mode,freq){
  var qc = QueryControl(props);
  qc.setPropertyStore();
  
  if(mode == MODES.RowRemovalTrigger.mode.set){
    qc.controlQueries("insert")
    .removeQueries();  
    if(freq == MODES.RowRemovalTrigger.freq.weekly)
      qc.setWeeklyTrigger();
    else if (MODES.RowRemovalTrigger.freq.daily)
      qc.setDailyTriggerDR();
  }
  else if(mode == MODES.RowRemovalTrigger.mode.unset){
    if(freq == MODES.RowRemovalTrigger.freq.weekly)
      qc.removeWeeklyTrigger();
    else if(freq == MODES.RowRemovalTrigger.freq.daily)
      qc.removeDailyTrigger();
  }
}

// Triggers

function weeklyTrigger(){
  removeRowsWeekly();
}

function setShtsForOpInCache(){
  var cache = CacheService.getDocumentCache();
  var props = JSON.parse(cache.get("GLOBAL_CSVE"));
  var dt,min,remainingTime,expireTime;
  if(cache.get("shtsForOp_CSVE")){
   sheetsForOp = JSON.parse(cache.get("shtsForOp_CSVE"));
  } else {
    dt = new Date();
    min = dt.getMinutes();
    remainingTime = 60 - min + parseInt(props["checkAgainLimit_CSVE"]);
    expireTime = remainingTime * 60 + 1800;
    cache.put("shtsForOp_CSVE",JSON.stringify(setSheetsForOp()),expireTime);
  }
}

function allShtsUpdated(){
  setPropertyStore();
  setShtsForOpInCache();
  
  var sLastUpdated;
  var sLastStarted;
  
  for(var i=0;i<sheetsForOp.length;i++){  
    shtProps = sheetsForOp[i];
    
    if(shtProps.lastStarted == undefined)
      throw EXCEPTIONS.LastStartedUndefinedExp();
    
    sLastStarted = shtProps.lastStarted;
    sLastUpdated = shtProps.lastUpdated || "1970-1-1";    
  
    if(sLastStarted != sLastUpdated)
      return false;
  }
  
  return true;
}

function dailyTrigger(){
  setPropertyStore();
  var cache;
  var props = JSON.parse(ps.getProperties()["GLOBAL_CSVE"]);
  var d = new Date();
  var min = d.getMinutes();
  var remainingTime = 60 - min + parseInt(props["checkAgainLimit_CSVE"]);
  var expireTime = remainingTime * 60 + 1800;
  cache = CacheService.getDocumentCache();
  cache.put("GLOBAL_CSVE", ps.getProperties()["GLOBAL_CSVE"], expireTime)
  
  updateSheets(MODES.running.dailyTrigger);
  
  if((props["checkAgain_CSVE"] === "true" || props["checkAgain_CSVE"])  && !allShtsUpdated()){
    if(props["checkAgainMins_CSVE"] != ""){      
      cache.put("recRemainingTime_CSVE", remainingTime.toString(), expireTime);
      cache.put("checkAgainTrgId_CSVE", 
                ScriptApp
                .newTrigger("recursiveTrigger")
                .timeBased()
                .everyMinutes(parseInt(props["checkAgainMins_CSVE"]))
                .create()
                .getUniqueId(), 
                expireTime);
    } 
  }
}

function recursiveTrigger(){
  mode = MODES.running.recursive;
  var cache = CacheService.getDocumentCache();
  var props = JSON.parse(cache.get("GLOBAL_CSVE"));  
  var remainingTime = parseInt(cache.get("recRemainingTime_CSVE")) - parseInt(props["checkAgainMins_CSVE"]);
  var expireTime = remainingTime * 60 + 1800;
  var trgId = cache.get("checkAgainTrgId_CSVE");
  cache.put("recRemainingTime_CSVE", remainingTime,expireTime);
  
  if(allShtsUpdated() || remainingTime < 0){
    ScriptApp.getProjectTriggers().map((function(trgId){ 
      return function(trigger){
      if(trigger.getUniqueId() == trgId)
        ScriptApp.deleteTrigger(trigger);
      }
    })(trgId));
  } else {
    updateSheets(MODES.running.recursive);
  }
}

// Front end functions

function getProperties(){  
  return PropertiesService.getDocumentProperties().getProperties(); 
}

function setSheetNames(){
  setSpreadsheet();
  shtNames = shtNames || ss.getSheets().map(function(sht){ return sht.getName(); });
  return shtNames;
}

//Reduce into key value pair
function reduce_sJson(sJson){
  for(var key in sJson){
   sJson[key] = JSON.stringify(sJson[key]); 
  }
  return sJson;
}

function setProperties(sJson){
  setPropertyStore().setProperties(reduce_sJson(sJson));  
}

function setDefaultProps(sJson){
  var currentProps,defaultProps,shtProps;  
  currentProps = setPropertyStore().getProperties();
  defaultProps = sJson["defaultProps_CSVE"];
  for(var sht in sJson){
    if(currentProps[sht]){
      shtProps = JSON.parse(currentProps[sht]);
      for(var prop in defaultProps)
        shtProps[prop] = defaultProps[prop];
      currentProps[sht] = JSON.stringify(shtProps);
    } else
      currentProps[sht] = JSON.stringify(sJson[sht]);
  }
  ps.setProperties(currentProps);
}

function getReports(sJson){
  var getProperties = (function(sJson,reduce_sJson){
    return function(){
      return reduce_sJson(sJson);
    }
  })(sJson,reduce_sJson);
  
  //Hacking ps so that temporary settings from ui can be used
  this.ps = { getProperties : getProperties };
  updateSheets(MODES.running.ui);
}

// Settings sheet creation
function createSheet(sheetName,headers){
  setSpreadsheet();
  return ss.getSheetByName(sheetName) || (function (sheetName,headers){    
    var sht = ss
    .insertSheet()
    .setName(sheetName);
    
    if(headers)
      sht
      .getRange(1,1,headers.length,headers[0].length)
      .setValues(headers);
    
    return sht;
  })(sheetName,headers);
}

function deleteSettingsSheet(){
  setSpreadsheet();
  var sht = ss.getSheetByName(CONFIG.settingsSheet.NAME);
  if(sht != undefined)
    ss.deleteSheet(sht);
}

function onOpen() {
  SpreadsheetApp.getUi().createAddonMenu().addItem("!Extraction settings", "showSettings").addToUi(); 
}

function showSettings() {
  var html = HtmlService.createTemplateFromFile('Settings').evaluate();
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showSidebar(html);
}

// Unit Testing
function testerPrintProperties(){
 setPropertyStore();
  Logger.log(JSON.stringify(ps.getProperties()));
}

function testerSettingsSheet(){
  deleteSettingsSheet();
  createSettingsSheet(); 
}

function testerCSVE(){
  setSheetNames()
  Logger.log(shtNames);
}

function testerProperties(){
//  removeDailyTrigger();
  Logger.log(PropertiesService.getDocumentProperties().getProperties());
}

function testerMainObj(){
 Logger.log(this.constructor.name); 
}

function testerQueryControl(){
var p = setPropertyStore().getProperties().weeklyTrigger_Meta_CSVE;
  ps;
  updateRowRemovalTrigger(JSON.parse(ps.getProperties().weeklyTrigger_Meta_CSVE).propsWeeklyTrigger,1,1);
}

function testerLastRow(){
 setSpreadsheet();
  Logger.log(ss.getSheetByName("Tester props").getLastRow());
}

function testerSetShtsForOp(){
  setSheetsForOp();
  Logger.log(JSON.stringify(sheetsForOp));
}

var testWeeklyTrg = {"sheetsMeta":[{"name":"Sheet1","dateCol":"2","daysWorth":"2","remove":true,"sortDel":true},{"name":"Settings CSVE","dateCol":"2","daysWorth":"1","remove":false,"sortDel":true},{"name":"Tester Props","dateCol":"","daysWorth":"","remove":true,"sortDel":false},{"name":"Query Host Sheet","dateCol":"","daysWorth":"","remove":true,"sortDel":false}],"queryHostSheet":"Query Host Sheet","isSetWeekly":false,"isSetDaily":true}

function testerWeeklyTrg(){
updateRowRemovalTrigger(testWeeklyTrg, 1, 1)
}

function testerUpdateSheets(){
updateSheets(MODES.running.testing);
}

function testerFormulaRegex(){
  var formula, colsInput;
  formula = " add(1 ,1,  123) ";
   
  Logger.log(formula.split(' ').join('').toUpperCase());
}

function testerGetReports(){
   getReports((function(props){ 
    for(var p in props){
      props[p] = JSON.parse(props[p]);
    }
    return props;
  })(PropertiesService.getDocumentProperties().getProperties())); 
}

function testerLastSheetUpdated(){
 Logger.log(allShtsUpdated()); 
}  