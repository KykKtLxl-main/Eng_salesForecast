// BigQueryAPI(V2)設定済
// 231226：動作確認済
// https://www.niandc.co.jp/tech/20220725_2213/

function insertBQ_byBranch(){
  const projectId = 'lixil-workspace';
  const datasetId = 'an1_extEng_salesForecast';
  const tableId = 't51_salesPlanValue_byBranch';
  const columns = [
    {name: 'N_LEVEL_CD',     type: 'string'},
    {name: 'N_LEVEL_NAME',   type: 'string'},
    {name: 'preYear_amount', type: 'numeric'},  //前年実績
    {name: 'salesPlan',      type: 'numeric'},  //計画値
    {name: 'salesPlan_correct',  type: 'numeric'} //計画値補正（手入力修正）
  ];
  sheetName = 'plan_byBranch'

  insertBQ(projectId, datasetId, tableId, columns);
}

function insertBQ_byOffice(){
  const projectId = 'lixil-workspace';
  const datasetId = 'an1_extEng_salesForecast';
  const tableId = 't52_salesPlanValue_byOffice';
  const columns = [
    {name: 'N_LEVEL_CD',     type: 'string'},
    {name: 'N_LEVEL_NAME',   type: 'string'},
    {name: 'OFFICE_CD',     type: 'string'},
    {name: 'OFFICE_NAME',   type: 'string'},
    {name: 'preYear_amount', type: 'numeric'},  //前年実績
    {name: 'salesPlan',      type: 'numeric'},  //計画値
    {name: 'salesPlan_correct',  type: 'numeric'} //計画値補正（手入力修正）
  ];
  sheetName = 'plan_byOffice'

  insertBQ(projectId, datasetId, tableId, columns);
}


function insertBQ(projectId, datasetId, tableId, columns) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("plan_byBranch");

  var table ={
    tableReference: {
      projectId: projectId,
      datasetId: datasetId,
      tableId: tableId
    },
    schema: {
      fields: columns
    }
  };

  try{
    BigQuery.Tables.remove(projectId, datasetId, tableId);
  } catch(e) {}
  table = BigQuery.Tables.insert(table, projectId, datasetId);

  var range = sheet.getDataRange();
  var blob = Utilities.newBlob(convCsv(range)).setContentType('application/octet-stream');
  var job = {
    configuration: {
      load: {
        destinationTable: {
          projectId: projectId,
          datasetId: datasetId,
          tableId: tableId
        },
        skipLeadingRows: 3 // ヘッダ行は無視
      }
    }
  };
  job = BigQuery.Jobs.insert(job, projectId, blob);
}

function convCsv(range) {
  try {
    var data = range.getValues();
    var ret = "";
    if (data.length > 1) {
      var csv = "";
      for (var i = 0; i < data.length; i++) {
        for (var j = 0; j < data[i].length; j++) {
          if (data[i][j].toString().indexOf(",") != -1) {
            data[i][j] = "\"" + data[i][j] + "\"";
          }
        }
        if (i < data.length-1) {
          csv += data[i].join(",") + "\r\n";
        } else {
          csv += data[i];
        }
      }
      ret = csv;
    }
    return ret;
  }
  catch(e) {
    Logger.log(e);
  }
}

