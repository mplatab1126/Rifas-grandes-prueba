function registrarDatosDiarios() {
  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  const DASHBOARD_SHEET_NAME = "Dashboard Principal";
  const HISTORICO_SHEET_NAME = "Histórico Diario";

  const dashboardSheet = SPREADSHEET.getSheetByName(DASHBOARD_SHEET_NAME);
  const historicoSheet = SPREADSHEET.getSheetByName(HISTORICO_SHEET_NAME);

  const ayer = new Date();
  ayer.setDate(ayer.getDate() - 1);
  
  dashboardSheet.getRange("C2").setValue(ayer);
  
  SpreadsheetApp.flush();
  Utilities.sleep(3000);

  const fecha = dashboardSheet.getRange("C2").getValue();
  const totalBoletas = dashboardSheet.getRange("C3").getValue();
  const conAbono = dashboardSheet.getRange("C4").getValue();
  const sinAbono = dashboardSheet.getRange("C5").getValue();

  historicoSheet.appendRow([fecha, totalBoletas, conAbono, sinAbono]);
}