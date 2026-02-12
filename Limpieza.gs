
function limpiarZombiesDiario() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  
  let numerosEnVarias = new Set();
  
 
  const hojasVarias = (typeof VARIAS_SHARDS !== 'undefined') ? VARIAS_SHARDS : [
    "VB1", "VB2", "VB3", "VB4", "VB5", 
    "VB6", "VB7", "VB8", "VB9", "VB10"
  ];
  const hojasVentas = (typeof VENTAS_SHARDS !== 'undefined') ? VENTAS_SHARDS : ["V1", "V2"];


  hojasVarias.forEach(nombre => {
    const sh = ss.getSheetByName(nombre);
    if(sh && sh.getLastRow() > 1) {
     
      const datos = sh.getRange(2, 2, sh.getLastRow()-1, 1).getValues().flat();
      datos.forEach(n => numerosEnVarias.add(Number(n)));
    }
  });

  if (numerosEnVarias.size === 0) {
    console.log("No hay datos en Varias Boletas (VB). Nada que limpiar.");
    return;
  }

  
  let eliminados = 0;

  hojasVentas.forEach(nombre => {
    const sh = ss.getSheetByName(nombre);
    if(sh && sh.getLastRow() > 1) {
      const ultimaFila = sh.getLastRow();
      
     
      const datos = sh.getRange(2, 2, ultimaFila-1, 1).getValues().flat();
      
     
      for (let i = datos.length - 1; i >= 0; i--) {
        const numeroEnVenta = Number(datos[i]);
        const filaReal = i + 2; 

        
        if (numerosEnVarias.has(numeroEnVenta)) {
          
          console.log(`🧟 Zombie detectado: Boleta ${numeroEnVenta} en hoja "${nombre}". Eliminando...`);
          sh.deleteRow(filaReal);
          eliminados++;
        }
      }
    }
  });

  if (eliminados > 0) {
    _registrarAuditoria("LIMPIEZA AUTOMÁTICA", "SISTEMA", "ROBOT", `Se eliminaron ${eliminados} registros duplicados (zombies) de Ventas.`);
  }
  
  console.log(`✅ Limpieza terminada. Zombies eliminados: ${eliminados}`);
}