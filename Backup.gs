const BACKUP_FOLDER_ID = "1iYlyAht5ZLWUngYOihhE-yQz2QJVee76";

function backupToFolder() {
  const file   = DriveApp.getFileById(SPREADSHEET_ID);
  const folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);

  const ts   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
  const name = `Backup Rifa - ${ts}`;

  file.makeCopy(name, folder);
  Logger.log("✅ Backup creado: " + name);
}

function pruneBackups_keepLastN() {
  const N = 60;
  const folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);

  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().startsWith("Backup Rifa - ")) files.push(f);
  }

  files.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());

  for (let i = N; i < files.length; i++) {
    files[i].setTrashed(true);
  }
  Logger.log(`🧹 Limpieza completa. Conservados: ${Math.min(N, files.length)}.`);
}