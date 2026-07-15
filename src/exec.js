function logAndReturnMigrationResult_(result) {
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function getMigrationSourceSpreadsheetIdForRun_() {
  return PropertiesService.getScriptProperties().getProperty('MIGRATION_SOURCE_SHEET_ID') || BillUtil.databaseSheetId;
}

function getMigrationTargetSpreadsheetNameForRun_() {
  return PropertiesService.getScriptProperties().getProperty('MIGRATION_TARGET_SHEET_NAME') || `AM Database Migrated ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;
}

function runValidateCurrentDatabase() {
  return logAndReturnMigrationResult_(validateNamedDatabaseSheets(BillUtil.databaseSheetId));
}

function runValidateMigrationSourceDatabase() {
  return logAndReturnMigrationResult_(validateNamedDatabaseSheets(getMigrationSourceSpreadsheetIdForRun_()));
}

function runMigrateDatabaseToNamedSheets() {
  return logAndReturnMigrationResult_(migrateDatabaseToNamedSheets(
    getMigrationSourceSpreadsheetIdForRun_(),
    getMigrationTargetSpreadsheetNameForRun_()
  ));
}

function runValidateFakePreMigrationDatabase() {
  return logAndReturnMigrationResult_(validateNamedDatabaseSheets('1Yx4UoqVl_8VAKFE1ffz0QSnLWnmPr_DaK4AG9ngEOyc'));
}

function runMigrateFakePreMigrationDatabase() {
  return logAndReturnMigrationResult_(migrateDatabaseToNamedSheets(
    '1Yx4UoqVl_8VAKFE1ffz0QSnLWnmPr_DaK4AG9ngEOyc',
    `移行後テスト ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`
  ));
}
