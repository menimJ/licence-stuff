function setupTriggers() {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const existingTriggers =
    ScriptApp.getProjectTriggers();

  const functionsToReplace = [
    'onFormSubmit',
    'processSystemJobs',
    'cleanupExpiredPaymentProofs'
  ];

  existingTriggers.forEach((trigger) => {
    const handler =
      trigger.getHandlerFunction();

    if (
      functionsToReplace.includes(handler)
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(
    'onFormSubmit'
  )
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger(
    'processSystemJobs'
  )
    .timeBased()
    .everyMinutes(1)
    .create();

  ScriptApp.newTrigger(
    'cleanupExpiredPaymentProofs'
  )
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();

  Logger.log(
    'CRFFN triggers created successfully.'
  );
}