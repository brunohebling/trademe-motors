function resumeFromLastBatch() {
  // Check what state was saved before the crash
  var props = PropertiesService.getScriptProperties();
  Logger.log('Resuming from: device=' + props.getProperty('zenserpDevice') + 
    ', batch=' + props.getProperty('zenserpBatchIndex'));
  
  // Mark as running again and schedule the next trigger
  props.setProperty('zenserpRunning', 'true');
  deleteZenserpBatchTriggers();
  ScriptApp.newTrigger('runZenserpAutomation')
    .timeBased().after(10 * 1000).create();
  
  Logger.log('✓ Resuming in 10 seconds...');
}
