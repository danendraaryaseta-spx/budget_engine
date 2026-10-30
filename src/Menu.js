function onOpen() {
  SpreadsheetApp.getUi().createMenu('Budget Sources').addItem('Initialize hub','setup').addItem('Create engine copy','createEngineCopy')
    .addSeparator().addItem('Discover imports','discoverImports').addItem('Refresh sources','startRefresh').addItem('Continue pending job','continueRefresh')
    .addItem('Retry failed checkpoint','retryFailedJob').addItem('Show status','showRefreshStatus').addSeparator().addItem('Install daily refresh','installRefreshTrigger').addItem('Remove refresh triggers','removeRefreshTriggers').addToUi();
}
function showRefreshStatus() { console.log(JSON.stringify(LH.state(),null,2)); }
function scheduledRefresh() {
  var s=LH.state();
  if (s && ['DISCOVER','PLAN','PULL','MANUAL','PUBLISH','COMMITTING'].indexOf(s.phase)>=0) continueRefresh();
  else if (!s || s.phase!=='FAILED') startRefresh();
}
function resumeScheduledRefresh() { var s=LH.state(); if (s && ['DISCOVER','PLAN','PULL','MANUAL','PUBLISH','COMMITTING'].indexOf(s.phase)>=0) continueRefresh(); }
function installRefreshTrigger() {
  LH.Core.assert(LH.state() && LH.state().phase==='PUBLISHED','Complete a manual refresh first'); removeRefreshTriggers();
  ScriptApp.newTrigger('scheduledRefresh').timeBased().everyDays(1).atHour(7).inTimezone(LH.Config.timezone).create();
  ScriptApp.newTrigger('resumeScheduledRefresh').timeBased().everyMinutes(5).create();
}
function removeRefreshTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (['scheduledRefresh','resumeScheduledRefresh'].indexOf(t.getHandlerFunction())>=0) ScriptApp.deleteTrigger(t); });
}
