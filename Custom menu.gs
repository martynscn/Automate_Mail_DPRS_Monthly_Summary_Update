function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menuEntries = [];
  menuEntries.push({
        name: "Trace Dependents",
        functionName: "traceDependents"
        });
  menuEntries.push({
        name: "Trace Dependents",
        functionName: "traceDependents"
        });
  var newMenu = ui.createMenu('New menu')
    .addItem('Send mail', 'sendEmails5')
    .addItem("Trace Dependents","traceDependents")
    .addItem('Trace dependents of named ranges', "traceDependentsOfNamedRanges");
    
  ui.createMenu('Custom menu')
  .addSubMenu(newMenu)
  .addToUi();
}