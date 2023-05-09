function showAlert() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Failed accessing sheet by ID',
     'Try again?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Retrying...');
    copyDataWithFilter();
  } else {
    // User clicked "No" or X in the title bar.
    //ui.alert('Closing...');
    console.log("Alert window closed");
    return
  }
}