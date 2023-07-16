function calculateScoreAndSendEmail() {
  // Get the active sheet
  var sheet = SpreadsheetApp.getActiveSheet();

  // Define the columns to use for calculating the scores
  var depressionColumns = [11,13,18,21,24,25,29];
  var anxietyColumns = [10,12,15,17,23,27,28];
  var stressColumns = [9,14,16,19,20,22,26];

  // Get the last row that was processed (stored in script properties)
  var lastProcessedRow = parseInt(PropertiesService.getScriptProperties().getProperty('lastProcessedRow')) || 0;

  // Get the last row with data
  var lastRow = sheet.getLastRow();

  // Loop through each row that hasn't been processed yet
  for (var i = lastProcessedRow + 1; i <= lastRow; i++) {
    // Get the values in the specified columns for the current row
    var depressionValues = depressionColumns.map(function(column) {
      return sheet.getRange(i, column).getValue();
    });
    var anxietyValues = anxietyColumns.map(function(column) {
      return sheet.getRange(i, column).getValue();
    });
    var stressValues = stressColumns.map(function(column) {
      return sheet.getRange(i, column).getValue();
    });

    // Calculate the sum of values for each category
    var depressionScore = depressionValues.reduce(function(sum, value) {
      return sum + value;
    }, 0);
    var anxietyScore = anxietyValues.reduce(function(sum, value) {
      return sum + value;
    }, 0);
    var stressScore = stressValues.reduce(function(sum, value) {
      return sum + value;
    }, 0);

    // Write the results to the appropriate cells
    sheet.getRange(i, 31).setValue(depressionScore);
    sheet.getRange(i, 32).setValue(anxietyScore);
    sheet.getRange(i, 33).setValue(stressScore);

    // Get the email address from column 30
    var email = sheet.getRange(i, 30).getValue();
    
    // Check if the email address is valid
    var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailPattern.test(email)) {
      Logger.log("Invalid email: " + email);
      continue;
    }

    // Send an email to the sender
    var subject = "DASS Questionnaire Results";
    var body = "Your DASS questionnaire results have been calculated.\n\n" +
               "Depression score: " + depressionScore + "\n" +
               "Anxiety score: " + anxietyScore + "\n" +
               "Stress score: " + stressScore + "\n\n\n" +
               "IF YOUR SCORE ABOVE following THRESHOLD, REFER YOUR COUNSELOR/ HR" + "\n\n" +
               "Depression Severe Level > 13" + "\n" +
               "Anxiety Severe Level > 11" + "\n" +
               "Stress Severe Level > 8" + "\n\n" +
               "Thank you for participating.";

    // Send email only if the current row is the latest row processed
    if (i == lastRow) {
      MailApp.sendEmail(email, subject, body);
    }
  }

  // Save the latest processed row number to script properties
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('lastProcessedRow', lastRow);
}
