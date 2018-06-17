function generateDocument() {
  // Form
  var form = FormApp.getActiveForm();
  var formResponses = form.getResponses();
  var lastResponse = formResponses.length - 1;
  var formResponse = formResponses[lastResponse];
  var itemResponses = formResponse.getItemResponses();
  var spreadsheetId = form.getDestinationId();
  var numberOfQuestions = itemResponses.length;
  var totalPoints = 0;
  // Sheet
  var sheetActive = SpreadsheetApp.openById(spreadsheetId);
  var sheet = sheetActive.getSheetByName("Feedbacks");
  var sheetOverallFeedback = sheetActive.getSheetByName("Feedback Geral");
  var overallFeedbackUpperBoundColumn = 1;
  var overallFeedbackColumn = 2;
  var searchColumnOne = 1;
  var searchColumnTwo = 2;
  var feedbackColumn = 4;
  var pointColumn = 3;
  var rangeLastColumn = sheet.getLastColumn();
  var rangeLastRow = sheet.getLastRow();
  var values = sheet.getRange(1, 1, rangeLastRow, rangeLastColumn).getValues();
  var sheetOverallFeedbackRangeLastColumn = sheetOverallFeedback.getLastColumn();
  var sheetOverallFeedbackRangeLastRow = sheetOverallFeedback.getLastRow();
  var sheetOverallFeedbackvalues = sheetOverallFeedback.getRange(1, 1, rangeLastRow, rangeLastColumn).getValues();
  //Document
  var date = new Date();
  var documentToPdf = DocumentApp.create(itemResponses[numberOfQuestions - 1].getResponse()+ " - " + date.toString());
  var body = documentToPdf.getBody();
  
  //Feedback for every question
  for (var j = 0; j < itemResponses.length; j++) {
    for (var i = 1; i < rangeLastRow; i++){
      if (values[i][searchColumnOne] == itemResponses[j].getItem().getTitle()){
        if (values[i][searchColumnTwo] == itemResponses[j].getResponse()){
          totalPoints = totalPoints + values[i][pointColumn];
          body.appendParagraph('A pergunta #' + values[i][0]
                              + ' que era "' + itemResponses[j].getItem().getTitle() 
                              + '" a sua respotas foi "' + itemResponses[j].getResponse()
                              + '" o nosso feedback é "' + values[i][feedbackColumn] + '"');
        }
      }
    }
  }
  
  //Overall Feedback of the form 
  if(totalPoints <= sheetOverallFeedbackvalues[1][overallFeedbackUpperBoundColumn]){
    body.appendParagraph('Você fez ' + totalPoints
              + '" o nosso feedback para a sua autoavaliação é "' + sheetOverallFeedbackvalues[1][overallFeedbackColumn] + '"');
  }else{
    if(totalPoints <= sheetOverallFeedbackvalues[2][overallFeedbackUpperBoundColumn]){
      body.appendParagraph('Você fez ' + totalPoints
                + '" o nosso feedback para a sua autoavaliação é "' + sheetOverallFeedbackvalues[2][overallFeedbackColumn] + '"');
      }else{
        body.appendParagraph('Você fez ' + totalPoints
                + '" o nosso feedback para a sua autoavaliação é "' + sheetOverallFeedbackvalues[3][overallFeedbackColumn] + '"');
      }
  }
  documentToPdf.saveAndClose();
  return {documentToPdf:documentToPdf,formResponse:formResponse};
}
