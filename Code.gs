function formResponseToPdf() {
  sendEmail(generatePdf(generateDocument()));
}
