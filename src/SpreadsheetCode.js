//-----------On Open-----------
function onOpen() {
  addUiButtons();
}

//-----------Modify UI-----------
function addUiButtons() {
  const ui = SpreadsheetApp.getUi();

  const menuConfig = ui.createMenu("✉️ Email Merge");
  menuConfig.addItem("⬅️ Show Sidebar", "showSideBar").addToUi();
}

function showSelectedTemplate(options){
  let utils = EmailMergeApp.getApp().newUtils();
  let result = utils.showSelectedTemplate(options);
  return result;
}

function sendEmails(options){
  let batchProcessor = EmailMergeApp.getApp()
                                .newBatchProcessor()
                                .loadTemplatesFromGmail()
                                .loadRecipientsFromSheet(options)
                                .validate()

  let result = batchProcessor.sendEmails(options);
  return result;
}

function refreshTemplates() {
  let batchProcessor = EmailMergeApp.getApp()
                                .newBatchProcessor()
                                .loadTemplatesFromGmail();

  let result = batchProcessor.refreshTemplates();
  return result;
};

function checkUndeliveredEmails(options) {
  let batchProcessor = EmailMergeApp.getApp()
                                .newBatchProcessor()
                                .loadTemplatesFromGmail()
                                .loadRecipientsFromSheet(options)
                                .validate()
  let result = batchProcessor.checkUndeliveredEmails(options)
  return result;
}

function checkResponsesToEmails(options) {
  let batchProcessor = EmailMergeApp.getApp()
                                .newBatchProcessor()
                                .loadTemplatesFromGmail()
                                .loadRecipientsFromSheet(options)
                                .validate()
  let result = batchProcessor.checkResponsesToEmails(options)
  return result;
}

function showSideBar() {
  let utils = EmailMergeApp.getApp().newUtils();
  utils.initialSetup();
  utils.showSidebar();
  return;
}

function updateUserProperties(data) {
  let utils = EmailMergeApp.getApp().newUtils();
  let result = utils.updateUserProperties(data);
  return result;
}

function initialSetup(){
  let utils = EmailMergeApp.getApp().newUtils();
  utils.initialSetup();
  return;
}

function cleanSheet(){
  let utils = EmailMergeApp.getApp().newUtils();
  let result = utils.cleanSheet();
  return result;
}

function include(filename) {
  let utils = EmailMergeApp.getApp().newUtils();
  let result = utils.include(filename);
  return result;
}
