function getApp() {
  return EmailMergeApp;
}

let EmailMergeApp = (function() {

  class Utils {
    constructor() {

      this.include = function(filename) {
        return HtmlService.createHtmlOutputFromFile(filename).getContent();
      };

      this.emailValidator = function(email) {
        email = email.trim();
        let emailRegex = /^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,26}$/;
        return emailRegex.test(email);
      };

      this.extractEmail = function(email) {
        let emailRegex = /<(.+)>/;
        let match = email.match(emailRegex);
        return match ? match[1] : email;
      };

      this.replaceAll = function(originalText, search, replacement) {
        return originalText.replace(new RegExp(search, 'g'), replacement);
      };

      this.capitalizeText = function(text) {
        text = text.trim();
        let regFirstLetter = /\b(\w)/g;
        let regOtherLetters = /\B(\w)/g;
        return text.replace(regFirstLetter, match => match.toLocaleUpperCase())
                   .replace(regOtherLetters, match => match.toLocaleLowerCase());
      };

      this.loadTemplatesFromGmail = function() {
        let drafts = GmailApp.getDraftMessages();
        return drafts.map(draft => ({
          subject: draft.getSubject(),
          body: draft.getPlainBody(),
          htmlBody: draft.getBody()
        }));
      };

      this.getLabelNamesFromGmail = function() {
        let labelNames = GmailApp.getUserLabels().map(label => label.getName());
        return labelNames;
      };

      this.cleanSheet = function() {
        let ui = SpreadsheetApp.getUi();
        let cleanSheetConfirmation = ui.alert(`Are you sure you want to clean the current sheet?`, ui.ButtonSet.OK_CANCEL);
        if (cleanSheetConfirmation === ui.Button.CANCEL) {
          return {title: '', message: '', result: ''};
        }
        let sheet = SpreadsheetApp.getActiveSheet();
        let range = sheet.getDataRange().offset(1, 0);
        range.deleteCells(SpreadsheetApp.Dimension.ROWS);
        let batchProcessor = new BatchProcessor();
        batchProcessor.loadTemplatesFromGmail()
                      .refreshTemplates();
        return {title: 'Clean Sheet', message: 'Sheet was successfully cleared!', result: 'success'};
      };

      this.showSelectedTemplate = function(options) {
        options = JSON.parse(options);
        let sheet = SpreadsheetApp.getActiveSheet();
        let [headers, ...data] = sheet.getDataRange().getValues();
        let row = sheet.getCurrentCell().getRow();
        let column = headers.indexOf('{{template}}') +1;
        let subject = sheet.getRange(row, column).getValue();
        if (!subject) {
          return {title: 'Show Selected Template', message: `You have to select the template subject you want to show up`, result: 'warning'};
        }
        let templates = this.loadTemplatesFromGmail();
        let templateFromGmail = templates.find(item => item.subject === subject);
        if (!templateFromGmail) {
          return {title: 'Show Selected Template', message: `Template not found on Gmail saved templates`, result: 'danger'};
        }
        let template = new Template().setSubject(templateFromGmail.subject)
                                       .setHtmlBody(templateFromGmail.htmlBody)
                                       .setSignatureFromGmail()
                                       .validate();

        let recipient = new Recipient().setEmail("noemail@email.com")
                                        .setTemplate(template);
        if(options?.replacePlaceholders === true){
          headers.forEach((header, index) => {
            let object = { placeholder: header, replacement: data[row - 2][index] };
            recipient.addPlaceholderAndReplacement(object);
          });
        }
        recipient.validate(JSON.stringify(options));
        let recipientTemplate = recipient.getTemplate();
        let body = recipientTemplate?.body ? recipientTemplate?.body : recipientTemplate?.htmlBody;
        if(options?.useSignature === true){
          body += '<br>' + '--';
          body += '<br>' + recipientTemplate?.signature;
        }
        let temp = HtmlService.createTemplateFromFile('html/template_viewer');
        temp.data = JSON.stringify({subject: recipient.getTemplate().subject, body: body});
        let html = temp.evaluate().setWidth(600).setHeight(455);
        SpreadsheetApp.getUi().showModalDialog(html, " ");
      };

      this.showSidebar = function() {
        let userProperties = PropertiesService.getUserProperties();
        let userData = userProperties.getProperties();
        let template = HtmlService.createTemplateFromFile("html/sidebar");
        template.labels = JSON.stringify(this.getLabelNamesFromGmail());
        template.userData = JSON.stringify(userData);
        let html = template.evaluate();
        SpreadsheetApp.getUi().showSidebar(html);
      };

      this.updateUserProperties = function(data) {
        let userProperties = PropertiesService.getUserProperties();
        userProperties.setProperties(JSON.parse(data));
        return {title: 'Save', message: 'Settings were successfully saved!', result: 'success'};
      };

      this.initialSetup = function() {
        let sheet = SpreadsheetApp.getActiveSheet();
        let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 2).getValues().flat();
        if (!headers.includes('{{email}}')) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue('{{email}}');
        }
        if (!headers.includes('{{template}}')) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue('{{template}}');
        }
      };
    }
  }

  class Recipient {
    constructor() {
      let _email = '';
      let _placeholdersAndReplacements = []; // {placeholder: '', replacement: ''}
      let _template = {};
      let _validated = false;

      this.setEmail = function(email) {
        let utils = new Utils();
        if (!utils.emailValidator(email)) {
          return {title: 'Email Error', message: `Invalid Email: ${email}`, result: 'danger'};
        }
        _email = email;
        return this;
      };

      this.setTemplate = function(template) {
        _template = template.toJSON();
        return this;
      };

      this.getTemplate = function() {
        return _template;
      };

      this.addPlaceholderAndReplacement = function(placeholder) {
        if(!placeholder?.placeholder) return;
        if (placeholder.placeholder && !placeholder.replacement) {
            throw Error(JSON.stringify({title: 'Placeholders Warning', message: `Take a look at replacement for placeholder ${placeholder.placeholder} on ${_email}`, result: 'warning'}));
        }
        _placeholdersAndReplacements.push(placeholder);
        return this;
      };

      this.toJSON = function() {
        return {
          email: _email,
          placeholdersAndReplacements: JSON.stringify(_placeholdersAndReplacements),
          template: _template,
          validated: _validated
        };
      };

      this.validate = function(options) {
        options = JSON.parse(options);
        if (!_template?.validated) {
          return {title: 'Template Error', message: `Invalid Template`, result: 'danger'};
        }
        let utils = new Utils();
        if(options?.replacePlaceholders === true && _placeholdersAndReplacements.length > 0){
          _placeholdersAndReplacements.forEach(item => {
            _template.subject = utils.replaceAll(
              _template?.subject, 
              item.placeholder, 
              options?.capitalizeText ? utils.capitalizeText(item.replacement) : item.replacement
            );
            _template.body = utils.replaceAll(
              _template?.body, 
              item.placeholder, 
              options?.capitalizeText ? utils.capitalizeText(item.replacement) : item.replacement
            );
            _template.htmlBody = utils.replaceAll(
              _template?.htmlBody, 
              item.placeholder, 
              options?.capitalizeText ? utils.capitalizeText(item.replacement) : item.replacement
            );
          });
        }
        _validated = true;
        return this;
      };

      this.sendEmail = function(options) {
        options = JSON.parse(options);
        let body = _template?.body ? _template?.body : _template?.htmlBody;
        if(options?.useSignature === true){
          body += '<br>' + '--';
          body += '<br>' + _template?.signature;
        }
        let userData = PropertiesService.getUserProperties().getProperties();
        let emailData = {
          name: userData?.name || "",
          to: _email,
          subject: _template?.subject,
          htmlBody: body
        };
        try {
          let thread = GmailApp.createDraft(null, null, null, emailData)
                             .send()
                             .getThread();
          if(options?.labels?.length > 0){
            options.labels.forEach( labelName => {
              let label = GmailApp.getUserLabelByName(labelName.trim()) 
              thread.addLabel(label)
            })
          }
        } catch (e) {
        }
      };
    }
  }

  class Template {
    constructor() {
      let _subject = '';
      let _body = '';
      let _htmlBody = '';
      let _signature = '';
      let _validated = false;

      this.setSubject = function(subject) {
        _subject = subject;
        return this;
      };

      this.setBody = function(body) {
        _body = `<div>${body}</div>`;
        return this;
      };

      this.setHtmlBody = function(htmlBody) {
        _htmlBody = htmlBody;
        return this;
      };

      this.setSignature = function(signature) {
        _signature = `<div>${signature}</div>`;
        return this;
      };

      this.setSignatureFromGmail = function() {
        let userEmail = Gmail.Users.getProfile('me').emailAddress;
        _signature = Gmail.Users.Settings.SendAs.get("me", userEmail).signature;
        return this;
      };

      this.toJSON = function() {
        return {
          subject: _subject,
          body: _body,
          htmlBody: _htmlBody,
          signature: _signature,
          validated: _validated
        };
      };

      this.validate = function() {
        if (!_subject) {
          console.log("Invalid Template Subject");
          return this;
        }
        if (!_body && !_htmlBody) {
          console.log("Invalid Template Body");
          return this;
        }
        if (!_signature) {
          console.log("Invalid Template Signature");
          return this;
        }
        _validated = true;
        return this;
      };
    }
  }

  class BatchProcessor {
    constructor() {
      let _recipients = [];
      let _templates = [];
      let _emailsColumnHeader = '{{email}}';
      let _templatesColumnHeader = '{{template}}';
      let _validated = false;

      this.loadTemplatesFromGmail = function() {
        _templates = new Utils().loadTemplatesFromGmail();
        return this;
      };

      this.validate = function() {
        if (_templates.length === 0) {
          throw Error("There are no templates loaded, call the loadRecipientsFromSheet function");
        }
        if (_recipients.length === 0) {
          throw Error("There are no recipients loaded, call the loadRecipientsFromSheet function");
        }
        if (_recipients.filter(recipient => !recipient.toJSON().validated).length > 0) {
          throw Error("Some recipients are invalid");
        }
        _validated = true;
        console.log("Batch Processor Valid!!");
        return this;
      };

      this.sendEmails = function(options) {
        if (!_validated) {
          throw Error("Invalid Operation, you must validate the Batch Processor first");
        }
        let ui = SpreadsheetApp.getUi();
        // Show confirmation dialog before sending emails
        let sendEmailsConfirmation = ui.alert(`You are about to send a mass email to ${_recipients.length} recipients. Are you sure you want to proceed?`, ui.ButtonSet.OK_CANCEL);
        if (sendEmailsConfirmation === ui.Button.OK) {
          _recipients.forEach(recipient => {
            recipient.sendEmail(options);
          });
          return {title: 'Send Emails', message: 'All emails were successfully sent!', result: 'success'};
        }
        return {title: 'Send Emails', message: 'No email was sent', result: 'warning'};
      };

      this.loadRecipientsFromSheet = function(options) {
        if (_templates.length === 0) {
          throw Error(JSON.stringify({title: 'Template Error', message: `You need to retrieve the templates from Gmail first`, result: 'danger'}));
        }
        if (!_emailsColumnHeader || !_templatesColumnHeader) {
          throw Error(JSON.stringify({title: 'Template Error', message: `You must define both the email column header and the templates column header first`, result: 'danger'}));
        }
        let sheet = SpreadsheetApp.getActiveSheet();
        let [headers, ...data] = sheet.getDataRange().getValues();
        if (data.length == 0) {
          throw Error(JSON.stringify({title: 'Recipients Error', message: `You need to add some recipients first`, result: 'danger'}));
        }
        let emailsColumn = headers.indexOf(_emailsColumnHeader);
        let templatesColumn = headers.indexOf(_templatesColumnHeader);
        for (let i = 0; i < data.length; i++) {
          let row = data[i];
          let subject = row[templatesColumn].trim();
          let email = row[emailsColumn].trim();
          console.log(`Working on row ${i + 1} (${email}: ${subject})`);
          if (!email) {
            throw Error(JSON.stringify({title: 'Email Error', message: `There is no email in row ${i + 1} of data`, result: 'danger'}));
          }
          if (!subject) {
            throw Error(JSON.stringify({title: 'Template Error', message: `There is no template selected in row ${i + 1} of data`, result: 'danger'}));
          }
          let templateFromGmail = _templates.find(item => item.subject === subject);
          if (!templateFromGmail) {
            throw Error(JSON.stringify({title: 'Template Error', message: `Template in row ${i + 1} not found on Gmail saved templates`, result: 'danger'}));
          }
          let template = new Template().setSubject(templateFromGmail.subject)
                                       .setHtmlBody(templateFromGmail.htmlBody)
                                       .setSignatureFromGmail()
                                       .validate();
          let recipient = new Recipient().setEmail(email)
                                         .setTemplate(template);
          if(JSON.parse(options)?.replacePlaceholders === true){
            headers.forEach((header, index) => {
              let object = { placeholder: header, replacement: row[index] };
              recipient.addPlaceholderAndReplacement(object);
            });
          }
          recipient.validate(options);
          _recipients.push(recipient);
        }
        return this;
      };

      this.refreshTemplates = function() {
        if (_templates.length === 0) {
          return {title: 'Refresh Templates', message: `You need to retrieve the templates from Gmail first`, result: 'danger'};
        }
        let sheet = SpreadsheetApp.getActiveSheet();
        let [headers, ...data] = sheet.getDataRange().getValues();
        let templatesColumn = headers.indexOf(_templatesColumnHeader) + 1;
        let subjects = _templates.map(item => item.subject);
        let validation = SpreadsheetApp.newDataValidation()
                                       .setAllowInvalid(false)
                                       .requireValueInList(subjects, true)
                                       .build();
        if (!templatesColumn) {
          return {title: 'Refresh Templates', message: `You must define the templates column header first`, result: 'danger'};
        }
        sheet.getRange(2, templatesColumn, data.length || 1 , 1).setDataValidation(validation);
        return {title: 'Refresh Templates', message: 'Templates successfully refreshed!', result: 'success'};
      };

      this.checkUndeliveredEmails = function(options) {
        options = JSON.parse(options)
        if (!_validated) {
          throw Error("Invalid Operation, you must validate the Batch Processor first");
        }
        let subjects = ["Delivery Status Notification (Failure)", "Undelivered Mail Returned to Sender", "Mail Delivery Subsystem"]
        let emailAddresses = _recipients.map( recipient => recipient.toJSON().email)
        let utils = new Utils();
        let query = `to:('${emailAddresses.join("' OR '")}')`;
        if(options?.startDate && options?.endDate){
          query += ` after:${options.startDate}` + ` before:${options.endDate}`;
        }
        let threads = GmailApp.search(query);
        let undeliveredEmails = [];
        threads.forEach(thread => {
          let messages = thread.getMessages();
          let firstMessage = messages[0];
          let firstMessageTo = firstMessage.getTo();
          let firstMessageSubject = firstMessage.getSubject();
          let index = _recipients.findIndex( recipient => {
            let subject = recipient.toJSON().template.subject;
            let email = recipient.toJSON().email;
            if(firstMessageTo.includes(email) && subject === firstMessageSubject){
              return recipient
            }
          });
          if(index == -1) return;
          messages.forEach(message => {
            let subject = message.getSubject();
            if (subjects.includes(subject)) {
              undeliveredEmails.push({
                index: index + 2,
                email: utils.extractEmail(firstMessageTo),
                subject: firstMessageSubject,
                date: message.getDate(),
                body: message.getPlainBody()
              });
            }
          });
        });
        let rangeList = undeliveredEmails.map( email => `${email.index}:${email.index}`)
        if(rangeList.length == 0) return;
        SpreadsheetApp.getActiveSheet().getRangeList(rangeList).setBackground('#f4cccc');
      };

      this.checkResponsesToEmails = function(options) {
        options = JSON.parse(options)
        if (!_validated) {
          throw Error("Invalid Operation, you must validate the Batch Processor first");
        }
        let emailAddresses = _recipients.map( recipient => recipient.toJSON().email)
        let utils = new Utils();
        let query = `to:('${emailAddresses.join("' OR '")}')`;
        if(options?.startDate && options?.endDate){
          query += ` after:${options.startDate}` + ` before:${options.endDate}`;
        }
        let threads = GmailApp.search(query);
        let respondedEmails = [];
        threads.forEach(thread => {
          let messages = thread.getMessages();
          let firstMessage = messages[0];
          let firstMessageTo = firstMessage.getTo();
          let firstMessageSubject = firstMessage.getSubject();
          let index = _recipients.findIndex( recipient => {
            let subject = recipient.toJSON().template.subject;
            let email = recipient.toJSON().email;
            if(firstMessageTo.includes(email) && subject === firstMessageSubject){
              return recipient
            }
          });
          if(index == -1) return;
          messages.forEach(message => {
            let fromEmail = utils.extractEmail(message.getFrom());
            if (emailAddresses.includes(fromEmail)) {
              respondedEmails.push({
                index: index + 2,
                email: fromEmail,
                subject: firstMessageSubject,
                date: message.getDate(),
                body: message.getPlainBody()
              });
            }
          });
        });
        let rangeList = respondedEmails.map( email => `${email.index}:${email.index}`)
        if(rangeList.length == 0) return;
        SpreadsheetApp.getActiveSheet().getRangeList(rangeList).setBackground('#d9ead3');
      };
    }
  }

  return {

    newRecipient: function() {
      return new Recipient();
    },

    newTemplate: function() {
      return new Template();
    },

    newBatchProcessor: function() {
      return new BatchProcessor();
    },

    newUtils: function() {
      return new Utils();
    }
    
  };
})();
