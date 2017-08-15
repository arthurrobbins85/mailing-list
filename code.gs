function changeLastCheck(){
  
  var last_check = ScriptProperties.getProperty("last_check");
  var d = new Date();
  var d2 = new Date();
  d2.setDate(d.getDate() - 10);
  
  
  //Logger.log(last_check);
  ScriptProperties.setProperty("last_check", d2);
  
}

var ss = SpreadsheetApp.getActiveSpreadsheet();
var dashboardSheet = ss.getSheetByName('Dashboard');
var english = ["Unsubscribe", "You can find those updates in attachment.", "Preview", "Link", "List"];
var french = ["Résilier mon abonnement", "Vous pouvez retrouver ces nouveautés en pièce-jointe.", "Aperçu", "Lien", "Liste"];

function send_changes() {
  var current_time = new Date();
  var last_check = ScriptProperties.getProperty("last_check");
  if (last_check == null) {
    last_check = current_time;
    ScriptProperties.setProperty("last_check", last_check);
  }
  var updates = '';
  var url = dashboardSheet.getRange('C5').getValue();
  var when = dashboardSheet.getRange('C8').getValue();
  var lang = dashboardSheet.getRange('G12').getValue();
  var lang = (dashboardSheet.getRange('G12').getValue().indexOf('English') != -1) ? english : french;
  var mailType = dashboardSheet.getRange('C10').getValue();
  var scope = dashboardSheet.getRange('C3').getValue();
  
  if (scope == 'Site') {
    var site = SitesApp.getSiteByUrl(url);
    var descendants = site.getAllDescendants();
    for (var j = 0; j < descendants.length; j++) {
      updates = createDigest_(descendants[j], updates, last_check, when, lang);
    }
    if (updates != "") sendMessage_(updates, lang, mailType);
  }
  else if (scope == 'Page and all subpages') {
    var page = SitesApp.getPageByUrl(url);
    var descendants = page.getAllDescendants();
    for (var j = 0; j < descendants.length; j++) {
      updates = createDigest_(descendants[j], updates, last_check, when, lang);
    }
    if (updates != "") sendMessage_(updates, lang, mailType);
  }
  else if (scope == 'Blog' && mailType != 'Digest') {
    var page = SitesApp.getPageByUrl(url);
    /*createNewsletter_ var coucou =*/ createNewsletter_(page, updates, last_check, when, lang);
  }
  else {
    var page = SitesApp.getPageByUrl(url);
    updates = createDigest_(page, updates, last_check, when, lang);
    if (updates != "") sendMessage_(updates, lang, mailType);
  }
  ScriptProperties.setProperty("last_check", current_time);
}

function sendMessage_(updates, lang, mailType) {
  // Remove duplicates, remove people who have unsubscribed, remove false email addresses
  var mailing = scrubData_();
  var subject = dashboardSheet.getRange('G5').getValue();
  
  // HEADER
  var header = HtmlService.createTemplateFromFile('header.html');
  header.intro = dashboardSheet.getRange('G7').getValue();
  header.newstitle = dashboardSheet.getRange('H3').getValue();


  
  header = header.evaluate().getContent();
  
  var body = header;
  
  
  if(mailType == 'Digest'){
    var subject = dashboardSheet.getRange('G5').getValue();
  }
  else{
    var subject = dashboardSheet.getRange('G5').getValue()+': '+mailType;
  }
  ///////////////////////////////////////////////////////////////////////////////////
  // If the number of updates is too important, send those updates as attachment
  ///////////////////////////////////////////////////////////////////////////////////
  if (updates.length > 200000) body += "*** " + lang[1] + " ***";
  else body += updates;
  
  
  // FOOTER
  
  // get the template from the HTML file
  var footer = HtmlService.createTemplateFromFile('footer.html');
  footer.closingtext = dashboardSheet.getRange('G10').getValue();
  
  // assign variable in the template
  footer.formURL = ss.getFormUrl();
  
  // Evaluate the content and get the HTML code
  body += footer.evaluate().getContent();
  
  
  ////////////////////
  // Batch emails   
  ////////////////////
  for (var j = 0; j < mailing.length; j = j + 40) {
    var bcc = "";
    var sender = "Patricia Vear";
    for (var k = j; k < j + 40; k++) {
      if (k < mailing.length && mailing[k][1] != "") {
        bcc += mailing[k][1] + ",";
      }
    }
    if (updates.length > 200000) {
      MailApp.sendEmail('notify@google.com', subject, body, {
        bcc: bcc,
        htmlBody: body,
        name: sender,
        attachments: [Utilities.newBlob(updates, "text/html", "Updates")]
      });
    }
    else {
      MailApp.sendEmail('notify@google.com', subject, body, {
        bcc: bcc,
        htmlBody: body
      });
    }
  }
  var emails_sent = dashboardSheet.getRange('G17').getValue();
  dashboardSheet.getRange('G17').setValue(emails_sent + 1);
}

function createDigest_(page, updates, last_check, when, lang) {
  var type = page.getPageType();
  //Logger.log(type);
  switch (type.toString()) {
    case 'AnnouncementsPage':
      //Logger.log('AnnouncementsPage');
      var newsfeed = page.getAnnouncements();
      for (var j = 0; j < newsfeed.length; j++) {
        if (chooseWhen_(newsfeed[j], when) > new Date(last_check).getTime()) {
          var content = HtmlService.createTemplateFromFile('content.html');
          content.TextContent = newsfeed[j].getTextContent().substr(0, 280) + "...";
          content.Title = newsfeed[j].getTitle();
          content.PageType = "Blog";
          content.PageTitle = page.getTitle();
          content.Url = newsfeed[j].getUrl();
          updates += content.evaluate().getContent();
        }
      }
      break;
    case 'FileCabinetPage':
      //Logger.log('FCPage');
      var attachments = page.getAttachments();
      for (var j = 0; j < attachments.length; j++) {
        if (chooseWhen_(attachments[j], when) > new Date(last_check).getTime()) {
          var content = HtmlService.createTemplateFromFile('content.html');
          content.TextContent = attachments[j].getDescription().substr(0, 280) + "...";
          content.Title = page.getTitle();
          content.PageType = "File Cabinet";
          content.PageTitle = attachments[j].getTitle();
          content.Url = attachments[j].getUrl();
          updates += content.evaluate().getContent();
          
        }
      }
      break;
    case 'WebPage':
      //Logger.log('WPage');
      if (chooseWhen_(page, when) > new Date(last_check).getTime()) {
        var content = HtmlService.createTemplateFromFile('content.html');
        content.TextContent = page.getTextContent().substr(0, 280) + "...";
        content.Title = page.getTitle();
        content.PageType = "Web Page";
        content.Url = page.getUrl();
        updates += content.evaluate().getContent();
      }
      break;
    case 'ListPage':
      var listItems = page.getListItems();
      
      var listUpdated = false;
      for (var j = 0; j < listItems.length; j++) {
        if (chooseWhen_(listItems[j], when) > new Date(last_check).getTime()) {
          listUpdated = true;
        }
      }

      //désolé Nicolas ce n'est pas très clair... Édouard
      var col = page.getColumns();
      var item = page.getListItems();
      var title = [];
      var arr = [];
      for(var i in col){
        title.push(col[i].getName());
      }
      arr.push(title);
      for (var i in item){
        var arrite = [];
        //var itemo = item[i].getName();
        for(var j in col){
          arrite.push(item[i].getValueByName(col[j].getName()));
          
        }
        arr.push(arrite);
        
      }
      var ligne = '';
      var cologne = '';
      for (var i=0;i<3;i++){
      
        for(var j in arr[i]){
        
          ligne+='<td style="padding-top:10px;padding-bottom:10px;border-bottom: 1px solid rgba(0,0,0,0.08);">'+arr[i][j]+'</td>';
          
        }
        cologne += '<tr>'+ligne+'</tr>';
        ligne='';
      }
      var tableau = '<table style="width:100%;border-spacing:inherit;">'+cologne+'</table>'
      
      if (listUpdated) {
      var content = HtmlService.createTemplateFromFile('content.html');
      content.TextContent = tableau;//page.getHtmlContent();
        content.Title = page.getTitle();
        content.PageType = "List Page";
        content.Url = page.getUrl();
        updates += content.evaluate().getContent();
    }
    break;
  }
  //Logger.log(updates);
  return updates;
}

function createNewsletter_(page, updates, last_check, when, lang) {
  var type = page.getPageType();
  //Logger.log(type);
  switch (type.toString()) {
    case 'AnnouncementsPage':
      var newsfeed = page.getAnnouncements();
      for (var j = 0; j < newsfeed.length; j++) {
        if (chooseWhen_(newsfeed[j], when) > new Date(last_check).getTime()) {
          var updates = '';
          
          
          var content = HtmlService.createTemplateFromFile('content.html');
          content.TextContent = newsfeed[j].getHtmlContent();
          content.PageType = "Announcement";
          content.Title = newsfeed[j].getTitle();
          content.PageTitle = page.getTitle();
          content.Url = newsfeed[j].getUrl();
          updates += content.evaluate().getContent();
          sendMessage_(updates, lang, newsfeed[j].getTitle());
        }
      }
      break;
  }
}

function chooseWhen_(item, when) {
  var time = 0;
  if (when == 'Item is updated') {
    time = item.getLastUpdated().getTime();
  }
  else {
    time = item.getDatePublished().getTime();
  }
  return time;
}


function test(){
  
  //Logger.log(ss.getFormUrl());
  var footer = HtmlService.createTemplateFromFile('footer.html');
  
  footer.formURL = ss.getFormUrl();
  
  //Logger.log(footer.evaluate().getContent());
  var content = HtmlService.createTemplateFromFile('content.html');
  content.TextContent = newsfeed[j].getTextContent().substr(0, 280) + "...";
  content.Title = newsfeed[j].getTitle();
  content.PageTitle = page.getTitle();
  content.Url = newsfeed[j].getUrl();
  (content.evaluate().getContent());
  
}
