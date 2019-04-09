/***********************************************************************************/
/*Script de notification des modifications d'une Google sheet dans un bot Telegram */
/*Dans Google Script app, ajouter un déclencheur basé sur la feuille de calcul,    */
/*lors d'une modification pour exécuter la fonction sendText                       */
/*Licence : cc0                                                                    */
/***********************************************************************************/

/*déclarer des variables*/
var token = ""; // id du bot Telegram fourni par botFather
var telegramUrl = "https://api.telegram.org/bot" + token; // url de l'api de Telegram + variable de l'id du bot 
var webAppUrl = ""; // url du script une fois publié
var ssId = ""; // id de la feuille de calcul à utiliser
var range = ""; // Plage surveillée par le script
var isNotifActive ; // Déclarer une variable pour l'état des notifications

// Ajout du menu bot dans la feuille de calcul
function onOpen() {
 isNotifActive = PropertiesService.getScriptProperties().setProperty('telegram_bot_active', 'activées');
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var Sousmenu = [];
   Sousmenu.push({name: "Afficher l'état du bot", functionName: "Message"}); 
   Sousmenu.push({name: "Activer/Désactiver les notifications", functionName: "switchNotif"}); 
   Sousmenu.push({name: "À propos", functionName: "Message2"}); 
  
   ss.addMenu("Notifications Telegram", Sousmenu);
}

function Message() {
  var isNotifActive = PropertiesService.getScriptProperties().getProperty('telegram_bot_active');
  Browser.msgBox ('Notifications vers Telegram '+ isNotifActive +".");
}

function Message2() {
  Browser.msgBox ("********** Script de notification des modifications d'une Google sheet dans Telegram **********\\n\\nChaque modification de la Google sheet est envoyée sur le bot Telegram.\\n\\nVous pouvez désactiver le bot lors de modifications de masse. Mais pensez à le réactiver par la suite !\\n\\nLicence : cc0\\n2019");
}

function switchNotif() {
    var isNotifActive = PropertiesService.getScriptProperties().getProperty('telegram_bot_active');
  if('activées' === isNotifActive) {
    var isNotifActive = PropertiesService.getScriptProperties().setProperty('telegram_bot_active', 'désactivées');
  } else {
    var isNotifActive = PropertiesService.getScriptProperties().setProperty('telegram_bot_active', 'activées');
  }
  var isNotifActive = PropertiesService.getScriptProperties().getProperty('telegram_bot_active');
  Browser.msgBox ('Les notifications vers Telegram sont ' + isNotifActive +".");
}

/*Se connecter au webhook Telegram*/
function setWebhook () {
  var url = telegramUrl + "/setWebhook?url=" + webAppUrl;
  var response = UrlFetchApp.fetch(url);
  Logger.log(response.getContentText());
}

/*Envoyer une notification*/
function sendText() {
    var isNotifActive = PropertiesService.getScriptProperties().getProperty('telegram_bot_active');
  if ('activées' === isNotifActive) {
    var id = -1001459431946 ; // si notifications désactivées, envoyer la notif dans le bot
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SpreadsheetApp.getActiveSheet();
    var editRange = sheet.getActiveRange();
    var editRow = editRange.getRow();
    var editCol = editRange.getColumn();
    var testrange = sheet.getRange(range);
    var rangeRowStart = testrange.getRow();
    var rangeRowEnd = rangeRowStart + testrange.getHeight()-1;
    var rangeColStart = testrange.getColumn();
    var rangeColEnd = rangeColStart + testrange.getWidth()-1;
    if (editRow >= rangeRowStart && editRow <= rangeRowEnd 
        && editCol >= rangeColStart && editCol <= rangeColEnd)
    { 
      var sheet = SpreadsheetApp.getActiveSheet();
      var body = sheet.getName() + ' a été mis à jour : «' + editRange.getValue() + '».';
      var url = telegramUrl + "/sendMessage?chat_id=" + id + "&text=" + body ;
      var response = UrlFetchApp.fetch(url);
    }
  }
  else {
    }
  }