
function doPost(e) {

  if(typeof e !== 'undefined')
  var role = e.parameter.role;
  var workplace = e.parameter.workplace;
  var team = e.parameter.team;
  var name = e.parameter.name_001.replace(/^\s+|\s+$/g, '');
  var email = e.parameter.email_001.replace(/^\s+|\s+$/g, '');
  var phone = e.parameter.phone_001.replace(/^\s+|\s+$/g, '');
  var leavingdate = e.parameter.leavingdate;
  var startingdate = e.parameter.startingdate;
  var addrline1 = e.parameter.addrline1_001.replace(/^\s+|\s+$/g, '');
  var addrline2 = e.parameter.addrline2_001.replace(/^\s+|\s+$/g, '');
  var town = e.parameter.town_001.replace(/^\s+|\s+$/g, '');
  var county = e.parameter.county_001.replace(/^\s+|\s+$/g, '');
  var postcode = e.parameter.postcode_001.replace(/^\s+|\s+$/g, '');
  var reasonforleaving = e.parameter.reasonforleaving;
  var outstandingexpenses = e.parameter.outstandingexpenses;
  var charityequipment = e.parameter.charityequipment;
  var days_in_post = e.parameter.days_in_post;

  sendLetter(role,workplace,team,name,email,phone,startingdate,leavingdate,addrline1,addrline2,town,county,postcode,reasonforleaving,outstandingexpenses,charityequipment,days_in_post);

  var html = HtmlService.createTemplateFromFile('Thankyou');

  return html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
 .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}

function getScriptUrl() {
 var url = ScriptApp.getService().getUrl();
 return url;
}

function doGet(e) {
  var currentuser = Session.getActiveUser().getEmail();
  var authusers = [someUser@someEmail.org.uk];//users removed for anonimity
  var isauthorised = authusers.indexOf(currentuser);

  if ( isauthorised >= 0) {
   return HtmlService.createHtmlOutputFromFile('Form');
  } else {
  var html = HtmlService.createTemplateFromFile('Error');
  html.currentuser= currentuser;
  return html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
 }
}

function ordinal_suffix_of(i) {
    var j = i % 10,
        k = i % 100;
    if (j == 1 && k != 11) {
        return "st";
    }
    if (j == 2 && k != 12) {
        return "nd";
    }
    if (j == 3 && k != 13) {
        return "rd";
    }
    return "th";
}

function getLongDateFormat(dateArray) {
  var ordinal;
  var day;
  if (dateArray[0].substring(0,1)  == "0") {
  ordinal =  ordinal_suffix_of(dateArray[0].substring(1))
  day = dateArray[0].substring(1);
} else {
  ordinal =  ordinal_suffix_of(dateArray[0])
  day = dateArray[0];
 }
 var monthname;
 switch(dateArray[1]) {

       case "01":
       monthname = "January";
       break;

       case "02":
       monthname = "February";
       break;

       case "03":
       monthname = "March";
       break;

       case "04":
       monthname = "April";
       break;

       case "05":
       monthname = "May";
       break;

       case "06":
       monthname = "June";
       break;

       case "07":
       monthname = "July";
       break;

       case "08":
       monthname = "August";
       break;

       case "09":
       monthname = "September";
       break;

       case "10":
       monthname = "October";
       break;

       case "11":
       monthname = "November";
       break;

       case "12":
       monthname = "December";
       break;
    }
    return day + " " + monthname + " " + dateArray[2];
}

function sendLetter(role,workplace,team,name,email,phone,startingdate,leavingdate,addrline1,addrline2,town,county,postcode,reasonforleaving,outstandingexpenses,charityequipment,days_in_post) {
  var roletype;
  switch(role) {

    case "Volunteer":
    roletype = "volunteering role";
    break;

    case "Staff":
    roletype = "employment";
    break;
}

var dstFolderId = DriveApp.getFolderById("1lVsjyYWp2VdvSqEF4lgfgMXF6kEK-FcN");

var startdateArray = startingdate.split("/");
var leavingdateArray = leavingdate.split("/");
var startdate = new Date(startdateArray[2] + "/" + startdateArray[1] + "/" + startdateArray[0]);
var leavingdate = new Date(leavingdateArray[2] + "/" + leavingdateArray[1] + "/" + leavingdateArray[0]);
var todayArray = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy").split("/");;
var leavingdateLongFormat = getLongDateFormat(leavingdateArray);
var todayLongFormat = getLongDateFormat(todayArray);

var address;
  if (addrline2.length >0) {
    if (county.length > 1) {
        address = addrline1 + '\r' + addrline2 + '\r' + town + '\r' + county  + '\r' + postcode;
        } else {
        address = addrline1 + '\r' + addrline2 + '\r' + town + '\r' + postcode;
        }
        } else {

    if (county.length > 1) {
        address = addrline1 + '\r' + town + '\r' + county  + '\r' + postcode;
        } else {
        address = addrline1 + "\r" + town + "\r" + postcode;
        }
    }
