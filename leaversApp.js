
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
