function doPost(e) {

    if (typeof e !== 'undefined')
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

    sendLetter(role, workplace, team, name, email, phone, startingdate, leavingdate, addrline1, addrline2, town, county, postcode, reasonforleaving, outstandingexpenses, charityequipment, days_in_post);

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
    var authusers = [someUser @someEmail.org.uk]; //users removed for anonimity
    var isauthorised = authusers.indexOf(currentuser);

    if (isauthorised >= 0) {
        return HtmlService.createHtmlOutputFromFile('Form');
    } else {
        var html = HtmlService.createTemplateFromFile('Error');
        html.currentuser = currentuser;
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
    if (dateArray[0].substring(0, 1) == "0") {
        ordinal = ordinal_suffix_of(dateArray[0].substring(1))
        day = dateArray[0].substring(1);
    } else {
        ordinal = ordinal_suffix_of(dateArray[0])
        day = dateArray[0];
    }
    var monthname;
    switch (dateArray[1]) {

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

function sendLetter(role, workplace, team, name, email, phone, startingdate, leavingdate, addrline1, addrline2, town, county, postcode, reasonforleaving, outstandingexpenses, charityequipment, days_in_post) {
    var roletype;
    switch (role) {

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
    if (addrline2.length > 0) {
        if (county.length > 1) {
            address = addrline1 + '\r' + addrline2 + '\r' + town + '\r' + county + '\r' + postcode;
        } else {
            address = addrline1 + '\r' + addrline2 + '\r' + town + '\r' + postcode;
        }
    } else {

        if (county.length > 1) {
            address = addrline1 + '\r' + town + '\r' + county + '\r' + postcode;
        } else {
            address = addrline1 + "\r" + town + "\r" + postcode;
        }
    }

    if (days_in_post >= 180) {

        var docid = DriveApp.getFileById("1n3zWEttQA8LiwADEBFH-a5oHzU6ycGBhz0KM3l_d2fs").makeCopy("Leavers_Letter_" + Utilities.formatDate(new Date(), "GMT+1", "dd-MMM-yyyy") + "_" + name, dstFolderId).getId()
        var doc = DocumentApp.openById(docid);

        var body = doc.getActiveSection();
        body.replaceText("%fullname%", name);
        body.replaceText("%addrline1%", address);
        body.replaceText("%town%", town);
        body.replaceText("%county%", county);
        body.replaceText("%postcode%", postcode);
        //body.replaceText("%name%", name.split(" ")[0]);
        body.replaceText("%leavingdate%", leavingdateLongFormat);
        body.replaceText("%today%", todayLongFormat);
        body.replaceText("%team%", team);
        body.replaceText("%area%", workplace);
        body.replaceText("%roletype%", roletype);

        doc.saveAndClose();

        GmailApp.sendEmail( /*CEO email hidden */ , 'Leaver Notification (LETTER)', 'Please see the attached letter.', {
            bcc: /*manager email hidden */ ,
            attachments: [doc.getAs(MimeType.PDF)],
            name: 'Leaver Notification Letter'
        });
        //GmailApp.sendEmail('email hidden', 'Leaver Notification Letter - TEST PLEASE IGNORE', 'Please see the attached letter.', {attachments: [doc.getAs(MimeType.PDF)],name: 'Leaver Notification - TEST PLEASE IGNORE' });
    }

    var docid2 = DriveApp.getFileById("1dm7XshpRE1jsHhUkV45X61-eYA1IjW6ut2H6g1mxhzE").makeCopy("Leavers_Report_" + Utilities.formatDate(new Date(), "GMT+1", "dd-MMM-yyyy") + "_" + name, dstFolderId).getId();
    var doc2 = DocumentApp.openById(docid2);

    if (phone.length < 1) {
        phone = "Not recorded";
    }
    if (reasonforleaving.length < 1) {
        reasonforleaving = "Not recorded";
    }
    if (outstandingexpenses.length < 1) {
        outstandingexpenses = "None stated";
    }
    if (charityequipment.length < 1) {
        charityequipment = "None stated";
    }
    var body2 = doc2.getActiveSection();

    body2.replaceText("%fullname%", name);
    body2.replaceText("%addrline1%", address)
    body2.replaceText("%town%", town);
    body2.replaceText("%county%", county);
    body2.replaceText("%postcode%", postcode);
    body2.replaceText("%name%", name.split(" ")[0]);
    body2.replaceText("%startdate%", Utilities.formatDate(startdate, "GMT+1", "dd-MMM-yyyy"));
    body2.replaceText("%leavingdate%", Utilities.formatDate(leavingdate, "GMT+1", "dd-MMM-yyyy"));
    body2.replaceText("%team%", team);
    body2.replaceText("%email%", email);
    body2.replaceText("%phone%", phone);
    body2.replaceText("%role%", role);
    body2.replaceText("%workplace%", workplace);
    body2.replaceText("%reasonforleaving%", reasonforleaving);
    body2.replaceText("%outstandingexpenses%", outstandingexpenses);
    body2.replaceText("%charityequipment%", charityequipment);

    doc2.saveAndClose();

    switch (workplace) {
        case "Leyland":
            GmailApp.sendEmail( /*email hidden*/ , 'Leaver Notification (REPORT)', 'Please see the attached leaver report.', {
                cc: /*emails hidden*/ ,
                attachments: [doc2.getAs(MimeType.PDF)],
                name: 'Leaver Notification Report'
            });
            break;

        case "Chorley":
            GmailApp.sendEmail( /*email hidden*/ , 'Leaver Notification (REPORT)', 'Please see the attached leaver report.', {
                cc: /*emails hidden*/ ,
                attachments: [doc2.getAs(MimeType.PDF)],
                name: 'Leaver Notification Report'
            });

        case "West Lancashire":
            GmailApp.sendEmail( /*email hidden*/ , 'Leaver Notification (REPORT)', 'Please see the attached leaver report.', {
                cc: /*emails hidden*/ ,
                attachments: [doc2.getAs(MimeType.PDF)],
                name: 'Leaver Notification Report'
            });
            break;

        case "Wyre":
            GmailApp.sendEmail( /*email hidden*/ , 'Leaver Notification (REPORT)', 'Please see the attached leaver report.', {
                cc: /*emails hidden*/ ,
                attachments: [doc2.getAs(MimeType.PDF)],
                name: 'Leaver Notification Report'
            });
            //  GmailApp.sendEmail('email hidden', 'Leaver Notification (REPORT) - TEST PLEASE IGNORE', 'Please see the attached leaver report.', {attachments: [doc2.getAs(MimeType.PDF)],name: 'Leaver Notification Report - TEST PLEASE IGNORE' });
            break;

        case "Blackburn with Darwen":
            GmailApp.sendEmail( /*email hidden*/ , 'Leaver Notification (REPORT)', 'Please see the attached leaver report.', {
                cc: /*emails hidden*/ ,
                attachments: [doc2.getAs(MimeType.PDF)],
                name: 'Leaver Notification Report'
            });

            break;
    }
}
