var aws = require('aws-sdk');
var ses = new aws.SES({
   accessKeyId: 'AWS USER ACCESS KEY ID',
   secretAccesskey: 'AWS USER SECRET ACCESS KEY',
   region: 'us-east-1'
});

exports.handler = function(event, context) {
    console.log("Incoming: ", event);

    /* Update these to work in your environment */
    var rgAssitantEmail = 'messageFrom@example.com';
    var rgAdminEmail = 'adminEmail@example.com';    //This address is used to log basic errors.

    var rgPayloadId = event.rgPayloadId;
    var rgPayloadResourceName = event.rgPayloadResourceName;
    var rgEmail = event.rgPayloadResourceEmail;
    var rgPayloadProjectName = event.rgPayloadProjectName;
    var rgPayloadClient = event.rgPayloadClient;
    var rgPayloadNotes = event.rgPayloadNotes;
    var rgProjectManager = event.rgPayloadBookerName;
    var rgPayloadStartDate = event.rgPayloadStartDate;
    var rgPayloadEndDate = event.rgPayloadEndDate;
    var rgPayloadType = event.rgPayloadType;
    var rgPayloadAction = event.rgPayloadAction;
    var params = {};
    var message = [];

    if (rgPayloadResourceName === undefined) { rgPayloadResourceName = 'Unknown' }
    if (rgPayloadProjectName === undefined) { rgPayloadProjectName = 'Unknown' }
    if (rgPayloadClient === undefined) { rgPayloadClient = 'Unknown' }
    if (rgPayloadNotes === undefined) { rgPayloadNotes = 'Unknown' }
    if (rgProjectManager === undefined) { rgProjectManager = 'Unknown' }

    var errorMessage = {};
    //If we are missing any of these values, we should fail with a message indicating the reason.
    if ((rgEmail === undefined) || (rgPayloadStartDate === undefined) || (rgPayloadEndDate === undefined) || (rgPayloadId === undefined)) {
        errorMessage.error = true;
        errorMessage.event = event;
        errorMessage.message = 'We were missing some critical information';
    }

    //Convert the Resource Guru data into something iCalendar can understand
    function convertResourceGuruDate(rgStartDate, rgEndDate) {
      rgStartDate = (rgStartDate).replace(/-/gi,'');
      rgEndDate =  (rgEndDate).replace(/-/gi,'');
      var result = {
        startTime:  rgStartDate + "T080000",
        endTime: rgEndDate + "T170000"
      };
      return result;
    }

    if (errorMessage.error !== true) {

        /* Test Values. Use these values to test the function directly from Lambda
        var rgPayloadId = '4917672';
        var rgPayloadResourceName = 'Bob';
        var rgEmail = 'bob@bobsburgers.com';
        var rgPayloadProjectName = 'Make me a burger bob!';
        var rgPayloadClient = 'Prestige World Wide';
        var rgPayloadNotes = 'Here are some notes and address of the location';
        var rgProjectManager = 'PM Dawn';
        var rgPayloadStartDate = '2016-06-08';
        var rgPayloadEndDate = '2016-06-12';
        var rgPayloadType = 'booking';
        var rgPayloadAction = 'update';   //Resource Guru sends create/update/delete
        */

        var inviteStatus = 'CONFIRMED';
        var emailSubject = 'New Project Assignment: [' + rgPayloadClient + ']';

        //Based on the Resource Guru action, we will create/update/cancel the invite accordingly.
        if (rgPayloadAction == 'delete') {
            inviteStatus = 'CANCELLED';
            emailSubject = 'CANCELLED: Project Assignment: [' + rgPayloadClient + ']';
        } else if (rgPayloadAction == 'update') {
            emailSubject = 'UPDATE: Project Assignment: [' + rgPayloadClient + ']';
        }

        icsDateRange = convertResourceGuruDate(rgPayloadStartDate, rgPayloadEndDate);

        //Format the HTML content for the Email
        var htmlDetails = [
          "<html><head></head><body><img src=http://www.siriuscom.com/wp-content/uploads/2016/02/SiriusLogo240x79.png>",
          "<h1>Project Details</h1><hr>",
          "<br><b>Project Name: </b>" + rgPayloadProjectName,
          "<br><b>Engineer: </b>" + rgPayloadResourceName,
          "<br><b>Customer: </b>" + rgPayloadClient,
          "<br><b>Project Start Date: </b>" + rgPayloadStartDate,
          "<br><b>Project End Date: </b>" + rgPayloadEndDate,
          "<br><b>Project Manager: </b>" + rgProjectManager,
          "<br><b>Notes: </b><br>" + rgPayloadNotes,
          "<hr>",
          "<a href=https://siriuscomputersolutions.resourceguruapp.com/schedule#" + rgPayloadStartDate + ">Open in ResourceGuru</a>"
        ];
        htmlDetails = htmlDetails.join('');

        //Format the iCalendar invite details
        var details = [
          "Project Details ----------------------------------------------------------------------------------",
          "Project Name: " + rgPayloadProjectName,
          "Engineer: " + rgPayloadResourceName,
          "Customer: " + rgPayloadClient,
          "Project Start Date: " + rgPayloadStartDate,
          "Project End Date: " + rgPayloadEndDate,
          "Project Manager: " + rgProjectManager,
          "Note: " + rgPayloadNotes,
          "--------------------------------------------------------------------------------------------------",
          "Open in ResourceGuru: https://siriuscomputersolutions.resourceguruapp.com/schedule#" + rgPayloadStartDate
        ];
        details = details.join('\r\n');

        //Create the actual iCalendar Invite. May want to look into using the ical.js in the future
        var invite = [
        'BEGIN:VCALENDAR',
        'METHOD:REQUEST',
        'PRODID:Microsoft Exchange Server 2010',
        'VERSION:2.0',
        'BEGIN:VTIMEZONE',
        'TZID:Eastern Standard Time',
        'BEGIN:STANDARD',
        'DTSTART:16010101T020000',
        'TZOFFSETFROM:-0400',
        'TZOFFSETTO:-0500',
        'RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=1SU;BYMONTH=11',
        'END:STANDARD',
        'BEGIN:DAYLIGHT',
        'DTSTART:16010101T020000',
        'TZOFFSETFROM:-0500',
        'TZOFFSETTO:-0400',
        'RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=2SU;BYMONTH=3',
        'END:DAYLIGHT',
        'END:VTIMEZONE',
        'BEGIN:VEVENT',
        'ORGANIZER;CN=RG Assistant:MAILTO:' + rgAssitantEmail,
        'ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=' + rgPayloadResourceName + ':MAILTO:'+ rgEmail,
        'DESCRIPTION:\n' + details,
        'UID:' + rgPayloadId,                       //This is the unique ID that will allow us to update an existing invite.
        'SUMMARY;LANGUAGE=en-US:' + emailSubject,
        'DTSTART;TZID=Eastern Standard Time:' + icsDateRange.startTime,
        'DTEND;TZID=Eastern Standard Time:' + icsDateRange.endTime,
        'CLASS:PUBLIC',
        'PRIORITY:5',
        'DTSTAMP:20160606T230601Z',
        'TRANSP:OPAQUE',
        'STATUS:' + inviteStatus,
        'SEQUENCE:0',
        'X-MICROSOFT-CDO-APPT-SEQUENCE:0',
        'X-MICROSOFT-CDO-OWNERAPPTID:2114325233',
        'X-MICROSOFT-CDO-BUSYSTATUS:TENTATIVE',
        'X-MICROSOFT-CDO-INTENDEDSTATUS:BUSY',
        'X-MICROSOFT-CDO-ALLDAYEVENT:FALSE',
        'X-MICROSOFT-CDO-IMPORTANCE:1',
        'X-MICROSOFT-CDO-INSTTYPE:0',
        'X-MICROSOFT-DISALLOW-COUNTER:FALSE',
        'BEGIN:VALARM',
        'DESCRIPTION:REMINDER',
        'TRIGGER;RELATED=START:-PT15M',
        'ACTION:DISPLAY',
        'END:VALARM',
        'END:VEVENT',
        'END:VCALENDAR',
        ];

        //Create the RAW HTML required to embed the calendar invite.
        //This creates a multi part MIME message with the HTML content and the iCalendar attachment
        message = [
            "To: \"" + rgPayloadResourceName + "\" <" + rgEmail + ">",
            "Subject: " + emailSubject,
            "MIME-Version: 1.0",
            "Content-Type: multipart/mixed; boundary=\"XXXXboundary text\"\n",
            "This is a multipart message in MIME format.\n",
            "--XXXXboundary text",
            "Content-type: text/html; charset=iso-8859-1\n",
            htmlDetails + '\n',
            "--XXXXboundary text",
            "Content-Type: text/calendar; charset=\"utf-8\"; method=REQUEST",
            "Content-Transfer-Encoding: base64;\n",
            (new Buffer(invite.join("\r\n")).toString('base64')),
            "--XXXXboundary text--"
            ];

        message = message.join('\n');

        params = {
          RawMessage: {
            Data: message
          },
          Destinations: [ rgEmail.toLowerCase() ],
          Source: rgAssitantEmail
        };
    } else {
        //Something wasn't sent correctly from Resource Guru. Lets notify the admin email.
        console.log('There was an error, we are emailing the details to the administrator');
        message = [
            "Subject: Calendar",
            "MIME-Version: 1.0",
            "Content-Type: text/plain;\n",
            'We ran into an issue processing the following request:',
            'Error Message: ' + errorMessage.message,
            'Event Data: ' + JSON.stringify(errorMessage.event)
            ];
        message = message.join('\n');
        params = {
          RawMessage: {
            Data: message
          },
          Destinations: [ rgAdminEmail ],
          Source: rgAdminEmail
        };
    }
    var email = ses.sendRawEmail(params, function(err, data) {
      if (err) {
        console.log(err);
      } else {
        console.log('Message sent!');
        console.log(data);
      }
      context.succeed(event);
    });
};
