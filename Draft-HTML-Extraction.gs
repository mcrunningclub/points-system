
/**
 * User function to execute `generateHtmlFromDraft_`.
 * 
 * Must updated subject line as needed.
 * 
 */

function saveDraftAsHtml() {
  const subjectLine = 'Here\'s your post-run report! 🙌';
  generateHtmlFromDraft_(subjectLine);
}


/**
 * Generate html version of email found in draft using its subject line.
 * 
 * @param {string} subjectLine  Subject line of target draft.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 */

function generateHtmlFromDraft_(subjectLine) {
  // Prevent email sent by wrong user
  if (getCurrentUserEmail_() != 'mcrunningclub@ssmu.ca') {
    throw new Error ('Please switch to the McRUN Google Account before sending emails');
  }

  const datetime = Utilities.formatDate(new Date(), TIMEZONE, 'MMM-dd\'T\'hh.mm');
  const baseName = subjectLine.replace(/ /g, '-').toLowerCase();

  // Create filename for html file
  const fileName = `${baseName}-html-${datetime}`;

  // Find template in drafts and get email objects
  const emailTemplate = getGmailTemplateFromDrafts(subjectLine);
  const msgObj = fillInTemplateFromObject_(emailTemplate.message, {});

  // Save html file in drive
  DriveApp.createFile(fileName, msgObj.html);
}





function testRuntime() {
  const recipient = 'andrey.gonzalez@mail.mcgill.ca';
  const startTime = new Date().getTime();

  // Runtime if using DriveApp call : 1200ms
  // If caching images in script properties once: 550 ms
  sendSamosaEmailFromHTML(recipient, 'Test 5 samosa sale');

  //sendSamosaEmail();    // around 3000ms
  
  // Record the end time
  const endTime = new Date().getTime();
  
  // Calculate the runtime in milliseconds
  const runtime = endTime - startTime;
  
  // Log the runtime
  Logger.log(`Function runtime: ${runtime} ms`);
}



/**
 * Sends email using member information.
 * 
 * @author  Martin Hawksey (2022)
 * @update  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * 
 * @param {{key:value<string>}} memberInformation  Information to populate email draft
 * @return {{message:string, isError:bool}}  Status of sending email.
*/
function sendEmail_(memberInformation) {
  // Gets the draft Gmail message to use as a template
  const subjectLine = DRAFT_SUBJECT_LINE;
  const emailTemplate = getGmailTemplateFromDrafts(subjectLine);

  try {
    const memberEmail = memberInformation['EMAIL'];
    const msgObj = fillInTemplateFromObject_(emailTemplate.message, memberInformation);

    //DriveApp.createFile('TestFile3b', msgObj.html);

    MailApp.sendEmail(
      'andrey.gonzalez@mail.mcgill.ca',
      msgObj.subject,
      msgObj.text,
      {
        htmlBody: msgObj.html,
        from: 'mcrunningclub@ssmu.ca',
        name: 'McGill Students Running Club',
        replyTo: 'mcrunningclub@ssmu.ca',
        attachments: emailTemplate.attachments,
        inlineImages: emailTemplate.inlineImages
      }
    );

  } catch(e) {
    // Log and return error
    console.log(`(sendEmail) ${e.message}`);
    throw new Error(e);
  }
  // Return success message
  return {message: 'Sent!', isError : false};
}


/**
 * Get a Gmail draft message by matching the subject line.
 * 
 * @author  Martin Hawksey (2022)
 * @update  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * 
 * @param {string} subjectLine to search for draft message
 * @return {object} containing the subject, plain and html message body and attachments
*/

function getGmailTemplateFromDrafts(subjectLine = DRAFT_SUBJECT_LINE){
  // Verify if McRUN draft to search
  if (Session.getActiveUser().getEmail() != MCRUN_EMAIL) {
    return Logger.log('Change Gmail Account');
  }

  try {
    // Get the target draft, then message object
    const drafts = GmailApp.getDrafts();
    const filteredDrafts = drafts.filter(subjectFilter_(subjectLine));

    if (filteredDrafts.length > 1) {
      throw new Error (`Too many drafts with subject line '${subjectLine}. Please review.`);
    }

    const draft = filteredDrafts[0];
    const msg = draft.getMessage();

    // Handles inline images and attachments so they can be included in the merge
    // Based on https://stackoverflow.com/a/65813881/1027723
    // Gets all attachments and inline image attachments
    //const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
    const htmlBody = msg.getBody();
    //DriveApp.createFile('testFile3a', htmlBody);

    const allInlineImages = draft.getMessage().getAttachments({
      includeInlineImages: true,
      includeAttachments:false
    });

    //Initiate the allInlineImages object
    var inlineImagesObj = {}
    //Regexp to search for all string positions 
    var regexp = RegExp('img data-surl=\"cid:', 'g');
    var indices = htmlBody.matchAll(regexp)

    //Iterate through all matches
    var i = 0;
    for (const match of indices){
      //Get the start position of the CID
      var thisPos = match.index + 19
      //Get the CID
      var thisId = htmlBody.substring(thisPos, thisPos + 15).replace(/"/,"").replace(/\s.*/g, "")
      //Add to object
      inlineImagesObj[thisId] = allInlineImages[i];
      i++
    }

    /* // Creates an inline image object with the image name as key 
    // (can't rely on image index as array based on insert order)
    const imgObj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

    for(const [key, value] of Object.entries(imgObj)) {
      console.log(key + " has value... " + value);
    }

    // Regexp searches for all img string positions with cid
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];

    // Initiates the allInlineImages object
    const inlineImagesObj = {};
    // built an inlineImagesObj from inline image matches
    matches.forEach(match => {
      console.log(match);
      inlineImagesObj[match[1]] = imgObj[match[2]];
    })

    console.log(matches); */
       

    const draftObj = {
      message: {
        subject: subjectLine, 
        text: msg.getPlainBody(), 
        html:htmlBody
      }, 
      //attachments: attachments, 
      inlineImages: inlineImagesObj 
    };

    console.log(inlineImagesObj);
    for(const [key, value] of Object.entries(inlineImagesObj)) {
      console.log(key + " has value... " + value);
    }


    return draftObj;
     
  } catch(e) {
    throw new Error("Oops - can't create template from draft. Error: " + e.message);
  }
}


/**
 * Filter draft objects with the matching subject line message by matching the subject line.
 * 
 * @author  Martin Hawksey (2022)
 * @update  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
 * 
 * @param {string} subjectLine to search for draft message
 * @return {object} GmailDraft object
*/

function subjectFilter_(subjectLine){
  return function(element) {
    if (element.getMessage().getSubject() === subjectLine) {
      return element;
    }
  }
}