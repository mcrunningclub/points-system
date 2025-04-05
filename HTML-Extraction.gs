/**
 * Extracts all literal tags from HTML file using regex.
 * 
 * Used to debug HTML file.
 * 
 * @author  [Andrey Gonzalez] (<andrey.gonzalez@mail.mcgill.ca>) & ChatGPT
 * 
 * @date Apr 2, 2025
 * @update Apr 3, 2025
 */

function extractTagsFromProjectFile() {
  // Get the content of email.html from the script project
  const filename = 'Points Email V2';
  var htmlContent = HtmlService.createHtmlOutputFromFile(filename).getContent();

  // Decode &lt; and &gt; back to < and >
  var fileContent = htmlContent.replace(/&lt;/g, "<").replace(/&gt;/g, ">");

  // Extract all occurrences of <?= TAG ?>
  var matches = fileContent.match(/<\?=\s*([A-Z_]+)\s*\?>/g);

  // Remove `<?=` and `?>` to extract just the tag names
  var tags = matches ? matches.map(match => match.replace(/<\?=\s*|\s*\?>/g, '')) : [];

  // Get unique tags
  const unique = tags.filter((item, i, ar) => ar.indexOf(item) === i);

  Logger.log("Extracted Tags: " + JSON.stringify(unique));
}


function emailWebApp() {
  const template = HtmlService.createTemplateFromFile('test');
  const filledTemplate = template.evaluate();
  
  MailApp.sendEmail(
    message = {
      to : 'andrey.gonzalez@mail.mcgill.ca',
      name: EMAIL_SENDER_NAME,
      subject: 'Email Web App Test',
      htmlBody: filledTemplate.getContent(),
    }
  );
}



function createInlineImage_(fileUrl, blobKey) {
  // Extract file id from url. 
  // For cases like .../file/d/FILE_ID/... and ...&id=FILE_ID...
  const match = fileUrl.match(/(?:\/file\/d\/|[?&]id=)([\w-]+)/);
  const fileId = match[1];

  // Fetch the image using file id and set name for reference
  // Create inline image object with assigned CID
  const mapFile = DriveApp.getFileById(fileId);
  return mapFile.getBlob().setName(blobKey);
}


/**
 * Generate html version of email found in draft using its subject line.
 * 
 * Must updated subject line as needed.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 */

function saveDraftAsHtml() {
  const subjectLine = 'Here\'s your post-run report! 🙌';
  generateHtmlFromDraft_(subjectLine);

  function generateHtmlFromDraft_(subjectLine) {
    const datetime = Utilities.formatDate(new Date(), TIMEZONE, 'MMM-dd\'T\'hh.mm');
    const baseName = subjectLine.replace(/ /g, '-').toLowerCase();

    // Create filename for html file
    const fileName = `${baseName}-${datetime}.html`;

    // Find template in drafts and get email objects
    const emailTemplate = getGmailTemplateFromDrafts(subjectLine);
    const msgObj = fillInTemplateFromObject_(emailTemplate.message, {});

    // Save html file in drive
    DriveApp.createFile(fileName, msgObj.html);
    Logger.log(`Created HTML file '${fileName}'`);
  }
}


function testRuntime() {
  const recipient = 'andrey.gonzalez@mail.mcgill.ca';
  const startTime = new Date().getTime();

  // Runtime if using DriveApp call : 
  //  - 1 recipient: 5878ms
  //  - 3 recipients : 8503 ms (forming blob for every recipient)
  //  - 3 recipients: 6307ms (generating blob once)

  sendStatsEmail();
  
  // Record the end time
  const endTime = new Date().getTime();
  
  // Calculate the runtime in milliseconds
  const runtime = endTime - startTime;
  
  // Log the runtime
  Logger.log(`Function runtime: ${runtime} ms`);
}


/**
 * Get a Gmail draft message by matching the subject line.
 * 
 * @author  Martin Hawksey (2022)
 * @author2  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
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
 * @author2  [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>) (2025)
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

