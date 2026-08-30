/*
Copyright 2025 Andrey Gonzalez (for McGill Students Running Club)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/


/**
 * Create the PNG image of the run route from Strava activity from its polyline
 * data and saves the public image URL in the row.
 * 
 * @param {number} row  GSheet row to target. Defaults to last row.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Sep 17, 2025
 * @update  Oct 27, 2025
 */
function createMapForRow(row = getValidLastRow_(LOG_SHEET)){
  const activity = getExistingStravaActivity_(row);
  if(!activity) throw Error("No activity detected");

  const timestamp = getTimestampInRow_(row);
  activity.map = extractPolyline(activity.map);
  if(!activity.map) throw Error("No polyline detected");

  // Store created map and print url
  const updated = createAndAppendMap_(timestamp, activity);
  setMapUrlInRow(row, updated?.mapUrl);
  Logger.log(updated?.mapUrl);

  /** Since `JSON.parse` will not work with values stored in
   *  GSheet col `MAP_POLYLINE`, we use a regex to extract it instead */
  function extractPolyline(str) {
    // Case 1: polyline is delimited by comma, i.e. another property to its right
    const regexUntilComma = /summary_polyline=(.*)\,/;

    // Case 2: polyline is delimited by end closing bracket `}`
    const regexUntilBracket = /summary_polyline=(.*)\}/;

    // Return polyline for both cases without delimeter (comma or bracket)
    const match = regexUntilComma.exec(str) ?? regexUntilBracket.exec(str);
    return match ? {'summary_polyline': match[1]} : null;
  }

  /** Only set mapUrl if cell empty */
  function setMapUrlInRow(row, mapUrl) {
    const logSheet = GET_LOG_SHEET();
    const targetCell = logSheet.getRange(row, LOG_COL.MAP_URL);
    if (targetCell.getValue() == "") targetCell.setValue(mapUrl);
  }
}


/**
 * Gets the polyline data from given activity and calls helper function to create a
 * PNG image for it, then adds the image URL to the activity.
 * 
 * Previous iterations of map creation include `MAP.newStaticMap()`, embedding GDrive download url
 * in email (access restricted after some time), and adding map as inline image (email becomes too heavy).
 * 
 * @param {Object} activity  Strava activity with "map" key containing polyline data
 * @param {Date} timestamp  Recorded timestamp of event.
 * @param {string} fileName  Name to save map with
 * @returns {Object}  Strava activity with appended map url under the "mapUrl" key (or '' if unsuccessful)
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Mar 27, 2025
 * @update  Apr 4, 2025
 */
function createMapForActivity_(activity, timestamp) {
  // Extract polyline and save headrun route as map
  const polyline = activity['map']['polyline'] ?? activity['map']['summary_polyline'];

  if (polyline) {
    // Create file name for the map image
    const formatName = (str) => str.replace(/\s+/g, '-') || "";
    const formatTimestamp = (ts) => Utilities.formatDate(ts, TIMEZONE, "EEE-d-MMM-yyyy-k-mm-ss");
    const fileName = `headrun-${ formatName(activity?.name) }map-${ formatTimestamp(timestamp)}.png`;

    const response = convertPolylineToMap_(polyline, fileName, "580x420").getHeaders();

    // Get file by id or name, then set permission to allow downloading
    const file = response['file_id'] ? getFileById_(response['file_id']) : getFileByName_(fileName);
    //file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);

    const mapBlob = file.getBlob();
    // Upload image to Google Cloud Storage and get sharing link
    activity['mapUrl'] = uploadImageToCloudStorageBucket_(mapBlob, filename);
  }
  else {
    // If polyline not found, don't create image
    activity['mapUrl'] = '';
  }

  return activity;
}


/**
 * Save polyline as image using Google Static Map API and Make.com automation.
 * 
 * @param {string} polyline  Encoded Google Map polyline string.
 * @param {string} name  Name for map.
 * @param {string} imgSize  Size of map image, e.g "400x300"
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 27, 2025
 * @update  Apr 4, 2025
 */
function convertPolylineToMap_(polyline, name, imgSize) {
  // Construct postUrl for Make webhook
  const propertyStore = PropertiesService.getScriptProperties();
  const apiKey = propertyStore.getProperty(SCRIPT_PROPERTY_KEYS.googleMapAPI); // Replace with your API Key

  const googleCloudMapId = '30aa339d7a505f10aca29579' //'bfeadd271a2b0a58';  //'2ff6c54f4dd84b16';
  const pathColor = '0xEB4E3D';

  const queryObj = {
    size: imgSize,
    map_id: googleCloudMapId,
    key: apiKey,
    //path: `color:${pathColor}` + '|' + `enc:${polyline}`,
    path: `enc:${polyline}`,
  }

  const postUrl = MAPS_BASE_URL + queryObjToString_(queryObj);

  // Call Make webhook 
  const webhookUrl = "https://hook.us1.make.com/8obb3hb6bzwgi7s4nyi8yfghb3kxsksc";
  const payload = JSON.stringify({ url: postUrl, name: name });

  const options = {
    method: "post",
    contentType: "application/json",
    payload: payload
  };

  const response = UrlFetchApp.fetch(webhookUrl, options);
  logAsPL_(`Make Webhook Response: ${response.getContentText()}`, convertPolylineToMap_.name);

  return response;
}


function testCloudUpload() {
  const fileId = "1XhuP7peNPTWCnNhs3MGC-IzGNi0Yu-7X";   //"14csoxHqwHnnN7KFhsEgi5o55x1Sajvbh";
  const blob = DriveApp.getFileById(fileId).getBlob();

  const time = Utilities.formatDate(new Date(), TIMEZONE, "EEE-d-MMM-yyyy-k-mm-ss");
  const imageName = "headrun-map-" + time + '.png';

  console.log(imageName);
  return;
  
  try {
    const imageUrl = uploadImageToCloudStorageBucket_(blob, imageName);
    Logger.log("Uploaded image URL: " + imageUrl);
  } catch (e) {
    Logger.log("Error during upload: " + e);
  }
}


/**
 * Uploads given image blob to cloud storage under the provided name,
 * and returns the resulting URL.
 * 
 * @param {string} imageBlob  Image data as a blob.
 * @param {string} imageName  Name to save the image under.
 * 
 * @return {string|null}  URL of the image in cloud storage, or null if error occurred.
 */
function uploadImageToCloudStorageBucket_(imageBlob, imageName) {
  // Get service key to access cloud storage
  const store = PropertiesService.getScriptProperties();
  const propertyName = SCRIPT_PROPERTY_KEYS.googleCloudKey;
  const serviceAccountKey = JSON.parse(store.getProperty(propertyName));

  // Authenticate using the Service Account
  const token = getServiceAccountAccessToken_(serviceAccountKey);

  // Construct the upload URL
  const uploadUrl = `${BASE_UPLOAD_URL}/${STORAGE_BUCKET_NAME}/o?uploadType=media&name=${imageName}`;

  // Set up the options for the UrlFetchApp request
  const options = {
    'method': 'post',
    'contentType': 'image/jpeg',
    'payload': imageBlob.getBytes(),
    'headers': {
      'Authorization': 'Bearer ' + token
    },
    'muteHttpExceptions': true // Allows you to see error responses
  };

  // Make the upload request
  const response = UrlFetchApp.fetch(uploadUrl, options);
  Logger.log(response.getContentText()); // Log the response for debugging

  // Check for errors
  if (response.getResponseCode() >= 400) {
    Logger.log('Error uploading image,');
    return null;
  }

  logAsPL_('Image uploaded successfully!', uploadImageToCloudStorageBucket_.name);
  return `https://storage.googleapis.com/${STORAGE_BUCKET_NAME}/${imageName}`; // Return the public URL
}


/**
 * Helper function to get an access token using the service account key
 * 
 * @param {Object}  Key for service account for Google Cloud.
 * 
 * @return {string}  Access token to cloud storage.
 */
function getServiceAccountAccessToken_(key) {
  var jwt = Utilities.base64EncodeWebSafe(JSON.stringify({
    "alg": "RS256",
    "typ": "JWT"
  }));

  var now = Math.floor(Date.now() / 1000);
  var claim = Utilities.base64EncodeWebSafe(JSON.stringify({
    "iss": key.client_email,
    "scope": "https://www.googleapis.com/auth/devstorage.full_control",
    "aud": "https://oauth2.googleapis.com/token",
    "exp": now + 3600,
    "iat": now
  }));

  var signature = Utilities.computeRsaSha256Signature(jwt + "." + claim, key.private_key);
  signature = Utilities.base64EncodeWebSafe(signature);

  var assertion = jwt + "." + claim + "." + signature;

  var payload = {
    "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
    "assertion": assertion
  };

  var options = {
    "method": "post",
    "payload": payload
  };

  var response = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", options);
  var json = JSON.parse(response.getContentText());
  return json.access_token;
}