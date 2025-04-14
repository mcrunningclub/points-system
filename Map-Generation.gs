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

const MAPS_FOLDER = 'run_maps';
const MAPS_BASE_URL = "https://maps.googleapis.com/maps/api/staticmap";


/**
 * Create a PNG image of run route to include in email from polyline.
 * 
 * Google Static Map API + Make Automations.
 * 
 * Previous iterations of map creation include `MAP.newStaticMap()`, embedding GDrive download url
 * in email (access restricted after some time), and adding map as inline image (email becomes too heavy).
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Mar 27, 2025
 * @update  April 4, 2025
 * 
 */

function createStravaMap_(activity, name) {
  // Extract polyline and save headrun route as map
  const polyline = activity['map']['polyline'] ?? activity['map']['summary_polyline'];

  if (polyline) {
    const response = saveMapForRun_(polyline, name).getHeaders();

    // Get file by id or name, then set permission to allow downloading
    const file = response['file_id'] ? getFileById_(response['file_id']) : getFileByName_(name);
    //file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);

    return file.getBlob();
  }

  return '';
}


/**
 * Save polyline as image using Google Map API and Make.com automation.
 * 
 * @param {string} polyline  Encoded Google Map polyline string.
 * @param {string} name  Name for map.
 * 
 * @author [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * @date  Mar 27, 2025
 * @update  Apr 4, 2025
 */

function saveMapForRun_(polyline, name) {
  const postUrl = buildPostUrl_(polyline, "580x420");
  return postToMakeWebhook_(postUrl, name);
}

/** Helper 1: Construct postUrl for Make webhook */
function buildPostUrl_(polyline, imgSize = "580x420") {
  const propertyStore = PropertiesService.getScriptProperties();
  const apiKey = propertyStore.getProperty(SCRIPT_PROPERTY_KEYS.googleMapAPI); // Replace with your API Key

  const googleCloudMapId = 'bfeadd271a2b0a58';  //'2ff6c54f4dd84b16';
  const pathColor = '0xEB4E3D';

  const queryObj = {
    size: imgSize,
    map_id: googleCloudMapId,
    key: apiKey,
    //path: `color:${pathColor}` + '|' + `enc:${polyline}`,
    path: `enc:${polyline}`,
  }

  return MAPS_BASE_URL + queryObjToString_(queryObj);
}

/** Helper 2: Call Make webhook */
function postToMakeWebhook_(postUrl, mapName) {
  const webhookUrl = "https://hook.us1.make.com/8obb3hb6bzwgi7s4nyi8yfghb3kxsksc";
  const payload = JSON.stringify({ url: postUrl, name: mapName });

  const options = {
    method: "post",
    contentType: "application/json",
    payload: payload
  };

  const response = UrlFetchApp.fetch(webhookUrl, options);
  Logger.log("[PL] Make Webhook Response: " + response.getContentText());
  return response;
}


/**
 * Save the map of a Strava activity as `MAPS_FOLDER/<Unix Epoch timestamp>.png`
 * 
 * @deprecated  Use Make Webhook instead.
 * 
 * @param {string} polyline  A encoded polyline representing a path.
 * @param {integer} timestamp  Name to save file as.
 * 
 * @author [Jikael Gagnon](<jikael.gagnon@mail.mcgill.ca>)
 * @author2 [Andrey Gonzalez](<andrey.gonzalez@mail.mcgill.ca>)
 * 
 * @date  Dec 1, 2024
 * @update  Mar 27, 2025
 */

function saveMapAsBlob_(polyline, timestamp) {
  if (!polyline) {
    return Logger.log('Map cannot be created: no polyline found for this activity');
  }

  const runMap = Maps.newStaticMap()
    .setSize(800, 600) // Adjust size as needed
    .setFormat(Maps.StaticMap.Format.PNG)
    .setMapType(Maps.StaticMap.Type.ROADMAP) // Use ROADMAP for a minimalist style
    .setPathStyle(3, "0x1155cc", "0x00000000") // Thin black route line, transparent fill
    .addPath(polyline)
    ;

  // Get save locati on using timestamp
  const saveLocation = getSaveLocation(timestamp);

  // Save runMap as as image to specified location
  const mapBlob = Utilities.newBlob(runMap.getMapImage(), 'image/png', saveLocation);
  const file = DriveApp.createFile(mapBlob);

  // Display success message
  Logger.log(`Successfully saved map as ${saveLocation}`);

  // Set permission to allow downloading
  file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  return file.getDownloadUrl();

  /** Helper function to get location of file to save */
  function getSaveLocation(submissionTime) {
    return MAPS_FOLDER + '/' + submissionTime.toString() + '.png'
  }
}



function testCloudUpload() {
  const fileId = "14csoxHqwHnnN7KFhsEgi5o55x1Sajvbh";
  const blob = DriveApp.getFileById(fileId).getBlob();

  const time = Utilities.formatDate(new Date(), TIMEZONE, "EEE-d-MMM-yyyy-k\'h\'mm");
  const imageName = "headrun-map-" + time + '.png';

  try {
    const imageUrl = uploadImageToBucket_(blob, imageName);
    Logger.log("Uploaded image URL: " + imageUrl);
  } catch (e) {
    Logger.log("Error during upload: " + e);
  }
}


function uploadImageToBucket_(imageBlob, imageName) {
  // Name of bucket in Google Cloud Storage
  const STORAGE_BUCKET_NAME = 'run-map-storage.firebasestorage.app';
  const BASE_UPLOAD_URL = "https://storage.googleapis.com/upload/storage/v1/b";

  // Get service key to access cloud storage
  const store = PropertiesService.getScriptProperties();
  const propertyName = SCRIPT_PROPERTY_KEYS.googleCloudKey;
  const SERVICE_ACCOUNT_KEY = JSON.parse(store.getProperty(propertyName));

  // Authenticate using the Service Account
  const token = getGSCAccessToken_(SERVICE_ACCOUNT_KEY);

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

  Logger.log('[PL] Image uploaded successfully!');
  return `https://storage.googleapis.com/${STORAGE_BUCKET_NAME}/${imageName}`; // Return the public URL
}


// Helper function to get an access token using the service account key
function getGSCAccessToken_(key) {
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



/**
 * Adds a Blob to a Google Cloud Storage bucket.
 *
 * @param {Blob} BLOB - The Blob of the file to add to the bucket.
 * @param {string} BUCKET_NAME - The name of the bucket containing to add the file to.
 * @param {string} OBJECT_NAME - The name/path of the file to add.
 * @return {Object} objects#resource
 *
 * @example
 * const fileBlob = addBucketFile_(driveBlob, 'my-bucket', 'path/to/my/file.txt');
 */
function addBucketFile_(BLOB, BUCKET_NAME, OBJECT_NAME) {
  const bytes = BLOB.getBytes();
  const baseUrl = "https://www.googleapis.com/upload/storage/v1/b";

  // Base URL for Cloud Storage API
  const url = `${baseUrl}/${BUCKET_NAME}/o?uploadType=media&name=${encodeURIComponent(OBJECT_NAME)}`;
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentLength: bytes.length,
      contentType: BLOB.getContentType(),
      payload: bytes,
      headers: {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken()
      }
    });
    return JSON.parse(response.getContentText());
  } catch (error) {
    console.error('Error getting file:', error);
  }
}