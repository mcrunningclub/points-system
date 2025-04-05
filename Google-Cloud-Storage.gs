function testUpload() {
  const fileId = "1iPoOHCiUju-BjqrN7nhZVs_nZXxDwhOa";
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  
  const imageName = "headrun-map-" + fileId;

  try {
    const imageUrl = uploadImageToBucket(blob, imageName);
    Logger.log("Uploaded image URL: " + imageUrl);
  } catch (e) {
    Logger.log("Error during upload: " + e);
  }
}


// `https://www.googleapis.com/upload/storage/v1/b/${BUCKET_NAME}/o?uploadType=media&name=${encodeURIComponent(OBJECT_NAME)}`;
// https://storage.googleapis.com/run-map-storage.firebasestorage.app/Tue-5-Nov-2024-06_07.png

function uploadImageToBucket(imageBlob, imageName) {
  // Name of bucket in Google Cloud Storage
  const STORAGE_BUCKET_NAME = 'run-map-storage.firebasestorage.app';

  // Get service key to access cloud storage
  const store = PropertiesService.getScriptProperties();
  const SERVICE_ACCOUNT_KEY = JSON.parse(store.getProperty("GOOGLE_ACCOUNT_KEY"));

  // Authenticate using the Service Account
  const token = getAccessToken(SERVICE_ACCOUNT_KEY);

  // Construct the upload URL
  const uploadUrl = "https://storage.googleapis.com/upload/storage/v1/b/" +
  `${STORAGE_BUCKET_NAME}/o?uploadType=media&name=${imageName}`;

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

  Logger.log('Image uploaded successfully!');
  return `https://storage.googleapis.com/${STORAGE_BUCKET_NAME}/${imageName}`; // Return the public URL
}


// Helper function to get an access token using the service account key
function getAccessToken(key) {
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
 
  // Base URL for Cloud Storage API
  const url = `https://www.googleapis.com/upload/storage/v1/b/${BUCKET_NAME}/o?uploadType=media&name=${encodeURIComponent(OBJECT_NAME)}`;
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