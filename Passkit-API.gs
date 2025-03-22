const PROPERTY_STORE = PropertiesService.getScriptProperties();
const PK_API_KEY = PROPERTY_STORE.getProperty('PK_API_KEY');
const PK_API_SECRET = PROPERTY_STORE.getProperty('PK_API_SECRET');

function getPasskitResponse() {
  // Generate PassKit auth token for an API call
  const apiKey = PK_API_KEY;
  const apiSecret = PK_API_SECRET;
  var token = generateJWT_(apiKey, apiSecret);
  //Logger.log(token);
  
  var pk_url = 'https://api.pub2.passkit.io/';
  var options = {
    headers: {
      'method' : 'get',
      'authorization' : token,
      'contentType' : 'application/json'
    }
  }
  
  // Call PassKit API. Documentation available on https://docs.passkit.io/.
  // This example makes a get request to obtain an account profile.
  var endpoint = pk_url + 'user/profile';

  var url2 = "https://api-pass.passkit.net/v3/campaigns";
  //Logger.log(url);
  var response = UrlFetchApp.fetch(url2, options);
  var respText = response.getContentText();
  
  Logger.log(respText);
}


function passkitPost_() {
  // Make a POST request with a JSON payload.
  var data = {
    'name': 'Bob Smith',
    'age': 35,
    'pets': ['fido', 'fluffy']
  };
  var options = {
    'method' : 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload' : JSON.stringify(data)
  };
  UrlFetchApp.fetch('https://httpbin.org/post', options);
}


// Generate PassKit API token with your API key and secret
function generateJWT_(key, secret) {		
    var body = {
        "uid": key,
        "exp": Math.floor(new Date().getTime() / 1000) + 3600,
        "iat": Math.floor(new Date().getTime() / 1000),
        "web": true,
    };
    header = {
        "alg": "HS256",
        "typ": "JWT"
    };
    var token = [];
    token[0] = base64url_(JSON.stringify(header));
    token[1] = base64url_(JSON.stringify(body));
    token[2] = genTokenSign_(token, secret);
   return token.join(".");
}


function genTokenSign_(token, secret) {
    if (token.length != 2) {
        return;
    }
    var hash = Utilities.computeHmacSha256Signature(token.join("."), secret);
    var base64Hash = Utilities.base64Encode(hash);
    return urlConvertBase64_(base64Hash);
}


function base64url_(input) {
    var base64String = Utilities.base64Encode(input);
    return urlConvertBase64_(base64String);
}


function urlConvertBase64_(input) {
    var output = input.replace(/=+$/, '');
    output = output.replace(/\+/g, '-');
    output = output.replace(/\//g, '_');
   return output;
}
