function newTwitterV2() {
  return new TwitterV2();
}

function jsonFetch(options) {
  const fetchOptions = {
    method: options.method || "get",
    contentType: "application/json",
    muteHttpExceptions: true,
    headers: options.headers || {},
  };

  if (options.payload) {
    fetchOptions.payload = options.payload;
  }

  const response = UrlFetchApp.fetch(options.url, fetchOptions);
  const code = response.getResponseCode();
  const body = response.getContentText();

  let json;
  try {
    json = JSON.parse(body);
  } catch (e) {
    json = { error: "Invalid JSON", raw: body };
  }

  return {
    status: code,
    body: json,
  };
}

class TwitterV2 {
  static oauth2() {
    const clientId = TWIITER_CLIENT_ID;
    const clientSecret = TWITTER_CLIENT_SECRET;
    let codeVerifier = PropertiesService.getUserProperties().getProperty("code_verifier");
    let codeChallenge = PropertiesService.getUserProperties().getProperty("code_challenge");
    if (!codeVerifier) {
      const chrs = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~";
      codeVerifier = "";
      for (var i = 0; i < 128; i++) {
        codeVerifier += chrs.charAt(Math.floor(Math.random() * chrs.length));
      }

      codeChallenge = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, codeVerifier))
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=+$/, '')
      PropertiesService.getUserProperties().setProperty("code_verifier", codeVerifier);
      PropertiesService.getUserProperties().setProperty("code_challenge", codeChallenge);
    }

    return OAuth2.createService('t')
      .setAuthorizationBaseUrl('https://twitter.com/i/oauth2/authorize')
      .setTokenUrl('https://api.twitter.com/2/oauth2/token?code_verifier=' + codeVerifier)
      .setClientId(clientId)
      .setClientSecret(clientSecret)
      .setCallbackFunction(authCallback.name)
      .setPropertyStore(PropertiesService.getUserProperties())
      .setScope('users.read follows.read tweet.write tweet.read offline.access')
      .setParam('response_type', 'code')
      .setParam('code_challenge_method', 'S256')
      .setParam('code_challenge', codeChallenge)
      .setTokenHeaders({
        'Authorization': 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret),
        'Content-Type': 'application/x-www-form-urlencoded'
      })
  }

  constructor(callback = (url) => console.log(url)) {
    this.service = TwitterV2.oauth2();
    this.callback = callback;
  }

  okAuth_() {
    if (this.service.hasAccess()) {
      return true;
    } else {
      this.callback(this.service.getAuthorizationUrl());
      return false;
    }
  }

  appendHeaders_(options) {
    options.headers = {
      ...options.headers,
          'Authorization': `Bearer ${this.service.getAccessToken()}`,
    };
    return options;
  }

  postMessage(message) {
    if (!this.okAuth_()) {
      return;
    }

    return jsonFetch(this.appendHeaders_({
      url: "https://api.twitter.com/2/tweets",
      payload: JSON.stringify({ text: message }),
      method: "post",
    }));
  }

  me() {
    if (!this.okAuth_()) {
      return;
    }

    return jsonFetch(this.appendHeaders_({
      url: "https://api.twitter.com/2/users/me",
    }))
  }

  getUser(username) {
    if (!this.okAuth_()) {
      return;
    }

    return jsonFetch(this.appendHeaders_({
      url: `https://api.twitter.com/2/users/by/username/${username}`,
    }))
  }
}

function authCallback(request) {
  return TwitterV2.oauth2().handleCallback(request)
    ? HtmlService.createHtmlOutput('Success!')
    : HtmlService.createHtmlOutput('Denied.');
}

