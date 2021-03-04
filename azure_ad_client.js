// Request AzureAd credentials for the user
// @param options {optional}
// @param credentialRequestCompleteCallback {Function} Callback function to call on
//   completion. Takes one argument, credentialToken on success, or Error on
//   error.
AzureAd.requestCredential = function (options, credentialRequestCompleteCallback) {
  // support both (options, callback) and (callback).
  if (!credentialRequestCompleteCallback && typeof options === 'function') {
    credentialRequestCompleteCallback = options;
    options = {};
  } else if (!options) {
    options = {};
  }

  var config = AzureAd.getConfiguration(true);
  if (!config) {
    credentialRequestCompleteCallback &&
      credentialRequestCompleteCallback(new ServiceConfiguration.ConfigError());
    return;
  }

  var prompt = '&prompt=login';
  if (typeof options.loginPrompt === 'string') {
    if (options.loginPrompt === '') prompt = '';
    else prompt = '&prompt=' + options.loginPrompt;
  }

  var loginStyle = OAuth._loginStyle('azureAd', config, options);
  var credentialToken = Random.secret();

  var tenant = options.tenant || 'common';

  var baseUrl = 'https://login.microsoftonline.com/' + tenant + '/oauth2/v2.0/authorize?';

  var redirectUrl = encodeURIComponent(OAuth._redirectUri('azureAd', config, null, options));

  var loginUrl =
    baseUrl +
    'response_type=code' +
    prompt +
    '&scope=' +
    config.scope +
    '&client_id=' +
    config.clientId +
    '&response_mode=query' +
    '&state=' +
    OAuth._stateParam(loginStyle, credentialToken, redirectUrl) +
    '&redirect_uri=' +
    redirectUrl;

  OAuth.launchLogin({
    loginService: 'azureAd',
    loginStyle: loginStyle,
    loginUrl: loginUrl,
    credentialRequestCompleteCallback: credentialRequestCompleteCallback,
    credentialToken: credentialToken,
    popupOptions: { height: 600 },
  });
};
