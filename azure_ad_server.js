AzureAd.whitelistedFields = [
  'id',
  'userPrincipalName',
  'mail',
  'displayName',
  'surname',
  'givenName',
];

OAuth.registerService('azureAd', 2, null, function (query) {
  var tokens = getTokensFromCode(redirectUrlFromQuery(query), query.code);
  var graphUser = AzureAd.resources.graph.getUser(tokens.accessToken);
  var serviceData = {
    ...tokens,
    expiresAt: +new Date() + 1000 * parseInt(tokens.expiresIn, 10),
  };

  var fields = _.pick(graphUser, AzureAd.whitelistedFields);

  _.extend(serviceData, fields);

  // only set the token in serviceData if it's there. this ensures
  // that we don't lose old ones (since we only get this on the first
  // log in attempt)
  if (tokens.refreshToken) serviceData.refreshToken = tokens.refreshToken;

  var emailAddress = graphUser.mail || graphUser.userPrincipalName;

  var options = {
    profile: {
      name: graphUser.displayName,
    },
  };

  if (!!emailAddress) {
    options.emails = [
      {
        address: emailAddress,
        verified: true,
      },
    ];
  }

  return { serviceData: serviceData, options: options };
});

function redirectUrlFromQuery(query) {
  const state = OAuth._stateFromQuery(query) || {};
  return decodeURIComponent(state.redirectUrl);
}

function getTokensFromCode(redirectUrl, code) {
  return AzureAd.http.getAccessTokensBase(redirectUrl, {
    grant_type: 'authorization_code',
    code: code,
  });
}

AzureAd.retrieveCredential = function (credentialToken, credentialSecret) {
  return OAuth.retrieveCredential(credentialToken, credentialSecret);
};
