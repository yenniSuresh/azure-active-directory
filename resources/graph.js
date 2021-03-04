AzureAd.resources.graph = {};
AzureAd.resources.graph.friendlyName = 'graph';
AzureAd.resources.graph.resourceUri = 'https://graph.microsoft.com/';

AzureAd.resources.graph.getUser = function (accessToken, version = 'v1.0') {
  var config = AzureAd.getConfiguration();
  var url = AzureAd.resources.graph.resourceUri + version + '/me';

  return AzureAd.http.callAuthenticated('GET', url, accessToken);
};

AzureAd.resources.graph.get = function (resource, accessToken) {
  var config = AzureAd.getConfiguration();
  var version = config.verion || 'v1.0';
  var url = AzureAd.resources.graph.resourceUri + version + '/' + resource;

  return AzureAd.http.callAuthenticated('GET', url, accessToken);
};

if (Meteor.isServer) {
  Meteor.startup(function () {
    AzureAd.resources.registerResource(
      AzureAd.resources.graph.friendlyName,
      AzureAd.resources.graph.resourceUri
    );
  });
}
