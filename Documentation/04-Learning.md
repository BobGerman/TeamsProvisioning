# Teams Provisioning Sample

## Part 4: Learning from the code

This article highlights some interesting areas in the Azure Functions that may be useful in other projects.

### Handling app settings

App settings can be read using

~~~JavaScript
process.env["name"]
~~~

The environment is smart enough to pull the settings from your local.settings.json file when debugging, or from the Azure Function Application settings in Azure. The file Services/Settings/settings.js reads in these values and also includes a few constants that are used throughout the solution.

### Running with App Permissions in Node

One of the design goals of this project was to use application permissions rather than relying on a user account. This sample shows how to do that in Node.

If you look at the package.json file, you'll notice that we have two dependencies on npm packages:

~~~JSON
  "dependencies": {
    "adal-node": "^0.1.28",
    "request": "^2.88.0"
  },
~~~

* adal-node is the package that obtains an access token using app permissions (client credential flow). This token is included in all Graph API calls to authenticate our application

* request is a package that lets us make web service calls from Node

Now check out the file Services/Token/getToken.js. This function returns a promise of an access token which can be used in every Graph call. There's not a lot of code here - very simple!

~~~JavaScript
  return new Promise((resolve, reject) => {

    // Get the ADAL client
    const authContext = new adal.AuthenticationContext(
      settings().AUTH_URL + settings().TENANT);
      
    authContext.acquireTokenWithClientCredentials(
      settings().GRAPH_URL, settings().CLIENT_ID, settings().CLIENT_SECRET, 
      (err, tokenRes) => {
        if (err) {
          reject(err); 
        } else {
          resolve(tokenRes.accessToken);
        }
    });
  });

~~~

Then, to call the Graph, this token is added to the heading:

~~~JavaScript
request.get(url, {
    'auth': {
        'bearer': token
    }
}, (error, response, body) => {
    // ... callback from request
}
~~~

### The right way to Create or Clone

The Create and Clone calls in Graph are long-running - that is, they probably won't finish immediately. Rather than return a result, these calls return an "operation URL" which you're supposed to call periodically until the operation is done.

I've noticed many examples that call the Clone or Create API and then they move on without checking the operation URL. This might be OK sometimes, but there's no way to be sure your request will ultimately succeed, and someone might try to access the Team before it's fully created!

Check out Services/Team/createTeam.js (or cloneTeam.js). There you will see a function called PollUntilDone(); it's called by the callback from the create or clone request, and it calls itself recursively until it gets a completion (success or failure), or until it exceeds a retry count.

### Reading a file from SharePoint

If you've never read a SharePoint file from the Graph API, you've never lived! When you're running on a SharePoint page, this is trivial (just request the URL). But in the Graph API I found myself going through quite a number of steps:

1. Given the site URL, get the SharePoint site ID

1. Given the SharePoint Site ID and library name, get a Drive ID

1. Given the Drive ID and the filename, get the Item, which includes a download URL

1. Reuqest the download URL and get the contents of the file

See Services/Template/getTemplate.js for the code.
