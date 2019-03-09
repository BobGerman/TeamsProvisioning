module.exports = () => {
    var settings = settings ||
        {
            AUTH_URL: "https://login.microsoftonline.com/",
            GRAPH_URL: "https://graph.microsoft.com/",
            TEMPLATE_FILE_EXTENSION: ".json.txt",
            TENANT: process.env["TENANT"],
            TENANT_NAME: process.env["TENANT_NAME"],
            CLIENT_ID: process.env["CLIENT_ID"],
            CLIENT_SECRET: process.env["CLIENT_SECRET"],
            TEMPLATE_SITE_URL: process.env["TEMPLATE_SITE_URL"],
            TEMPLATE_LIB_NAME: process.env["TEMPLATE_LIB_NAME"]
        };
    return settings;
}