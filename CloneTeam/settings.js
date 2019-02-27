module.exports = () => {
    var settings = settings ||
        {
            AUTH_URL: "https://login.microsoftonline.com/",
            GRAPH_URL: "https://graph.microsoft.com/",
            TENANT: process.env["TENANT"],
            CLIENT_ID: process.env["CLIENT_ID"],
            CLIENT_SECRET: process.env["CLIENT_SECRET"]
        };
    return settings;
}