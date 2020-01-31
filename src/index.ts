import '@k2oss/k2-broker-core';

metadata = {
	"systemName": "MSGraphSharePoint",
	"displayName": "Microsoft Graph - SharePoint Broker",
	"description": "Sample broker for SharePoint using MSGraph"
};

ondescribe = function() {
    postSchema({ objects: {
                "com.k2.sample.msgraph.sharepoint.list": {
                    displayName: "List",
                    description: "SharePoint List",
                    properties: {
                        "com.k2.sample.msgraph.sharepoint.list.id": {
                            displayName: "ID",
                            type: "string" 
                        },
                        "com.k2.sample.msgraph.sharepoint.list.name": {
                            displayName: "Name",
                            type: "string" 
                        }
                    },
                    methods: {
                        "com.k2.sample.msgraph.sharepoint.list.get": {
                            displayName: "Get List",
                            type: "list",
                            outputs: [ "com.k2.sample.msgraph.sharepoint.list.id", "com.k2.sample.msgraph.sharepoint.list.name" ]
                        }
                    }
                }
            }
        }
    )};

onexecute = function(objectName, methodName, parameters, properties) {
    switch (objectName)
    {
        case "com.k2.sample.msgraph.sharepoint.list": onexecuteList(methodName, parameters, properties); break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

function onexecuteList(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName)
    {
        case "com.k2.sample.msgraph.sharepoint.list.get": onexecuteListGet(parameters, properties); break;
        default: throw new Error("The method " + methodName + " is not supported.");
    }
}

function onexecuteListGet(parameters: SingleRecord, properties: SingleRecord) {
    var xhr = new XMLHttpRequest();

    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);

        //console.log(xhr.responseText);
        var obj = JSON.parse(xhr.responseText);
        for (var key in obj) {
            postResult({
                "com.k2.sample.msgraph.sharepoint.list.id": obj[key].id,
                "com.k2.sample.msgraph.sharepoint.list.name": obj[key].name});
                // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/map
            }
    };

    var url = "https://graph.microsoft.com/v1.0/sites/root/lists";
    //console.log(url);

    xhr.open("GET", url);
    // Authentication Header
    // Use .withCredentials to use service instance configured OAuth (Bearer) or Static (Basic)
    // Anything else, don't set .withCredentials and use .setRequestHeader to set the Authentication header
    xhr.withCredentials = true;
    //xhr.setRequestHeader("Accept", "application/json");
    xhr.send();
}