import '@k2oss/k2-broker-core';

metadata = {
	"systemName": "MSGraphSharePoint",
	"displayName": "Microsoft Graph - SharePoint Broker",
	"description": "Sample broker for SharePoint using MSGraph"
};

ondescribe = function() {
    postSchema({ objects: {
                "list": {
                    displayName: "List",
                    description: "SharePoint List",
                    properties: {
                        "id": {
                            displayName: "ID",
                            type: "string" 
                        },
                        "name": {
                            displayName: "Name",
                            type: "string" 
                        }
                    },
                    methods: {
                        "get": {
                            displayName: "Get List",
                            type: "list",
                            outputs: [ "id", "name" ]
                        }
                    }
                }
            }
        }
    )};

onexecute = function(objectName, methodName, parameters, properties) {
    switch (objectName)
    {
        case "list": onexecuteList(methodName, parameters, properties); break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

function onexecuteList(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName)
    {
        case "get": onexecuteListGet(parameters, properties); break;
        default: throw new Error("The method " + methodName + " is not supported.");
    }
}

function onexecuteListGet(parameters: SingleRecord, properties: SingleRecord) {
    var xhr = new XMLHttpRequest();

    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);

        var obj = JSON.parse(xhr.responseText);
        postResult(obj.map(x => {
            return {
                "id": x.id,
                "name": x.name
            }
        }));
    };

    var url = "https://graph.microsoft.com/v1.0/sites/root/lists";

    xhr.open("GET", url);
    // Authentication Header
    // Use .withCredentials to use service instance configured OAuth (Bearer) or Static (Basic)
    // Anything else, don't set .withCredentials and use .setRequestHeader to set the Authentication header
    xhr.withCredentials = true;
    xhr.send();
}