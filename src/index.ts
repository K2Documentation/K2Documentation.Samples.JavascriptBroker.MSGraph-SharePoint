import '@k2oss/k2-broker-core';

metadata = {
    "systemName": "MSGraphSharePoint",
    "displayName": "Microsoft Graph - SharePoint Broker",
    "description": "Sample broker for SharePoint using MSGraph"
};

ondescribe = async function (): Promise<void> {
    postSchema({
        objects: {
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
                        outputs: ["id", "name"]
                    }
                }
            }
        }
    }
    )
};

onexecute = async function (objectName, methodName, parameters, properties): Promise<void> {
    switch (objectName) {
        case "list": await onexecuteList(methodName, parameters, properties); break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

async function onexecuteList(methodName: string, parameters: SingleRecord, properties: SingleRecord): Promise<void> {
    switch (methodName) {
        case "get": await onexecuteListGet(parameters, properties); break;
        default: throw new Error("The method " + methodName + " is not supported.");
    }
}

function onexecuteListGet(parameters: SingleRecord, properties: SingleRecord): Promise<void> {
    return new Promise<void>((resolve, reject) => {
        var xhr = new XMLHttpRequest();

        xhr.onreadystatechange = function () {
            try {
                if (xhr.readyState !== 4) return;
                if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);

                var obj = JSON.parse(xhr.responseText);
                postResult(obj.map(x => {
                    return {
                        "id": x.id,
                        "name": x.name
                    }
                }));
                resolve();
            } catch (error) {
                reject();
            }
        };

        var url = "https://graph.microsoft.com/v1.0/sites/root/lists";

        xhr.open("GET", url);
        // Authentication Header
        // Use .withCredentials to use service instance configured OAuth (Bearer) or Static (Basic)
        // Anything else, don't set .withCredentials and use .setRequestHeader to set the Authentication header
        xhr.withCredentials = true;
        xhr.send();
    });
}