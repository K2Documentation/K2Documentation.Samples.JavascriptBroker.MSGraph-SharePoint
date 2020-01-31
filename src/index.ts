import '@k2oss/k2-broker-core';

metadata = {
    systemName: "com.k2.example",
    displayName: "Example Broker",
    description: "An example broker that accesses JSONPlaceholder."
};

ondescribe = async function(): Promise<void> {
    postSchema({
        objects: {
            "com.k2.todo": {
                displayName: "TODO",
                description: "Manages a TODO list",
                properties: {
                    "com.k2.todo.id": {
                        displayName: "ID",
                        type: "number"
                    },
                    "com.k2.todo.userId": {
                        displayName: "User ID",
                        type: "number"
                    },
                    "com.k2.todo.title": {
                        displayName: "Title",
                        type: "string"
                    },
                    "com.k2.todo.completed": {
                        displayName: "Completed",
                        type: "boolean"
                    }
                },
                methods: {
                    "com.k2.todo.get": {
                        displayName: "Get TODO",
                        type: "read",
                        inputs: [ "com.k2.todo.id" ],
                        outputs: [ "com.k2.todo.id", "com.k2.todo.userId", "com.k2.todo.title", "com.k2.todo.completed" ]
                    }
                }
            }
        }
    });
}

onexecute = async function(objectName, methodName, _parameters, properties): Promise<void> {
    switch (objectName)
    {
        case "com.k2.todo": await onexecuteTodo(methodName, properties); break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

async function onexecuteTodo(methodName: string, properties: SingleRecord): Promise<void> {
    switch (methodName)
    {
        case "com.k2.todo.get": await onexecuteTodoGet(properties); break;
        default: throw new Error("The method " + methodName + " is not supported.");
    }
}

function onexecuteTodoGet(properties: SingleRecord): Promise<void> {
    return new Promise<void>((resolve, reject) =>
    {
        var xhr = new XMLHttpRequest();
        xhr.onreadystatechange = function() {
            try {
                if (xhr.readyState !== 4) return;
                if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);

                var obj = JSON.parse(xhr.responseText);
                postResult({
                    "com.k2.todo.id": obj.id,
                    "com.k2.todo.userId": obj.userId,
                    "com.k2.todo.title": obj.title,
                    "com.k2.todo.completed": obj.completed
                });
                resolve();
            } catch (e) {
                reject(e);
            }
        };

        xhr.open("GET", 'https://jsonplaceholder.typicode.com/todos/' + properties["com.k2.todo.id"]);
        xhr.setRequestHeader('test', 'test value');
        xhr.send();
    });
}