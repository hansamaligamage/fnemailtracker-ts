# Http trigger written in Typescript to read mails from MS Graph API and store them in Cosmos DB using SQL API
This is a http trigger function written in Typescript in Visual Studio Code as the editor. It reads emails from your mail account using Microsoft Graph API and store them to Cosmos DB on SQL API

## Technology stack  
* Typescript version 3.8.3 *(npm i typescript)* https://www.npmjs.com/package/typescript 
* Azure functions for typescript version 1.2 *(npm i @azure/functions)* https://www.npmjs.com/package/@azure/functions 
* Azure Cosmos 3.7 to connect to the Cosmos DB SQL API *(npm i @azure/cosmos)* https://www.npmjs.com/package/@azure/cosmos
* isomorphic-fetch version 2.2.1 used to fetch data from Microsoft graph REST API *(npm i isomorphic-fetch es6-promise)* https://www.npmjs.com/package/isomorphic-fetch
* Microsoft graph types 1.12 to get the intellisence support for graph objects, in this example its emails *(npm i @microsoft/microsoft-graph-types)* https://www.npmjs.com/package/@microsoft/microsoft-graph-types

## How to run the solution

 * You have to create a Cosmos DB account with SQL API and go to the Connection String section and get the endpoint and primary key to connect to the database
 * Open the solution from Visual Studio code, install all the packages from npm i command and run the solution

## Code snippets
### Package references in the main file
```
import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
require('isomorphic-fetch');
import { AppConfiguration } from "read-appsettings-json";
```

### Read emails using MS Graph API
```
const accessToken = AppConfiguration.Setting().accessToken;
const key = AppConfiguration.Setting().key;
const endpoint  = AppConfiguration.Setting().endpoint;
const database = AppConfiguration.Setting().container;
const container = AppConfiguration.Setting().container;
    
let url = "https://graph.microsoft.com/v1.0/me/messages/";
    let request = new Request(url, {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + accessToken
        })
    });
     
    fetch(request)
    .then((response) => {
        response.json().then(async (res) => {
            if(res.error){
                console.error("code : " + res.error.code + " message : " + res.error.message);
            }
            else
            {
                let messages:[MicrosoftGraph.Message] = res.value;
                await saveemails(key, endpoint, database, container, messages);
            }
        });
    })
    .catch((error) => {
        console.error(error);
    });
```

### Create a cosmos client
```
const { CosmosClient } = require("@azure/cosmos");
                
const client = new CosmosClient({ endpoint, key });
console.log("Cosmos client is created");
```

### Create a database in the Cosmos account
```
const { database } = await client.databases.createIfNotExists({ id: databaseId })
console.log(database.id);
```

### Create a container in the database
```
const { container } = await database.containers.createIfNotExists({ id: containerId});
console.log(container.id);
```

### Create a document in the database
```
for (const email of emails) {
    container.items.create(email);
}
 ```
