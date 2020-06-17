import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
require('isomorphic-fetch');
import { AppConfiguration } from "read-appsettings-json";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');

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
};

export default httpTrigger;

async function saveemails (key, endpoint, databaseId, containerId, emails){
    
    const { CosmosClient } = require("@azure/cosmos");
                
    const client = new CosmosClient({ endpoint, key });
    console.log("Cosmos client is created");

    const { database } = await client.databases.createIfNotExists({ id: databaseId });
    console.log(database.id);

    const { container } = await database.containers.createIfNotExists({ id: containerId});
    console.log(container.id);

    for (const email of emails) {
        container.items.create(email);
      }
}