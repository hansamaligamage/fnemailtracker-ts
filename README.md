# Http trigger written in Typescript to read mails from MS Graph API and store them in Cosmos DB using SQL API
This is a http trigger function written in Typescript in Visual Studio Code as the editor. It reads emails from your mail account using Microsoft Graph API and store them to Cosmos DB on SQL API

## Technology stack  
* Typescript version 3.8.3 *(npm i typescript)* https://www.npmjs.com/package/typescript 
* Azure functions for typescript version 1.2 *(npm i @azure/functions)* https://www.npmjs.com/package/@azure/functions 
* Azure Cosmos 3.7 to connect to the Cosmos DB SQL API *(npm i @azure/cosmos)* https://www.npmjs.com/package/@azure/cosmos
* isomorphic-fetch version 2.2.1 used to fetch data from Microsoft graph REST API *(npm i isomorphic-fetch es6-promise)* https://www.npmjs.com/package/isomorphic-fetch
* Microsoft graph types 1.12 to get the intellisence support for graph objects, in this example its emails *(npm i @microsoft/microsoft-graph-types)* https://www.npmjs.com/package/@microsoft/microsoft-graph-types
