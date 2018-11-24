/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var restify = require('restify');
var builder = require('botbuilder');
var botbuilder_azure = require("botbuilder-azure");
var azurest = require('azure-storage');
var tableService = azurest.createTableService("botdyesa01", "+F+IpcFtKyi6jrCm05KMCPYfIuG2J+ezhnAgqTvtwVAYKb/rmJvOKp4KuJ+Q44ie0HhPMKaFk3sSjvweQ/31Kw==");

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    openIdMetadata: process.env.BotOpenIdMetadata 
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

/*----------------------------------------------------------------------------------------
* Bot Storage: This is a great spot to register the private state storage for your bot. 
* We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
* For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
* ---------------------------------------------------------------------------------------- */

var tableName = 'botdata';
var azureTableClient = new botbuilder_azure.AzureTableClient(tableName, process.env['AzureWebJobsStorage']);
var tableStorage = new botbuilder_azure.AzureBotStorage({ gzipData: false }, azureTableClient);

// Create your bot with a function to receive messages from the user
var bot = new builder.UniversalBot(connector);
bot.set('storage', tableStorage);


var Choice = {
    Si: 'Sí',
    No: 'No'
 };
// El díalogo principal inicia aquí
bot.dialog('/', [
    
    function (session, results, next) {
        // Primer diálogo    
        builder.Prompts.text(session, '¿Cuál es el número de serie que deseas revisar?')
    },
    function (session, results) {
        // Segundo diálogo
        session.dialogData.ticket = results.response;
        builder.Prompts.text(session, '¿Cuál es el nombre del asociado?')
    },
    function (session, results) {
        session.dialogData.asociado = results.response;
        // Tercer diálogo
        tableService.retrieveEntity("botdyesatb01", session.dialogData.asociado, session.dialogData.ticket, function(error, result, response) {
            // var unlock = result.Status._;
            if(!error ) {
    
                session.send(`Esta es la información relacionada: \n **Número de Serie: ${session.dialogData.ticket} \n Asociado: ${result.PartitionKey._}  \n Proyecto: ${result.Proyecto._} \n Estatus: ${result.Status._}.**`);
                builder.Prompts.choice(session, 'Hola ¿Esta información es correcta?', [Choice.Si, Choice.No], { listStyle: builder.ListStyle.button });
            }
            else{
                session.endDialog("**Error:** Los datos son incorrectos, intentalo nuevamente.");
            }
        });
    },
    function (session, results) {
        var selection = results.response.entity;
        switch (selection) {
            // El díalogo desbloqueo inicia si el usuario presiona Desbloquear cuenta
            case Choice.Si:
            // return session.beginDialog('viaticos');
            
            tableService.retrieveEntity("botdyesatb01", session.dialogData.asociado, session.dialogData.ticket, function(error, result, response) {
                // var unlock = result.Status._;
                if(!error ) {
        
                    session.send(`¿Deseas adjutar documentación o evidencia?`);
                    session.endDialog(`Hemos terminado gracias`);
                }
                else{
                    session.endDialog("**Error:**");
                }
            });
            break;
            // El díalogo existe inicia si el usuario presiona Resetear contraseña
            case Choice.No:
            session.endDialog('Por favor vuelve a introducir correctamente la información.');
            break;
        }
        
    }
]);
