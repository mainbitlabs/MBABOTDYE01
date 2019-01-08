/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var restify = require('restify');
var builder = require('botbuilder');
var botbuilder_azure = require("botbuilder-azure");
var nodeoutlook = require('nodejs-nodemailer-outlook');
var image2base64 = require('image-to-base64');
var azurest = require('azure-storage');
var config = require('./config');
var tableService = azurest.createTableService(config.storageA, config.accessK);
var blobService = azurest.createBlobService(config.storageA,config.accessK);

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
//    console.log('%s listening to %s', server.name, server.url); 
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

 var Motivos = {
    Uno: 'El usuario rechaza el servicio y/o no disponible',
    Dos: 'El usuario no se localiza en el sitio',
    Tres: 'El usuario no quiere firmar',
    Cuatro: 'Problema de infraestructura',
    Cinco: 'Equipo no se encuentra en sitio',
 };

 var Opts = {
    Resguardo : 'Resguardo',
    Check: 'Check',
    Borrado: 'Borrado',
    Baja: 'Baja',
    Pospuesto: 'Incidente de Servicio'
 };
 
 
 var time;
 // Variable Discriptor para actualizar tabla
 var Discriptor = {};
 // El díalogo principal inicia aquí
 bot.dialog('/', [
     function (session) {
         // Primer diálogo    
         session.send(`Hola bienvenido al Servicio Automatizado de Mainbit.`);
         session.send(`**Sugerencia:** Recuerda que puedes cancelar en cualquier momento escribiendo **"cancelar".** \n\n **Importante:** este bot tiene un ciclo de vida de 5 minutos, te recomendamos concluir la actividad antes de este periodo.`);
         builder.Prompts.text(session, 'Por favor, **escribe el Número de Serie del equipo.**');
         time = setTimeout(() => {
             session.endConversation(`**Lo sentimos ha transcurrido el tiempo estimado para completar esta actividad. Intentalo nuevamente.**`);
         }, 300000);
     },
     function (session, results) {
         // Segundo diálogo
         session.dialogData.serie = results.response;
         builder.Prompts.text(session, '¿Cuál es tu **Clave de Asociado**?')
     },
     function (session, results) {
         session.dialogData.asociado = results.response;
         // Tercer diálogo
         tableService.retrieveEntity(config.table1, session.dialogData.asociado, session.dialogData.serie, function(error, result, response) {
             if(!error && result.Resguardo._ === 'Resguardo Adjunto' && result.Baja._ === 'Baja Adjunto' && result.Check._ === 'Check Adjunto' && result.Borrado._ === 'Borrado Adjunto') {
                 var Estatus = {
                     PartitionKey : {'_': session.dialogData.asociado, '$':'Edm.String'},
                     RowKey : {'_': session.dialogData.serie, '$':'Edm.String'},
                     Status : {'_': 'Completado', '$':'Edm.String'}
                 };
                 console.log(Estatus);
                 tableService.mergeEntity(config.table1, Estatus, function(err, res, respons){
                     if (!err){
                         console.log(`Status Completado`);
                         Estatus = {};
                     }
                     else{err}
                 });
             } 
             else{
                 clearTimeout(time);
                 // session.endConversation("**Error** 1");
             }
         });
         session.sendTyping();
             // Envíamos un mensaje al usuario para que espere.
             session.send('Estamos atendiendo tu solicitud. Por favor espera un momento...');
             setTimeout(() => {
         tableService.retrieveEntity(config.table1, session.dialogData.asociado, session.dialogData.serie, function(eror, result, response) {
             if (!eror) {                    
                 session.dialogData.proyecto= result.Proyecto._;
                 session.send(`**Proyecto:** ${result.Proyecto._} \n\n **Número de Serie**: ${result.RowKey._} \n\n **Asociado:** ${result.PartitionKey._}  \n\n  **Descripción:** ${result.Descripcion._} \n\n  **Localidad:** ${result.Localidad._} \n\n  **Inmueble:** ${result.Inmueble._} \n\n  **Servicio:** ${result.Servicio._} \n\n \n\n **Estatus:** ${result.Status._} \n\n \n\n **Resguardo:** ${result.Resguardo._} \n\n  **Check:** ${result.Check._} \n\n  **Borrado:** ${result.Borrado._} \n\n  **Baja:** ${result.Baja._}`);
                 builder.Prompts.choice(session, 'Hola ¿Esta información es correcta?', [Choice.Si, Choice.No], { listStyle: builder.ListStyle.button });          
             }
             else{
                 clearTimeout(time);
                 session.endConversation("**Error** La serie no coincide con el Asociado.");
             }
             });
         }, 5000);
     },
     function (session, results) {
         // Cuarto diálogo
         var selection = results.response.entity;
         switch (selection) {
             
             case Choice.Si:
             builder.Prompts.choice(session, '¿Deseas adjuntar Evidencia o Documentación?', [Choice.Si, Choice.No], { listStyle: builder.ListStyle.button });
             break;
 
             case Choice.No:
             clearTimeout(time);
             session.endConversation("Por favor valida con tu soporte que el Número de Serie esté asignado a tu Asociado");
             break;
         }
         
     },
     function (session, results) {
         // Cuarto diálogo
         var selection3 = results.response.entity;
         switch (selection3) {
             
             case Choice.Si:
             builder.Prompts.choice(session, '¿Que tipo de Evidencia o Documentación?', [Opts.Resguardo, Opts.Check, Opts.Borrado, Opts.Baja, Opts.Pospuesto], { listStyle: builder.ListStyle.button });  
             break;
 
             case Choice.No:
             clearTimeout(time);
             session.endConversation("De acuerdo, hemos terminado por ahora");
             break;
         }
         
     },
     function (session, results) {
         // Quinto diálogo
         var selection2 = results.response.entity;
         session.dialogData.tipo = selection2;
         session.dialogData.Discriptor ={};
         switch (selection2) {
 
             case Opts.Resguardo:
             function appendResguardo() {
                 Discriptor.PartitionKey = {'_': session.dialogData.asociado, '$':'Edm.String'};
                 Discriptor.RowKey = {'_': session.dialogData.serie, '$':'Edm.String'};
                 Discriptor.Resguardo = {'_': 'Resguardo Adjunto', '$':'Edm.String'};
             };
             appendResguardo();
             builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Resguardo}**`);
             break;
 
             case Opts.Borrado:
             function appendBorrado() {
                 Discriptor.PartitionKey = {'_': session.dialogData.asociado, '$':'Edm.String'};
                 Discriptor.RowKey = {'_': session.dialogData.serie, '$':'Edm.String'};
                 Discriptor.Borrado = {'_': 'Borrado Adjunto', '$':'Edm.String'};
             };
             appendBorrado();
             builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Borrado}**`);
             break;
 
             case Opts.Baja:
             function appendBaja() {
                 Discriptor.PartitionKey = {'_': session.dialogData.asociado, '$':'Edm.String'};
                 Discriptor.RowKey = {'_': session.dialogData.serie, '$':'Edm.String'};
                 Discriptor.Baja = {'_': 'Baja Adjunto', '$':'Edm.String'};
             };
             appendBaja();
             builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Baja}**`);
             break;
 
             case Opts.Check:
             function appendCheck() {
                 Discriptor.PartitionKey = {'_': session.dialogData.asociado, '$':'Edm.String'};
                 Discriptor.RowKey = {'_': session.dialogData.serie, '$':'Edm.String'};
                 Discriptor.Check = {'_': 'Check Adjunto', '$':'Edm.String'};
                 
             };
             appendCheck();
             builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Check}**`);
             break;
 
             case Opts.Pospuesto:
             // Comentar detalles de Servicio Pospuesto
             builder.Prompts.choice(session, `**Elije el motivo por el cual se pospone el servicio.**`,[Motivos.Uno, Motivos.Dos, Motivos.Tres, Motivos.Cuatro, Motivos.Cinco], { listStyle: builder.ListStyle.button });
             break;
         }
         
     },
     function (session, results) {
         // Sexto diálogo
         session.dialogData.comentarios = results.response.entity;
         var msg = session.message;
         if (msg.attachments && msg.attachments.length > 0) {
             // Echo back attachment
             var attachment = msg.attachments[0];
             session.send({
                 "attachments": [
                     {
                     "contentType": attachment.contentType,
                     "contentUrl": attachment.contentUrl,
                     "name": attachment.name
                     }
                 ],});
 
             var stype = attachment.contentType.split('/');
             var ctype = stype[1];
             var url = attachment.contentUrl;
             image2base64(url)
                 .then(
                     (response) => {
                         // console.log(response); //iVBORw0KGgoAAAANSwCAIA...
                         var buffer = new Buffer(response, 'base64');
                     blobService.createBlockBlobFromText(config.blobcontainer, session.dialogData.proyecto+'_'+session.dialogData.serie+'_'+session.dialogData.tipo+'_'+session.dialogData.asociado+'.'+ctype, buffer,  function(error, result, response) {
                         if (!error) {
                             console.log(Discriptor);
                             tableService.mergeEntity(config.table1, Discriptor, function(err, res, respons) {
                                 if (!err) {
                                     console.log(`entity property ${session.dialogData.tipo} updated`);
                                 Discriptor = {};
                                 }
                                 else{err}
                             });
                            
                             session.send(`El archivo **${session.dialogData.proyecto}_${session.dialogData.serie}_${session.dialogData.tipo}.${ctype}** se ha subido correctamente`);
                             builder.Prompts.choice(session, '¿Deseas adjuntar Evidencia o Documentación?', [Choice.Si, Choice.No], { listStyle: builder.ListStyle.button });
                         }
                         else{
                             console.log('Hubo un error: '+ error);
                             
                         }
                     });
                     }
                 )
                 .catch(
                     (error) => {
                         console.log(error); //Exepection error....
                     }
                 );
         } else {
                tableService.retrieveEntity(config.table1, session.dialogData.asociado, session.dialogData.serie, function(eror, result, response) {
                    if (!eror) {                    
                        nodeoutlook.sendEmail({
                            auth: {
                                user: `${config.email1}`,
                                pass: `${config.pass}`,
                            }, from: `${config.email1}`,
                            to: `${config.email3}, ${config.email2} `,
                            subject: 'Incidente de Servicio',
                            html: `<p>El servicio se pospuso por el siguiente motivo: <br><h3> <blockquote>${session.dialogData.comentarios}</blockquote> <br> <b>Asociado: ${session.dialogData.asociado}</b>  <br> <b>Proyecto: ${session.dialogData.proyecto}</b>  <br> <b>Serie: ${session.dialogData.serie}</b> <br> <b>Servicio: ${result.Servicio._}</b> <br> <b>Localidad: ${result.Localidad._}</b> <br> <b>Inmueble: ${result.Inmueble._}</b> </h3> </p><br><p>Saludos.</p>`
                           });
                    }
                    else{
                        clearTimeout(time);
                        session.endConversation("**Error** La serie no coincide con el Asociado.");
                    }
                });
                 // Echo back users text
                     function appendPospuesto() {
                         Discriptor.PartitionKey = {'_': session.dialogData.asociado, '$':'Edm.String'};
                         Discriptor.RowKey = {'_': session.dialogData.serie, '$':'Edm.String'};
                         Discriptor.Pospuesto = {'_': session.dialogData.comentarios, '$':'Edm.String'};
                         
                     };
                     appendPospuesto();
                     tableService.mergeEntity(config.table1, Discriptor, function(err, res, respons) {
                         if (!err) {
                             console.log(`entity property ${session.dialogData.tipo} updated`);
                         Discriptor = {};
                         }
                         else{err}
                     });
                clearTimeout(time);
                session.endConversation("**Hemos terminado por ahora, Se enviarán tus observaciones por correo.**");
         }
     },
     function (session, results) {
         // Cuarto diálogo
         var selection3 = results.response.entity;
         switch (selection3) {
             
             case Choice.Si:
             builder.Prompts.choice(session, '¿Que tipo de Evidencia o Documentación?', [Opts.Resguardo, Opts.Check, Opts.Borrado, Opts.Baja], { listStyle: builder.ListStyle.button });  
             break;
 
             case Choice.No:
             clearTimeout(time);
             session.endConversation("De acuerdo, hemos terminado por ahora");
             break;
         }
         
     },
     function (session, results) {
         // Quinto diálogo
         var selection2 = results.response.entity;
         session.dialogData.tipo = selection2;
         session.dialogData.Discriptor ={};
         switch (selection2) {
 
             case Opts.Resguardo:
             function appendResguardo() {
                 Discriptor.PartitionKey = {'_': session.dialogData.asociado, '$':'Edm.String'};
                 Discriptor.RowKey = {'_': session.dialogData.serie, '$':'Edm.String'};
                 Discriptor.Resguardo = {'_': 'Resguardo Adjunto', '$':'Edm.String'};
             };
             appendResguardo();
             builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Resguardo}**`);
             break;
 
             case Opts.Borrado:
             function appendBorrado() {
                 Discriptor.PartitionKey = {'_': session.dialogData.asociado, '$':'Edm.String'};
                 Discriptor.RowKey = {'_': session.dialogData.serie, '$':'Edm.String'};
                 Discriptor.Borrado = {'_': 'Borrado Adjunto', '$':'Edm.String'};
             };
             appendBorrado();
             builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Borrado}**`);
             break;
 
             case Opts.Baja:
             function appendBaja() {
                 Discriptor.PartitionKey = {'_': session.dialogData.asociado, '$':'Edm.String'};
                 Discriptor.RowKey = {'_': session.dialogData.serie, '$':'Edm.String'};
                 Discriptor.Baja = {'_': 'Baja Adjunto', '$':'Edm.String'};
             };
             appendBaja();
             builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Baja}**`);
             break;
 
             case Opts.Check:
             function appendCheck() {
                 Discriptor.PartitionKey = {'_': session.dialogData.asociado, '$':'Edm.String'};
                 Discriptor.RowKey = {'_': session.dialogData.serie, '$':'Edm.String'};
                 Discriptor.Check = {'_': 'Check Adjunto', '$':'Edm.String'};
                 
             };
             appendCheck();
             builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Check}**`);
             break;
         }
         
     },
     function (session) {
         // Sexto diálogo
         var msg = session.message;
         if (msg.attachments && msg.attachments.length > 0) {
             // Echo back attachment
             var attachment = msg.attachments[0];
             session.send({
                 "attachments": [
                     {
                     "contentType": attachment.contentType,
                     "contentUrl": attachment.contentUrl,
                     "name": attachment.name
                     }
                 ],});
 
             var stype = attachment.contentType.split('/');
             var ctype = stype[1];
             var url = attachment.contentUrl;
             image2base64(url)
                 .then(
                     (response) => {
                         // console.log(response); //iVBORw0KGgoAAAANSwCAIA...
                         var buffer = new Buffer(response, 'base64');
                     blobService.createBlockBlobFromText(config.blobcontainer, session.dialogData.proyecto+'_'+session.dialogData.serie+'_'+session.dialogData.tipo+'_'+session.dialogData.asociado+'.'+ctype, buffer,  function(error, result, response) {
                             if (!error) {
                             console.log(Discriptor);
                             tableService.mergeEntity(config.table1, Discriptor, function(err, res, respons) {
                                 if (!err) {
                                     console.log(`entity property ${session.dialogData.tipo} updated`);
                                 Discriptor = {};
                                 }
                                 else{err}
                             });
                             
                             session.send(`El archivo **${session.dialogData.proyecto}_${session.dialogData.serie}_${session.dialogData.tipo}.${ctype}** se ha subido correctamente`);
                             session.endConversation('Hemos terminado por ahora.');
                         }
                         else{
                             console.log('Hubo un error: '+ error);
                             
                         }
                     });
                     }
                 )
                 .catch(
                     (error) => {
                         console.log(error); //Exepection error....
                     }
                 );
         } else {
                 // Echo back users text
                 session.send("Enviaste esto en ves de una imagen: %s", session.message.text);
         }
 
     }
 ]);
 // Cancela la operación en cualquier momento
 bot.dialog('cancel',
     function (session) {
         clearTimeout(time);
         session.endConversation('No hay problema, volvamos a iniciar de nuevo.');
         session.replaceDialog('/');
     }
 ).triggerAction(
     {matches: /(cancel|cancelar)/gi}
 );