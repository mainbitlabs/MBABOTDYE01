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
var Docs = {
    Evidencia: 'Adjuntar Documentación',
    Incidente: 'Reportar Incidente'
 };
var optsbutton = [];
 var Motivos = {
    Uno: 'Usuario',
    Dos: 'Documentos',
    Tres: 'Infraestructura',
    Cuatro: 'Equipo',
    Cinco: 'Servicio',
 };

 var Opts = {};

 var time;
 // Variable Discriptor para actualizar tabla
 var Discriptor = {};
 // El díalogo principal inicia aquí
 bot.dialog('/', [
     function (session) {
         // Primer diálogo    
         session.send(`Hola bienvenido al Servicio Automatizado de Mainbit.`);
         session.send(`**Sugerencia:** Recuerda que puedes cancelar en cualquier momento escribiendo **"cancelar".** \n\n **Importante:** este bot tiene un ciclo de vida de 10 minutos, te recomendamos concluir la actividad antes de este periodo.`);
         builder.Prompts.text(session, 'Por favor, **escribe el Número de Serie del equipo.**');
         time = setTimeout(() => {
             session.endConversation(`**Lo sentimos ha transcurrido el tiempo estimado para completar esta actividad. Intentalo nuevamente.**`);
         }, 600000);
     },
     function (session, results) {
         // Segundo diálogo
         session.privateConversationData.serie = results.response;
         builder.Prompts.text(session, '¿Cuál es tu **Clave de Asociado**?');
     },
     function (session, results) {
         session.privateConversationData.asociado = results.response;
         // Tercer diálogo
         tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(error, result, response) {
             if(!error && result.Resguardo._ === 'Resguardo Adjunto' && result.Baja._ === 'Baja Adjunto' && result.Check._ === 'Check Adjunto' && result.Borrado._ === 'Borrado Adjunto'  && result.HojaDeServicio._ === 'Hoja de Servicio Adjunto') {
                 var Estatus = {
                     PartitionKey : {'_': session.privateConversationData.asociado, '$':'Edm.String'},
                     RowKey : {'_': session.privateConversationData.serie, '$':'Edm.String'},
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
                //  clearTimeout(time);
                //  session.endConversation("**Error** Los ");
             }
         });
         session.sendTyping();
             // Envíamos un mensaje al usuario para que espere.
             session.send('Estamos atendiendo tu solicitud. Por favor espera un momento...');
             setTimeout(() => {
         tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(eror, result, response) {
             if (!eror) {                    
                 session.privateConversationData.proyecto= result.Proyecto._;
                 session.send(`**Proyecto:** ${result.Proyecto._} \n\n **Número de Serie**: ${result.RowKey._} \n\n **Asociado:** ${result.PartitionKey._}  \n\n  **Descripción:** ${result.Descripcion._} \n\n  **Localidad:** ${result.Localidad._} \n\n  **Inmueble:** ${result.Inmueble._} \n\n  **Servicio:** ${result.Servicio._} \n\n  **Resguardo:** ${result.Resguardo._} \n\n  **Check:** ${result.Check._} \n\n  **Borrado:** ${result.Borrado._} \n\n  **Baja:** ${result.Baja._} \n\n  **Hoja de Servicio:** ${result.HojaDeServicio._}`);
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
             builder.Prompts.choice(session, '¿Qué deseas realizar?', [Docs.Evidencia, Docs.Incidente], { listStyle: builder.ListStyle.button });
             break;
 
             case Choice.No:
             clearTimeout(time);
             session.endConversation("Por favor valida con tu soporte que el Número de Serie esté asignado a tu Asociado");
             break;
         }
         
     },
     function (session, results) {
         // Quinto diálogo
         var selection3 = results.response.entity;
         switch (selection3) {
             
             case Docs.Evidencia:
             tableService.retrieveEntity(config.table4, "Proyecto", session.privateConversationData.proyecto, function(error, result, response) {
                if(!error) {
                    Opts.Ubicacion = "Reportar llegada a Sitio";
                    optsbutton.push(Opts.Ubicacion);
                    if (result.Baja._ == "X") {
                        Opts.Baja="Baja";
                        optsbutton.push(Opts.Baja);
                    }
                    if (result.Borrado._ == "X") {
                        Opts.Borrado="Borrado";
                        optsbutton.push(Opts.Borrado);
                    }
                    if (result.Check._ == "X") {
                        Opts.Check="Check";
                        optsbutton.push(Opts.Check);
                    }
                    if (result.Resguardo._ == "X") {
                        Opts.Resguardo ="Resguardo";
                        optsbutton.push(Opts.Resguardo);
                    }
                    if (result.HojaDeServicio._ == "X") {
                        Opts.Hoja ="HojaDeServicio";
                        optsbutton.push(Opts.Hoja);
                    }
                  console.log(optsbutton);
                  console.log(Opts);
                  builder.Prompts.choice(session, 'Que tipo de Evidencia o Documentación deseas adjuntar: ', optsbutton, {listStyle: builder.ListStyle.button});
                  
                } 
                else{
                   //  clearTimeout(time);
                   //  session.endConversation("**Error**");
                }
            });
            //  session.send('aqui deben ir las opciones');
            //  builder.Prompts.choice(session, '¿Que tipo de Evidencia o Documentación?', [Opts.Resguardo, Opts.Check,  Opts.Baja, Opts.Borrado, Opts.Hoja, Opts.Pospuesto], { listStyle: builder.ListStyle.button });  
            Opts={};
            optsbutton=[]; 
            break;
 
             case Docs.Incidente:
             session.beginDialog("incidente");
             break;
         }
         
     },
     function (session, results) {
         // Sexto diálogo
         var selection2 = results.response.entity;
         session.privateConversationData.tipo = selection2;
         session.privateConversationData.Discriptor ={};
         switch (selection2) {
 
             case Opts.Resguardo:
             tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(eror, result, response) {
                if (!eror) {
                    if(result.Resguardo._== "Aprobado" ){
                        clearTimeout(time);
                        session.endConversation("**No puedes adjuntar el archivo, este documento ya ha sido aprobado.** \n**Hemos concluido por ahora.**");
                    }
                    // else if(result.Resguardo._== "En validacion"){
                    //     clearTimeout(time);
                    //     session.endConversation("**No puedes adjuntar el archivo, el documento ya esta en proceso de validación.** \n**Hemos concluido por ahora.**");
                    // }
                    else{
                        builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Resguardo}**`);
                    }                  
                }
                else{
                    clearTimeout(time);
                    session.endConversation("**Error** La serie no coincide con el Asociado.");
                }
                });
             break;
 
             case Opts.Borrado:
             tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(eror, result, response) {
                if (!eror) {
                    if(result.Borrado._== "Aprobado" ){
                        clearTimeout(time);
                        session.endConversation("**No puedes adjuntar el archivo, este documento ya ha sido aprobado.** \n**Hemos concluido por ahora.**");
                    }
                    else{
                        session.beginDialog('borrado', session.privateConversationData.sborrado);//Llama al dialogo externo "borrado"
                    }                  
                }
                else{
                    clearTimeout(time);
                    session.endConversation("**Error** La serie no coincide con el Asociado.");
                }
                });
             break;
 
             case Opts.Baja:
             tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(eror, result, response) {
                if (!eror) {
                    if(result.Baja._== "Aprobado" ){
                        clearTimeout(time);
                        session.endConversation("**No puedes adjuntar el archivo, este documento ya ha sido aprobado.** \n**Hemos concluido por ahora.**");
                    }
                    else{
                        builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Baja}**`);
                    }                  
                }
                else{
                    clearTimeout(time);
                    session.endConversation("**Error** La serie no coincide con el Asociado.");
                }
                });
             break;
 
             case Opts.Check:
             tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(eror, result, response) {
                if (!eror) {
                    if(result.Check._== "Aprobado" ){
                        clearTimeout(time);
                        session.endConversation("**No puedes adjuntar el archivo, este documento ya ha sido aprobado.** \n**Hemos concluido por ahora.**");
                    }
                    else{
                        builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Check}**`);
                    }                  
                }
                else{
                    clearTimeout(time);
                    session.endConversation("**Error** La serie no coincide con el Asociado.");
                }
                });
             break;
             
             case Opts.Hoja:
             tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(eror, result, response) {
                if (!eror) {
                    if(result.HojaDeServicio._== "Aprobado" ){
                        clearTimeout(time);
                        session.endConversation("**No puedes adjuntar el archivo, este documento ya ha sido aprobado.** \n**Hemos concluido por ahora.**");
                    }
                    else{
                        builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Hoja}**`);
                    }                  
                }
                else{
                    clearTimeout(time);
                    session.endConversation("**Error** La serie no coincide con el Asociado.");
                }
                });
             break;
             case Opts.Ubicacion:
                // session.send("[Ubicación actual](https://mainbitlabs.github.io/)");
                session.send("Comparte tu ubicación actual");
                session.beginDialog('location');
                break;
             
             case Motivos.Uno:
            session.privateConversationData.X = Motivos.Uno;
             // Comentar detalles de Servicio Pospuesto
            builder.Prompts.text(session, 'Escribe tus observaciones');
             break;

             case Motivos.Dos:
             session.privateConversationData.X = Motivos.Dos;
             // Comentar detalles de Servicio Pospuesto
             builder.Prompts.text(session, 'Escribe tus observaciones');
             break;
             
             case Motivos.Tres:
             session.privateConversationData.X = Motivos.Tres;
             // Comentar detalles de Servicio Pospuesto
             builder.Prompts.text(session, 'Escribe tus observaciones');
             break;
            
             case Motivos.Cuatro:
             session.privateConversationData.X = Motivos.Cuatro;
             // Comentar detalles de Servicio Pospuesto
             builder.Prompts.text(session, 'Escribe tus observaciones');
             break;
             
             case Motivos.Cinco:
             session.privateConversationData.X = Motivos.Cinco;
             // Comentar detalles de Servicio Pospuesto
             builder.Prompts.text(session, 'Escribe tus observaciones');
             break;
         }
         
     },
     function (session, results, next) {
        // si el tipo es = borrado usa este diálogo
        if (session.privateConversationData.tipo == 'Borrado') {
            session.privateConversationData.sborrado = results.response;
            var Serie = {}; //Objeto para actualizar serie borrada
            function borrado() {
                Serie.PartitionKey = {'_': session.privateConversationData.asociado, '$':'Edm.String'};
                Serie.RowKey = {'_': session.privateConversationData.serie, '$':'Edm.String'};
                Serie.SerieBorrada = {'_': session.privateConversationData.sborrado, '$':'Edm.String'};
                
            };
            borrado();
            tableService.mergeEntity(config.table1, Serie, function(err, res, respons) {
                if (!err) {
                    console.log(`entity property ${session.privateConversationData.tipo} updated`);
                    function appendBorrado() {
                        Discriptor.PartitionKey = {'_': session.privateConversationData.asociado, '$':'Edm.String'};
                        Discriptor.RowKey = {'_': session.privateConversationData.serie, '$':'Edm.String'};
                        Discriptor.Borrado = {'_': 'Borrado Adjunto', '$':'Edm.String'};
                    };
                    // appendBorrado();
                    //Vacía el descriptor para volver a ser utilizado
                    Discriptor = {};
                    builder.Prompts.attachment(session, `**Adjunta aquí documento de ${Opts.Borrado}**`);
                }
                else{err} 
            });
            // session.send('elegiste Borrado');
        }
        else{ //si el tipo != borrado salta este diálogo
            session.privateConversationData.comentarios = results.response;
            console.log('Comentarios ', session.privateConversationData.comentarios);
            console.log('Next function');
            next();
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
                     blobService.createBlockBlobFromText(config.blobcontainer, session.privateConversationData.proyecto+'_'+session.privateConversationData.serie+'_'+session.privateConversationData.tipo+'_'+session.privateConversationData.asociado+'.'+ctype, buffer,  function(error, result, response) {
                         if (!error) {
                            //  console.log(Discriptor);
                            //  tableService.mergeEntity(config.table1, Discriptor, function(err, res, respons) {
                            //      if (!err) {
                            //          console.log(`entity property ${session.privateConversationData.tipo} updated`);
                            //     //  vacia el contenido del Discriptor para volver a ser usado
                            //      Discriptor = {};
                            //      }
                            //      else{err}
                            //  });
                            
                             session.send(`El archivo **${session.privateConversationData.proyecto}_${session.privateConversationData.serie}_${session.privateConversationData.tipo}.${ctype}** se ha subido correctamente`);
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
                tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(eror, result, response) {
                    if (!eror) {  
                        // Update Comentarios Azure
                        var now = new Date();
                        now.setHours(now.getHours()-5);
                        var dateNow = now.toLocaleString();
                        function appendPospuesto() {
                            Discriptor.PartitionKey = {'_': session.privateConversationData.asociado, '$':'Edm.String'};
                            Discriptor.RowKey = {'_': session.privateConversationData.serie, '$':'Edm.String'};
                            Discriptor.Pospuesto = {'_':dateNow +' '+session.privateConversationData.X +' '+ session.privateConversationData.comentarios+'\n'+result.Pospuesto._, '$':'Edm.String'};
                            
                        };
                        appendPospuesto();
                        tableService.mergeEntity(config.table1, Discriptor, function(err, res, respons) {
                            if (!err) {
                                // Correo de Incidentes 
                                nodeoutlook.sendEmail({
                                    auth: {
                                        user: `${config.email1}`,
                                        pass: `${config.pass}`,
                                    }, from: `${config.email1}`,
                                    to: `${config.email1}, ${config.email2}, ${config.email3}, ${config.email4}  `,
                                    subject: `${result.Proyecto._} Incidente de ${session.privateConversationData.X}: ${result.RowKey._} / ${result.Servicio._}`,
                                    html: `<p>El servicio se pospuso por el siguiente motivo:</p> <br> <b>${session.privateConversationData.X}</b> <br> <b><blockquote>${session.privateConversationData.comentarios}</blockquote></b> <br> <b>Proyecto: ${result.Proyecto._}</b>  <br> <b>Serie: ${result.RowKey._}</b> <br> <b>Servicio: ${result.Servicio._}</b> <br> <b>Localidad: ${result.Localidad._}</b> <br> <b>Inmueble: ${result.Inmueble._}</b> <br> <b>Nombre de Usuario: ${result.NombreUsuario._}</b> <br> <b>Area: ${result.Area._}</b>`
                                   });
                                
                                    console.log(`Incidente de "${session.privateConversationData.tipo}" actualizado y enviado por correo.`);
                                   Discriptor = {};
                                   clearTimeout(time);
                                   session.endConversation("**Hemos terminado por ahora, Se enviarán tus observaciones por correo.**");
                                   }
                                   else{
                                   console.log("Merge Entity Error: ",err);
                                       
                                   }
                               });
                    }
                    else{
                        clearTimeout(time);
                        session.endConversation("**Error** La serie no coincide con el Asociado.");
                    }
                });
                
         }
     },
     function (session, results) {
         // Cuarto diálogo
         var selection3 = results.response.entity;
         switch (selection3) {
             
             case Choice.Si:
             builder.Prompts.choice(session, '¿Que tipo de Evidencia o Documentación?', optsbutton, { listStyle: builder.ListStyle.button });  
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
         session.privateConversationData.tipo = selection2;
         session.privateConversationData.Discriptor ={};
         switch (selection2) {
 
            case Opts.Resguardo:
            tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(eror, result, response) {
               if (!eror) {
                   if(result.Resguardo._== "Aprobado" ){
                       clearTimeout(time);
                       session.endConversation("**No puedes adjuntar el archivo, este documento ya ha sido aprobado.** \n**Hemos concluido por ahora.**");
                   }
                   // else if(result.Resguardo._== "En validacion"){
                   //     clearTimeout(time);
                   //     session.endConversation("**No puedes adjuntar el archivo, el documento ya esta en proceso de validación.** \n**Hemos concluido por ahora.**");
                   // }
                   else{
                       builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Resguardo}**`);
                   }                  
               }
               else{
                   clearTimeout(time);
                   session.endConversation("**Error** La serie no coincide con el Asociado.");
               }
               });
            break;
 
            case Opts.Borrado:
            tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(eror, result, response) {
               if (!eror) {
                   if(result.Borrado._== "Aprobado" ){
                       clearTimeout(time);
                       session.endConversation("**No puedes adjuntar el archivo, este documento ya ha sido aprobado.** \n**Hemos concluido por ahora.**");
                   }
                   else{
                       session.beginDialog('borrado', session.privateConversationData.sborrado);//Llama al dialogo externo "borrado"
                   }                  
               }
               else{
                   clearTimeout(time);
                   session.endConversation("**Error** La serie no coincide con el Asociado.");
               }
               });
            break;
 
            case Opts.Baja:
            tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(eror, result, response) {
               if (!eror) {
                   if(result.Baja._== "Aprobado" ){
                       clearTimeout(time);
                       session.endConversation("**No puedes adjuntar el archivo, este documento ya ha sido aprobado.** \n**Hemos concluido por ahora.**");
                   }
                   else{
                       builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Baja}**`);
                   }                  
               }
               else{
                   clearTimeout(time);
                   session.endConversation("**Error** La serie no coincide con el Asociado.");
               }
               });
            break;
 
            case Opts.Check:
            tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(eror, result, response) {
               if (!eror) {
                   if(result.Check._== "Aprobado" ){
                       clearTimeout(time);
                       session.endConversation("**No puedes adjuntar el archivo, este documento ya ha sido aprobado.** \n**Hemos concluido por ahora.**");
                   }
                   else{
                       builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Check}**`);
                   }                  
               }
               else{
                   clearTimeout(time);
                   session.endConversation("**Error** La serie no coincide con el Asociado.");
               }
               });
            break;
            
            case Opts.Hoja:
            tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(eror, result, response) {
               if (!eror) {
                   if(result.HojaDeServicio._== "Aprobado" ){
                       clearTimeout(time);
                       session.endConversation("**No puedes adjuntar el archivo, este documento ya ha sido aprobado.** \n**Hemos concluido por ahora.**");
                   }
                   else{
                       builder.Prompts.attachment(session, `**Adjunta aquí ${Opts.Hoja}**`);
                   }                  
               }
               else{
                   clearTimeout(time);
                   session.endConversation("**Error** La serie no coincide con el Asociado.");
               }
               });
            break;
         }
         
     },
     function (session, results, next) {
        // si el tipo es = borrado usa este diálogo
        if (session.privateConversationData.tipo == 'Borrado') {
            session.privateConversationData.sborrado = results.response;
            var Serie = {}; //Objeto para actualizar serie borrada
            function borrado() {
                Serie.PartitionKey = {'_': session.privateConversationData.asociado, '$':'Edm.String'};
                Serie.RowKey = {'_': session.privateConversationData.serie, '$':'Edm.String'};
                Serie.SerieBorrada = {'_': session.privateConversationData.sborrado, '$':'Edm.String'};
                
            };
            borrado();
            tableService.mergeEntity(config.table1, Serie, function(err, res, respons) {
                if (!err) {
                    console.log(`entity property ${session.privateConversationData.tipo} updated`);
                    function appendBorrado() {
                        Discriptor.PartitionKey = {'_': session.privateConversationData.asociado, '$':'Edm.String'};
                        Discriptor.RowKey = {'_': session.privateConversationData.serie, '$':'Edm.String'};
                        Discriptor.Borrado = {'_': 'Borrado Adjunto', '$':'Edm.String'};
                    };
                    // appendBorrado();
                    //Vacía el descriptor para volver a ser utilizado
                    Discriptor = {};
                    builder.Prompts.attachment(session, `**Adjunta aquí documento de ${Opts.Borrado}**`);
                }
                else{err} 
            });
            // session.send('elegiste Borrado');
        }
        else{ //si el tipo != borrado salta este diálogo
            session.privateConversationData.comentarios = results.response;
            console.log('Comentarios ', session.privateConversationData.comentarios);
            console.log('Next function');
            next();
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
                     blobService.createBlockBlobFromText(config.blobcontainer, session.privateConversationData.proyecto+'_'+session.privateConversationData.serie+'_'+session.privateConversationData.tipo+'_'+session.privateConversationData.asociado+'.'+ctype, buffer,  function(error, result, response) {
                             if (!error) {
                            //  console.log(Discriptor);
                            //  tableService.mergeEntity(config.table1, Discriptor, function(err, res, respons) {
                            //      if (!err) {
                            //          console.log(`entity property ${session.privateConversationData.tipo} updated`);
                            //      Discriptor = {};
                            //      }
                            //      else{err}
                            //  });
                             
                             session.send(`El archivo **${session.privateConversationData.proyecto}_${session.privateConversationData.serie}_${session.privateConversationData.tipo}.${ctype}** se ha subido correctamente`);
                             clearTimeout(time);
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
 bot.dialog('incidente',
    function (session, next) {
        builder.Prompts.choice(session, `**Elije el motivo por el cual se pospone el servicio.**`,[Motivos.Uno, Motivos.Dos, Motivos.Tres, Motivos.Cuatro, Motivos.Cinco], { listStyle: builder.ListStyle.button });
    }
 );
 bot.dialog('borrado', //dialogo externo "borrado"
 function (session) {
     builder.Prompts.text(session, 'Ingresa el número de serie de borrado');
 },
 function (session, results) {
     session.privateConversationData.sborrado = results.response;
     session.endDialogWithResult({ response: session.privateConversationData.sborrado });}
);
bot.dialog("location", [
    function (session) {
       
       if (session.message.text == "") { 
        //    console.log("<<< Imposible >>>", session.message.entities);
        // console.log("<<< core_company: "+ session.privateConversationData.company);
        // console.log("<<< typeof_company: "+ typeof(session.privateConversationData.company));
        // console.log("<<< Session >>>", session);
           console.log("<<< Session.message.user >>>", session.message.user);
           console.log("<<< Latitude >>>", session.message.entities[0].geo.latitude);
           console.log("<<< Longitude >>>", session.message.entities[0].geo.longitude);

           var d = new Date();
            var m = d.getMonth() + 1;
            var fecha = d.getFullYear() + "-" + m + "-" + d.getDate() + "-" + d.getHours() + ":" + d.getMinutes() + ":" + d.getSeconds();
            // var descriptor = {
            //     PartitionKey: {'_': session.privateConversationData.asociado, '$':'Edm.String'},
            //     RowKey: {'_': session.privateConversationData.serie, '$':'Edm.String'},
            //     Fecha: {'_': fecha, '$':'Edm.String'},
            //     Latitud: {'_': session.message.entities[0].geo.latitude, '$':'Edm.String'},
            //     Longitud: {'_': session.message.entities[0].geo.longitude, '$':'Edm.String'},
            //     Historico: {'_': fecha +" "+ session.message.entities[0].geo.latitude + " " + session.message.entities[0].geo.longitude+"\n", '$':'Edm.String'},
            //     GPS: {'_': 'https://www.google.com.mx/maps/search/ '+ session.message.entities[0].geo.latitude + "," + session.message.entities[0].geo.longitude, '$':'Edm.String'},
            // };
            
    
    tableService.retrieveEntity(config.table1, session.privateConversationData.asociado, session.privateConversationData.serie, function(error, result, response) {
        if (!error) {
            
            var now = new Date();
            now.setHours(now.getHours()-5);
            var dateNow = now.toLocaleString();
            var lat = session.message.entities[0].geo.latitude.toString();
            var long = session.message.entities[0].geo.longitude.toString();
            console.log("<<< Type of Latitude >>>", typeof lat);
            console.log("<<< Type of Longitude >>>", typeof long);
            var merge = {
                PartitionKey: {'_': session.privateConversationData.asociado, '$':'Edm.String'},
                RowKey: {'_': session.privateConversationData.serie, '$':'Edm.String'},
                Latitud: {'_': lat, '$':'Edm.String'},
                Longitud: {'_': long, '$':'Edm.String'},
                GPS: {'_': dateNow +' '+ 'https://www.google.com.mx/maps/search/'+ session.message.entities[0].geo.latitude + "," + session.message.entities[0].geo.longitude+'\n' + result.GPS._, '$':'Edm.String'}

            };
            

                tableService.mergeEntity(config.table1, merge, function(err, res, respons) {
                    if (!err) {
                        // Correo Reporte de llegada a Sitio
                        nodeoutlook.sendEmail({
                            auth: {
                                user: `${config.email1}`,
                                pass: `${config.pass}`,
                            }, from: `${config.email1}`,
                            to: `${config.email3}, ${config.email2}`,
                            subject: `${session.privateConversationData.proyecto} Check-In: ${session.privateConversationData.serie} / ${result.Servicio._}`,
                            html: `<p>Se reporta la llegada al sitio:</p> <br><br> <b>Proyecto: ${session.privateConversationData.proyecto}</b>  <br> <b>Serie: ${session.privateConversationData.serie}</b> <br> <b>Servicio: ${result.Servicio._}</b> <br> <b>Localidad: ${result.Localidad._}</b> <br> <b>Inmueble: ${result.Inmueble._}</b> <br> <b>Nombre de Usuario: ${result.NombreUsuario._}</b> <br> <b>Area: ${result.Area._}</b> <br> <b>Mapa: <a href="https://www.google.com.mx/maps/search/${session.message.entities[0].geo.latitude}, ${session.message.entities[0].geo.longitude}">Ir al mapa</a> </b> `
                           });
                       console.log("Merge Entity Latitud y Longitud");
                       clearTimeout(time);
                       session.endConversation("Gracias, tu ubicación ha sido registrada.");
                    }
                    else{
                    console.log(err);
                    } 
                });
                
          
    
        } else {
            console.log(error);
            
        }
    });


                    
                }else{
                    console.log("<<< Session.message >>>", session.message);

                    
                }
          
            
       } 
    
    
]);