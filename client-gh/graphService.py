from configparser import SectionProxy
from datetime import datetime
from msgraph import GraphServiceClient
from azure.identity.aio import ClientSecretCredential
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import SendMailPostRequestBody
from msgraph.generated.models.message import Message
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.email_address import EmailAddress
from kiota_abstractions.native_response_handler import NativeResponseHandler
from kiota_http.middleware.options import ResponseHandlerOption
from msgraph.generated.models.single_value_legacy_extended_property import SingleValueLegacyExtendedProperty


class Graph:
    settings: SectionProxy # Objeto que recibe client_id, tenant_id, client_secret
    client_credential: ClientSecretCredential #Objeto de tipo azure_identity para utlizarse en el GraphServiceClient
    app_client: GraphServiceClient #Objeto que contiene todos los metodos necesarios para hacer llamadas e interacturar con Graph
    
    #Configurar Service client con credenciales recibidas desde main
    def __init__(self, config: SectionProxy):
        self.settings = config #Asignar 'config', recibido por parametros al Objeto 'settings'
        
        #Comenzar a extraer los atributos del objeto 'config'. Este objeto se recibe mediante la función SectionProxy
        client_id = self.settings['clientId'] # extraer atributo 'clientId' recibido desde main a través del param 'config'
        tenant_id = self.settings['tenantId'] # extraer atributo 'tenantId' recibido desde main a través del param 'config'
        client_secret = self.settings['clientSecret'] # extraer atributo 'clientSecret' recibido desde main a través del param 'config'
        
        #Configurar objeto 'client_credential' con los parametros de autenticación
        self.client_credential = ClientSecretCredential(tenant_id, client_id, client_secret)
        
        #Configurar el GraphServiceClient declarado más arriba como 'app_client'
        self.app_client = GraphServiceClient(self.client_credential)
        
        
        #TODO
        #solicitar el scope 'default' lo que significa que los permisos que tenga configurados la API de Graph
        graph_scope = 'https://graph.microsoft.com/.default'
        
    async def get_users(self):
        #devolverá todos los parametros en crudo de usuarios
        users = await self.app_client.users.get()
        return users
        
    
    async def send_mail_message(self, mail_to):
        
        now = datetime.now()
        formatted_date = now.strftime('%y%m%d %H%M%S')
                
        subject_string = "This message sent with ms-graph-sdk"
        request_body = SendMailPostRequestBody(
            message = Message(
                subject = formatted_date + " - " + subject_string,
                body = ItemBody(
                    content_type = BodyType.Text,
                    content = "hello There! This is simple message sent with python 🐍",
                ),
                to_recipients = [
                    Recipient(
                        email_address = EmailAddress(
                            address = mail_to,
                        ),
                    ),
                ],
            ),
            save_to_sent_items = False,
        )
        await self.app_client.users.by_user_id('adelev@tms.365enespanol.com').send_mail.post(request_body)
        
    async def get_parent_folder_id(self, user_id):
        result = await self.app_client.users.by_user_id(user_id).mail_folders.by_mail_folder_id('Inbox').get()
        return result
    
    async def get_target_messages(self, user_target_id):
        target_messages = await self.app_client.users.by_user_id(user_target_id).messages.get()
        return target_messages
        
    async def copy_mail_message(self, user_id_source, user_id_target):
        
        #Responses array
        responses = []
        
        #Get user_source mail messages
        source_messages = await self.app_client.users.by_user_id(user_id_source).messages.get()
        
        #verify values
        if source_messages and source_messages.value:
            #Start copy
            for s_msg_index in range(10):

                source_msg = source_messages.value[s_msg_index]
                #print('MessageId: ', source_msg.id)
                #Parsing message
                #sender = EmailAddress()
                #sender.address = source_msg.
                
                this_parent_folder = await self.get_parent_folder_id(user_id_source)
                
                request_body = Message(
                    id = source_msg.id,
                    subject = source_msg.subject,
                    from_=source_msg.from_,
                    sender = source_msg.sender,
                    reply_to = source_msg.reply_to,
                    to_recipients = source_msg.to_recipients,
                    created_date_time = source_msg.created_date_time,
                    last_modified_date_time = source_msg.last_modified_date_time,
                    attachments = source_msg.attachments,
                    cc_recipients = source_msg.cc_recipients,
                    bcc_recipients = source_msg.bcc_recipients,
                    body = source_msg.body,
                    conversation_id=source_msg.conversation_id,
                    importance = source_msg.importance,
                    is_draft = source_msg.is_draft,
                    is_read = source_msg.is_read,
                    parent_folder_id = this_parent_folder.parent_folder_id,
                    received_date_time = source_msg.received_date_time,
                    internet_message_headers = source_msg.internet_message_headers,
                    internet_message_id = source_msg.internet_message_id,
                    web_link = source_msg.web_link,
                    #Avoids create message as a draft
                    single_value_extended_properties = [
                        SingleValueLegacyExtendedProperty(
                            id = "Integer 0x0E07",
                            value = "4",
                        ),
                    ],
                )
                
                action_copy_response = await self.app_client.users.by_user_id(user_id_target).mail_folders.by_mail_folder_id(this_parent_folder.display_name).messages.post(request_body)
                
                responses.append(action_copy_response)
            
        return responses
                