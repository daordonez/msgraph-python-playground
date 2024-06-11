from configparser import SectionProxy
from datetime import datetime
import base64
from msgraph import GraphServiceClient # type: ignore
from azure.identity.aio import ClientSecretCredential # type: ignore
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import SendMailPostRequestBody # type: ignore
from msgraph.generated.models.message import Message # type: ignore
from msgraph.generated.models.item_body import ItemBody # type: ignore
from msgraph.generated.models.body_type import BodyType # type: ignore
from msgraph.generated.models.recipient import Recipient # type: ignore
from msgraph.generated.models.email_address import EmailAddress # type: ignore
from kiota_abstractions.native_response_handler import NativeResponseHandler # type: ignore
from kiota_http.middleware.options import ResponseHandlerOption # type: ignore
from msgraph.generated.models.single_value_legacy_extended_property import SingleValueLegacyExtendedProperty # type: ignore
from msgraph.generated.models.file_attachment import FileAttachment
from msgraph.generated.users.item.messages.item.attachments.item.attachment_item_request_builder import AttachmentItemRequestBuilder
from kiota_abstractions.base_request_configuration import RequestConfiguration
from MessageAttachment import MessageAttachment




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
        await self.app_client.users.by_user_id('diego@tms.365enespanol.com').send_mail.post(request_body)
        
    async def get_parent_folder_id(self, user_id):
        result = await self.app_client.users.by_user_id(user_id).mail_folders.by_mail_folder_id('Inbox').get()
        return result
    
    async def get_target_messages(self, user_target_id):
        target_messages = await self.app_client.users.by_user_id(user_target_id).messages.get()
        return target_messages
    
    async def get_message_attachments(self, user_id_att, message_id):
        
        #Expand file properties
        query_params = AttachmentItemRequestBuilder.AttachmentItemRequestBuilderGetQueryParameters(
		expand = ["microsoft.graph.itemattachment/item"],
        )

        request_configuration = RequestConfiguration(
        query_parameters = query_params,
        )
        
        this_attachments = await self.app_client.users.by_user_id(user_id_att).messages.by_message_id(message_id).attachments.get()
        
        return this_attachments
    
    async def attach_message_files(self, user_id_source, user_id_target, message_id_source, message_id_target):
        
        all_attachments = []
        
        #Get message attachments in source
        msg_attachments = await self.get_message_attachments(user_id_source, message_id_source)
                    
                    
        for attach in msg_attachments.value:
            this_attachment = FileAttachment(
                odata_type = attach.odata_type,
                name = attach.name,
                content_bytes = attach.content_bytes,
                )
            #append all attachments
            attachment_response = await self.app_client.users.by_user_id(user_id_target).messages.by_message_id(message_id_target).attachments.post(this_attachment)
            all_attachments.append(attachment_response)
        
        return all_attachments
    
 
    async def copy_mail_message(self, user_id_source, user_id_target):
        
        #Responses array
        responses = []
        #Messages Array
        all_user_messages = []
        #messages with attachments
        messages_with_attachments = []
        
        #Get user_source mail messages
        source_messages = await self.app_client.users.by_user_id(user_id_source).messages.get()
        this_parent_folder = await self.get_parent_folder_id(user_id_source)
        
        #verify values
        if source_messages and source_messages.value:
            #Start copy
            for s_msg_ in source_messages.value:
                
                all_user_messages.append(s_msg_)
        
        if source_messages.odata_next_link:
                    
            has_next_link = source_messages.odata_next_link
                
            while has_next_link:
                
                print('New @odata.nextLink >>>')
                
                source_messages_page = await self.app_client.users.by_user_id(user_id_source).messages.with_url(has_next_link).get()
                
                has_next_link = source_messages_page.odata_next_link
                
                for s_msg_page in source_messages_page.value:
                    
                    all_user_messages.append(s_msg_page)
            
            #Messages creation
            for msg_to_copy in all_user_messages:
                
                if msg_to_copy.has_attachments:
                    msg_with_attch = MessageAttachment(msg_to_copy.internet_message_id, msg_to_copy.id ,msg_to_copy.has_attachments)
                    messages_with_attachments.append(msg_with_attch)
                               
                request_body = Message(
                    id = msg_to_copy.id,
                    subject = msg_to_copy.subject,
                    from_=msg_to_copy.from_,
                    sender = msg_to_copy.sender,
                    reply_to = msg_to_copy.reply_to,
                    to_recipients = msg_to_copy.to_recipients,
                    created_date_time = msg_to_copy.created_date_time,
                    last_modified_date_time = msg_to_copy.last_modified_date_time,
                    has_attachments = msg_to_copy.has_attachments,
                    attachments = msg_to_copy.attachments,
                    cc_recipients = msg_to_copy.cc_recipients,
                    bcc_recipients = msg_to_copy.bcc_recipients,
                    body = msg_to_copy.body,
                    conversation_id=msg_to_copy.conversation_id,
                    importance = msg_to_copy.importance,
                    is_draft = msg_to_copy.is_draft,
                    is_read = msg_to_copy.is_read,
                    parent_folder_id = this_parent_folder.parent_folder_id,
                    received_date_time = msg_to_copy.received_date_time,
                    internet_message_headers = msg_to_copy.internet_message_headers,
                    internet_message_id = msg_to_copy.internet_message_id,
                    web_link = msg_to_copy.web_link,
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
                
            #Messages attachments
            print('Launching attachments process over created messages')
            for msg_with_attch in messages_with_attachments:
                #get from all messages with attachment the one with the same internet_message_id
                id_to_attach_target = MessageAttachment.get_message_id_(msg_with_attch.internet_message_id, responses)
                #id to extract from source
                id_to_extract_source = msg_with_attch.source_message_id
                #attach files
                response_attach = await self.attach_message_files(user_id_source, user_id_target, id_to_extract_source, id_to_attach_target)
                    
                print(response_attach)
                
            
        return responses
                