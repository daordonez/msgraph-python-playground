from configparser import SectionProxy
from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.users_request_builder import UsersRequestBuilder
from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder

class Graph:
    settings: SectionProxy
    client_credential: ClientSecretCredential
    app_client: GraphServiceClient

    def __init__(self, config: SectionProxy):
        self.settings = config
        client_id = self.settings['clientId']
        tenant_id = self.settings['tenantId']
        client_secret = self.settings['clientSecret']

        self.client_credential = ClientSecretCredential(tenant_id, client_id, client_secret)
        self.app_client = GraphServiceClient(self.client_credential) # type: ignore
        
    async def get_app_only_token(self):
        graph_scope = 'https://graph.microsoft.com/.default'
        access_token = await self.client_credential.get_token(graph_scope)
        return access_token.token
    
    async def get_users(self):
        query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            select = ['displayName', 'id', 'mail'],
            top = 25,
            orderby = ['displayName']
        )
        
        request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )
        
        users = await self.app_client.users.get(request_configuration=request_config)
        
        return users

    async def get_users_messages(self):
        request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration()
        
        request_config.headers.add('prefer', 'outlook.body-content-type=text')
        
        messages = await self.app_client.users.by_user_id('adelev@tms.365enespanol.com').messages.get(request_configuration=request_config)
        
        return messages