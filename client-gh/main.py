import time
import asyncio
import configparser
from graphService import Graph
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

async def main():
    # Construir objeto de tipo de ConfigParser capaz de leer el fichero 'config.cfg'
    config = configparser.ConfigParser()
    #cargar fichero de configuracion con parametros de conexión. En este caso esta en la raiz del script
    # y se llama 'config.cfg'
    config.read(['config.cfg'])
    #Indicar a config parser que debe traer el objeto de tipo ['azure']. La configuracion puede tener más objetos
    azure_settings = config['azure']
    
    #Crear un objeto de tipo instanciando la clase Graph y pasandole los parametros de autenticación
    #la clase los recibe y ejecutará el método '__init__' 
    #En adelante 'graph_client' se usara en todas las funciones que interactuen con MS Graph
    graph_client: Graph = Graph(azure_settings)
    
    
    choice = -1

    while choice != 0:
        print('Please choose one of the following options:')
        print('0. Exit')
        print('1. ListUsers')
        print('2. SendMailMessage')
        print('3. CopyMailMessages')
        print('4. GetMessages')
        print('5. ListUserMailboxFolder')

        try:
            choice = int(input())
        except ValueError:
            choice = -1

        try:
            if choice == 0:
                print('Goodbye...')
            elif choice == 1:
                await list_users(graph_client)
            elif choice == 2:
                mail_to = input('Enter mail_to address: ')
                iterations = int(input('Enter total message count to be sent: '))
                print('Send Mail message to '+ mail_to +'\n')
                await send_mail_message(graph_client, mail_to, iterations)
            elif choice == 3:
                print('CopyMessages!\n')
                #copy_from = input('Enter Source UPN: ')
                copy_from = 'adelev@tms.365enespanol.com'
                #copy_to = input('Enter Target UPN: ')
                copy_to = 'diego@tms.365enespanol.com'
                print('Messages will be copied from:'+ copy_from + 'to: '+ copy_to)
                await copy_mail_messages(graph_client, copy_from, copy_to)
            elif choice == 4:
                user_messages = input('Enter Target User: ')
                await get_user_messages(graph_client, user_messages)
            elif choice == 5:
                user_folders = input('Enter Target User: ')
                await get_folders_id(graph_client, user_folders)
            else:
                print('Invalid choice!\n')
        except ODataError as odata_error:
            print('Error:')
            if odata_error.error:
                print(odata_error.error.code, odata_error.error.message)

async def list_users(graph_client: Graph):
    users_page = await graph_client.get_users()
    
    if users_page and users_page.value:
        for user in users_page.value:
            print('User:', user.display_name)
            print(' ID:', user.id)
            print(' Email:', user.mail)
        
        #if @odata.nextLink is present
        more_available = users_page.odata_next_link is not None
        print('\nMore users available?', more_available, '\n')
        
async def send_mail_message(graph_client: Graph, user_id, iterations_count):
    
    for i in range(iterations_count):
        await graph_client.send_mail_message(user_id)
        print (f"Mail {i+1} sent")
        time.sleep(1)
        
async def get_folders_id(graph_client: Graph, user_id):
    folders = await graph_client.get_parent_folder_id(user_id)
    
    print(folders)
    
    ''' if folders and folders.value:
        for folder in folders.value:
            print('FolderName:', folder.display_name)
            print('FolderID: ', folder.id) '''

async def get_user_messages(graph_client: Graph, user_target_id):
    t_messages = await graph_client.get_target_messages(user_target_id)
    
    if t_messages and t_messages.value:
        for t_msg in t_messages.value:
            print('MessageSubject: ', t_msg.subject)
            print('ParentFolderID: ',t_msg.parent_folder_id)

async def copy_mail_messages(graph_client: Graph, user_source, user_target):
    source_messages = await graph_client.copy_mail_message(user_source, user_target)
    
    if source_messages:
        for msg in source_messages:
            print('MessageID: ',msg.id)
            print('MessageSubject: ',msg.subject)
            print('###############################MESSAGECOPIED')

        

asyncio.run(main())