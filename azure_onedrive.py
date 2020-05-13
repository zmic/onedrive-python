import os
import socket
import re

# -------------------------------------------------------------------------
#
#  HOW TO SET UP
#
#  Login to your Microsoft Azure account and go to the portal
#  Go to service "Azure Active Directory"
#  Go to App registration -> Register an application
#      Give it a name
#      Select the option with "... and personal Microsoft accounts (e.g. Skype, Xbox)"
#      Set Redirect URI to some port on local host, for instance http://localhost:9001
#  Go to API permissions and click "Add a permission"
#      Add Microsoft Graph, then "Delegated Permissions"
#      Add these permissions: profile, Files.ReadWrite, offline_access
#  Go to "Certificates & secrets" -> "New client secret"
#      Store the value of the secret in environment variable AZURE_CLIENT_SECRET
#

# -------------------------------------------------------------------------
#
#  pip install microsoftgraph-python
#
import microsoftgraph.client

# -------------------------------------------------------------------------
#
#  Store your client ID in an environment variable AZURE_CLIENT_ID
#  You can find the client ID of your App on the overview page.
#
client_id = os.environ["AZURE_CLIENT_ID"]  
client_secret = os.environ["AZURE_CLIENT_SECRET"]  

# -------------------------------------------------------------------------
#
#  this redirect uri should be registered in your application
#
redirect_uri = 'http://localhost:9001'

# -------------------------------------------------------------------------
#
#  Try to open the refresh_token saved to disk in a previous session.
#  If it's still valid, we don't need to login.
#  You only get refresh tokens if your app has permission "offline_access"
#
def try_refresh_token():
    try:
        refresh_token = open('refresh_token.txt', 'r').read()
    except FileNotFoundError:
        return None        
    try:
        return client.refresh_token(redirect_uri, refresh_token)
    except microsoftgraph.exceptions.BaseError as e:
        print(e)
        
# -------------------------------------------------------------------------
#
#  Run a "webserver" to catch the redirect from the login page
#
def mini_webserver():
    HOST = 'localhost'
    PORT = int(redirect_uri.split(':')[-1].replace('/', ''))
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind((HOST, PORT))
        s.listen(1)
        conn, addr = s.accept()
        with conn:
            data = conn.recv(1024)
            return str(data)
            
# -------------------------------------------------------------------------
#
#  Let's go
#
client = microsoftgraph.client.Client(client_id, client_secret=client_secret, account_type='common')
token = try_refresh_token()     
    
if token is None:    
    # 
    #  refreshing the token did not work so we need to log in.
    #
    scopes=['files.readwrite', 'user.read', 'offline_access']    
    url = client.authorization_url(redirect_uri, scopes, state=None)
    # launch url in a browser
    os.startfile(url)
    # and catch the redirect
    response = mini_webserver()
    # GET request contains the code in the form
    # code=xxxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxx
    code = re.search("code=([\w^-]+)", response).group(1)
    print("Got code", code)
    # code has to be exchanged for the actual token
    token = client.exchange_code(redirect_uri, code)
    print("Got token")

# save the refresh token to disk so we don't need to login next time 
if 'refresh_token' in token:
    open('refresh_token.txt', 'w').write(token['refresh_token'])
    
client.set_token(token)

# -------------------------------------------------------------------------
#
#  From this point on we can use the app
#

# Get basic info about my account
me = client.get_me()
print(me)

# Get folders and files at the root of my onedrive
root_children_items = client.drive_root_children_items()
items = root_children_items['value']
for x in items:
    print(x['name'], x['size'])

#item = client.get("/me/drive/root:/data/mydata.zip")
#print(item)

