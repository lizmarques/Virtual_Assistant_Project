import webbrowser
from msal import ConfidentialClientApplication, PublicClientApplication

#client_secret = "e9bc7db7-98d3-482d-a049-0fd834f31447"
app_id = "41b6ad65-0042-4be8-bb5c-e2bc91d1c935"

SCOPES = ["basic", "calendar"]

client = PublicClientApplication(client_id=app_id)
flow = client.initiate_device_flow(scopes=SCOPES)
print(flow)
#print("user code: " + flow["user_code"])
#webbrowser.open(flow["verification_uri"]

#authorization_code =
#acess_token = client.acquire_token_by_authorization_code((code=authorization_code, scopes = SCOPES))
#print(acess_token)