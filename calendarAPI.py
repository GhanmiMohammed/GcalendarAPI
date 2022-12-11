


from O365 import Account, MSGraphProtocol

CLIENT_ID = 'a5a8098b-be18-48a1-8e95-e1a049969e34'
SECRET_ID = 'oUD8Q~vugmGlkwd3nA1ur3ry5HuYJmxugxo3Tcbp'

credentials = (CLIENT_ID, SECRET_ID)

print(CLIENT_ID)


protocol = MSGraphProtocol()

scopes =['Calenders.Read.Shared']

account = Account(credentials, protocol=protocol)

if account.authenticate(scopes=scopes):
    print("Authenticated")
else:
    print('Not Authitecated')




