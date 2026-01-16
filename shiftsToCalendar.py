# currently just a skeleton; unable to actually call into teams until I get
# the app registered.  for demo purposes, a JSON from the graph explorer is to be used as a data source

# lets load our secrets!
import json
with open("secrets.json") as fp:
    secretsData:dict[str,str] = json.load(fp)



# selected "client secret" authentication flow, see source of code: https://learn.microsoft.com/en-us/graph/sdks/choose-authentication-providers?tabs=python#client-credentials-provider
import azure.identity.aio as auth
from  msgraph.graph_service_client import GraphServiceClient

# The client credentials flow requires that you request the
# /.default scope, and pre-configure your permissions on the
# app registration in Azure. An administrator must grant consent
# to those permissions beforehand.
scopes = ['https://graph.microsoft.com/.default']

# Values from app registration
tenant_id = 'YOUR_TENANT_ID'
client_id = 'YOUR_CLIENT_ID'
certificate_path = 'YOUR_CERTIFICATE_PATH'

# azure.identity.aio
credential = auth.ClientSecretCredential(
    tenant_id=secretsData["tenant_id"],
    client_id=secretsData["client_id"],
    client_secret=secretsData["client_secret"])

graph_client = GraphServiceClient(credential, scopes)

# based on snippet GraphExplorer

from msgraph.generated.teams.item.schedule.shifts.shifts_request_builder import ShiftsRequestBuilder
from kiota_abstractions.base_request_configuration import RequestConfiguration

async def writeShiftsToJson():
    request_configuration = RequestConfiguration(
        query_parameters = ShiftsRequestBuilder.ShiftsRequestBuilderGetQueryParameters(
            filter = "sharedShift/startDateTime ge 2026-01-15T00:00:00.000Z",
        )
    )

    result = await graph_client.teams.by_team_id('team-id').schedule.shifts.get(request_configuration = request_configuration)
    with open("shiftsData.json", "rw") as fp:
        json.dump(result,fp)

