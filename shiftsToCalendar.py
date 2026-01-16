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

    result = await graph_client.teams.by_team_id(secretsData["teamID"]).schedule.shifts.get(request_configuration = request_configuration)

    with open("shiftsData.json", "rw") as fp:
        json.dump(result,fp)

# load the json shifts back to a fancypants object
from msgraph.generated.models.shift_collection_response import ShiftCollectionResponse 
def loadJsonShifts() -> ShiftCollectionResponse:
    with open("shiftsData.json", "rb") as fp:
        shiftsData = fp.read()
    # i spent a chunk of time seeing if I could read the JSON back into MS Graph SDK objects
    # and finally got it to work by reverse engineering the underdocumented pile of nonsense that it is
    
    from kiota_serialization_json.json_parse_node_factory import JsonParseNodeFactory
    rootNode = JsonParseNodeFactory().get_root_parse_node("application/json",shiftsData)
    value = rootNode.get_object_value(ShiftCollectionResponse)
    return value
    

import icalendar as ical
from collections import defaultdict
def createCalendars(shiftCollection: ShiftCollectionResponse):
    eventLists:dict[str,list[ical.Event]] = defaultdict(list)
    if shiftCollection.value is None:
        raise RuntimeError("missing value dict, is JSON valid?")
    
    for shift in shiftCollection.value:
        if shift is None:
            raise RuntimeError("no shifts in collection!")
        if shift.user_id is None:
            raise RuntimeError("no user ID!")
        sharedShift = shift.shared_shift
        if (sharedShift := shift.shared_shift) is None:
            raise RuntimeError("shift with no SharedShift attribute")
        event = ical.Event()
        event.DTSTART = sharedShift.start_date_time
        event.DTEND = sharedShift.end_date_time
        eventLists[shift.user_id].append(event)

    
