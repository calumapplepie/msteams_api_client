# currently just a skeleton; unable to actually call into teams until I get
# the app registered.  for demo purposes, a JSON from the graph explorer is to be used as a data source

# lets load our secrets!
import json
with open("secrets.json") as fp:
    secretsData:dict[str,str] = json.load(fp)

print(secretsData)
from pathlib import Path

# selected "client secret" authentication flow, see source of code: https://learn.microsoft.com/en-us/graph/sdks/choose-authentication-providers?tabs=python#client-credentials-provider
from  msgraph.graph_service_client import GraphServiceClient

def initalize_auth() -> GraphServiceClient:
    import azure.identity as auth


    # The client credentials flow requires that you request the
    # /.default scope, and pre-configure your permissions on the
    # app registration in Azure. An administrator must grant consent
    # to those permissions beforehand.
    scopes = [
              'https://graph.microsoft.com/User.ReadBasic.All'
]

    # azure.identity.aio
    credential = auth.InteractiveBrowserCredential(
        tenant_id=secretsData["tenant_id"],
        client_id=secretsData["client_id"],
        redirect_uri="https://localhost:5000",
        client_credential=secretsData["client_secret"]
    )

    return GraphServiceClient(credential, scopes)

from kiota_abstractions.serialization.parsable import Parsable
from msgraph.generated.models.base_collection_pagination_count_response import BaseCollectionPaginationCountResponse
"""
Given an initial response of paginated data, continue to request the next page,
and then serialize the assembled data out to a single .json file.
"""
async def writePagedKiotaDataToJson(graph_client: GraphServiceClient, outputData: BaseCollectionPaginationCountResponse, filename:Path|str):
    from kiota_serialization_json.json_serialization_writer import JsonSerializationWriter
    from msgraph_core.tasks import PageIterator

    # if given just a name, not a path, write to current directory
    if isinstance(filename,str):
        filename = Path.cwd()/filename

    # prepare the writer
    writer = JsonSerializationWriter()
    
    # we may have gotten a paginated response; use a PageIterator to output it
    iter = PageIterator(outputData,graph_client.request_adapter)
    # if the MS people who wrote the SDK used python, I could just do this: "async for page in iter"
    # instead, we need to do the following to serialize all the objects in the list
    depth = 0
    while True:
        # serialize all the values of the current page
        assert iter.current_page.value is not None
        for i in iter.current_page.value:
            i.serialize(writer)
        print(f"serialized page of {len(iter.current_page.value)} values")
        if iter.has_next is False:
            break
        if depth > 10:
            print("too many pages!")
            break
        depth += 1
        await iter.next()
        


    #write out the data
    with open(filename, "bw") as fp:
        fp.write(writer.get_serialized_content())


# based on snippet GraphExplorer
async def writeShiftsToJson(graph_client: GraphServiceClient):
    from msgraph.generated.teams.item.schedule.shifts.shifts_request_builder import ShiftsRequestBuilder
    from kiota_abstractions.base_request_configuration import RequestConfiguration

    request_configuration = RequestConfiguration(
        query_parameters = ShiftsRequestBuilder.ShiftsRequestBuilderGetQueryParameters(
            filter = "sharedShift/startDateTime ge 2026-01-15T00:00:00.000Z",
        )
    )

    result = await graph_client.teams.by_team_id(secretsData["team_id"]).schedule.shifts.get(request_configuration = request_configuration)
    assert result is not None
    await writePagedKiotaDataToJson(graph_client, result, "shiftsData.json")


async def writeUsersToJson(graph_client: GraphServiceClient):
    result = await graph_client.users.get()
    assert result is not None
    await writePagedKiotaDataToJson(graph_client, result, "userData.json");



# load the json shifts back to a fancypants object
from msgraph.generated.models.shift_collection_response import ShiftCollectionResponse 
def loadJsonShifts(fileName) -> ShiftCollectionResponse:
    with open(fileName, "rb") as fp:
        shiftsData = fp.read()
    
    # i spent a chunk of time seeing if I could read the JSON back into MS Graph SDK objects
    # and finally got it to work by reverse engineering the underdocumented pile of nonsense that it is
    # in other words, we leverage the parser from the MSGraph SDK to parse the cached response/what we got from the Graph Explorer request
    
    from kiota_serialization_json.json_parse_node_factory import JsonParseNodeFactory
    rootNode = JsonParseNodeFactory().get_root_parse_node("application/json",shiftsData)
    value = rootNode.get_object_value(ShiftCollectionResponse)
    # TODO; generalize this function so initalizeUsers can use it
    return value

from collections import defaultdict
import namer
"""
dictionary of user names from IDs, defaulting to generated nonsense if a name isnt known
originally because I was going to only put in a few IDs manually, but then I decided to throw together initalizeUserd
"""
userIdToNameDict = defaultdict(lambda sep=" ": namer.generate(separator=sep, style="title"))

"""
Set up our dictionary of users; currently using hardcoded values
"""
def initializeUsers():
    from kiota_serialization_json.json_parse_node_factory import JsonParseNodeFactory
    from msgraph.generated.models.user_collection_response import UserCollectionResponse
   
    # two userData arrays, because the json response was paginated; this is a hardcoded limit, if I was writing to be mainfainable id assert() the second json is the last one
    with open("testUserData1.json", "rb") as fp:
        userData1 = fp.read()
    with open("testUserData2.json", "rb") as fp:
        userData2 = fp.read()
    

    rootNode = JsonParseNodeFactory().get_root_parse_node("application/json",userData1)
    userCollection1 = rootNode.get_object_value(UserCollectionResponse)   
    assert userCollection1.value is not None
    rootNode = JsonParseNodeFactory().get_root_parse_node("application/json",userData2)
    userCollection2 = rootNode.get_object_value(UserCollectionResponse)   
    assert userCollection2.value is not None
    userCollectionFull = userCollection1.value.copy() # we wont ever user userCol1 again, but still.
    userCollectionFull.extend(userCollection2.value)

    for user in userCollectionFull:
        assert user.id is not None
        if user.surname is None or user.given_name is None:
            continue
        userIdToNameDict[user.id] = f"{user.given_name} {user.surname}"


# now we get into the calendar building section
import icalendar as ical
"""
add location markers to calender event based on Teams-formatted notes
"""
def addLocation(event:ical.Event, notes:str):
    if "Meet @ Office" in notes or "Shop" in notes:
            event.add("LOCATION","Powerstation Events")
            # an attempt to get Apple to put a location dot in the calender, for travel time; this didnt work
            #TODO: get the components right for calendar apps to fill in the location for nav purposes
            event.add('GEO', (41.550091, -72.890677))



"""
read parsed shift collection to make calendars for each mentioned employee
"""
def createCalendars(shiftCollection: ShiftCollectionResponse):
    eventLists:dict[str,list[ical.Event]] = defaultdict(list) 
    if shiftCollection.value is None:
        raise RuntimeError("missing value dict, is JSON valid?")
    
    # convert shifts into event collections indexed by usedID
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
        assert (notes := sharedShift.notes) is not None
        event.add("SUMMARY", notes)
        addLocation(event, notes)
        eventLists[shift.user_id].append(event)

    
    #write out prepared event collections to calenders
    ical_folder = Path("./calendars")
    for userid, eventList in eventLists.items():
        cal = ical.Calendar()
        userName = userIdToNameDict[userid]
        print(f"writing calendar for {userName}")
        cal.calendar_name = userName
        for event in eventList:
            cal.add_component(event)
        with open(ical_folder/f"{userName}.ical", "wb") as fp:
            fp.write(cal.to_ical())



"""
# actually run some code
shifts = loadJsonShifts("testShiftsData.json")
userIdToNameDict |= {}
initializeUsers()
createCalendars(shifts)
"""
import asyncio
graph_client: GraphServiceClient = initalize_auth()
print("authInitialized!")
asyncio.run(writeUsersToJson(graph_client))

