import win32file, win32evtlog, win32event, pywintypes
import time, json, requests, configparser, os, subprocess
import ldap3, re

mainFolder = os.path.dirname(os.path.abspath(__file__))+"/"
config = configparser.ConfigParser()
config.read(mainFolder+'settings.ini')

airtableURL = config['Other']['airtable_url']
slackURL = config['Other']['slack_url']
airtableURLFields = config['Other']['airtable_url_fields']
airtableAPIKey = config['Other']['airtable_api_key']
ADIP = config['Other']['AD_IP']
ADIP2 = config['Other']['AD_IP2']
ADusername = config['Other']['AD_Username']
ADdomain = config['Other']['AD_Domain_Name']
ADuserpass = config['Other']['AD_User_Password']
# enableAirtablePosts = config['Other']['enable_Airtable_posts']
enableAirtablePosts = True

allUserADSearchParams = {
    'objectClass':'user',
    'objectCategory':'person'
}

allADSearchAttributes = [
    'objectGUID',
    'mail',
    'givenName',
    'sn',
    'l',
    'info',
    'sAMAccountName'
]


AirtableAPIHeaders = {
    "Authorization":str("Bearer "+airtableAPIKey),
    "User-Agent":"Python Script",
    "Content-Type":"application/json"
}

eventIDs = {
    4720:"A user account was created.",
    4722:"A user account was enabled.",
    4723:"An attempt was made to change the password of an account.",
    4724:"An attempt was made to reset password of an account.",
    4725:"A user account was disabled.",
    4726:"A user account was deleted.",
    4738:"A user account was changed.",
    4741:"A user account was created.",
    4742:"A user account was changed.",
    4743:"A user account was deleted.",
    5136:"A user account was changed"
}

convADAttributeToAirtableHeader = {
    'objectGUID':'objectGUID',
    'info':'Employee Number',
    'sn':'Last Name',
    'givenName':'First Name',
    'mail':'Email Address',
    'sAMAccountName':'Login Name',
    'l':'Location'

}

# xmlQuery = "*[System[(EventID=4720 or (EventID >= 4722 and EventID <= 4726) or (EventID=4738) or (EventID >= 4741 and EventID <= 4743))] and EventData[Data[@Name='SubjectDomainName']!='NT AUTHORITY'] and EventData[Data[@Name='ObjectClass']!='computer']]"
xmlQuery = "*[System[({})] and EventData[Data[@Name='SubjectDomainName']!='NT AUTHORITY'] and EventData[Data[@Name='ObjectClass']!='computer']]".format(' or '.join("EventID="+str(x) for x in eventIDs))

class airtable:
    def __init__(self):
        self.records = {}   # formatted as such: {'objectGUID':'recordID', 'objectGUID2':'recordID2'}
        self.reloadRecords()

    def reloadRecords(self):
        self.lastRefreshTime = time.time()
        for x in self.retrieveRecordsFromAirtable():
            self.records[x['fields']['objectGUID']] = x['id']

    def retrieveRecordsFromAirtable(self, offset=None):
        while True:
            try:
                # if enableAirtablePosts != True:
                #     return "Airtable connection disabled."
                if offset == None:
                    x = requests.get(airtableURL+airtableURLFields, data=None, headers=AirtableAPIHeaders)
                else:
                    x = requests.get(airtableURL+airtableURLFields+"&offset="+offset, data=None, headers=AirtableAPIHeaders)

                records = x.json()['records']
                if 'offset' in json.loads(x.text):
                    records.extend(self.retrieveRecordsFromAirtable(json.loads(x.text)['offset']))
                return records
                    
            except ConnectionError:
                print("Could not connect to airtable.com")
                time.sleep(30)


def getAirtableRecordID(objectGUID, ATRecordsList):
    for x in ATRecordsList:
        if objectGUID in x['fields']:
            return x['id']


def postOrUpdate(content, sendType):
    if sendType == "Post":
        return requests.post(airtableURL,data=None,json=content,headers=AirtableAPIHeaders)
    elif sendType == "Update":
        return requests.patch(airtableURL,data=None,json=content,headers=AirtableAPIHeaders)


def uploadDataToAirtable(content, sendType):                 # uploads the data to Airtable
    if enableAirtablePosts != True:
        return "Airtable connection disabled."
    x = postOrUpdate(content, sendType)
    print("\nPost HTTP code:", x.status_code, "  |   Send type:",sendType)
    if x.status_code == 200:                                 # if Airtable upload successful, move PDF files to Done folder
        print("Success! Sent via "+sendType+"\n")
        return json.loads(x.text)
    else:
        print(str(json.loads(x.text)['error']['message']))
        return {'content':str(content), 'status code: ':str(x.status_code), 'failureText':str(json.loads(x.text)['error']['message'])}

def retrieveRecordsFromAD(ADSearchParams, ADSearchAttributes):
    ADserver = ldap3.Server(ADdomain, use_ssl=False) # NOTE: PLEASE GET SECURE LDAP RUNNING ON THE DOMAIN CONTROLLER, THEN CHANGE use_ssl TO True
    ADconnection = ldap3.Connection(ADserver, user=ADdomain+"\\"+ADusername, password=ADuserpass, authentication=ldap3.NTLM, auto_bind=True)

    ADconnection.search('",".join("dc="+x for x in ADdomain.split("."))', '(&({}))'.format(')('.join('{}={}'.format(key, value) for key, value in ADSearchParams.items())), attributes=ADSearchAttributes)
    adlist = {}
    for x in ADconnection.entries:
        loadedAttributes = json.loads(x.entry_to_json())['attributes']

        if 'objectGUID' in loadedAttributes:
            GUID = loadedAttributes['objectGUID'][0]
            adlist[GUID] = {}
        
            for attributeName, attributeContents in loadedAttributes.items():
                correctedAttributeName = convADAttributeToAirtableHeader[attributeName]
                if attributeContents == []:
                    adlist[GUID][correctedAttributeName] = ''
                else:
                    adlist[GUID][correctedAttributeName] = attributeContents[0]
    ADconnection.unbind()
    # print(adlist)
    return adlist

def getInfoFromGUID(GUID):
    return retrieveRecordsFromAD({'objectGUID':GUID}, allADSearchAttributes)[GUID]




def main():

    evtSessionCredentials = (ADIP, ADusername, ADdomain, ADuserpass, win32evtlog.EvtRpcLoginAuthDefault)
    evtSessionCredentials2 = (ADIP2, ADusername, ADdomain, ADuserpass, win32evtlog.EvtRpcLoginAuthDefault)
    evtSession = win32evtlog.EvtOpenSession(evtSessionCredentials, win32evtlog.EvtRpcLogin, 0, 0)
    evtSession2 = win32evtlog.EvtOpenSession(evtSessionCredentials2, win32evtlog.EvtRpcLogin, 0, 0)
    # eventPulse = win32event.CreateEvent(None, 0, 0, None)
    ADAccountsChanged = set()
    ATRecords = airtable()

    def eventTriggered(evt1, evt2, eventContent):   #evt1: int specifying why the function was called | evt2: context object (5th parameter in EvtSubscribe)
        print('triggered')
        xmlData = win32evtlog.EvtRender(eventContent, win32evtlog.EvtRenderEventXml)
        GUID = re.search(r'<Data Name=\'ObjectGUID\'>(.*?)</Data>', xmlData).group(1)
        badGUID = bool('$' in re.search(r'<Data Name=\'SubjectUserName\'>(.*?)</Data.', xmlData).group(1))
        if not badGUID:
            ADAccountsChanged.add(GUID)
        # win32event.PulseEvent(eventPulse)


    eventLog = win32evtlog.EvtSubscribe('Security', win32evtlog.EvtSubscribeToFutureEvents, None, eventTriggered, None, xmlQuery, evtSession, None)
    eventLog2 = win32evtlog.EvtSubscribe('Security', win32evtlog.EvtSubscribeToFutureEvents, None, eventTriggered, None, xmlQuery, evtSession2, None)

    while True:
        try:
            time.sleep(10)  # every 10 seconds, check if a user has been updated and reconcile records

            if len(ADAccountsChanged) == 0:
                continue

            if time.time() - ATRecords.lastRefreshTime > 3600:
                ATRecords.reloadRecords()

            ADAccountsChanged_copy = [x for x in ADAccountsChanged]
            for GUID in ADAccountsChanged_copy:
                print(GUID)
                if GUID in ATRecords.records:
                    print({"id":ATRecords.records[GUID],"fields":getInfoFromGUID(GUID), "typecast":True})
                    y = uploadDataToAirtable({"records":[{"id":ATRecords.records[GUID],"fields":getInfoFromGUID(GUID)}], "typecast":True}, "Update")
                else:
                    y = uploadDataToAirtable({"fields":getInfoFromGUID(GUID), "typecast":True}, "Post")
                    if type(y) == dict:
                        ATRecords.records[GUID] = json.loads(y['id'])
                # print(getInfoFromGUID(GUID))
                ADAccountsChanged.remove(GUID)

            # trigger = win32event.WaitForSingleObject(eventPulse, 60000)
            # if trigger == win32event.WAIT_TIMEOUT:
            #     hasTimedOut = True
            # elif trigger == win32event.WAIT_OBJECT_0:
            #     hasTimedOut = False
        except Exception:
            eventLog.CloseEventLog()
            evtSession.CloseEventLog()
            eventLog2.CloseEventLog()
            evtSession2.CloseEventLog()
        except KeyboardInterrupt:
            eventLog.CloseEventLog()
            evtSession.CloseEventLog()
            eventLog2.CloseEventLog()
            evtSession2.CloseEventLog()

main()