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
    'sAMAccountName',
    'telephoneNumber',
    'proxyAddresses',
    'title',
    'department',
    'company',
    'manager',
    'description'

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
    'l':'Location',
    'telephoneNumber':'Telephone Number',
    'proxyAddresses':'Proxy Addresses',
    'title':'Title',
    'department':'Department',
    'company':'Company',
    'manager':'Manager',
    'description':'Description'

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
    elif sendType == "Remove":          # Content must be a string containing a single record ID.
        return requests.delete(airtableURL+"/"+content,headers=AirtableAPIHeaders)


def changeDataInAirtable(content, sendType):                 # uploads the data to Airtable
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

def retrieveRecordsFromAD(ADSearchAttributes, ADSearchParams):
    ADserver = ldap3.Server(ADdomain, use_ssl=False) # NOTE: PLEASE GET SECURE LDAP RUNNING ON THE DOMAIN CONTROLLER, THEN CHANGE use_ssl TO True
    ADconnection = ldap3.Connection(ADserver, user=ADdomain+"\\"+ADusername, password=ADuserpass, authentication=ldap3.NTLM, auto_bind=True)

    ADconnection.search('",".join("dc="+x for x in ADdomain.split("."))', '(&({}))'.format(')('.join('{}={}'.format(key, value) for key, value in ADSearchParams.items())), attributes=ADSearchAttributes)
    adlist = {}
    for x in ADconnection.entries:
        loadedAttributes = json.loads(x.entry_to_json())['attributes']

        if 'objectGUID' in loadedAttributes and 'givenName' in loadedAttributes and len(loadedAttributes['givenName']) != 0:     # Skip any entries where the givenName attribute is blank.
            GUID = loadedAttributes['objectGUID'][0]
            adlist[GUID] = {}
            for attributeName, attributeContents in loadedAttributes.items():
                correctedAttributeName = convADAttributeToAirtableHeader[attributeName]
                if attributeContents == []:
                    adlist[GUID][correctedAttributeName] = ''
                else:
                    adlist[GUID][correctedAttributeName] = attributeContents[0]
        else:
            print("Ignoring: "+loadedAttributes['sAMAccountName'][0])
            pass
    ADconnection.unbind()
    # print(adlist)
    return adlist

def getInfoFromGUID(GUID):
    return retrieveRecordsFromAD(allADSearchAttributes, {'objectGUID':GUID})[GUID]

def initialCheck(ATRecords):
    ADrecords = retrieveRecordsFromAD(allADSearchAttributes, allUserADSearchParams)
    recordsToUpdate = []
    recordsToSend = []  # list of lists, with those sublists containing up to 10 entries to send to Airtable
    for x in ADrecords:
        if x not in ATRecords.records:
            recordsToSend.append({"fields":ADrecords[x]})
        else:
            recordsToUpdate.append({"id":ATRecords.records[x], "fields":ADrecords[x]})

    print("Removing entries from Airtable that do not exist in AD.")
    for x in ATRecords.records:
        if x not in ADrecords:
            print('id to remove: '+ATRecords.records[x])
            changeDataInAirtable(ATRecords.records[x], "Remove")
    print("Done removing bad records")

    print("Updating current entries in Airtable.")
    for x in range(0, len(recordsToUpdate), 10):
        y = changeDataInAirtable({"records":[z for z in recordsToUpdate[x:x+10]], "typecast":True}, "Update")
    print("Done updating.")

    print("Uploading to Airtable any missing entries.")
    for x in range(0, len(recordsToSend), 10):
        y = changeDataInAirtable({"records":[z for z in recordsToSend[x:x+10]], "typecast":True}, "Post")
        for z in y['records']:
            ATRecords.records[z['fields']['objectGUID']] = z['id']
    print("Done uploading.")


def main():

    evtSessionCredentials = (ADIP, ADusername, ADdomain, ADuserpass, win32evtlog.EvtRpcLoginAuthDefault)
    evtSessionCredentials2 = (ADIP2, ADusername, ADdomain, ADuserpass, win32evtlog.EvtRpcLoginAuthDefault)
    evtSession = win32evtlog.EvtOpenSession(evtSessionCredentials, win32evtlog.EvtRpcLogin, 0, 0)
    evtSession2 = win32evtlog.EvtOpenSession(evtSessionCredentials2, win32evtlog.EvtRpcLogin, 0, 0)
    # eventPulse = win32event.CreateEvent(None, 0, 0, None)
    ADAccountsChanged = set()
    ATRecords = airtable()
    initialCheck(ATRecords)                         # Verify all AD records are present in AT on script startup, add missing records

    print("Now watching Event Viewer for new AD Users and for AD User updates.")

    def eventTriggered(evt1, evt2, eventContent):   #evt1: int specifying why the function was called | evt2: context object (5th parameter in EvtSubscribe)
        print('triggered')
        xmlData = win32evtlog.EvtRender(eventContent, win32evtlog.EvtRenderEventXml)
        GUID = re.search(r'<Data Name=\'ObjectGUID\'>(.*?)</Data>', xmlData).group(1)
        subjectusername = re.search(r'<Data Name=\'SubjectUserName\'>(.*?)</Data.', xmlData).group(1)
        badGUID = bool('$' in subjectusername or len(subjectusername) == 0)  # badGUID = True if the SubjectUserName has a $ in it OR has no text. False otherwise.
        if not badGUID:
            ADAccountsChanged.add(GUID.lower())
        # win32event.PulseEvent(eventPulse)


    eventLog = win32evtlog.EvtSubscribe('Security', win32evtlog.EvtSubscribeToFutureEvents, None, eventTriggered, None, xmlQuery, evtSession, None)
    eventLog2 = win32evtlog.EvtSubscribe('Security', win32evtlog.EvtSubscribeToFutureEvents, None, eventTriggered, None, xmlQuery, evtSession2, None)

    while True:
        try:
            time.sleep(10)  # every 10 seconds, check if a user has been updated and reconcile records
                            # Maybe: Remove this, use trigger below. If pulse comes, wait 1 second. If another pulse comes during that time, reset the wait. If one doesn't, continue processing

            if len(ADAccountsChanged) == 0:
                continue

            if time.time() - ATRecords.lastRefreshTime > 3600:
                ATRecords.reloadRecords()

            ADAccountsChanged_copy = [x for x in ADAccountsChanged]
            for GUID in ADAccountsChanged_copy:
                print(GUID)
                if GUID in ATRecords.records:
                    print({"id":ATRecords.records[GUID],"fields":getInfoFromGUID(GUID), "typecast":True})
                    y = changeDataInAirtable({"records":[{"id":ATRecords.records[GUID],"fields":getInfoFromGUID(GUID)}], "typecast":True}, "Update")
                else:
                    y = changeDataInAirtable({"fields":getInfoFromGUID(GUID), "typecast":True}, "Post")
                    if type(y) == dict:
                        ATRecords.records[GUID] = y['id']
                # print(getInfoFromGUID(GUID))
                ADAccountsChanged.remove(GUID)

            # trigger = win32event.WaitForSingleObject(eventPulse, 60000)
            # if trigger == win32event.WAIT_TIMEOUT:
            #     hasTimedOut = True
            # elif trigger == win32event.WAIT_OBJECT_0:
            #     hasTimedOut = False
        except Exception:
            win32evtlog.CloseEventLog(eventLog)     #not sure which one is correct, haven't tested yet
            evtSession.CloseEventLog()
            win32evtlog.CloseEventLog(eventLog2)
            evtSession2.CloseEventLog()
        except KeyboardInterrupt:
            eventLog.CloseEventLog()
            evtSession.CloseEventLog()
            eventLog2.CloseEventLog()
            evtSession2.CloseEventLog()

main()