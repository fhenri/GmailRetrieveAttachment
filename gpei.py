
from __future__ import print_function
import base64
import configparser
import httplib2
import json
import os, shutil, sys
import time
import zipfile

from apiclient import discovery
from apiclient import errors
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from os.path import basename

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# visit https://developers.google.com/gmail/api/auth/about-auth?authuser=1 
# to get API token

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-api-python.json
# see https://developers.google.com/gmail/api/auth/scopes
SCOPES              = 'https://mail.google.com/'
CLIENT_SECRET_FILE  = '.credentials/client_secret.json'
GMAIL_API_FILE      = '.credentials/gmail-api-python.json'
APPLICATION_NAME    = 'Gmail retrieve attachment'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """

    credential_path = os.path.abspath(GMAIL_API_FILE)
    client_path     = os.path.abspath(CLIENT_SECRET_FILE)

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(client_path, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def ListLabels(service, user_id):
    results = service.users().labels().list(userId=user_id).execute()
    labels = results.get('labels', [])

    if not labels:
        print('No labels found.')
    else:
        print('Labels:')
        for label in labels:
            print(label['id'] + ":" + label['name'])

def GetLabelFromName(service, user_id, labelName):
    results = service.users().labels().list(userId=user_id).execute()
    labels = results.get('labels', [])

    if not labels:
        print('No labels found.')
    else:
        for label in labels:
            if labelName in label['name']:
                return label['id']

def listMessagesMatchingQuery(service, user_id, query=''):
  """List all Messages of the user's mailbox matching the query.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    query: String used to filter messages returned.
    Eg.- 'from:user@some_domain.com' for Messages from a particular sender.

  Returns:
    List of Messages that match the criteria of the query. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate ID to get the details of a Message.
  """
  try:
    response = service.users().messages().list(
        userId=user_id,
        q=query).execute()
    messages = []
    if 'messages' in response:
        messages.extend(response['messages'])

    while 'nextPageToken' in response:
        page_token = response['nextPageToken']
        response = service.users().messages().list(userId=user_id, q=query,
                                         pageToken=page_token).execute()
        messages.extend(response['messages'])

    return messages
  except errors.HttpError as error:
    print('An error occurred: {0}'.format(error))


def ListMessagesWithLabel(service, user_id, label_ids=[]):
  """List all Messages of the user's mailbox with label_ids applied.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    label_ids: Only return Messages with these labelIds applied.

  Returns:
    List of Messages that have all required Labels applied. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate id to get the details of a Message.
  """
  try:
    response = service.users().messages().list(userId=user_id,
                                               labelIds=label_ids).execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id,
                                                 labelIds=label_ids,
                                                 pageToken=page_token).execute()
      messages.extend(response['messages'])

    return messages
  except errors.HttpError as error:
    print('An error occurred: {0}'.format(error))

def GetAttachments(service, user_id, msg_id, store_dir):
  """Get and store attachment from Message with given id.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: ID of Message containing attachment.
    store_dir: The directory used to store attachments.
  """
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()

    for part in message['payload']['parts']:
        if part['filename']:
            if 'data' in part['body']:
                data=part['body']['data']
            else:
                attachment_id=part['body']['attachmentId']
                attachment = service.users().messages().attachments().get(
                    id=attachment_id,
                    userId=user_id, 
                    messageId=msg_id).execute()
                data = attachment['data']

            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            path = os.path.join(store_dir, part['filename'])

            with open(path, 'wb') as f:
                f.write(file_data)


  except errors.HttpError as error:
      print('An error occurred: {0}'.format(error))


def DeleteMessage(service, user_id, msg_id):
  """Delete a Message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: ID of Message to delete.
  """
  try:
      service.users().messages().delete(userId=user_id, id=msg_id).execute()
      print('Message with id: {0} deleted successfully.'.format(msg_id))
  except errors.HttpError as error:
      print('An error occurred: {0}'.format(error))


def Zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            fileToZip = os.path.join(root, file)
            ziph.write(fileToZip, basename(fileToZip))

def Cleandir(path):
    for file in os.listdir(path):
        file_path = os.path.join(path, file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as error:
            print('An error occurred: {0}'.format(error))

def main():
    """
    """
    credentials = get_credentials()
    http        = credentials.authorize(httplib2.Http())
    service     = discovery.build('gmail', 'v1', http=http)
    
    # check if config provides labelId else search from Name
    settings  = json.load(open('gmail.json'))
    gmailConf = settings["GMail"]
    if 'LabelId' in gmailConf:
        label_id = gmailConf["LabelId"]
    else:
        label_id = GetLabelFromName(service, 'me', gmailConf['LabelName'])

    if not label_id:
        sys.exit("No label found in GMail")

    messages = ListMessagesWithLabel(service, 'me', label_id)
    storageFolder = settings["Settings"]["StorageFolder"]
    for message in messages:
        #print (json.dumps(message, indent=4, sort_keys=True))

        GetAttachments(service, 'me', message['id'], storageFolder)
        # once processed check if we can remove the message
        if settings["Settings"]["RemoveMessages"]:
            print("remove messages")
            #DeleteMessage(service, 'me', message['id'])

    # zip folder with attachment
    zipf = zipfile.ZipFile(
        'gpei_files_{0}.zip'.format(time.strftime("%Y%m%d-%H%M%S")), 
        'w', zipfile.ZIP_DEFLATED)
    Zipdir(storageFolder, zipf)

    # clean folder
    if settings["Settings"]["CleanFolder"]:
        print("clean directory.")
        Cleandir(storageFolder)

if __name__ == '__main__':
    main()