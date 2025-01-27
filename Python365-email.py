import asyncio
import base64

import configparser

from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder)
from msgraph.generated.users.item.messages.item.attachments.item.attachment_item_request_builder import AttachmentItemRequestBuilder

config = configparser.ConfigParser()
config.read(['config.cfg', 'config.dev.cfg'])
credential = config['azure']

# scope read from Admin portal
scopes = ['https://graph.microsoft.com/.default']

# Create an API client with the credentials and scopes
client = GraphServiceClient(credentials=credential, scopes=scopes)

# To fetch e-mail(s) by filter
async def get_emails():
    # Filter settings
    query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            # Only request specific properties
            select=['from', 'isRead', 'receivedDateTime', 'subject'],
            # Get at most 25 results
            top=25,
            # Sort by received time, newest first
            orderby=['receivedDateTime DESC']
        )
    # Building filter request
    request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters= query_params
        )
    # Search message with above filter in selected e-mail
    messages = await client.users.by_user_id(userID).mail_folders.by_mail_folder_id('inbox').messages.get(
                request_configuration=request_config)
    return messages

# After fetching all e-mail from inbox, show info nice
async def list_inbox():
    message_page = await get_emails()
    if message_page and message_page.value:
        # Output each message's details
        for message in message_page.value:
            print('Message ID:', message.id)            # Used to fetch single e-mail
            print('Message:', message.subject)
            if (
                message.from_ and
                message.from_.email_address
            ):
                print('  From:', message.from_.email_address.name , 'E-mail: ',message.from_.email_address.address  or 'NONE')
            else:
                print('  From: NONE')
            print('  Status:', 'Read' if message.is_read else 'Unread')
            print('  Received:', message.received_date_time)
        # If @odata.nextLink is present
        more_available = message_page.odata_next_link is not None
        print('\nMore messages', more_available, '\n')

# Fetch single e-mail by ID
async def get_email(idnum):
    message = await client.users.by_user_id(userID).messages.by_message_id(idnum).get()
    # Most e-mail are in HTML, you can check this by message.body.content_type.
    # Then you can use html2text module to make it readable.
    print(message.body.content) # This is RAW
    ###  Following field can be useful:
    # print('Importance: ',message.importance)
    #if message.bcc_recipients:
    #    print('BCC: ', message.bcc_recipients)
    #print('CC: ',message.cc_recipients)
    #print('categories: ',message.categories) # = List
    #print('from: ',message.from_)
    #print('delivery request: ',message.is_delivery_receipt_requested)
    #print('read request: ',message.is_read_receipt_requested)
    #print('reply to: ',message.reply_to)    # = List
    #print('sender: ',message.sender)
    # print('Contains attachement(s): ',message.has_attachments)
    ### Following can be used for attachments
    # if message.has_attachments:
    #    attachments = await client.users.by_user_id(userID).messages.by_message_id(email).attachments.get()
    #    print('Attachement count: ', len(attachments.value))
    #    for attachment in attachments.value:
    #        print('Attachement ID: ', attachment.id)
    #        print('Attachement name: ', attachment.name)
    #        print('Attachement content: ', attachment.content_type)
    #        print('Attachement size: ', attachment.size)
    #        await getAttachment(email,attachment.id)
    ### End additional fields 
    # Some above field can be viewed from LIST_INBOX function
    return message

# Can be called from single e-mail function, to extract attachement to file.
async def getAttachment(idnum,attachID):
    attachment = await client.users.by_user_id(userID).messages.by_message_id(idnum).attachments.by_attachment_id(attachID).get()
    f = open(attachment.name, 'w+b')
    f.write(base64.b64decode(attachment.content_bytes))
    f.close

asyncio.run(list_inbox())