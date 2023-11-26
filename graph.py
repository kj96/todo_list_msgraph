# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

import os
import json
from datetime import datetime, timedelta

import aiohttp

# <UserAuthConfigSnippet>
from configparser import SectionProxy
from azure.identity import DeviceCodeCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.item.user_item_request_builder import UserItemRequestBuilder
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder)
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
    SendMailPostRequestBody)
from msgraph.generated.models.message import Message
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.email_address import EmailAddress

class Graph:
    settings: SectionProxy
    device_code_credential: DeviceCodeCredential
    user_client: GraphServiceClient

    def __init__(self, config: SectionProxy):
        self.settings = config
        client_id = self.settings['clientId']
        tenant_id = self.settings['tenantId']
        graph_scopes = self.settings['graphUserScopes'].split(' ')

        self.device_code_credential = DeviceCodeCredential(client_id, tenant_id = tenant_id)
        self.user_client = GraphServiceClient(self.device_code_credential, graph_scopes)
# </UserAuthConfigSnippet>

    # <GetUserTokenSnippet>
    async def get_user_token(self):
        graph_scopes = self.settings['graphUserScopes']
        access_token = self.device_code_credential.get_token(graph_scopes)
        return access_token.token
    # </GetUserTokenSnippet>

    # <GetUserSnippet>
    async def get_user(self):
        # Only request specific properties using $select
        query_params = UserItemRequestBuilder.UserItemRequestBuilderGetQueryParameters(
            select=['displayName', 'mail', 'userPrincipalName']
        )

        request_config = UserItemRequestBuilder.UserItemRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        user = await self.user_client.me.get(request_configuration=request_config)
        return user
    # </GetUserSnippet>

    # <GetInboxSnippet>
    async def get_inbox(self):
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            # Only request specific properties
            select=['from', 'isRead', 'receivedDateTime', 'subject'],
            # Get at most 25 results
            top=25,
            # Sort by received time, newest first
            orderby=['receivedDateTime DESC']
        )
        request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters= query_params
        )

        messages = await self.user_client.me.mail_folders.by_mail_folder_id('inbox').messages.get(
                request_configuration=request_config)
        return messages
    # </GetInboxSnippet>

    # <SendMailSnippet>
    async def send_mail(self, subject: str, body: str, recipient: str):
        message = Message()
        message.subject = subject

        message.body = ItemBody()
        message.body.content_type = BodyType.Text
        message.body.content = body

        to_recipient = Recipient()
        to_recipient.email_address = EmailAddress()
        to_recipient.email_address.address = recipient
        message.to_recipients = []
        message.to_recipients.append(to_recipient)

        request_body = SendMailPostRequestBody()
        request_body.message = message

        await self.user_client.me.send_mail.post(body=request_body)
    # </SendMailSnippet>

    # <MakeGraphCallSnippet>
    async def make_graph_call(self, method, url, data=None, headers=None):
        # Use aiohttp for asynchronous HTTP requests
        async with aiohttp.ClientSession() as session:
            # Prepare headers, including the Authorization header for the access token
            if headers is None:
                headers = {}
            if 'Authorization' not in headers:
                # token = await self.get_access_token()  # Method to obtain the access token
                token = await self.get_user_token()
                headers['Authorization'] = f'Bearer {token}'

            # Make the API call
            async with session.request(method, url, data=data, headers=headers) as response:
                # Check response status
                if response.status == 200:
                    return await response.json()  # Return JSON response
                else:
                    # Handle error responses
                    raise Exception(f"Graph API call failed: {response.status}")
    # </MakeGraphCallSnippet>

    async def get_task_lists(self):
        url = 'https://graph.microsoft.com/v1.0/me/todo/lists'  # Directly set the URL
        return await self.make_graph_call('GET', url)

    async def get_tasks_in_list(self, list_id):
        # Directly set the URL to get tasks in a specific list
        url = f'https://graph.microsoft.com/v1.0/me/todo/lists/{list_id}/tasks'
        return await self.make_graph_call('GET', url)

    async def save_token(self, graph_o):
        try:
            access_token = await graph_o.get_user_token()  # Assuming this is an async method
            refresh_token = await graph_o.refresh_token()  # Assuming this is an async method
            expires_on = (datetime.utcnow() + timedelta(days=30)).strftime('%Y-%m-%dT%H:%M:%S.%fZ')  # This should be a datetime or timestamp

            token = {
                'access_token': access_token,
                'refresh_token': refresh_token,
                'expires_on': expires_on.isoformat() if isinstance(expires_on, datetime) else expires_on
            }

            with open('token.json', 'w') as token_file:
                json.dump(token, token_file)
        except Exception as e:
            # Handle exceptions, maybe log them or print an error message
            print(f"Error saving token: {e}")

    def load_token(self):
        if os.path.exists('token.json'):
            with open('token.json', 'r') as token_file:
                token = json.load(token_file)
                # Check if token is expired
                expiry_time = datetime.strptime(token['expires_on'], '%Y-%m-%dT%H:%M:%S.%fZ')
                if expiry_time - datetime.utcnow() > timedelta(minutes=5):
                    return token
        return None

    # Add a method to refresh the token if needed
    async def refresh_token(self):
        # Implement token refresh logic
        pass