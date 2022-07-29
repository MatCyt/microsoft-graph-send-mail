import os
import logging
import base64
from io import BytesIO
import requests

import msal
import azure.functions as func
from azure.storage.blob import BlobServiceClient


# COLLECT CREDENTIALS AND ENV VARIABLES
c_id = os.environ['client_id']
c_secret = os.environ['client_secret']
tenant_id = os.environ['tenant_id']
sender_email = os.environ['sender_email']
sa_cs = os.environ['storage_connection_string']


# DEFINE GENERAL FUNCTIONS
def get_access_token(client_id: str, client_secret: str, tenant_id: str):
    """
    Generate client access token using MSAL library

    :param client_id: Service Principal client_id
    :param client_secret: Service Principal secret
    :param tenant_id

    :return: A dict representing the json response from AAD:
        - A successful response would contain "access_token" key,
        - an error response would contain "error" and "error_description".
    """

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["https://graph.microsoft.com/.default"]

    app = msal.ConfidentialClientApplication(
        client_id=client_id, client_credential=client_secret, authority=authority)

    access_response = None
    access_response = app.acquire_token_silent(scopes, account=None)

    if not access_response:
        access_response = app.acquire_token_for_client(scopes=scopes)

    return access_response


def get_attachment_from_blob(connection_string: str, container_name: str, blob_name: str):
    """
    Load file from Azure storage and encode it returing attachment data

    :param connection_string: connection string for storage account
    :param container_name: container name containing the blob
    :param blob_name: blob name including the path

    :return enconded_data: base64 encoded data from azure blob to be used as email attachment
    """

    blob_service_client = BlobServiceClient.from_connection_string(connection_string)
    blob_client = blob_service_client.get_blob_client(container_name, blob_name)

    stream_object_from_blob = blob_client.download_blob()
    stream = BytesIO()
    stream_object_from_blob.download_to_stream(stream)
    encoded_data = base64.b64encode(stream.getvalue()).decode()

    return encoded_data


# AZURE FUNCTION
def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('>> mailing request received')

    try:
        req_params = req.get_json()
        email_to = req_params['email_to']
        subject = req_params['subject']
        body = req_params['body']
        importance = req_params['importance']
        include_attachment = req_params['include_attachment']

    except Exception as e:
        logging.error(f'! failed to parse request parameters \n\t error message: {str(e)}')
        return func.HttpResponse("Failed to parse request parameters. Check AZ function logs for details.")

    # Generate email message from request parameters
    recipients = [{"EmailAddress": {"Address": email}} for email in email_to.split(";")]

    email_msg = {
        "message": {
            "toRecipients": recipients,
            "subject": subject,
            "importance": importance,
            "body": {"ContentType": "HTML", "Content": body}
        },
        "saveToSentItems": "true"
    }

    # If required - generate single or multiple attachment objects
    if include_attachment == "yes":

        try:
            container_name = req_params['container_name']
            attachments = req_params['blob_name']

            blob_list = attachments.split(';')
            attachments_list = []

            for blob in blob_list:
                attachment_data = get_attachment_from_blob(sa_cs, container_name, blob)
                file_name = blob.rsplit("/", maxsplit=1)[-1]
                attachment_item = {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": file_name,
                    "contentType": "text/csv",
                    "contentBytes": attachment_data,
                }

                attachments_list.append(attachment_item)

            email_msg["message"]["attachments"] = attachments_list
            logging.info('all attachments processed')

        except Exception as e:
            logging.error(f'! failed to create attachment object \n\t error message: {str(e)}')
            return func.HttpResponse(f"Failed to load blob and create attachment - response error: {str(e)}")

    # Get MSAL token used to authorize MS Graph API
    try:
        access_reponse = get_access_token(c_id, c_secret, tenant_id)
        logging.info('access token generated')

    except Exception as e:
        logging.error(f'! failed to generate access token \n\t error message: {str(e)}')
        return func.HttpResponse(f"Failed to generate access token - response error: {str(e)}")

    # Send email through MS graph API
    endpoint = f"https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail"

    try:
        r = requests.post(
            endpoint,
            headers={"Authorization": "Bearer " +
                     access_reponse["access_token"]},
            json=email_msg)
        logging.info(f'POST request to MS Graph API sent - status code: {r.status_code}')

        r.raise_for_status()

    except Exception as e:
        logging.error(f'! failed to call MS Graph API \n\t error message: {str(e)}')
        return func.HttpResponse(f"Failed to call MS Graph API - response error: {str(e)}")

    return func.HttpResponse("Email sent - function executed succesfully", status_code=200)
