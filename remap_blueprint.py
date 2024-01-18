import azure.functions as func
import logging
import os
from io import BytesIO
import json
import base64
import pandas as pd
from openai import AzureOpenAI

remap_blueprint = func.Blueprint()

@remap_blueprint.route(route="remap_excel")
def remap_excel(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a remap_excel request.')
    
    try:
        # get the request json body
        req_body = req.get_json()
        
        # get the headers
        headers = req.headers
        openai_apikey = headers["OpenAI-APIKey"]
        openai_host = headers["OpenAI-Host"]        
        openai_version = headers["OpenAI-Version"]        
        openai_deployment = headers["OpenAI-Deployment"]        
        if not openai_apikey or not openai_host or not openai_version or not openai_deployment:
            return func.HttpResponse(
                "Please provide the OpenAI API Key, Host, Version, and Deployment headers",
                status_code=400
            )
        
        # get the standard columns
        standard_columns = req_body["standardColumns"]
        if not standard_columns:
            return func.HttpResponse(
                "Please provide a comma separated list of standard headers.",
                status_code=400
            )
        # remove all spaces
        standard_columns = standard_columns.replace(" ", "")
        # split to an array
        standard_columns = standard_columns.split(",")
        
        # get the excel file
        base64_excel = req_body.get('excelFile')
        if not base64_excel:
            return func.HttpResponse(
                "Please provide a base64 encoded Excel document.",
                status_code=400
            )
        excel_data = base64.b64decode(base64_excel)
        excel_file = BytesIO(excel_data)
        excelColumns = read_headers(excel_file)
        
        # Use OpenAI to create the create_mapping
        mapping = create_mapping(excelColumns, standard_columns, openai_apikey, openai_host, openai_version, openai_deployment)

        # create the remapped excel file
        standardized_df = standardize_df(excel_file, standard_columns, mapping)
        # convert the base64 excel file
        base64_excel = excel_to_base64(standardized_df)
        
        # create the response body
        response_body = {}
        response_body["standardColumns"] = standard_columns
        response_body["excelColumns"] = excelColumns.tolist()
        response_body["mapping"] = mapping
        response_body["excelFile"] = base64_excel.decode("utf-8")
    
        return func.HttpResponse(
            json.dumps(response_body),
            mimetype="application/json",
            status_code=200
        )
    except Exception as e:
        return func.HttpResponse(
            f"An error occurred: {str(e)}",
            status_code=500
        )
        

# return the column headers        
def read_headers(excel_bytes):
    df = pd.read_excel(excel_bytes)
    return df.columns

def create_mapping(headers, standard_columns, openai_apikey, openai_host, openai_version, openai_deployment):
    # create the OpenAI client
    client = AzureOpenAI(
        api_key=openai_apikey,  
        api_version=openai_version,
        azure_endpoint="https://" + openai_host + "/"
    )
    deployment_name=openai_deployment
    #The chat message to send to the model
    message_text = [
        {
            "role":"system","content":"You are an expert at manipulating excel spreadsheets."
        }
        ,
        {
            "role":"user","content":"Create a JSON object to remap columns from the given format to the standard format.\n\nThe given headers are: " + str(headers) + "\n\nThe standard headers are: " + str(standard_columns) + "\n\nGive your response in JSON as in the example below, ONLY RETURN THE JSON.\n\nExample: { ""source header 1"":""standard header1"",""source header2"":""standard header 2""}"
        }
    ]
    #call the model
    completion = client.chat.completions.create(
        model=deployment_name,
        messages = message_text,
        temperature=0.7,
        max_tokens=800,
        top_p=0.95,
        frequency_penalty=0,
        presence_penalty=0,
        stop=None
    )
    #return the JSON mapping object from the response
    print(completion.choices[0].message.content)
    return json.loads(completion.choices[0].message.content)

def standardize_df(excel_file, standard_columns, column_mapping):
    # Read the Excel file
    df = pd.read_excel(excel_file)
    # Rename columns based on the provided mapping
    df.rename(columns=column_mapping, inplace=True)
    # Reorder the columns to match the standard format
    df = df[standard_columns]
    return df

def excel_to_base64(excel_file):
    excel_output = BytesIO()
    excel_file.to_excel(excel_output, index=False)
    excel_output.seek(0)
    return base64.b64encode(excel_output.read())