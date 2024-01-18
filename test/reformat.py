import os
import pandas as pd
import json
from openai import AzureOpenAI

# Define your standard column headers
standard_columns = ['Number', 'Section', 'Question', 'Response', 'Notes', 'Reference']

def read_headers(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)
    # Return the column headers
    return df.columns

def create_mapping(headers, standard_columns):
    # create the OpenAI client
    client = AzureOpenAI(
        api_key=os.getenv("AZURE_OPENAI_KEY"),  
        api_version=os.getenv("AZURE_OPENAI_API_VERSION"),
        azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
    )
    deployment_name=os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME")
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

# Function to read an Excel file and standardize its format
def standardize_excel(file_path, column_mapping):
    # Read the Excel file
    df = pd.read_excel(file_path)
    # Rename columns based on the provided mapping
    df.rename(columns=column_mapping, inplace=True)
    # Reorder the columns to match the standard format
    df = df[standard_columns]
    # Handle missing data if necessary
    # e.g., df.fillna(0) or df.dropna()
    return df

# Example of column mapping for a specific file
mapping = {
    'Question Number': 'Number',
    'Section Title': 'Section',
    'Question': 'Question',
    'Answer': 'Response',
    'Comments': 'Notes',
    'Reference Links': 'Reference'
}

# load the environment variables
load_dotenv()

# Read the headers
headers = read_headers(os.getenv('INPUT_FILE_PATH'))

# Call OpenAI to create the mapping
mapping = create_mapping(headers, standard_columns)

# Standardize the format of the Excel file
standardized_df = standardize_excel(os.getenv('INPUT_FILE_PATH'), mapping)

# Write the new excel file to disk
standardized_df.to_excel(os.getenv('OUTPUT_FILE_PATH'), index=False)