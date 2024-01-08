import pandas as pd
from pymongo import MongoClient
from dotenv import load_dotenv
from transformers import pipeline


client = MongoClient(
    "mongodb+srv://intern:JeUDstYbGTSczN4r@interntest.i7decv0.mongodb.net/")
db = client.intern
collection = db.papers

# Fetch data from MongoDB
collection_data = collection.find({})


input_excel_file = 'Genoshi_Intern_Test_Input_Excel_Sheet.xlsx'
df = pd.read_excel(input_excel_file)

# Initialize GPT-3 pipeline from transformers library
gpt3 = pipeline("text-generation", model="gpt-3.5-turbo",
                token="OPENAI_API_TOKEN")

#  context for the AI model
context = "This is the context for GPT-3."

# Iterate through cells
filled_text_list = []
for index, row in df.iterrows():
    row_heading = row['Row Heading']
    col_heading = row['Column Heading']

    # Query the collection based on row and column
    relevant_data = [data['transcription'] for data in collection_data if data['row']
                     == row_heading and data['column'] == col_heading]

    # If data is found, use GPT-3 to process the relevant data
    if relevant_data:
        # Generate text using GPT-3 pipeline
        filled_info = gpt3(
            context + "\n\n" + relevant_data[0], max_length=50, num_return_sequences=1)[0]['generated_text']
        filled_text_list.append(filled_info)

# Insert generated text into the DataFrame
df['Generated_Text'] = filled_text_list

# Save updated DataFrame back to Excel
output_excel_file = 'Output_Excel_File.xlsx'
df.to_excel(output_excel_file, index=False)
