from google.cloud import bigquery
import os

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'danger-222103-c1071aec0cda.json'

client = bigquery.Client()

# Perform a query.
QUERY = (
    'SELECT * FROM `bigquery-public-data.covid19_jhu_csse.confirmed_cases` where province_state="Hong Kong"')
query_job = client.query(QUERY)  # API request
rows = query_job.result()  # Waits for query to finish

dataframe = (
    query_job
    .result()
    .to_dataframe()
)

daily_confirmed_cases_accumulate = dataframe.values[0][5:]

daily_confirmed_cases = [0] + list(daily_confirmed_cases_accumulate[1:] - daily_confirmed_cases_accumulate[:-1])


from openpyxl import load_workbook
wb = load_workbook(filename='Danger-Restaurant-Revenue.xlsx')
sheet1 = wb["Sheet1"]
sheet1[f'D1'] = "Covid number"
for i in range(2, 22):
    sheet1[f'D{i}'] = 0
for i, value in enumerate(daily_confirmed_cases):
    sheet1[f'D{i + 22}'] = value

wb.save(filename='Danger-Restaurant-Revenue-output.xlsx')
