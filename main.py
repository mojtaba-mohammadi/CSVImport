import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd


def read_data_from_gsheet(spreadSheetID, rangeName):
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

    # The ID and range of a sample spreadsheet.
    SAMPLE_SPREADSHEET_ID = spreadSheetID
    SAMPLE_RANGE_NAME = rangeName

    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        dataframe = pd.DataFrame(values[1:], columns=values[0])
        return dataframe


# Creating T-shirts

# Reading data from google sheet
spreadSheetID = '1n329f-UfjCc71QrxxAOkVisRXZKXqL-iq9mm_sedB1A'
rangeName = 'ARTs!A1:Q10'
df = read_data_from_gsheet(spreadSheetID, rangeName)
columns = df.columns

# Configuration
chunkSize = 20
mainCol = ['sku', 'meta:_sku', 'post_title', 'post_status', 'post_type', 'Type: product_type',
           'tax:product_type', 'tax:product_cat', 'tax:product_tag', 'regular_price', 'meta:_price',
           'sale_price', 'manage_stock', 'stock_status', 'images', 'attribute:pa_size', 'attribute_default:pa_size',
           'attribute_data: pa_size', 'attribute:pa_sex', 'attribute_default:pa_sex', 'attribute_data: pa_sex',
           'attribute:pa_color', 'attribute_data: pa_color', 'attribute:pa_material', 'attribute_data:pa_material',
           'attribute:pa_print', 'attribute_data:pa_print', 'attribute:pa_collar', 'attribute_data:pa_collar',
           'attribute:pa_wash', 'attribute_data:pa_wash', 'meta:_stock_status', 'meta:_wc_average_rating',
           'meta:_wc_rating_count', 'post_content', 'post_excerpt']

a = ['MainCat', 'Category', 'ArtCode', 'Title', 'Hashtag', 'Tshirt-سفید',
     'Tshirt-مشکی', 'Tshirt-خاکستری', 'Tshirt-آبی', 'Tshirt-زرشکی',
     'Tshirt-زرد', 'Tshirt-سبز', 'Hoodie-ملانژ', 'Hoodie-مشکی',
     'Hoodie-زرشکی', 'Hoodie-نارنجی', 'Hoodie-سرمه‌ای']

mainSheet = pd.DataFrame([], columns=mainCol)

tshirtColor = ['سفید', 'مشکی', 'خاکستری', 'آبی', 'زرشکی', 'زرد', 'سبز']
tshirtSize = ['M', 'L', 'XL', '2XL', '3XL']
hoodieColor = ['ملانژ', 'مشکی', 'زرشکی', 'نارنجی', 'سرمه‌ای']
hoodieSize = ['M', 'L', 'XL']
colorIndex = pd.Series(
    ['white', 'black', 'charcoal', 'blue', 'burgundy', 'yellow', 'greygreen', 'melange', 'orange', 'navy'],
    index=['سفید', 'مشکی', 'خاکستری', 'آبی', 'زرشکی', 'زرد', 'سبز', 'ملانژ', 'نارنجی', 'سرمه‌ای']
)

i = 0
for artCode in df['ArtCode']:
    # create T-shirts
    sku = df.loc[[i], ['ArtCode']] + '-TS'
    post_title = df.loc[[i], ['Title']]
    post_status = 'publish'
    post_type = 'product'
    product_type = 'variable'
    product_cat = df.loc[[i], ['Category']]
    product_cat = product_cat.replace(" - ", ">")
    product_cat = product_cat.replace("- ", ">")
    product_cat = product_cat.replace(" -", ">")
    product_cat = product_cat.replace("-", ">")
    product_tag = df.loc[[i], ['Hashtag']]
    regular_price = 120000
    manage_stock = 'no'
    stock_status = 'instock'

    images = sku + '-' + colorIndex[0] + '.jpg'
    images = images + '|' + sku + '-' + colorIndex[0] + '-men' + '.jpg'
    images = images + '|' + sku + '-' + colorIndex[0] + '-women' + '.jpg'
    for color in tshirtColor[1:]:
        images = images + '|' + sku + '-' + colorIndex[color] + '.jpg'
        images = images + '|' + sku + '-' + colorIndex[color] + '-men' + '.jpg'
        images = images + '|' + sku + '-' + colorIndex[color] + '-women' + '.jpg'




    i += 1
