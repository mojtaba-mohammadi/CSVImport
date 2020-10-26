import csv
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
import random


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


mainCol = ['sku', 'meta:_sku', 'post_title', 'post_status', 'post_type', 'Type: product_type',
           'tax:product_type', 'tax:product_cat', 'tax:product_tag', 'regular_price', 'sale_price',
           'meta:_price', 'manage_stock', 'stock_status', 'images',
           'attribute:pa_style', 'attribute_default:pa_style', 'attribute_data:pa_style',
           'attribute:pa_sex', 'attribute_default:pa_sex', 'attribute_data:pa_sex',
           'attribute:pa_color', 'attribute_default:pa_color', 'attribute_data:pa_color',
           'attribute:pa_size', 'attribute_default:pa_size', 'attribute_data:pa_size',
           'attribute:pa_material', 'attribute_data:pa_material',
           'attribute:pa_print', 'attribute_data:pa_print',
           'attribute:pa_wash', 'attribute_data:pa_wash',
           'attribute:pa_default_style', 'attribute_data:pa_default_style',
           'meta:_wc_average_rating', 'meta:_wc_rating_count', 'meta:_wc_review_count',
           'meta:_stock_status', 'post_content', 'post_excerpt']

mainSheet = pd.DataFrame([], columns=mainCol)

tshirtColor = ['color1', 'color2', 'color3', 'color4', 'color5', 'color6', 'color7']
tshirtSize = ['M', 'L', 'XL', '2XL', '3XL']
hoodieColor = ['color8', 'color2', 'color5', 'color9', 'color10']
hoodieSize = ['M', 'L', 'XL']
colorIndex = pd.Series(
    ['سفید', 'مشکی', 'خاکستری', 'آبی', 'زرشکی', 'زرد', 'سبز خاکی', 'ملانژ', 'نارنجی', 'سرمه‌ای'],
    index=['color1', 'color2', 'color3', 'color4', 'color5', 'color6', 'color7', 'color8', 'color9', 'color10']
)

# Configuration
chunkSize = 100
artCount = 657
spreadSheetID = '1n329f-UfjCc71QrxxAOkVisRXZKXqL-iq9mm_sedB1A'
# Reading data from google sheet
rangeName = 'ARTs!A1' + ':Q' + str(artCount + 1)
df = read_data_from_gsheet(spreadSheetID, rangeName)

i = 0
for artCode in df['ArtCode']:
    # create T-shirts
    sku = df.loc[i]['ArtCode'] + '-TS'
    post_title = df.loc[i]['Title']
    post_status = 'publish'
    post_type = 'product'
    product_type = 'variable'
    product_cat = df.loc[i]['Category']
    product_cat = product_cat.replace(" - ", ">")
    product_cat = product_cat.replace("- ", ">")
    product_cat = product_cat.replace(" -", ">")
    product_cat = product_cat.replace("-", ">")
    product_tag = df.loc[i]['Hashtag']
    regular_price = 120000
    manage_stock = 'no'
    stock_status = 'instock'

    pa_color = ''
    j = 0
    tempColorList = []
    for color in tshirtColor:
        if int(df.loc[i]['Tshirt-' + tshirtColor[j]]) == 1:
            pa_color = pa_color + ('|' if j > 0 else '') + colorIndex[tshirtColor[j]]
            tempColorList.append(tshirtColor[j])
        j += 1

    j = 0
    for color in hoodieColor:
        if int(df.loc[i]['Hoodie-' + hoodieColor[j]]) == 1:
            if hoodieColor[j] not in tempColorList:
                pa_color = pa_color + '|' + colorIndex[hoodieColor[j]]
        j += 1

    pa_color_default_index = random.choice(tempColorList)
    pa_color_default = colorIndex[pa_color_default_index]
    pa_color_data = '3|1|1'

    j = 0
    images = ""
    for color in tshirtColor:
        if int(df.loc[i]['Tshirt-' + tshirtColor[j]]) == 1:
            images = images + ('|' if j > 0 else '') + df.loc[i]['ArtCode'] + '-TS-' + tshirtColor[j] + '.jpg'
            images = images + '|' + df.loc[i]['ArtCode'] + '-TS-' + tshirtColor[j] + '-men' + '.jpg'
            images = images + '|' + df.loc[i]['ArtCode'] + '-TS-' + tshirtColor[j] + '-women' + '.jpg'
        j += 1

    j = 0
    for color in hoodieColor:
        if int(df.loc[i]['Hoodie-' + hoodieColor[j]]) == 1:
            images = images + '|' + df.loc[i]['ArtCode'] + '-HD-' + hoodieColor[j] + '.jpg'
            images = images + '|' + df.loc[i]['ArtCode'] + '-HD-' + hoodieColor[j] + '-men' + '.jpg'
            images = images + '|' + df.loc[i]['ArtCode'] + '-HD-' + hoodieColor[j] + '-women' + '.jpg'
        j += 1

    pa_style = 'تیشرت|هودی'
    pa_style_default = 'تیشرت'
    pa_style_data = '1|1|1'

    pa_sex = 'خانم‌ها|آقایان'
    pa_sex_default = random.choice(['آقایان', 'خانم‌ها'])
    pa_sex_data = '2|1|1'

    pa_size = ''
    j = 0
    tempSizeList = []
    for size in tshirtSize:
        pa_size = pa_size + ('|' if j > 0 else '') + tshirtSize[j]
        tempSizeList.append(tshirtSize[j])
        j += 1
    j = 0
    for size in hoodieSize:
        if hoodieSize[j] not in tempSizeList:
            pa_size = pa_size + '|' + hoodieSize[j]
        j += 1

    pa_size_default = tshirtSize[1]
    pa_size_data = '4|1|1'

    pa_material = '۱۰۰ درصد نخ پنبه'
    pa_material_data = '5|1|0'

    pa_print = 'چاپ دیجیتال با دوام جوهرافشان با استفاده از تکنولوژی مدرن DTG'
    pa_print_data = '6|1|0'

    pa_wash = 'جهت شست و شو لباس را پشت و رو کرده و با آب زیر ۴۰ درجه بشویید، ترجیحا از پودر صابون جهت شست‌ و شو استفاده کنید'
    pa_wash_data = '7|1|0'

    pa_default_style = 'تیشرت'
    pa_default_style_data = '8|1|0'

    wc_average_rating = random.randint(40, 50) / 10
    wc_rating_count = random.randint(2, 10)
    wc_review_count = wc_rating_count

    post_content = "<img src='https://artshirt.ir/wp-content/uploads/sizechart/TS-Size-Chart.jpg' style='display: block;margin-left: auto;margin-right: auto;'/>"
    post_excerpt = "<ul><li><span style='font-size: 12px; color: #808080;'>جنس  ۱۰۰ درصد نخ پنبه</span></li><li><span style='font-size: 12px; color: #808080;'>چاپ دیجیتال با دوام با استفاده از تکنولوژی مدرن DTG</span></li></ul>"

    row = [sku, sku, post_title, post_status, post_type, post_type, product_type, product_cat,
           product_tag, regular_price, '', regular_price, manage_stock, stock_status, images,
           pa_style, pa_style_default, pa_style_data, pa_sex, pa_sex_default, pa_sex_data,
           pa_color, pa_color_default, pa_color_data, pa_size, pa_size_default, pa_size_data,
           pa_material, pa_material_data, pa_print, pa_print_data, pa_wash, pa_wash_data,
           pa_default_style, pa_default_style_data, wc_average_rating, wc_rating_count,
           wc_review_count, stock_status, post_content, post_excerpt
           ]

    fileName = 'csvfile-TS-' + str(divmod(i, chunkSize)[0]) + '.csv'
    with open(fileName, "a", newline='') as file:
        wr = csv.writer(file)
        if divmod(i, chunkSize)[1] == 0:
            wr.writerow(mainCol)
        wr.writerow(row)
    i += 1

i = 0
for artCode in df['ArtCode']:
    # create Hoodie
    sku = df.loc[i]['ArtCode'] + '-HD'
    post_title = df.loc[i]['Title']
    post_status = 'publish'
    post_type = 'product'
    product_type = 'variable'
    product_cat = df.loc[i]['Category']
    product_cat = product_cat.replace(" - ", ">")
    product_cat = product_cat.replace("- ", ">")
    product_cat = product_cat.replace(" -", ">")
    product_cat = product_cat.replace("-", ">")
    product_tag = df.loc[i]['Hashtag']
    regular_price = 120000
    manage_stock = 'no'
    stock_status = 'instock'

    j = 0
    images = ""
    for color in tshirtColor:
        if int(df.loc[i]['Tshirt-' + tshirtColor[j]]) == 1:
            images = images + ('|' if j > 0 else '') + df.loc[i]['ArtCode'] + '-TS-' + tshirtColor[j] + '.jpg'
            images = images + '|' + df.loc[i]['ArtCode'] + '-TS-' + tshirtColor[j] + '-men' + '.jpg'
            images = images + '|' + df.loc[i]['ArtCode'] + '-TS-' + tshirtColor[j] + '-women' + '.jpg'
        j += 1

    j = 0
    for color in hoodieColor:
        if int(df.loc[i]['Hoodie-' + hoodieColor[j]]) == 1:
            images = images + '|' + df.loc[i]['ArtCode'] + '-HD-' + hoodieColor[j] + '.jpg'
            images = images + '|' + df.loc[i]['ArtCode'] + '-HD-' + hoodieColor[j] + '-men' + '.jpg'
            images = images + '|' + df.loc[i]['ArtCode'] + '-HD-' + hoodieColor[j] + '-women' + '.jpg'
        j += 1

    pa_style = 'تیشرت|هودی'
    pa_style_default = 'هودی'
    pa_style_data = '1|1|1'

    pa_sex = 'خانم‌ها|آقایان'
    pa_sex_default = random.choice(['آقایان', 'خانم‌ها'])
    pa_sex_data = '2|1|1'

    pa_color = ''
    j = 0
    tempColorList = []
    for color in hoodieColor:
        if int(df.loc[i]['Hoodie-' + hoodieColor[j]]) == 1:
            pa_color = pa_color + '|' + colorIndex[hoodieColor[j]]
            tempColorList.append(colorIndex[hoodieColor[j]])
        j += 1

    j = 0
    for color in tshirtColor:
        if int(df.loc[i]['Tshirt-' + tshirtColor[j]]) == 1:
            if colorIndex[tshirtColor[j]] not in tempColorList:
                pa_color = pa_color + ('|' if j > 0 else '') + colorIndex[tshirtColor[j]]
        j += 1

    pa_color_default = random.choice(tempColorList)
    pa_color_data = '3|1|1'

    pa_size = ''
    j = 0
    tempSizeList = []
    for size in tshirtSize:
        pa_size = pa_size + ('|' if j > 0 else '') + tshirtSize[j]
        tempSizeList.append(tshirtSize[j])
        j += 1
    j = 0
    for size in hoodieSize:
        if hoodieSize[j] not in tempSizeList:
            pa_size = pa_size + '|' + hoodieSize[j]
        j += 1

    pa_size_default = hoodieSize[1]
    pa_size_data = '4|1|1'

    pa_material = '۱۰۰ درصد نخ پنبه'
    pa_material_data = '5|1|0'

    pa_print = 'چاپ دیجیتال با دوام جوهرافشان با استفاده از تکنولوژی مدرن DTG'
    pa_print_data = '6|1|0'

    pa_wash = 'جهت شست و شو لباس را پشت و رو کرده و با آب زیر ۴۰ درجه بشویید، ترجیحا از پودر صابون جهت شست‌ و شو استفاده کنید'
    pa_wash_data = '7|1|0'

    pa_default_style = 'هودی'
    pa_default_style_data = '8|1|0'

    wc_average_rating = random.randint(40, 50) / 10
    wc_rating_count = random.randint(2, 10)
    wc_review_count = wc_rating_count

    post_content = "<img src='https://artshirt.ir/wp-content/uploads/sizechart/TS-Size-Chart.jpg' style='display: block;margin-left: auto;margin-right: auto;'/>"
    post_excerpt = "<ul><li><span style='font-size: 12px; color: #808080;'>جنس ۱۰۰ درصد نخ پنبه</span></li><li><span style='font-size: 12px; color: #808080;'>چاپ دیجیتال با دوام با استفاده از تکنولوژی مدرن DTG</span></li></ul>"

    row = [sku, sku, post_title, post_status, post_type, post_type, product_type, product_cat,
           product_tag, regular_price, '', regular_price, manage_stock, stock_status, images,
           pa_style, pa_style_default, pa_style_data, pa_sex, pa_sex_default, pa_sex_data,
           pa_color, pa_color_default, pa_color_data, pa_size, pa_size_default, pa_size_data,
           pa_material, pa_material_data, pa_print, pa_print_data, pa_wash, pa_wash_data,
           pa_default_style, pa_default_style_data, wc_average_rating, wc_rating_count,
           wc_review_count, stock_status, post_content, post_excerpt
           ]

    fileName = 'csvfile-HD-' + str(divmod(i, chunkSize)[0]) + '.csv'
    with open(fileName, "a", newline='') as file:
        wr = csv.writer(file)
        if divmod(i, chunkSize)[1] == 0:
            wr.writerow(mainCol)
        wr.writerow(row)
    i += 1