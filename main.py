import streamlit as st
import time
from io import BytesIO
from datetime import date
from openpyxl import load_workbook
from dateutil.relativedelta import relativedelta
import re
from calendar import month_abbr
import json

st.set_page_config(
     page_title="SEO Projections",
     page_icon="📈"
)

with st.sidebar:
    st.markdown("""
        - [NLP Tool](https://nlptool2.ey.r.appspot.com/)
    """)


# st.title('SEO Projections')
st.markdown("<h1 style='color: #F2C300'>SEO Projections Tool</h1>", unsafe_allow_html=True)

help = st.checkbox('Help')
if help:
    st.header('How to:')
    st.write("""
    Lorem ipsum dolor sit amet, consectetur adipiscing elit. Cras tortor felis, auctor quis lacus eget, dignissim imperdiet elit. Aliquam porta volutpat velit dignissim faucibus. Nulla vel urna augue. Vestibulum iaculis purus ac tempor faucibus. Vestibulum id venenatis nisi. Maecenas imperdiet urna a risus dictum, sed sagittis sem rutrum. Nulla facilisi. Proin interdum massa semper, aliquet lacus id, facilisis odio. Donec ultrices ornare purus, malesuada tempus neque bibendum eget. Nullam hendrerit efficitur purus accumsan posuere.

    Aenean ultricies sit amet risus nec consectetur. Duis elit metus, tempor eget dolor eget, laoreet sollicitudin ex. Quisque metus ante, convallis at quam vitae, auctor bibendum orci. Nulla gravida malesuada neque, vitae condimentum metus. Fusce lobortis laoreet enim quis posuere. Donec aliquet malesuada nisl, ac congue leo ornare id. Fusce facilisis ipsum vitae nulla elementum, vitae feugiat ipsum auctor.
    """)
    with open('H:\Code\python\misc\streamlit\input-example.xlsx', 'rb') as file:
        btn = st.download_button(
            label = 'Sample File',
            data = file,
            file_name = 'example.xlsx'
        )



with st.form('projections_form'):
    file_upload = st.file_uploader('Upload')
    date = st.date_input('Starting Date')
    length = st.select_slider('Length', [12, 24, 36])
    rankCap = st.slider('Rank Cap', 1, 20)
    submittion = st.form_submit_button('Submit')

def projections(wb, progress_bar, ankCap=6, projectionLength=12):

    keyPhrases = {}
    fields = ['Category', 'Sub Cat 1', 'Other', 'Internal comp?', 'Cross-Domain Competition', 'Existing Page']
    clickThroughRates = [i.value for i in wb['Ctr']['B']][1:]

    #turn sheet into dict
    rankImprovementRates = {}
    for i, row in enumerate(skip_first(wb['RankImprovement'].values)):
        #I'm sorry for putting it all on one line
        rankImprovementRates[row[0]] = dict(zip([f'Month {i}' for i in range(1, len(row[1:]) + 1)], row[1:]))

    for row in skip_first(wb['KeyPhraseList'].values):
        keyPhrases[row[0]] = dict(zip(fields, row[1:]))

    fields = [cell.value for cell in wb['RankingData'][1]]

    rows = len([i for i in skip_first(wb['RankingData'].values)])
    for i, row in enumerate(skip_first(wb['RankingData'].values)):
        progress_bar.progress(progress_calc(i, rows))
        keyPhrase = row[1]
        #For some reason there are keywords in RankingData that aren't in the KeyPhraseList?
        if keyPhrases.get(keyPhrase):
            keyPhrases[keyPhrase]['Average'] = row[17]
            #So many magic numbers, I'm sorry
            for i, date in [[i, f'{month_abbr[int(field[14:])]}-{field[9:13]}'] for i, field in enumerate(fields) if re.search('^Volume - \d{4}-\d{2}$', field)]:
                keyPhrases[keyPhrase][date] = row[i] if row[i] else 0

            keyPhrases[keyPhrase]['Ranking'] = row[3]

            keyPhrases[keyPhrase]['Starting Rank'] = 99 if keyPhrases[keyPhrase]['Existing Page'] == 'New Page' else keyPhrases[keyPhrase]['Ranking']
            for i in range(1, projectionLength + 1):
                rank = int(round(keyPhrases[keyPhrase]['Starting Rank'] * rankImprovementRates[keyPhrases[keyPhrase]['Category']][f'Month {i}']))
                keyPhrases[keyPhrase][f'Rank month {i}'] = rank if rank > rankCap else rankCap

            print(keyPhrase)
            dump(keyPhrases[keyPhrase], False)
    st.success('Rank Projections Complete')
    #Traffic

def generate_dates():
    return [(date.today() + relativedelta(months = i, day = 21)).isoformat() for i in range(0, (12 * 5))]

#Skips the first row which is the fields.
#Uses next() rather than a slice [1:] to avoid loading the entire sheet into memory at once as a list
#Done often enough that function was simpler
def skip_first(i):
    j = iter(i)
    next(j)
    return j

def progress_calc(i, l):
    return round((i / l) * 100)

def dump(l, pause=True):
    print(json.dumps(l, indent=4))
    if pause:
        input()


if submittion:
    if not file_upload:
        st.warning('No File Uploaded')
        st.stop()

    wb = load_workbook(filename=BytesIO(file_upload.read()))
    progress_bar = st.progress(0)
    projections(wb, progress_bar, rankCap, length)
    st.info('Work in Progress')
    # progress_bar = st.progress(0)
    # for j in range(3):
    #     time.sleep(1)
    #     for i in range(100):
    #         time.sleep(0.1)
    #         progress_bar.progress(i + 1)
    #     st.success(f'Key Phrase {j + 1} - Complete')
    st.download_button('Download', 'example')
