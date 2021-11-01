import streamlit as st
import numpy as np
import pandas as pd
import datetime as dt
import re
import base64
from config import doc_dict

st.title('Lag importfil til IFS')

username = st.text_input(label='Steg 1: Skriv inn ditt kortnavn')

if username:
    st.success(f'Kortnavn lagret som: {username.upper()}')

facility = st.text_input(label='Steg 2: Skriv inn kortnavn på anlegg')

if facility:
    st.success(f'Du har valgt {facility}')

def main():

    excel_files = st.file_uploader("Steg 3: Last opp excel eksport fil(er) fra Omega FDV krav", type="xlsx", accept_multiple_files=True)

    if excel_files:
        df1 = pd.concat([pd.read_excel(file).astype(str) for file in excel_files])
        df1.rename(columns=replace_columns_df, inplace=True)
        df2 = df1.copy()
        df1 = df1[df1_cols]
        df1 = create_doc_attributes(df1)
        df1 = group_columns(df1,  'Dokumentnummer','Dokumenttype', 'Ifs klasse', 'Ifs format')
        df2 = df2[df2_cols]
        df2.drop_duplicates(subset='Dokumentnummer', keep='first', inplace=True)
        st.dataframe(df2.merge(df1))
        st.success('Import av eksportfil fra Omega 365 var velykket')
        
    merge_button = st.button('Lag lastefil til IFS')

    if merge_button:

        try:
            df = merge_csv_and_excel(df1, df2, 'Dokumentnummer')
            df = create_new_document_titles(df)
            IMPORT_FILE = create_import_file(df)
            st.dataframe(IMPORT_FILE)
            date = dt.datetime.today().strftime("%d.%m.%y")
            number_of_files = IMPORT_FILE.FILE_TYPE.count() + IMPORT_FILE.FILE_TYPE2.count()
            st.success("Opprettelse av importfil til IFS var vellykket!")
            csv1 = IMPORT_FILE.to_csv(sep=';', encoding='latin-1', index=True)
            b641 = base64.b64encode(csv1.encode('latin-1')).decode()
            href1 = f'<a href="data:file/csv;base64,{b641}" download="Importfil til IFS med {number_of_files} filer ({date}).csv">Last ned CSV for filimport til IFS med {number_of_files} filer</a>'
            st.markdown(href1, unsafe_allow_html=True)

            error_file = create_error_file(df)

            IMPORT_FILE_MCH_CODES = import_file_mch_codes(df, error_file)

            if IMPORT_FILE_MCH_CODES.empty != True:
                st.dataframe(IMPORT_FILE_MCH_CODES)
                st.success('Opprettelse av importfil for objektkoblinger til IFS var vellykket!')
                number_of_objects = IMPORT_FILE_MCH_CODES.MCH_CODE.count()
                number_of_objects_files = IMPORT_FILE_MCH_CODES.reset_index()['FILE_NAME'].nunique()
                csv2 = IMPORT_FILE_MCH_CODES.to_csv(sep=';', encoding='latin-1', index=True)
                b642 = base64.b64encode(csv2.encode('latin-1')).decode()
                href2 = f'<a href="data:file/csv;base64,{b642}" download="IMPORT_FILE IFS med {number_of_objects} objektkoblinger ({date}).csv">Last ned CSV fil for import av {number_of_objects} objektkoblinger til IFS for {number_of_objects_files} filer</a>'
                st.markdown(href2, unsafe_allow_html=True)

        except UnboundLocalError:
            st.error('Fil mangler')
        
        

def get_doc_attributes(doc_type, index=0):
    """Omega 365 dokumentkode (key value) henter dokumenttype fra IFS (default index=0), IFS klasse (index=1) og IFS format (index=2) fra liste i doc_dict"""

    if doc_type in doc_dict.keys():
        return doc_dict.get(doc_type)[index]
    else:
        return np.nan


def create_doc_attributes(df):
    """Bruker get_doc_attributes til å fylle ut dokumenttype, klasse og format på df """

    df['Dokumenttype'] = df['Dokumenttype'].apply(lambda x: x.split(' ')[0])
    df['Ifs klasse'] = df['Dokumenttype'].apply(get_doc_attributes, index=1)
    df['Ifs format'] = df['Dokumenttype'].apply(get_doc_attributes, index=2)
    df['Dokumenttype'] = df['Dokumenttype'].apply(get_doc_attributes , index=0)

    return df

def create_new_document_titles(df):
    """Lager ny dokumenttittel bsasert på dokumenttype, leverandørs tittel, dokumentnummer og anleggskode"""

    df['Dokumenttittel'] = df['Dokumenttype'] + '_ ' + df['Title'] + ', ' +  df['Dokumentnummer'] + f', {facility} '

    return df

def group_columns(df, *columns):
    """Behold alle rader i første kolonne som har dokumentnummer og slå sammen rader som har identisk dokumentnummer"""

    df.rename(columns={df.columns[0]: "Dokumentnummer"}, inplace=True)

    mch_code = 'Mch Code'
    object_description = 'Mch Name'

    df['Dokumentnummer'] = df['Dokumentnummer'].apply(str)

    new_columns = list()

    for column in columns:
        new_columns.append(df[column])

    mch_df = df[mch_code].groupby(new_columns).apply(set).reset_index()
    obj_df = df[object_description].groupby(new_columns).apply(set).reset_index()

    new_df = mch_df.merge(obj_df)

    join_strings = lambda x: '; '.join(x)

    new_df[object_description] = new_df[object_description].apply(join_strings)
    new_df[mch_code] = new_df[mch_code].apply(join_strings)

    new_df['Antall sammenslåtte krav'] = new_df[mch_code].apply(lambda x: x.split('; ')).apply(list).apply(len)

    return new_df

def split_rows(df, column, sep=',', keep=False):
    """Lag 1 ny rad i dataframe per filnavn eller objektkobling dersom det ligger flere verdier i samme kolonne"""

    indexes = list()
    new_values = list()
    df = df.dropna(subset=[column])

    for i, presplit in enumerate(df[column].astype(str)):
        values = presplit.split(sep)
        if keep and len(values) > 1:
            indexes.append(i)
            new_values.append(presplit)
        for value in values:
            indexes.append(i)
            new_values.append(value)
    new_df = df.iloc[indexes, :].copy()
    new_df[column] = new_values

    return new_df


def merge_csv_and_excel(df1, df2, column):

    df1 = df1.copy()
    df2 = df2.copy()

    return df1.merge(df2, on=column)


file_extensions = ('docx', 'pdf', 'xlsx', 'jpg', 'dwg', 'zip') # 'PDF', 'XLSX', 'JPG', 'DWG', 'ZIP')


def create_filetype(filename):
    """Fyll ut felt FILE_TYPE basert på filekstensjon"""

    filename = str(filename).lower()

    if filename.endswith(file_extensions):
        return [word for word in filename.split('.')][-1].upper()
    else:
        return ''

def import_documents(df):

    IMPORT_FILE = pd.DataFrame(columns=['DOC_CLASS', 'DOC_NO', 'DOC_SHEET', 'DOC_REV', 'FORMAT_SIZE', 'REV_NO',
                                       'TITLE', 'DOC_TYPE', 'INFO', 'FILE_NAME', 'LOCATION_NAME', 'PATH',
                                       'FILE_TYPE', 'FILE_NAME2', 'FILE_TYPE2', 'DOC_TYPE2', 'FILE_NAME3',
                                       'FILE_TYPE3', 'DOC_TYPE3', 'DT_CRE', 'USER_CREATED', 'ROWSTATE',
                                       'MCH_CODE', 'CONTRACT', 'REFERANSE'])

    IMPORT_FILE.TITLE = df['Dokumenttittel']
    IMPORT_FILE.FILE_NAME = list(df['FileName'].apply(lambda x: x.strip()))
    IMPORT_FILE.DOC_CLASS = df['Ifs klasse']
    IMPORT_FILE.DOC_NO = np.nan
    IMPORT_FILE.DOC_SHEET = 1
    IMPORT_FILE.DOC_REV = 1
    IMPORT_FILE.FORMAT_SIZE = df['Ifs format']
    IMPORT_FILE.REV_NO = 1
    IMPORT_FILE.DOC_TYPE = 'ORIGINAL'
    IMPORT_FILE.INFO = np.nan
    IMPORT_FILE.LOCATION_NAME = 'XXXX'
    IMPORT_FILE.PATH = 'YYYY'
    IMPORT_FILE.FILE_TYPE = IMPORT_FILE.FILE_NAME.apply(create_filetype)
    IMPORT_FILE.FILE_NAME2 = np.nan
    IMPORT_FILE.FILE_TYPE2 = np.nan
    IMPORT_FILE.DOC_TYPE2 = np.nan
    IMPORT_FILE.FILE_NAME3 = np.nan
    IMPORT_FILE.FILE_TYPE3 = np.nan
    IMPORT_FILE.DOC_TYPE3 = np.nan
    IMPORT_FILE.DT_CRE = dt.datetime.today().strftime("%d.%m.%y")
    IMPORT_FILE.USER_CREATED = username.upper()
    IMPORT_FILE.ROWSTATE = 'Frigitt'
    IMPORT_FILE.MCH_CODE = df['Mch Code']
    IMPORT_FILE.CONTRACT = 10
    IMPORT_FILE.REFERANSE = np.nan
    IMPORT_FILE.dropna(subset=['DOC_CLASS', 'FORMAT_SIZE'], inplace=True)
    IMPORT_FILE.set_index('DOC_CLASS', inplace=True)
    IMPORT_FILE.MCH_CODE = IMPORT_FILE.MCH_CODE.apply(str).apply(lambda x: x.split(';')[0]) # hent den første mchkoden fra liste

    return IMPORT_FILE


def import_drawings(df):

    IMPORT_FILE = pd.DataFrame(columns=['DOC_CLASS', 'DOC_NO', 'DOC_SHEET', 'DOC_REV', 'FORMAT_SIZE', 'REV_NO',
       'TITLE', 'DOC_TYPE', 'INFO', 'FILE_NAME', 'LOCATION_NAME', 'PATH',
       'FILE_TYPE', 'FILE_NAME2', 'FILE_TYPE2', 'DOC_TYPE2', 'FILE_NAME3',
       'FILE_TYPE3', 'DOC_TYPE3', 'DT_CRE', 'USER_CREATED', 'ROWSTATE',
       'MCH_CODE', 'CONTRACT', 'REFERANSE'])

    IMPORT_FILE.MCH_CODE = df['Mch Code']
    IMPORT_FILE.TITLE = df['Dokumenttittel']
    IMPORT_FILE.DOC_CLASS = df['Ifs klasse']
    IMPORT_FILE.FORMAT_SIZE = df['Ifs format']

    IMPORT_FILE.FILE_NAME = df['FileName'][df['FileName'].str.contains('dwg') | df['FileName'].str.contains('DWG')]
    IMPORT_FILE.FILE_NAME2 = df['FileName'][df['FileName'].str.contains('pdf') | df['FileName'].str.contains('PDF')]

    IMPORT_FILE = IMPORT_FILE.groupby(['TITLE', ]).first().reset_index()

    IMPORT_FILE.DOC_NO = np.nan
    IMPORT_FILE.DOC_SHEET = 1
    IMPORT_FILE.DOC_REV = 1
    IMPORT_FILE.REV_NO = 1
    IMPORT_FILE.DOC_TYPE = 'ORIGINAL'
    IMPORT_FILE.INFO = np.nan
    IMPORT_FILE.LOCATION_NAME = 'XXXX'
    IMPORT_FILE.PATH = 'YYYY'
    IMPORT_FILE.FILE_TYPE = IMPORT_FILE.FILE_NAME.apply(create_filetype)
    IMPORT_FILE.FILE_TYPE2 = IMPORT_FILE.FILE_NAME2.apply(create_filetype)
    IMPORT_FILE.DOC_TYPE2 = 'VIEW'
    IMPORT_FILE.FILE_NAME3 = np.nan
    IMPORT_FILE.FILE_TYPE3 = np.nan
    IMPORT_FILE.DOC_TYPE3 = np.nan
    IMPORT_FILE.DT_CRE = dt.datetime.today().strftime("%d.%m.%y")
    IMPORT_FILE.USER_CREATED = username.upper()
    IMPORT_FILE.ROWSTATE = 'Frigitt'
    IMPORT_FILE.CONTRACT = 10
    IMPORT_FILE.REFERANSE = np.nan
    IMPORT_FILE.dropna(subset=['DOC_CLASS', 'FORMAT_SIZE'], inplace=True)
    IMPORT_FILE.set_index('DOC_CLASS', inplace=True)
    IMPORT_FILE.MCH_CODE = IMPORT_FILE.MCH_CODE.apply(str).apply(lambda x: x.split(';')[0])  # hent den første mchkoden fra liste
    return IMPORT_FILE


def get_unique_rows(df, column):
    return df.drop_duplicates(subset=[column], keep=False)

def get_duplicate_rows(df, column):
    return df[df.duplicated(subset=[column], keep=False)]

def find_rows_with_multiple_document_types(df1, column_1, column_2, get_rows=True):

    df1 = df1[(df1.duplicated(subset=[column_1], keep=False))].copy()
    df2 = df1[(df1.duplicated(subset=[column_2], keep=False))].copy()

    if get_rows == True:
        df3 = df1[df1.index.isin(df2.index)]
    else:
        df3 = df1[~df1.index.isin(df2.index)]

    return df3


def create_import_file(df):
    """Lager import fil til IFS"""

    documents = import_documents(get_unique_rows(df, 'Dokumentnummer'))
    drawings = import_drawings(get_duplicate_rows(df, 'Dokumentnummer'))

    if drawings.empty == False and documents.empty == False:
        return documents.append(drawings)
    elif drawings.empty == True:
        return documents
    elif documents.empty == True:
        return drawings
    else:
        return None

def create_error_file(df):
    """Lag dataframe med filer som markert opp til å svare ut mer enn 1 dokumentttype"""

    duplicate_document_types = find_rows_with_multiple_document_types(df, 'Dokumenttype', 'FileName', get_rows=True)
    duplicate_document_types = import_documents(get_duplicate_rows(duplicate_document_types, 'Dokumentnummer'))

    return duplicate_document_types


def remove_first_value(items, sep=';'):

    """Fjern første MchKode fra listen. Dersom det kun finns en MchKode returner 0"""
    items = str(items)
    items = [item for item in items.split(sep)][1:]

    if len(items) == 0:
        return '0'
    else:
        return ';'.join(items)

def import_file_mch_codes(df, error):
    """Lager import fil for objektkobler"""

    df = df.copy()

    # fjerner eventuelle rader som er med i feilrapport - disse er heller ikke med i IMPORT_FILE
    df = pd.concat([df, error]).drop_duplicates(keep=False)

    # fjerner første mchkode siden den allerede er blitt brukt i IMPORT_FILE
    df['Mch Code'] = df['Mch Code'].apply(remove_first_value)

    # oppretter mal for importfil
    IMPORT_FILE_MCH_CODES = pd.DataFrame(columns=['FILE_NAME', 'CONTRACT', 'MCH_CODE'])

    # fjerner linjer som ikke inneholder sammenslåtte krav
    df = df[df['Antall sammenslåtte krav'] != 1]

    # splitter rad dersom den inneholder flere MchKoder
    df = split_rows(df, 'Mch Code', ';')

    IMPORT_FILE_MCH_CODES['FILE_NAME'] = df['FileName']
    IMPORT_FILE_MCH_CODES['MCH_CODE'] = df['Mch Code']
    IMPORT_FILE_MCH_CODES['MCH_CODE'] = IMPORT_FILE_MCH_CODES['MCH_CODE'].apply(lambda x: x.strip())
    IMPORT_FILE_MCH_CODES['CONTRACT'] = 10
    IMPORT_FILE_MCH_CODES.dropna(subset=['FILE_NAME'], inplace=True)
    IMPORT_FILE_MCH_CODES.set_index('FILE_NAME', inplace=True)

    return IMPORT_FILE_MCH_CODES

df1_cols = ['Dokumentnummer','Mch Code', 'Mch Name', 'Dokumenttype',]
df2_cols = ['Dokumentnummer', 'Title', 'FileName', ]
replace_columns_df = {'ObjectName': 'Mch Code', 'DocType': 'Dokumenttype', 'ObjectDescription': 'Mch Name', 'ContractorDocumentNo':'Dokumentnummer'}

if __name__ == "__main__":
    main()

