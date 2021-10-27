import streamlit as st
import numpy as np
import pandas as pd
import datetime as dt
import re
import base64

st.title('Lag lastefil transformatorstasjon')

username = st.text_input(label='Steg 1: Skriv inn navn')

if username:
    st.success(f'Navn lagret')

substation = st.text_input(label='Steg 2: Skriv inn navn på stasjon: eg. AND Andeby')

if substation:
    st.success(f'Du har valgt {substation} stasjon')

def main():

    excel_file = st.file_uploader("Steg 3: Last opp kravsliste", type="xlsx")

    if excel_file is not None:
        df1 = pd.read_excel(excel_file)
        df1 = df1.loc[:, ~df1.columns.str.match('Unnamed')]
        df1.rename(columns=replace_columns_kravliste, inplace=True)
        df1 = split_rows(df1, 'Dokumentnummer', ';')
        df1 = group_columns(df1, 'Dokumentnummer','Dokumenttype', 'Ifs klasse', 'Ifs format')
        df1['Dokumentnummer'] = df1['Dokumentnummer'].apply(lambda x: x.strip())
        st.dataframe(df1)
        st.success('Import av kravsliste vellykket')


    csv_file = st.file_uploader("Steg 4: Last opp csv fil med csv export", type='csv', )

    if csv_file is not None:
        df2 = pd.read_csv(csv_file, sep=';', encoding='latin-1', engine='python')
        df2 = df2.loc[:, ~df2.columns.str.match('Unnamed')]
        df2 = df2.rename(columns=replace_columns_csv_export)
        df2['file name'] = df2['file name'].apply(str).apply(remove_comments_regex)
        df2 = split_rows(df2, 'file name', ',')
        df2['file name'] = df2['file name'].apply(drop_non_files)
        df2['Dokumentnummer'] = df2['Dokumentnummer'].apply(str).apply(lambda x: x.strip())
        df2['Dokumentnummer'] = df2['Dokumentnummer'].apply(lambda x: x.split(' ')[0])
        df2 = df2[pd.notnull(df2['file name'])]
        st.dataframe(df2)
        st.success('Import av csv export vellykket')

    merge_button = st.button('Slå sammen kravsliste og csv export')

    if merge_button:

        try:
            df = merge_csv_and_excel(df1, df2, 'Dokumentnummer')
            df = create_new_title(df)
            lastefil = create_import_file(df)
            st.dataframe(lastefil)
            date = pd.datetime.today().strftime("%d.%m.%y")
            number_of_files = lastefil.FILE_TYPE.count() + lastefil.FILE_TYPE2.count()
            csv1 = lastefil.to_csv(sep=';', encoding='latin-1', index=True)
            b641 = base64.b64encode(csv1.encode('latin-1')).decode()
            href1 = f'<a href="data:file/csv;base64,{b641}" download="lastefil IFS med {number_of_files} filer ({date}).csv">Last ned CSV for filimport til IFS med {number_of_files} filer</a>'
            st.markdown(href1, unsafe_allow_html=True)
            st.success("Opprettelse av lastefil var vellykket!")

            error_file = create_error_file(df)

            lastefil_objektkoblinger = import_file_mch_codes(df, error_file)

            if lastefil_objektkoblinger.empty != True:
                st.dataframe(lastefil_objektkoblinger)
                st.success('Opprettelse av lastefil for objektkoblinger av vellykket!')
                number_of_obj = lastefil_objektkoblinger.MCH_CODE.count()
                number_of_obj_files = lastefil_objektkoblinger.reset_index()['FILE_NAME'].nunique()
                csv2 = lastefil_objektkoblinger.to_csv(sep=';', encoding='latin-1', index=True)
                b642 = base64.b64encode(csv2.encode('latin-1')).decode()
                href2 = f'<a href="data:file/csv;base64,{b642}" download="lastefil IFS med {number_of_obj} objektkoblinger ({date}).csv">Last ned CSV fil for import av {number_of_obj} objektkoblinger til IFS for {number_of_obj_files} filer</a>'
                st.markdown(href2, unsafe_allow_html=True)

        except UnboundLocalError:
            st.error('Fil mangler')


def create_new_title(df):
    if 'Komponent' in df.columns:
        df['new file name'] = df['Dokumenttype'] + '_ ' + df['Komponent'] + ' - ' +df['Dokumenttittel'] + ', ' +  df['Dokumentnummer'] + f', {substation} '
    if 'Fabrikant' in df.columns:
        df['new file name'] = df['Dokumenttype'] + '_ ' + df['Dokumenttittel'] + ', ' + df['Fabrikant'] + ', ' + df['Type'] + f', {substation} '
    else:
        df['new file name'] = df['Dokumenttype'] + '_ ' + df['Dokumenttittel'] + ', ' +  df['Dokumentnummer'] + f', {substation} '

    return df

def group_columns(df, *columns):
    """Behold alle rader i første kolonne som har dokumentnummer og slå sammen rader som har identisk dokumentnummer"""

    df.rename(columns={df.columns[0]: "Dokumentnummer"}, inplace=True)

    mch_code = 'MchKode'
    superior_mch_code = 'Overordnet MchKode'
    object_description = 'Objektbeskr.'

    df['Dokumentnummer'] = df['Dokumentnummer'].apply(str)

    new_columns = list()

    for column in columns:
        new_columns.append(df[column])

    mch_df = df[mch_code].groupby(new_columns).apply(set).reset_index()
    obj_df = df[object_description].groupby(new_columns).apply(set).reset_index()
    sup_mch_df = df[superior_mch_code].groupby(new_columns).apply(set).reset_index()

    new_df = mch_df.merge(sup_mch_df)
    new_df = new_df.merge(obj_df)

    join_strings = lambda x: '; '.join(x)

    new_df[object_description] = new_df[object_description].apply(join_strings)
    new_df[mch_code] = new_df[mch_code].apply(join_strings)
    new_df[superior_mch_code] = new_df[superior_mch_code].apply(join_strings)

    new_df['Antall sammenslåtte krav'] = new_df[mch_code].apply(lambda x: x.split('; ')).apply(list).apply(len)

    return new_df

def split_rows(df, column, sep=',', keep=False):
    """Lag 1 ny rad per filnavn dersom det ligger flere filer i kolonnen """

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

replace_columns_kravliste = {'Dokumenttittel': 'Dokumentnummer', 'Dokument nummer': 'Dokumentnummer', 'Objekt id': 'MchKode',
                       'Doc Class': 'Ifs klasse', 'Title Mal': 'Dokumenttype', 'Format Size': 'Ifs format',
                      'Sup Mch Code': 'Overordnet MchKode', 'Overordnet Mch Kode': 'Overordnet MchKode', 'Mch Code': 'MchKode',
                       'Komponent': 'Objektbeskr.', 'Mch Name': 'Objektbeskr.', 'Lev.dok.nr.': 'Dokumentnummer',
                             'Contractor Doc.no': 'Dokumentnummer', 'Title': 'Dokumenttittel',
                             'Leverandørs dok. nr.': 'Dokumentnummer', 'Tittel': 'Dokumenttittel',
                             'Status': 'Dokumentnummer', 'Kommentar': 'Dokumentnummer'}


file_extensions = ('pdf', 'xlsx', 'jpg', 'dwg', 'zip', 'PDF', 'XLSX', 'JPG', 'DWG', 'ZIP')

def remove_comments_regex(files):
    """Fjern internt kommentarområde fra df['file name'] i csv fil"""

    if 'Z' or 'C' in files:
        file = re.sub(r'((\S+ - \S+ \S+ - )(.)*?([0-9][0-9][A-Z]))', '', files)
    else:
        file = re.sub(r'(\S+ - \S+ \S+ - )\d\d\d\d\d(.)*(_\d\d)', '', files)

    file = file.strip().rstrip(',').lstrip(', ')

    if file:
        return file
    else:
        return np.nan

def drop_non_files(filename):
    filename = str(filename)

    if filename.endswith(file_extensions):
        return filename
    else:
        return np.nan


replace_columns_csv_export = {'Doc.no.':'Dokumentnummer', 'Document title': 'Dokumenttittel', 'Lev.dok.nr.': 'Dokumentnummer',
                            'Tittel': 'Dokumenttittel', 'Subcontractors doc.no.': 'Dokumentnummer',
                                'Suppliers doc.no.': 'Dokumentnummer', 'Supplier doc. no.' : 'Dokumentnummer', 'Title': 'Dokumenttittel',
                                "Contractor's Doc. no": "Dokumentnummer", 'Kommentar': 'Dokumentnummer',
                                "Contractor's Doc. No" : 'Dokumentnummer', 'Contractor Doc.no': 'Dokumentnummer'}


def create_filetype(filename):
    """Fyll ut felt FILE_TYPE basert på filekstensjon"""

    filename = str(filename)

    if filename.endswith(file_extensions):
        return [word for word in filename.split('.')][-1].upper()
    else:
        return ''

def import_documents(df):

    lastefil = pd.DataFrame(columns=['DOC_CLASS', 'DOC_NO', 'DOC_SHEET', 'DOC_REV', 'FORMAT_SIZE', 'REV_NO',
       'TITLE', 'DOC_TYPE', 'INFO', 'FILE_NAME', 'LOCATION_NAME', 'PATH',
       'FILE_TYPE', 'FILE_NAME2', 'FILE_TYPE2', 'DOC_TYPE2', 'FILE_NAME3',
       'FILE_TYPE3', 'DOC_TYPE3', 'DT_CRE', 'USER_CREATED', 'ROWSTATE',
       'MCH_CODE', 'CONTRACT', 'REFERANSE'])

    lastefil.TITLE = df['new file name']
    lastefil.FILE_NAME = list(df['file name'].apply(lambda x: x.strip()))
    lastefil.DOC_CLASS = df['Ifs klasse']
    lastefil.DOC_NO = pd.np.nan
    lastefil.DOC_SHEET = 1
    lastefil.DOC_REV = 1
    lastefil.FORMAT_SIZE = df['Ifs format']
    lastefil.REV_NO = 1
    lastefil.DOC_TYPE = 'ORIGINAL'
    lastefil.INFO = pd.np.nan
    lastefil.LOCATION_NAME = 'XXXX'
    lastefil.PATH = 'YYYY'
    lastefil.FILE_TYPE = lastefil.FILE_NAME.apply(create_filetype)
    lastefil.FILE_NAME2 = pd.np.nan
    lastefil.FILE_TYPE2 = pd.np.nan
    lastefil.DOC_TYPE2 = pd.np.nan
    lastefil.FILE_NAME3 = pd.np.nan
    lastefil.FILE_TYPE3 = pd.np.nan
    lastefil.DOC_TYPE3 = pd.np.nan
    lastefil.DT_CRE = dt.datetime.today().strftime("%d.%m.%y")
    lastefil.USER_CREATED = username.upper()
    lastefil.ROWSTATE = 'Frigitt'
    lastefil.MCH_CODE = df['MchKode']
    lastefil.CONTRACT = 10
    lastefil.REFERANSE = pd.np.nan
    lastefil.dropna(subset=['DOC_CLASS', 'FORMAT_SIZE'], inplace=True)
    lastefil.set_index('DOC_CLASS', inplace=True)
    lastefil.MCH_CODE = lastefil.MCH_CODE.apply(str).apply(lambda x: x.split(';')[0]) # hent den første mchkoden fra liste

    return lastefil


def import_drawings(df):

    lastefil = pd.DataFrame(columns=['DOC_CLASS', 'DOC_NO', 'DOC_SHEET', 'DOC_REV', 'FORMAT_SIZE', 'REV_NO',
                                     'TITLE', 'DOC_TYPE', 'INFO', 'FILE_NAME', 'LOCATION_NAME', 'PATH',
                                     'FILE_TYPE', 'FILE_NAME2', 'FILE_TYPE2', 'DOC_TYPE2', 'FILE_NAME3',
                                     'FILE_TYPE3', 'DOC_TYPE3', 'DT_CRE', 'USER_CREATED', 'ROWSTATE',
                                     'MCH_CODE', 'CONTRACT', 'REFERANSE'])

    lastefil.MCH_CODE = df['MchKode']
    lastefil.TITLE = df['new file name']
    lastefil.DOC_CLASS = df['Ifs klasse']
    lastefil.FORMAT_SIZE = df['Ifs format']

    lastefil.FILE_NAME = df['file name'][df['file name'].str.contains('dwg') | df['file name'].str.contains('DWG')]
    lastefil.FILE_NAME2 = df['file name'][df['file name'].str.contains('pdf') | df['file name'].str.contains('PDF')]

    lastefil = lastefil.groupby(['TITLE', ]).first().reset_index()

    lastefil.DOC_NO = pd.np.nan
    lastefil.DOC_SHEET = 1
    lastefil.DOC_REV = 1
    lastefil.REV_NO = 1
    lastefil.DOC_TYPE = 'ORIGINAL'
    lastefil.INFO = pd.np.nan
    lastefil.LOCATION_NAME = 'XXXX'
    lastefil.PATH = 'YYYY'
    lastefil.FILE_TYPE = lastefil.FILE_NAME.apply(create_filetype)
    lastefil.FILE_TYPE2 = lastefil.FILE_NAME2.apply(create_filetype)
    lastefil.DOC_TYPE2 = 'VIEW'
    lastefil.FILE_NAME3 = pd.np.nan
    lastefil.FILE_TYPE3 = pd.np.nan
    lastefil.DOC_TYPE3 = pd.np.nan
    lastefil.DT_CRE = dt.datetime.today().strftime("%d.%m.%y")
    lastefil.USER_CREATED = username.upper()
    lastefil.ROWSTATE = 'Frigitt'
    lastefil.CONTRACT = 10
    lastefil.REFERANSE = pd.np.nan
    lastefil.dropna(subset=['DOC_CLASS', 'FORMAT_SIZE'], inplace=True)
    lastefil.set_index('DOC_CLASS', inplace=True)
    lastefil.MCH_CODE = lastefil.MCH_CODE.apply(str).apply(lambda x: x.split(';')[0])  # hent den første mchkoden fra liste
    lastefil = lastefil[['DOC_NO', 'DOC_SHEET', 'DOC_REV', 'FORMAT_SIZE', 'REV_NO',
                         'TITLE', 'DOC_TYPE', 'INFO', 'FILE_NAME', 'LOCATION_NAME', 'PATH',
                         'FILE_TYPE', 'FILE_NAME2', 'FILE_TYPE2', 'DOC_TYPE2', 'FILE_NAME3',
                         'FILE_TYPE3', 'DOC_TYPE3', 'DT_CRE', 'USER_CREATED', 'ROWSTATE',
                         'MCH_CODE', 'CONTRACT', 'REFERANSE']]
    return lastefil


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
    """Lag import fil til IFS" for dokumenter og tegninger"""

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

    duplicate_document_types = find_rows_with_multiple_document_types(df, 'Dokumenttype', 'file name', get_rows=True)
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
    """Lag import fil for mch koblinger"""
    
    df2 = df.copy()

    # fjerner eventuelle rader som er med i feilrapport - disse er heller ikke med i lastefil
    df2 = pd.concat([df2, error]).drop_duplicates(keep=False)

    # fjerner første mchkode siden den allerede er blitt brukt i lastefil
    df2['MchKode'] = df2['MchKode'].apply(remove_first_value)

    lastefil_objektkobling = pd.DataFrame(columns=['FILE_NAME', 'CONTRACT', 'MCH_CODE'])

    # fjerner
    df2 = df2[df2['Antall sammenslåtte krav'] != 1]

    # splitter enkelt rad opp i flere rader dersom den inneholder flere MchKoder
    df2 = split_rows(df2, 'MchKode', ';')

    lastefil_objektkobling['FILE_NAME'] = df2['file name']
    lastefil_objektkobling['MCH_CODE'] = df2['MchKode']
    lastefil_objektkobling['MCH_CODE'] = lastefil_objektkobling['MCH_CODE'].apply(lambda x: x.strip())
    lastefil_objektkobling['CONTRACT'] = 10
    lastefil_objektkobling.dropna(subset=['FILE_NAME'], inplace=True)
    lastefil_objektkobling.set_index('FILE_NAME', inplace=True)

    return lastefil_objektkobling

if __name__ == "__main__":
    main()

