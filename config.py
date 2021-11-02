file_extensions = ('docx', 'pdf', 'xlsx', 'jpg', 'dwg', 'zip')

column_dict = {'ObjectName': 'Mch Code', 'DocType': 'Dokumenttype', 'ObjectDescription': 'Mch Name', 'ContractorDocumentNo':'Dokumentnummer'}

doc_dict = {
            # DOKUMENTER
            'RP': ['Prøveprotokoll', 'ANLEGGSDOK', 'PRPROT'],
            'OM': ['Drift-, Montasje- og Vedlikeholdsmanual', 'ANLEGGSDOK', 'TEKDOK'],
            
            # TEGNINGER
            'XD': ['Målskisse', 'TEGNINGER', 'MONT'],
            'XF': ['Fundamenttegning', 'TEGNINGER', 'FUNDT'],
            'XK': ['Interne strømløpsskjema', 'TEGNINGER', 'SKJEMA'],
            'XQ': ['Stativtegning', 'TEGNINGER', 'MONT'],
            
                }
