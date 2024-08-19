import json
import re
import pandas as pd


def pre_process_text(text):
    """Removes all whitespace characters (spaces, tabs, newlines) and the □ sign from the text and replace placeholders with tags."""

    text = remove_whitespaces_newlines_tabs_signs(text)
    placeholders_tags_dict = [(r'ISIN:[_]+',f'ISIN:$ISIN$'), (r'\(Vertragsnummer[_]+\)', '$CONTRACT_NUMBER$'), (r'\(Datumergänzen\)', '$DATE$'), (r'[_]+€(\(aktuellerBestand-desFonds,derfürdenEntnahmeplanhinterlegtwirdundnichtdengesamtenDepotbestand\))?', '$Amount$'), (r'Einmalig[_]+', 'Einmalig$Amount$'),(r'Fonds[_]+','Fonds$FONDS$'),(r'\(zahlweise\)[_]*€?','$PAYMENT_TYPE$'), (r'ab\(Datum\)', 'ab$DATE$'), (r'AnzahlderAnteile[_]+', 'AnzahlderAnteile$AMOUNT_SHARES$'), (r'vom[_]+', 'vom$DATE$'), (r'am\(…\)', 'am$DATE$')]

    for (pattern, tag) in placeholders_tags_dict:
        text = replace_placeholders_with_tags(text, pattern, tag)


    return text

def remove_whitespaces_newlines_tabs_signs(text):
    """Removes all whitespace characters (spaces, tabs, newlines) and the □ sign from the text"""
    return re.sub(r'[\s□]+', '', text)


def replace_placeholders_with_tags(text, pattern, replacement_tag):
    """Replaces given patterns in a text with the replacement_tag"""

    pattern = re.compile(pattern)
    return pattern.sub(replacement_tag, text)


def excel_to_dict(path = 'gee_text_blocks.xlsx'):
    """Converts the specified Excel sheet to a dictionary."""

    excel_data = pd.ExcelFile(path)
    df = pd.read_excel(excel_data, sheet_name='Textbausteine')

    text_block_dict = {}
    rows = []

    for index, row in df.iterrows():
        name = row.iloc[0] # column A
        text = row.iloc[2] # column C

        if pd.notna(name) and pd.notna(text):
               cleaned_name = remove_whitespaces_newlines_tabs_signs(name)
               processed_text = pre_process_text(text)
               text_block_dict[cleaned_name] = processed_text
               rows.append(cleaned_name)

    print(rows)
    return text_block_dict


def save_as_json(text_dict, json_file_path='gee_text_blocks.json'):
    """Saves the text block dictionary as a JSON file."""

    try:
        with open(json_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(text_dict, json_file, ensure_ascii=False, indent=4)

        print(f"Data has been written to {json_file_path}")

    except Exception as e:
        print(f"An error occurred while writing the JSON file: {e}")


if __name__ == '__main__':
    gee_text_block_dict = excel_to_dict()
    save_as_json(gee_text_block_dict)

