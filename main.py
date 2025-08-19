import os
import json
import shutil
import zipfile
import random
import string
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.draw import Frame, Image
from odf.text import P
from odf.namespaces import XLINKNS
import xml.etree.ElementTree as ET
import config



def main():
    make_new_output_folder()
    loop_over_sheet_2()
    print("✅ Done! Images extracted and json file created.")


def make_new_output_folder():
    if os.path.exists(config.OUTPUT_FOLDER):
        shutil.rmtree(config.OUTPUT_FOLDER) 
    os.makedirs(config.OUTPUT_FOLDER)



def loop_over_sheet_2():
    ns = {
        'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
        'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'xlink': 'http://www.w3.org/1999/xlink'
    }

    with zipfile.ZipFile(config.ODS_FILE, 'r') as ods:
        content_xml = ods.read('content.xml')

        root = ET.fromstring(content_xml)

        sheets = root.findall('.//table:table', ns)

        sheet2 = sheets[1]  

        insects = []

        images = []

        for row_idx, row in enumerate(sheet2.findall('table:table-row', ns), start=1):
            cells = row.findall('table:table-cell', ns)
            if row_idx >= 78:
                continue

            insect_data = get_insect_data(cells, row_idx, ns)
            insects.append(insect_data)
            first_letters = get_first_letters(insect_data['generic_name'])

            image_data = export_and_rename_images(cells, first_letters, insect_data['insect_id'], row_idx, ns, ods)

            images.append(image_data['image1'])

            if row_idx in config.IGNORE_IMAGES:
                if 2 not in config.IGNORE_IMAGES[row_idx]:
                    images.append(image_data['image2'])
                if 3 not in config.IGNORE_IMAGES[row_idx]:
                    images.append(image_data['image3'])

    
    insects_json = json.dumps(insects, indent=4)
    with open(config.JSON_FILE, 'w') as json_file:
        json_file.write(insects_json)
    
    image_json = json.dumps(images, indent=4)
    with open('images.json', 'w') as image_file:
        image_file.write(image_json)
            
            
def get_insect_data(cells, row_idx, ns):
    generic_cell = cells[0] 
    specific_cell = cells[1]
    generic_name = None
    specific_name = None
    generic_name = get_name(generic_cell, ns)
    specific_name = get_name(specific_cell, ns)

    return {
        'insect_id': row_idx,
        'generic_name': generic_name,
        'specific_name': specific_name,
    }

def get_first_letters(name):
    parts = name.split()
    first_letters = ''.join(part[0].lower() for part in parts if part)
    return first_letters


def get_name(cell, ns):
    if cell is not None:
            text_elem = cell.find('.//text:p', ns)
            if text_elem is not None:
                return text_elem.text             
    return None 

def export_and_rename_images(cells, first_letters, insect_id, row_idx, ns, ods):

    image_1 = cells[2]
    image_2 = cells[3]
    image_3 = cells[4]

    image_2_url = None
    image_3_url = None

    s3_url = config.S3URL


    image_1_label = extract_and_label_image(image_1, first_letters, insect_id, ns, ods)
    image_1_url = s3_url + os.path.basename(image_1_label) if image_1_label else None


    if row_idx in config.IGNORE_IMAGES:
        if 2 not in config.IGNORE_IMAGES[row_idx]:
            image_2_label = extract_and_label_image(image_2, first_letters, insect_id, ns, ods)
            image_2_url = s3_url + os.path.basename(image_2_label) if image_2_label else None

        if 3 not in config.IGNORE_IMAGES[row_idx]:
            image_3_label = extract_and_label_image(image_3, first_letters, insect_id, ns, ods)
            image_3_url = s3_url + os.path.basename(image_3_label) if image_3_label else None

    


    return {'image1':{
        'insect_id': row_idx,
        'url': image_1_url,
    },

    'image2':{
        'insect_id': row_idx,
        'url': image_2_url,
    },

    'image3': {
        'insect_id': row_idx,
        'url': image_3_url,
    }
    }

def extract_and_label_image(image_cell,first_letters, insect_id, ns, ods):

    random_string = ''.join(random.choices(string.ascii_letters + string.digits, k=8)).lower()

    if image_cell is not None:
        image_elem = image_cell.find('.//draw:image', ns)
        if image_elem is not None:
            image_path = image_elem.attrib[f'{{{ns["xlink"]}}}href']
            image_data = ods.read(image_path)
            image_name = os.path.basename(image_path)
            file_type = image_name.split('.')[-1].lower()
            file_name = f"{insect_id}_{first_letters}_{random_string}.{file_type}"
            output_path = os.path.join(config.OUTPUT_FOLDER, file_name)
            with open(output_path, 'wb') as f:
                f.write(image_data)
            return file_name
    return None




if __name__ == "__main__":
    main()
