import json
import multiprocessing
from pathlib import Path

from MetaDataExtractor import MetaDataExtractor

metadata_extractor = MetaDataExtractor()


def process_file(file_path):
    metadata = {"Type": "file"}
    metadata.update(metadata_extractor.get_file_security_info(file_path))
    return {str(file_path.name): metadata}


def main(root_directory, output_path):
    directory_tree = {}
    pool = multiprocessing.Pool(multiprocessing.cpu_count())

    for directory_path in Path(root_directory).rglob('*'):
        if directory_path.is_dir():
            directory_data = {
                "Type": "folder",
                "Parent-dir": str(directory_path.parent),
                "Children": {}
            }
            directory_tree[str(directory_path.name)] = directory_data

            children_files = list(directory_path.iterdir())
            if children_files:  # only if the directory has children
                file_metadata = pool.map(process_file, (child for child in children_files if child.is_file()))
                for metadata in file_metadata:
                    directory_data["Children"].update(metadata)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(directory_tree, f, indent=2)


if __name__ == "__main__":
    root_directory = Path(r"C:\Users\dimit\Downloads")
    output_path = ".\\stored_data.json"

#     main(root_directory, output_path)
#
# # Load the JSON file
# with open('output.json', 'r') as f:
#     json_dict = json.load(f)
#
# # Load the 'reg' dictionary
# # Assuming 'reg' dictionary looks like this: {'dimit': 'Dimitri', 'samar': 'Sam'}
# reg_dict = {'dimit': 'Dimitri', 'samar': 'Sam'}
#
# def replace_owner(data):
#     if isinstance(data, dict):
#         for key, value in data.items():
#             if key == 'Owner' and value in reg_dict:
#                 data[key] = reg_dict[value]
#             else:
#                 replace_owner(value)
#     elif isinstance(data, list):
#         for item in data:
#             replace_owner(item)
#
# # Call the replace_owner function
# replace_owner(json_dict)
#
# # Write the updated dictionary back to the JSON file
# with open('output.json', 'w') as f:
#     json.dump(json_dict, f, indent=2)