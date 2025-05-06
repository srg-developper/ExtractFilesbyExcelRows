from photo_copier.copier import copy_photos_from_excel

# Update these paths or load from .env
excel_file_path = "EXCEL_FILE.xlsx"
source_folder = "SOURCE_FOLDER"
destination_folder = "DEST_FOLDER"

copy_photos_from_excel(excel_file_path, source_folder, destination_folder)
