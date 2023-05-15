'''
common google drive api's
'''
from services import get_drive_service


GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet"


def create_file_in_folder(folder_id, file_name, mime_type=""):
    '''
    create a empty file in folder [folder_id]

    return ID of folder
    '''
    service = get_drive_service()
    file_meta = {
        "name": file_name,
        "parents": [folder_id],
    }

    if mime_type:
        file_meta["mimeType"] = mime_type

    file = service.files().create(
        body=file_meta, supportsAllDrives=True, fields="id").execute()

    return file.get("id")


def move_doc_to_folder(doc_id, folder):
    drive_service = get_drive_service()
    file = drive_service.files().get(
        fileId=doc_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    file = drive_service.files().update(fileId=doc_id, addParents=folder,
                                        removeParents=previous_parents,
                                        fields='id, parents').execute()
    return file.get('parents')
