'''
spreadsheet common libraries
'''
from services import get_spreadsheet_service


def get_one_sheet_content(doc_id, sheet_name):
    '''
    return values of sheet contents
    '''
    service = get_spreadsheet_service()
    content = service.spreadsheets().values().get(
        spreadsheetId=doc_id,
        range=sheet_name,
    ).execute()

    return content.get("values")


def get_spreadsheet_meta(doc_id, fields=None):
    '''
    get meta data
    '''
    service = get_spreadsheet_service()
    return service.spreadsheets().get(
        spreadsheetId=doc_id,
        fields=fields
    ).execute()


def create_spreadsheet_file(body):
    service = get_spreadsheet_service()

    return service.spreadsheets().create(
        body=body
    ).execute()


def batch_update(doc_id, req):
    service = get_spreadsheet_service()
    return service.spreadsheets().batchUpdate(
        spreadsheetId=doc_id,
        body={"requests": req},
    ).execute()


def get_merge_cells_cmd(sheetId, rowStart, rowEnd, colStart, colEnd):
    return {
        "mergeCells": {
            "range": {
                "sheetId": sheetId,
                "startRowIndex": rowStart,
                "endRowIndex": rowEnd,
                "startColumnIndex": colStart,
                "endColumnIndex": colEnd,
            },
            "mergeType": "MERGE_ALL",
        }
    }


def get_color_by_code(color_code):
    return {
        "red": int(color_code[0:2], 16) / 255.0,
        "green": int(color_code[2:4], 16) / 255.0,
        "blue": int(color_code[4:6], 16) / 255.0

    }


def get_cell_value(value, textColor_code="000000", background_code="FFFFFF", try_use_number=False, number_format=None):
    textColor = get_color_by_code(textColor_code)
    backgroundColor = get_color_by_code(background_code)

    ret = {
        "userEnteredValue": {},
        "userEnteredFormat": {},
    }
    if try_use_number:
        ret["userEnteredValue"]["numberValue"] = value
        if number_format:
            ret["userEnteredFormat"]["numberFormat"] = number_format
    else:
        ret["userEnteredValue"]["stringValue"] = value

    if textColor is not None or backgroundColor is not None:
        if textColor:
            ret["userEnteredFormat"]["textFormat"] = {
                "foregroundColor": textColor}
        if backgroundColor:
            ret["userEnteredFormat"]["backgroundColor"] = backgroundColor

    return ret


def get_ge_rule(sheetId, threshold,
                foreground_rgb="FF0000", background_rgb="FFFFFF"):
    foreground = get_color_by_code(foreground_rgb)
    background = get_color_by_code(background_rgb)
    return {
        "ranges": [
            {
                "sheetId": sheetId,
                "startRowIndex": 0,
                "endRowIndex": 1000,
                "startColumnIndex": 0,
                "endColumnIndex": 1000,
            },
        ],
        "booleanRule": {
            "condition": {
                "type": "NUMBER_GREATER_THAN_EQ",
                "values": [
                        {
                            "userEnteredValue": threshold,
                        }
                ],
            },
            "format": {
                "backgroundColor": background,
                "textFormat": {
                    "foregroundColor": foreground,
                },
            }
        }
    }


def get_full_border(sheetId, rowStart, rowEnd, colStart, colEnd):
    return {
        'updateBorders': {
            'range': {
                'sheetId': sheetId,
                'startRowIndex': rowStart,
                'endRowIndex': rowEnd,
                'startColumnIndex': colStart,
                'endColumnIndex': colEnd,
            },
            'innerHorizontal': {
                'style': 'SOLID',
                'width': 1,
                'color': {
                    'red': 0.0,
                    'green': 0.0,
                    'blue': 0.0
                }
            },
            'innerVertical': {
                'style': 'SOLID',
                'width': 1,
                'color': {
                    'red': 0.0,
                    'green': 0.0,
                    'blue': 0.0
                }
            },
            'top': {
                'style': 'SOLID',
                'width': 1,
                'color': {
                    'red': 0.0,
                    'green': 0.0,
                    'blue': 0.0
                }
            },
            'bottom': {
                'style': 'SOLID',
                'width': 1,
                'color': {
                    'red': 0.0,
                    'green': 0.0,
                    'blue': 0.0
                }
            },
            'left': {
                'style': 'SOLID',
                'width': 1,
                'color': {
                    'red': 0.0,
                    'green': 0.0,
                    'blue': 0.0
                }
            },
            'right': {
                'style': 'SOLID',
                'width': 1,
                'color': {
                    'red': 0.0,
                    'green': 0.0,
                    'blue': 0.0
                }
            }
        }}
