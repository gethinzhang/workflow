'''
spreadsheet common libraries
'''
from services import get_spreadsheet_service
from decimal import Decimal


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


def delete_spreadsheet_sheets(doc_id):
    service = get_spreadsheet_service()

    ret = get_spreadsheet_meta(doc_id)
    sheet_ids = [sheet['properties']['sheetId'] for sheet in ret['sheets']]
    # Prepare the request to remove all sheets
    requests = []
    for sheet_id in sheet_ids:
        requests.append({
            'deleteSheet': {
                'sheetId': sheet_id
            }
        })

    return batch_update(doc_id, requests)


def update_cell_value(doc_id, sheetName, range, value):
    service = get_spreadsheet_service()
    service.spreadsheets().values().update(
        spreadsheetId=doc_id, range=sheetName + '!' + range,
        valueInputOption='RAW', body=value
    ).execute()


def clear_sheet(doc_id, sheetId):
    service = get_spreadsheet_service()
    return service.spreadsheets().values().clear(
        spreadsheetId=doc_id,
        range=sheetId,
    )


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
        if isinstance(value, Decimal):
            value = float(value)
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


def get_ho_align(sheetId, align, startRow=None, endRow=None, startColumn=None, endColumn=None):
    range = {
        "sheetId": sheetId,
    }
    if startRow:
        range["startRowIndex"] = startRow
    if endRow:
        range["endRowIndex"] = endRow
    if startColumn:
        range["startColumnIndex"] = startColumn
    if endColumn:
        range["endColumnIndex"] = endColumn

    return {
        'repeatCell': {
            'range': range,
            'cell': {
                'userEnteredFormat': {
                    'horizontalAlignment': align,
                },
            },
            'fields': 'userEnteredFormat(horizontalAlignment)'
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


def get_autosize(sheetId):
    merge_req = []

    merge_req.append({
        'autoResizeDimensions': {
            'dimensions': {
                'sheetId': sheetId,
                'dimension': 'COLUMNS',
            },
        }}
    )

    return merge_req