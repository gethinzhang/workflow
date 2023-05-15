
from drive import create_file_in_folder, GOOGLE_SHEET_MIME
import spreadsheet
import drive

HIDDEN_BY_USER_FIELD = 'sheets(data(columnMetadata(hiddenByUser))),sheets(data(rowMetadata(hiddenByUser))),sheets(properties)'
# PLH_FOLDER_ID = '1WkCP9Gl3IAunbqpNMOFluH6Nnz_MAfZW'
# PLH_FOLDER_ID = '1nagGqJQfdYV4Mgb0vUKj4bf-29mE68RR'
# PLH_FOLDER_ID = '1bO-hH5_dxY1Tdick7BsIAaEaw6rZs7zA'
PLH_FOLDER_ID = '1391PhT7Y9VX3bJxLRJen3Qz9iJ9TALAa'
# BASE_SHEET_ID = '16dQ0NOB7GMS1H19LVeKr4RpE507oEcJ7RoneiR5_LsU'
BASE_SHEET_ID = '1vnIYSguDfd4skdQtl1_ZILuK46xFwGZMf-V49oQQAPc'


def build_PLH_platform_usage_map():
    '''
    build a data map for
    {
        L1_PLH : {
            L2_PLH: {
                PLATFORM: {
                    indicator_name: "name"
                    tuple(9) of (quota, avg, usage) X 3
            }
        }
    }
    '''
    ss = spreadsheet.get_spreadsheet_meta(
        BASE_SHEET_ID,
        fields=HIDDEN_BY_USER_FIELD,
    )
    sheets = ss.get("sheets")

    m = {}

    for _, sheet in enumerate(sheets):
        properties = sheet.get("properties")
        data = sheet.get("data")
        platform = properties.get("title")
        if properties.get("hidden") is True:
            continue

        rowMeta = data[0]["rowMetadata"]
        colMeta = data[0]["columnMetadata"]

        rows = spreadsheet.get_one_sheet_content(BASE_SHEET_ID, platform)
        for j, row in enumerate(rows[2:], 2):
            if len(row) < 2:
                continue

            if "hiddenByUser" in rowMeta[j] and rowMeta[j]["hiddenByUser"] is True:
                continue
            v_row = []
            for z, c in enumerate(row):
                cm = colMeta[z]

                if "hiddenByUser" in cm and cm["hiddenByUser"] is True:
                    continue
                else:
                    v_row.append(c)

            if len(v_row) != 11:
                raise Exception(
                    F"unexpected line in {platform}\nrow {j}\nrow_content {row}\nfilter_info: {rowMeta[j]}\nv_row: {v_row}")

            product_line = v_row[0]
            if "." in product_line:
                l1, l2 = product_line.split(".")
            else:
                l1 = product_line
                l2 = product_line
            indicator = v_row[1]

            if l1 not in m:
                m[l1] = {}
            if l2 not in m[l1]:
                m[l1][l2] = {}

            m[l1][l2][platform] = {
                "data": v_row[2:],
                "indicator_name": indicator,
            }

    return m


def get_header_for_sheet():
    color = 'EE4D2D'
    white = 'FFFFFF'
    return [
        {
            "values": [
                spreadsheet.get_cell_value("", white, color),
                spreadsheet.get_cell_value("", white, color),
                spreadsheet.get_cell_value("April", white, color),
                spreadsheet.get_cell_value("", white, color),
                spreadsheet.get_cell_value("", white, color),
                spreadsheet.get_cell_value("March", white, color),
                spreadsheet.get_cell_value("", white, color),
                spreadsheet.get_cell_value("", white, color),
                spreadsheet.get_cell_value("Feb", white, color),
                spreadsheet.get_cell_value("", white, color),
                spreadsheet.get_cell_value("", white, color),
            ],
        },
        {
            "values": [
                spreadsheet.get_cell_value(
                    "Platform", white, color),
                spreadsheet.get_cell_value(
                    "Indicator", white, color),
                spreadsheet.get_cell_value(
                    "Quota", white, color),
                spreadsheet.get_cell_value(
                    "Quota Avg Usage (Monthly)", white, color),
                spreadsheet.get_cell_value(
                    "Quota Peak Usage (Monthly)", white, color),
                spreadsheet.get_cell_value(
                    "Quota", white, color),
                spreadsheet.get_cell_value(
                    "Quota Avg Usage (Monthly)", white, color),
                spreadsheet.get_cell_value(
                    "Quota Peak Usage (Monthly)", white, color),
                spreadsheet.get_cell_value(
                    "Quota", white, color),
                spreadsheet.get_cell_value(
                    "Quota Avg Usage (Monthly)", white, color),
                spreadsheet.get_cell_value(
                    "Quota Peak Usage (Monthly)", white, color),
            ]
        }
    ]


def write_to_plh_files(l1_name, l1_info):
    body = {
        "properties": {
            "title": F"APP Platform Quota Usage & Billing - {l1_name}",
        },
        "sheets": [
        ],
    }

    for l2_name, l2_info in l1_info.items():
        rows_data = get_header_for_sheet()
        for platform, row in l2_info.items():
            row_data = []
            row_data.append(spreadsheet.get_cell_value(platform))
            row_data.append(spreadsheet.get_cell_value(row["indicator_name"]))
            for col in row["data"]:
                if col[-1] == '%':
                    row_data.append(
                        spreadsheet.get_cell_value(float(col[:-1])/100.0,
                                                   try_use_number=True,
                                                   number_format={
                            'type': 'PERCENT',
                            'pattern': '#0.00%',
                        }))
                else:
                    row_data.append(spreadsheet.get_cell_value(col))

            rows_data.append({"values": row_data})

        body["sheets"].append(
            {
                "properties": {
                    "title": F"L2 - {l2_name}",
                },
                "data": {
                    "rowData": rows_data
                },
            },
        )

    ret = spreadsheet.create_spreadsheet_file(body)
    doc_id = ret["spreadsheetId"]

    merge_req = []
    # merge all sheet
    for sheet in ret["sheets"]:
        sheetId = sheet["properties"]["sheetId"]
        merge_req.append(spreadsheet.get_merge_cells_cmd(sheetId, 0, 1, 2, 5))
        merge_req.append(spreadsheet.get_merge_cells_cmd(sheetId, 0, 1, 5, 8))
        merge_req.append(spreadsheet.get_merge_cells_cmd(sheetId, 0, 1, 8, 11))
        merge_req.append(
            {
                "addConditionalFormatRule": {
                    "rule": spreadsheet.get_ge_rule(sheetId, "100%"),
                    "index": 0,
                },
            }
        )

    spreadsheet.batch_update(doc_id, merge_req)
    drive.move_doc_to_folder(ret["spreadsheetId"], PLH_FOLDER_ID)
    return ret


if __name__ == "__main__":
    m = build_PLH_platform_usage_map()
    # for plh, plh_info in m.items():
    #    write_to_plh_files(plh, plh_info)
    write_to_plh_files("recommendation", m["recommendation"])
