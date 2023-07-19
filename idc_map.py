from drive import create_file_in_folder, GOOGLE_SHEET_MIME
import spreadsheet
import drive

HIDDEN_BY_USER_FIELD = 'sheets(data(columnMetadata(hiddenByUser))),sheets(data(rowMetadata(hiddenByUser))),sheets(properties)'
SHEET_ID = "1EoORLTqwmRR1gFPdm_R3dNhIj-XYsh6Gf5urYy_ZeCI"
PLATS = set([
    "AZ",
    "DB",
    "Cache",
    "Log Platform",
    "MQ",
    "Distributed Analytics Engine",
    "Monitoring Platform SG",
    "Storage",
    "Monitoring Platform SZ",
    "Data Transmission Service",
    "MMDB",
    "Distributed Coordination",
    "Video Network",
    "Monitoring Platform",
    "IDbank_Infra&CorpIT",
    "NDRE",
])


def get_platform_map():
    ret = {}
    ss = spreadsheet.get_spreadsheet_meta(
        SHEET_ID,
        # fields=HIDDEN_BY_USER_FIELD
    )

    sheet = ss.get("sheets")[0]
    properties = sheet.get("properties")
    title = properties.get("title")

    rows = spreadsheet.get_one_sheet_content(SHEET_ID, title)
    for row in rows[1:]:
        try:
            date, bu, platform, region, idc, server_config, qty, category, location, BU = row[
                :10]
        except ValueError:
            print("hahaha, " + str(row))
            exit(0)
        #if BU != "Non-Bank":
        #    continue
        #if idc == "DC West":
        #    continue
        #if platform not in PLATS:
        #    print(F"plat {platform} filtered")
        #    continue

        if platform not in ret:
            ret[platform] = {}
        if server_config not in ret[platform]:
            ret[platform][server_config] = 0
        ret[platform][server_config] += int(qty)

    return ret


def write_to_final_files(r):
    body = {
        "properties": {
            "title": "For HX",
        },
        "sheets": [

        ],
    }

    whole_data = []
    merge_rows = []
    for platform, server_map in r.items():
        start = len(whole_data)
        for server_config, qty in server_map.items():
            row_data = [
                spreadsheet.get_cell_value(platform),
                spreadsheet.get_cell_value(server_config),
                spreadsheet.get_cell_value(qty, try_use_number=True)
            ]
            whole_data.append({"values": row_data})
        end = len(whole_data)
        merge_rows.append((start, end))

    body["sheets"].append(
        {
            "properties": {
                "title": "platform config map",
            },
            "data": {
                "rowData": whole_data
            }
        }
    )
    ret = spreadsheet.create_spreadsheet_file(body)
    doc_id = ret["spreadsheetId"]
    sheetId = ret["sheets"][0]["properties"]["sheetId"]

    merge_req = []
    for merge_row in merge_rows:
        merge_req.append(spreadsheet.get_merge_cells_cmd(
            sheetId, merge_row[0], merge_row[1], 0, 1)
        )

    spreadsheet.batch_update(doc_id, merge_req)


if __name__ == "__main__":
    r = get_platform_map()
    write_to_final_files(r)
