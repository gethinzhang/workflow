from drive import create_file_in_folder, GOOGLE_SHEET_MIME
import spreadsheet
import drive

HIDDEN_BY_USER_FIELD = 'sheets(data(columnMetadata(hiddenByUser))),sheets(data(rowMetadata(hiddenByUser))),sheets(properties)'
SHEET_ID = "1Khf357_9ztSmi5Hsf2WJI6T-Q0rvhpRgJDp5Zxm6G24"


def get_platform_map():
    ret = {}
    ss = spreadsheet.get_spreadsheet_meta(
        SHEET_ID,
        fields=HIDDEN_BY_USER_FIELD
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
        if BU != "Exclude Bank":
            continue
        if category != "APP":
            continue
        if bu == "seamoney" or bu == "shopee":
            bu_cate = bu
        else:
            continue

        if bu_cate not in ret:
            ret[bu_cate] = {}
        if platform not in ret[bu_cate]:
            ret[bu_cate][platform] = {}
        if server_config not in ret[bu_cate][platform]:
            ret[bu_cate][platform][server_config] = 0
        ret[bu_cate][platform][server_config] += int(qty)

    return ret


def write_to_final_files(r, bu_cate):
    body = {
        "properties": {
            "title": F"For HX ({bu_cate})",
        },
        "sheets": [

        ],
    }

    for platform, server_map in r.items():
        whole_data = []
        for server_config, qty in server_map.items():
            row_data = [
                spreadsheet.get_cell_value(server_config),
                spreadsheet.get_cell_value(qty, try_use_number=True)
            ]
            whole_data.append({"values": row_data})

        body["sheets"].append(
            {
                "properties": {
                    "title": platform
                },
                "data": {
                    "rowData": whole_data
                }
            }
        )
    ret = spreadsheet.create_spreadsheet_file(body)


if __name__ == "__main__":
    ret = get_platform_map()
    for cate, r in ret.items():
        write_to_final_files(r, cate)
