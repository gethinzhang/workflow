
from drive import create_file_in_folder, GOOGLE_SHEET_MIME
import spreadsheet
import drive

HIDDEN_BY_USER_FIELD = 'sheets(data(columnMetadata(hiddenByUser))),sheets(data(rowMetadata(hiddenByUser))),sheets(properties)'
# PLH_FOLDER_ID = '1WkCP9Gl3IAunbqpNMOFluH6Nnz_MAfZW'
# PLH_FOLDER_ID = '1nagGqJQfdYV4Mgb0vUKj4bf-29mE68RR'
# PLH_FOLDER_ID = '1bO-hH5_dxY1Tdick7BsIAaEaw6rZs7zA'
# PLH_FOLDER_ID = '1391PhT7Y9VX3bJxLRJen3Qz9iJ9TALAa'
# BASE_SHEET_ID = '16dQ0NOB7GMS1H19LVeKr4RpE507oEcJ7RoneiR5_LsU'
# BASE_SHEET_ID = '1vnIYSguDfd4skdQtl1_ZILuK46xFwGZMf-V49oQQAPc'
# BASE_SHEET_ID = '16dQ0NOB7GMS1H19LVeKr4RpE507oEcJ7RoneiR5_LsU'
BASE_SHEET_ID = '1ZWjzJSW1q0qm6ZABAhHpcOPOUE97cW8m901VL58A5Pk'
PLH_FOLDER_ID = '1PXuqqgbQwbZQ1IYOn-p-5lsXGBTXDzmm'

SPECIAL_PLH_PATH = {
    "marketplace.listing": "Listing",
    "marketplace.order": "Order",
    "marketplace.promotion": "Promotion",
    "marketplace.user": "User",
    "marketplace.chat": "Chat",
    "marketplace.noti": "Notification",
    "marketplace.seller": "Seller",
    "marketplace.app_and_mobile_fe": "Shopee App",
    "marketplace.web_fe_platform": "Shopee App",
    "marketplace.qa": "Shopee App",
    "marketplace.tech": "Marketplace Tech",
    "marketplace.tech_services": "Marketplace Tech",
    "marketplace.mpi": "MPI&D",
    "marketplace.intelligence": "MPI&D",
    "marketplace.data_mart": "MPI&D",
    "marketplace.data_product": "MPI&D",
    "marketplace.traffic_infra": "MPI&D",
    "marketplace.messaging": "MPI&D",
    "paidads": "Ads",
    "search": "Search",
    "recommendation": "Recommendation",
    "shopeevideo.shopeevideo_intelligence": "Recommendation",
    "machine_translation": "Machine Translation",
    "audio_service": "Audio AI",
    "off_platform_ads": "Off-Platform Ads",
    "id_crm": "CRM",
    "id_game": "Games",
    "game": "Games",
    "local_service_and_dp": "Digital Products & Local Services",
    "antifraud": "Marketplace Anti-Fraud",
    "merchant_service": "Merchant Service - Mitra",
    "shopeefood": "ShopeeFood",
    "foody_and_local_service_intelligence": "ShopeeFood Intelligence",
    "shopeevideo": "Shopee Video",
    "shopeevideo.shopeevideo_engineer": "Shopee Video",
    "shopeevideo.multimedia_center": "Shopee Video",
    "supply_chain": "Supply Chain",
    "map": "Map",
    "customer_service_and_chatbot": "Customer Servce & Chatbot",
    "internal_services": "Seatalk & Intenal System",
    "enterprise_efficiency": "Seatalk & Intenal System",
    "info_security": "Security",
}
IGNORE_SET = set([
    "ai_platform",
    "data_infrastructure",
    "engineering_infra",
    "finance",
    "labs",
    "lab",
    "sail",
    "shopeepay",
    "fin_products",
    "kyc",
    "bank",
])


def build_PLH_platform_usage_map():
    '''
    build a data map for
    {
        L1_PLH : {
            path: real_path
            product_line: name,
            l2 : {
                L2_PLH: {
                    origin_path: {

                    },
                    platforms: {
                        PLATFORM: {
                            indicator_name: "name"
                            tuple(9) of (quota, avg, usage) X 3
                    },
                },
            },
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
        if platform in ["tmp", "tmpp"]:
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

            product_line = v_row[0]
            if "." in product_line:
                l1, l2 = product_line.split(".")
            else:
                l1 = product_line
                l2 = product_line
            if l1 in IGNORE_SET:
                continue
            if len(v_row) != 14:
                raise Exception(
                    F"unexpected line in {platform}\nrow {j}\nrow_content {row}\nfilter_info: {rowMeta[j]}\nv_row({len(v_row)}): {v_row}")

            assert l1 in SPECIAL_PLH_PATH or product_line in SPECIAL_PLH_PATH, F"{platform} no exisits {product_line}"
            if product_line in SPECIAL_PLH_PATH:
                product_line_name = SPECIAL_PLH_PATH[product_line]
            else:
                product_line_name = SPECIAL_PLH_PATH[l1]

            indicator = v_row[1]

            if product_line_name not in m:
                m[product_line_name] = {
                    "path": l1,
                    "product_line": product_line_name,
                    "l2": {
                    }
                }
            if l2 not in m[product_line_name]["l2"]:
                m[product_line_name]["l2"][l2] = {
                    "path": v_row[0],
                    "platforms": {},
                }

            m[product_line_name]["l2"][l2]["platforms"][platform] = {
                "data": v_row[2:],
                "indicator_name": indicator,
            }

    return m


def get_header_for_sheet(l2_name, mons):
    color = 'EE4D2D'
    white = 'FFFFFF'
    black = '000000'
    # header_1 = spreadsheet.get_cell_value(l2_name, black, white),
    header_2 = []
    header_3 = [
        spreadsheet.get_cell_value("Platform", white, color),
        spreadsheet.get_cell_value("Indicator", white, color),
    ]

    for mon in mons:
        header_2.extend([
            spreadsheet.get_cell_value("", white, color),
            spreadsheet.get_cell_value("", white, color),
            spreadsheet.get_cell_value(mon, white, color),
        ])
        header_3.extend([
            spreadsheet.get_cell_value(
                "Quota", white, color),
            spreadsheet.get_cell_value(
                "Quota Avg Usage (Monthly)", white, color),
            spreadsheet.get_cell_value(
                "Quota Peak Usage (Monthly)", white, color),
        ])

    return [
        # {"values": header_1},
        {"values": header_2},
        {"values": header_3},
    ]


def write_to_plh_files(l1_info, overwrite=None):
    body = {
        "properties": {
            "title": F"APP Platform Quota Usage & Billing - {l1_info['product_line']}",
        },
        "sheets": [
        ],
    }
    mons = ["May", "April", "March", "Feb"]
    row_cnt = []
    col_cnt = []

    for l2_name, l2_info in l1_info["l2"].items():
        rows_data = get_header_for_sheet(l2_info["path"], mons)
        for platform, row in l2_info["platforms"].items():
            row_data = []
            row_data.append(spreadsheet.get_cell_value(platform))
            row_data.append(spreadsheet.get_cell_value(row["indicator_name"]))
            for col in row["data"]:
                # assert len(col) > 0, F"{l2_name} in {platform} is empty"
                if len(col) > 0 and col[-1] == '%':
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
        col_cnt.append(len(rows_data[-1]["values"]))
        row_cnt.append(len(rows_data))

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

    if overwrite is None:
        ret = spreadsheet.create_spreadsheet_file(body)
        doc_id = ret["spreadsheetId"]
    else:
        spreadsheet.batch_update(overwrite, body)
        

    merge_req = []
    # merge all sheet
    for i, sheet in enumerate(ret["sheets"]):
        sheetId = sheet["properties"]["sheetId"]
        for j, _ in enumerate(mons):
            merge_req.append(spreadsheet.get_merge_cells_cmd(
                sheetId, 0, 1, j*3+2, j*3+5))  # merge name line

        merge_req.append(
            {
                "addConditionalFormatRule": {
                    "rule": spreadsheet.get_ge_rule(sheetId, "100%"),
                    "index": 0,
                },
            }
        )

        merge_req.append({
            'autoResizeDimensions': {
                'dimensions': {
                    'sheetId': sheetId,
                    'dimension': 'COLUMNS',
                },
            }}
        )
        merge_req.append(spreadsheet.get_full_border(
            sheetId, 0, row_cnt[i], 0, col_cnt[i])
        )

    spreadsheet.batch_update(doc_id, merge_req)
    drive.move_doc_to_folder(ret["spreadsheetId"], PLH_FOLDER_ID)
    return ret


if __name__ == "__main__":
    m = build_PLH_platform_usage_map()
    for _, plh_info in m.items():
        write_to_plh_files(plh_info)
#    write_to_plh_files(m["Recommendation"])

    # write_to_plh_files(m["MPI&D"])
