
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
            if l1 in IGNORE_SET:
                continue

            assert l1 in SPECIAL_PLH_PATH or product_line in SPECIAL_PLH_PATH, F"no exisits {product_line}"
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


def get_header_for_sheet(l2_name):
    color = 'EE4D2D'
    white = 'FFFFFF'
    black = '000000'
    return [
        {
            "values": [
                spreadsheet.get_cell_value(l2_name, black, white),
            ]
        },
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


def write_to_plh_files(l1_info):
    body = {
        "properties": {
            "title": F"APP Platform Quota Usage & Billing - {l1_info['product_line']}",
        },
        "sheets": [
        ],
    }
    head_rows = []

    whole_data = []
    for _, l2_info in l1_info["l2"].items():
        rows_data = get_header_for_sheet(l2_info["path"])
        # which line need merge header for months
        head_rows.append(len(whole_data))
        for platform, row in l2_info["platforms"].items():
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

        rows_data.append({})  # empty row
        whole_data.extend(rows_data)

    body["sheets"].append(
        {
            "properties": {
                "title": F"{l1_info['product_line']}'s Bills",
            },
            "data": {
                "rowData": whole_data
            },
        },
    )

    ret = spreadsheet.create_spreadsheet_file(body)
    doc_id = ret["spreadsheetId"]

    merge_req = []
    # merge all sheet
    for sheet in ret["sheets"]:
        sheetId = sheet["properties"]["sheetId"]
        for head_row in head_rows:
            merge_req.append(spreadsheet.get_merge_cells_cmd(
                sheetId, head_row, head_row+1, 0, 3))  # merge name line
            merge_req.append(spreadsheet.get_merge_cells_cmd(
                sheetId, head_row+1, head_row+2, 2, 5))
            merge_req.append(spreadsheet.get_merge_cells_cmd(
                sheetId, head_row+1, head_row+2, 5, 8))
            merge_req.append(spreadsheet.get_merge_cells_cmd(
                sheetId, head_row+1, head_row+2, 8, 11))
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
    for _, plh_info in m.items():
        write_to_plh_files(plh_info)
    #write_to_plh_files(m["Recommendation"])

    #write_to_plh_files(m["MPI&D"])
