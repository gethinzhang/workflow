
from drive import create_file_in_folder, GOOGLE_SHEET_MIME
import spreadsheet
import drive
import itertools
import pprint
import re

HIDDEN_BY_USER_FIELD = 'sheets(data(columnMetadata(hiddenByUser))),sheets(data(rowMetadata(hiddenByUser))),sheets(properties)'
BASE_SHEET_ID = '16dQ0NOB7GMS1H19LVeKr4RpE507oEcJ7RoneiR5_LsU'
# BASE_SHEET_ID = '1ZWjzJSW1q0qm6ZABAhHpcOPOUE97cW8m901VL58A5Pk'
PLH_FOLDER_ID = '1PXuqqgbQwbZQ1IYOn-p-5lsXGBTXDzmm'
# MAP_SHEET_ID = '1V3SVl7pF2BD6t4oh-Eu9EaQRCYdrSXQuGmhPeCraUWA'
MAP_SHEET_ID = '1EnpZ8U-MAC-jGsE_KG-RFhuRFJqeGuCE0YCNPoXTsq0'

SPECIAL_PLH_PATH = {
    "paidads": "Ads",
    "search": "Search",
    "recommendation": "Recommendation",
    "shopeevideo.shopeevideo_intelligence": "Engineering & Architecture",
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
    "marketplace.mpi": "Marketplace Intelligence & Data",
    "marketplace.intelligence": "Marketplace Intelligence & Data",
    "marketplace.data_mart": "Marketplace Intelligence & Data",
    "marketplace.data_product": "Marketplace Intelligence & Data",
    "marketplace.traffic_infra": "Marketplace Intelligence & Data",
    "marketplace.messaging": "Marketplace Intelligence & Data",
    "machine_translation": "Machine Translation",
    "audio_service": "Audio AI",
    "off_platform_ads": "Off-platform Ads",
    "id_crm": "CRM",
    "id_game": "Games",
    "game": "Games",
    "local_service_and_dp": "Digital Products & Local Services",
    "antifraud": "Marketplace Anti-Fraud",
    "merchant_service": "Merchant Service - Mitra",
    "shopeefood": "ShopeeFood",
    "foody_and_local_service_intelligence": "ShopeeFood Intelligence",
    "shopeevideo": "ShopeeVideo",
    "shopeevideo.shopeevideo_engineer": "ShopeeVideo",
    "shopeevideo.multimedia_center": "ShopeeVideo",
    "supply_chain.fulfilment": "SBS",
    "supply_chain.sbs": "SBS",
    "supply_chain.wms": "SBS",
    "supply_chain.retail": "SBS",
    "supply_chain.spx": "SPX",
    "supply_chain.sls": "SLS",
    "map": "Map",
    "customer_service_and_chatbot": "Customer Service",
    "internal_services": "SeaTalk / Internal Systems",
    "enterprise_efficiency": "SeaTalk / Internal Systems",
    "info_security": "Seamoney Security (Ziyi)",
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

REV_LINK_MAP = {v: k for k, v in SPECIAL_PLH_PATH.items()}


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
            if product_line == '':
                continue

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
        ss_content = spreadsheet.create_spreadsheet_file(body)
    else:
        exitsting_ss = spreadsheet.get_spreadsheet_meta(overwrite)
        for es, sheet in itertools.zip_longest(exitsting_ss["sheets"], body["sheets"]):
            if es is None:
                es = spreadsheet.batch_update(overwrite, [
                    {
                        "addSheet": {
                            "properties": {
                                "title": sheet["properties"]["title"],
                            }
                        }
                    }
                ])
            else:  # clear the sheet
                spreadsheet.clear_sheet(overwrite, es["properties"]["sheetId"])

            spreadsheet.batch_update(overwrite, [
                {
                    'updateSheetProperties': {
                        "properties": {
                            "title": sheet["properties"]["title"],
                            "sheetId": es["properties"]["sheetId"],
                        },
                        'fields': 'title',
                    },
                },
                {
                    'updateCells': {
                        "range": {
                            "sheetId": es["properties"]["sheetId"],
                            "startRowIndex": 0,
                            "startColumnIndex": 0,
                        },
                        "fields": "*",
                        "rows": sheet["data"]["rowData"]},
                }]
            )
        ss_content = spreadsheet.get_spreadsheet_meta(overwrite)

    merge_req = []
    sheets = ss_content["sheets"]
    doc_id = ss_content["spreadsheetId"]

    # merge all sheet
    for i, sheet in enumerate(sheets):
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
        merge_req.append({
            'repeatCell': {
                'range': {
                    'sheetId': sheetId,
                    'startRowIndex': 0,
                    'startColumnIndex': 0,
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': 'RIGHT',
                    },
                },
                'fields': 'userEnteredFormat(horizontalAlignment)'
            }
        })
        merge_req.append(spreadsheet.get_full_border(
            sheetId, 0, row_cnt[i], 0, col_cnt[i])
        )

    spreadsheet.batch_update(doc_id, merge_req)
    if not overwrite:
        drive.move_doc_to_folder(doc_id, PLH_FOLDER_ID)
    return ss_content


def build_link_map():
    ret = {}
    rows = spreadsheet.get_one_sheet_content(MAP_SHEET_ID, "Sheet1")
    for i, row in enumerate(rows[2:]):
        division, l0, l1, mapping, link_1, new_link, cpo_link = row[:]
        ret[l1] = (i, link_1, new_link, cpo_link)
    return ret


def update_link_in_map_file(links, plh, lnk):
    assert plh in links
    cell = "E" + str(links[plh][0]+3)
    spreadsheet.update_cell_value(
        MAP_SHEET_ID,
        "Sheet1",
        cell,
        {"values": [[lnk]]},
    )


def extract_doc_id_from_url(url):
    regex_pattern = r"/d/([a-zA-Z0-9-_]+)"
    match = re.search(regex_pattern, url)
    if match:
        doc_id = match.group(1)
        return doc_id
    else:
        return None


if __name__ == "__main__":
    m = build_PLH_platform_usage_map()
    links = build_link_map()

    # for _, plh_info in m.items():
    #    ret = write_to_plh_files(plh_info)
    #    update_link_in_map_file(links, plh_info["product_line"], ret["spreadsheetUrl"])
    for _, plh_info in m.items():
        ret = write_to_plh_files(plh_info, extract_doc_id_from_url(
            links[plh_info["product_line"]][2]))
        update_link_in_map_file(
            links, plh_info["product_line"], ret["spreadsheetUrl"])

    # ret = write_to_plh_files(m["Recommendation"])
    # update_link_in_map_file(links, "Recommendation", ret["spreadsheetUrl"])
    # '1k6AwhOI4CAEJjg_rTIoh4OKZOivKEPn7jeskmW3mCew')

    # write_to_plh_files(m["MPI&D"])
