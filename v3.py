#!/usr/bin/env python

from termcolor import colored
from drive import create_file_in_folder, GOOGLE_SHEET_MIME
from collections import OrderedDict
from decimal import Decimal, ROUND_HALF_UP
import spreadsheet
import drive

import pprint

BILL_MONTH = "May"
BILL_SHEET_ID = "1VXbMo0fFjNANF02lxPIqGJxhdeRja7SVs3VtksdKfU8"
OUTPUT_FOLDER = "1ixZ-VtoPV2i6-SQ_q0lDsmx_9svrOvH1"
#OUTPUT_FOLDER = "1eCsCABMYJYw8M68-MKPPKYAD3l8U1C7b"

CPO_OFFICE_OVERALL_SHEET_NAME = "CPO Office Bill"
SERVER_MAP_SHEET_NAME = "ServerMap"
SERVER_PRICE_SHEET_NAME = "Server Pricing"
ADDITIONAL_STORAGE_MAP_SHEET_NAME = "Additional-SM-Storage"
ADDITIONAL_AZ_BAREMETAL_SHEET_NAME = "Additional-AZ-Baremetal"
ADDITIONAL_SEAMONEY_US_SHEET_NAME = "Additional-Seamoney-US"
ADDITIONAL_SEAMONEY_OTHERS_SHEET_NAME = "Additional-Seamoney-Others"
PRODUCT_LINE_MAP_SHEET_NAME = "Productline Mapping"
STANDARD_SERVER_CONFIG = "s1_v2"
NON_LIVE_DC = "DC West"
NON_BANK_FILTER = "Exclude Bank"

HIDDEN_BY_USER_FIELD = 'sheets(data(columnMetadata(hiddenByUser))),sheets(data(rowMetadata(hiddenByUser))),sheets(properties)'
MERGE_PLATFORMS = {
    "MMDB": "DB",
    "Data Transmission Service": "DB",
    "Video Network": "AZ",
}

SPLIT_MAP = {
    "Customer Service": "Chatbot",
    "Seamoney Security (Ziyi)": "Shopee Security (Patrick)"
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


def normalize_weights(weights, max_range=10000):
    weights = [Decimal(str(w)) for w in weights]
    sum_weights = sum(weights)
    normalized_array = [
        round((w / sum_weights) * max_range, 0)
        for w in weights
    ]

    diff = max_range - sum(normalized_array)
    if diff != 0:
        index = max(range(len(normalized_array)),
                    key=lambda i: (weights[i] / sum_weights) % 1)
        normalized_array[index] += diff

    return normalized_array


def get_key_sheets(meta):
    CPO_OVERALL_BILL_SHEET = None
    SERVER_QTY_SHEET = None
    SERVER_PRICE_SHEET = None
    ADDTIONAL_STORAGE_SM_SHEET = None
    ADDITIONAL_SEAMONEY_US_SHEET = None
    ADDITIONAL_SEAMONEY_OTHERS_SHEET = None
    BARE_METAL_SHEET = None
    PL_MAP = None
    PLATFORM_SHEETS = {}

    for sheet_meta in meta["sheets"]:
        properties = sheet_meta["properties"]
        title = properties["title"]
        if title == CPO_OFFICE_OVERALL_SHEET_NAME:
            assert CPO_OVERALL_BILL_SHEET is None, F"there are two sheets name called {CPO_OFFICE_OVERALL_SHEET_NAME}"
            CPO_OVERALL_BILL_SHEET = sheet_meta
        elif title == SERVER_MAP_SHEET_NAME:
            assert SERVER_QTY_SHEET is None, F"there are two sheets name called {SERVER_MAP_SHEET_NAME}"
            SERVER_QTY_SHEET = sheet_meta
        elif title == SERVER_PRICE_SHEET_NAME:
            assert SERVER_PRICE_SHEET is None, F"there are two sheets name called {SERVER_PRICE_SHEET_NAME}"
            SERVER_PRICE_SHEET = sheet_meta
        elif title.startswith(ADDITIONAL_STORAGE_MAP_SHEET_NAME):
            assert ADDTIONAL_STORAGE_SM_SHEET is None, F"there are two sheets name called {ADDITIONAL_STORAGE_MAP_SHEET_NAME}"
            ADDTIONAL_STORAGE_SM_SHEET = sheet_meta
        elif title.startswith(ADDITIONAL_AZ_BAREMETAL_SHEET_NAME):
            assert BARE_METAL_SHEET is None, F"there are two sheets name called {ADDITIONAL_AZ_BAREMETAL_SHEET_NAME}"
            BARE_METAL_SHEET = sheet_meta
        elif title.startswith(ADDITIONAL_SEAMONEY_OTHERS_SHEET_NAME):
            assert ADDITIONAL_SEAMONEY_OTHERS_SHEET is None, F"there are two sheets name called {ADDITIONAL_SEAMONEY_OTHERS_SHEET_NAME}"
            ADDITIONAL_SEAMONEY_OTHERS_SHEET = sheet_meta
        elif title.startswith(ADDITIONAL_SEAMONEY_US_SHEET_NAME):
            assert ADDITIONAL_SEAMONEY_US_SHEET is None, F"there are two sheets name called {ADDITIONAL_SEAMONEY_US_SHEET_NAME}"
            ADDITIONAL_SEAMONEY_US_SHEET = sheet_meta
        elif title == PRODUCT_LINE_MAP_SHEET_NAME:
            assert PL_MAP is None, F"there are two sheets name called {PRODUCT_LINE_MAP_SHEET_NAME}"
            PL_MAP = sheet_meta
        elif title.startswith("Platform-"):
            platform = title[len("Platform-"):]
            PLATFORM_SHEETS[platform] = sheet_meta

    return CPO_OVERALL_BILL_SHEET, SERVER_QTY_SHEET, SERVER_PRICE_SHEET, PL_MAP, \
        ADDTIONAL_STORAGE_SM_SHEET, ADDITIONAL_SEAMONEY_US_SHEET, ADDITIONAL_SEAMONEY_OTHERS_SHEET, BARE_METAL_SHEET, PLATFORM_SHEETS


def get_cpo_office_overall_bill(cpo_office_overall_sheet):
    '''
    expect format is
    Title Line 1 (ignored)
    Title Line 2 (9 columns)
    rows...

    return is {
        bu: {
            "others": {
                ""
            }
            "us": {

            }
        }
    }
    '''
    def _C():
        return {"power_opex": 0, "conn_opex": 0, "mw": 0, "server_capex": 0, "network_capex": 0, "server_count": 0}
    ret = {
        "others": _C(),
        "us": _C(),
    }

    properties = cpo_office_overall_sheet["properties"]
    title = properties["title"]
    rows = spreadsheet.get_one_sheet_content(BILL_SHEET_ID, title)
    headers = rows[1]
    assert len(headers) == 9, \
        f"""expect cpo office overall bill column is 9, format should be \
(Team,BU,Region,IDC Cost,Connectivity Cost,MW,Server Cost,Network Cost,Server Count)
but got {headers} in excel"""
    for row in rows[2:]:  # ignore the two line
        team, bu, region, power_opex, conn_opex, mw, server_capex, network_capex, server_count = row

        if bu != NON_BANK_FILTER:
            continue
        assert team == "Application", "team should be 'Application'"
        if region.upper() == "SG":
            region = "others"
        elif region.upper() == "US":
            region = "us"
        elif region == "Others":
            region = "others"
        else:
            print(colored(
                F"there is a empty region in cpo office bill {row}, will ignore", "red"))
            continue

        # clean up the data
        ret[region]["power_opex"] += Decimal(power_opex)
        ret[region]["conn_opex"] += Decimal(conn_opex)
        ret[region]["mw"] += Decimal(mw)
        ret[region]["server_capex"] += Decimal(server_capex)
        ret[region]["network_capex"] += Decimal(network_capex)
        ret[region]["server_count"] += Decimal(server_count)

    return ret


def get_platform_servers(server_qty_sheet, storage_addtional_sheet):
    '''
    expect server quantity sheet format is
    title
    rows <date, bu, platform, region, idc, server_config, qty, categoty, location, bu>

    returns {
    "us/others": {
        "platform_name": {server_config: qty..}
        "nonlive":
        "seamoney(others)": //ignored
        "seamoney(US)" : //ignored
        }
    }
    '''

    ret = {
        "us": {},
        "others": {
            "Storage-USS": {},
            "Storage-Ceph": {},
        }
    }

    storage_rows = spreadsheet.get_one_sheet_content(
        BILL_SHEET_ID, storage_addtional_sheet["properties"]["title"])
    for row in storage_rows:
        p, config, qty = row
        if p == "USS":
            ret["others"]["Storage-USS"][config.lower()] = int(qty)
        elif p == "Ceph":
            ret["others"]["Storage-Ceph"][config.lower()] = int(qty)
        else:
            assert F"unknow storage platform {p}"

    properties = server_qty_sheet.get("properties")
    title = properties.get("title")

    rows = spreadsheet.get_one_sheet_content(BILL_SHEET_ID, title)

    for row in rows[1:]:
        try:
            date, bu, platform, region, idc, server_config, qty, category, location, BU = row[
                :10]
        except ValueError:
            print(F"illegal row in server_quantity sheet {str(row)}")
            exit(-1)

        if category != "APP":  # ignore DI/AI
            continue
        if BU != NON_BANK_FILTER:  # ignore bank
            continue

        if location.lower() == "us":
            loc = "us"
        else:
            loc = "others"
        if idc == NON_LIVE_DC:
            platform = "nonlive"
        elif bu == "shopee" or bu == "seamoney":
            if bu == "seamoney":
                platform = "seamoney"
        else:
            continue  # ignore others, like seamoney etc.

        server_config = server_config.lower()

        if platform in MERGE_PLATFORMS:
            platform = MERGE_PLATFORMS[platform]

        if platform not in ret[loc]:
            ret[loc][platform] = {}
        if server_config not in ret[loc][platform]:
            ret[loc][platform][server_config] = 0
        ret[loc][platform][server_config] += int(qty)

    # validate CPO office's storage platform and split map
    for c, q in ret["others"]["Storage"].items():
        assert ret["others"]["Storage-USS"].get(c, 0) + ret["others"]["Storage-Ceph"].get(c, 0) == q, \
            F"additional uss, ceph serverconfig {c} mismatch count with CPO Office's bill q: "\
            F'additional value is uss: {ret["others"]["Storage-USS"].get(c, 0)}, ceph {ret["others"]["Storage-Ceph"].get(c, 0)}'

    del ret["others"]["Storage"]  # splited straoge
    return ret


def get_price_unit(pricing_sheet):
    '''
    first row ignored
    rows: <config_name, price, power>
    '''
    ret = {}

    properties = pricing_sheet["properties"]
    title = properties["title"]
    rows = spreadsheet.get_one_sheet_content(BILL_SHEET_ID, title)
    for row in rows[1:]:
        config_name, price, power = row
        ret[config_name.lower()] = {
            "price": Decimal(price) * 100,
            "power": Decimal(power),
        }

    return ret


def calculate_platform_cost(cpo_bill, server_qty, server_unit_price, bare_metal_sheet, seamoney_sheet_us, seamoney_sheet_others):
    '''
    output format is
    platform, server_count, total_capex, total_server_power, projected_server_capex, projected_network_device_capex, 
    projected_server_opex, projected_connectivity_opex, allocated_capex, allocated_opex

    '''
    ret = {
        "us": {},
        "others": {}
    }
    '''
    baremetal return is {
        product_line: {
            server_count:
            total_capex:
            total_power:
        }
    }
    '''
    def _summary_prices(server_config_map):
        capex = Decimal()
        server_power = Decimal()
        server_count = 0

        for server_config, qty in server_config_map.items():
            if server_config not in server_unit_price:
                # print(colored(F"server config name {server_config} not in map, so I chagne to {STANDARD_SERVER_CONFIG}", "green"))
                server_config = STANDARD_SERVER_CONFIG

            unit = server_unit_price[server_config]
            unit_price = unit["price"]
            unit_power = unit["power"]

            capex += unit_price * qty
            server_power += unit_power * qty
            server_count += qty
        return capex, server_power, server_count

    def _C():
        return {
            "total_server_capex": Decimal(),
            "total_server_count": 0,
            "total_power": Decimal(),
        }
    bare_metal_ret = {
        "us": {},
        "others": {}
    }

    # get bare metal productline map
    bare_metal_map = {"us": {}, "others": {}}
    bare_metal_rows = spreadsheet.get_one_sheet_content(
        BILL_SHEET_ID, bare_metal_sheet["properties"]["title"])
    for bare_metal_row in bare_metal_rows[1:]:
        try:
            product_line, location, server_config, qty = bare_metal_row
            server_config = server_config.lower()
        except ValueError:
            print(
                colored(F"abormal row for baremetal map: {bare_metal_row}", "red"))
            exit(-1)
        if location.lower() == "us":
            loc = "US"
        else:
            loc = "others"
        if product_line not in bare_metal_map[loc]:
            bare_metal_map[loc][product_line] = {}
        if server_config not in bare_metal_map[loc][product_line]:
            bare_metal_map[loc][product_line][server_config] = 0

        bare_metal_map[loc][product_line][server_config] += int(qty)

    for loc, platforms_qty in server_qty.items():
        for platform, server_config_map in platforms_qty.items():
            if platform not in ret[loc]:
                ret[loc][platform] = {}

            capex, server_power, server_count = _summary_prices(
                server_config_map)

            ret[loc][platform] = {
                "server_capex": capex if platform != "nonlive" else 0,
                "server_power": server_power if platform != "nonlive" else 0,
                "server_count": server_count,
                # final projected capex should be at most two decimal degits
                "projected_server_capex": 0,
                "projected_network_capex": 0,
                "projected_power_opex": 0,
                "projected_conn_opex": 0,
                "projected_capex": 0,
                "projected_opex": 0,
            }

    for loc, bmp in bare_metal_map.items():
        assert "AZ-Baremetal" not in ret[loc]
        ret[loc]["AZ-Baremetal"] = {
            "server_capex": Decimal(),
            "server_power": Decimal(),
            "server_count": 0,
            "projected_server_capex": 0,
            "projected_network_capex": 0,
            "projected_power_opex": 0,
            "projected_conn_opex": 0,
            "projected_capex": 0,
            "projected_opex": 0,
        }
        for pl, bm_config_map in bmp.items():
            pl_bm_capex, pl_bm_power, pl_bm_count = _summary_prices(
                bm_config_map)
            ret[loc]["AZ-Baremetal"]["server_capex"] += pl_bm_capex
            ret[loc]["AZ-Baremetal"]["server_power"] += pl_bm_power
            ret[loc]["AZ-Baremetal"]["server_count"] += pl_bm_count

            # reduce this number for az platform
            ret[loc]["AZ"]["server_capex"] -= pl_bm_capex
            ret[loc]["AZ"]["server_power"] -= pl_bm_power
            ret[loc]["AZ"]["server_count"] -= pl_bm_count

            bare_metal_ret[loc][pl] = {
                "server_capex": pl_bm_capex,
                "server_power": pl_bm_power,
                "server_count": pl_bm_count,
            }

    for loc, loc_details in ret.items():
        platforms = []
        platform_capex_weights = []
        platform_power_weights = []
        platform_sc_weights = []
        for platform, platform_cost in loc_details.items():
            if platform == "nonlive":
                continue
            platforms.append(platform)
            platform_capex_weights.append(platform_cost["server_capex"])
            platform_power_weights.append(platform_cost["server_power"])
            platform_sc_weights.append(platform_cost["server_count"])
        platform_capex_weights = normalize_weights(platform_capex_weights)
        platform_power_weights = normalize_weights(platform_power_weights)
        platform_sc_weights = normalize_weights(platform_sc_weights)

        assert sum(platform_capex_weights) == 10000
        assert sum(platform_power_weights) == 10000
        assert sum(platform_sc_weights) == 10000
        platform_capex_frac = dict(zip(platforms, platform_capex_weights))
        platform_power_frac = dict(zip(platforms, platform_power_weights))
        platform_sc_frac = dict(zip(platforms, platform_sc_weights))

        for platform, platform_cost in loc_details.items():
            if platform == "nonlive":
                continue
            details = cpo_bill[loc]

            ret[loc][platform]["projected_server_capex"] = details["server_capex"] * \
                platform_capex_frac[platform] / 10000
            ret[loc][platform]["projected_network_capex"] = details["network_capex"] * \
                platform_sc_frac[platform] / 10000
            ret[loc][platform]["projected_power_opex"] = details["power_opex"] * \
                platform_power_frac[platform] / 10000
            ret[loc][platform]["projected_conn_opex"] = details["conn_opex"] * \
                platform_sc_frac[platform] / 10000
            ret[loc][platform]["projected_capex"] = ret[loc][platform]["projected_server_capex"] + \
                ret[loc][platform]["projected_network_capex"]
            ret[loc][platform]["projected_opex"] = ret[loc][platform]["projected_power_opex"] + \
                ret[loc][platform]["projected_conn_opex"]

    pls = []
    pl_capex_weights = []
    pl_power_weights = []
    for pl, bmd in bare_metal_ret[loc].items():
        pls.append(pl)
        pl_capex_weights.append(bmd["server_capex"])
        pl_power_weights.append(bmd["server_power"])
    pl_capex_weights = normalize_weights(pl_capex_weights)
    pl_power_weights = normalize_weights(pl_power_weights)
    pl_capex_frac = dict(zip(pls, pl_capex_weights))
    pl_power_frac = dict(zip(pls, pl_power_weights))
    # calculate baremetal infos
    for pl, bmd in bare_metal_ret[loc].items():
        bmd["projected_capex"] = ret[loc]["AZ-Baremetal"]["projected_capex"] * \
            pl_capex_frac[pl] / 10000
        bmd["projected_opex"] = ret[loc]["AZ-Baremetal"]["projected_opex"] * \
            pl_power_frac[pl] / 10000

    for loc, lc in ret.items():
        capex_sum = Decimal()
        opex_sum = Decimal()
        for platform, pm in lc.items():
            capex_sum += pm["projected_capex"]
            opex_sum += pm["projected_opex"]
        # check point 1, sum of platform costs should nearly cpo office's bill
        capex_error = capex_sum - \
            cpo_bill[loc]["server_capex"] - cpo_bill[loc]["network_capex"]
        opex_error = opex_sum - \
            cpo_bill[loc]["power_opex"] - cpo_bill[loc]["conn_opex"]
        assert capex_error < 10, "{loc} capex error shoud under 10, now is {capex_error}"
        assert opex_error < 10, "{loc} opex error shoud under 10, now is {opex_error}"

    seamoney_ret = {"us": {}, "others": {}}
    # calculate seamoney

    def _get_seamoney_map(seamoney_sheet):
        ret = {}
        seamoney_rows = spreadsheet.get_one_sheet_content(
            BILL_SHEET_ID, seamoney_sheet["properties"]["title"])
        smpls = seamoney_rows[0][1:]
        for smpl in smpls:
            ret[smpl] = {}

        for seamoney_row in seamoney_rows[1:]:
            assert len(seamoney_row) > len(
                smpls), F"invalid server config line {seamoney_rows} in seamoney sheet"
            server_config = seamoney_row[0].lower()
            for i in range(len(smpls)):
                if server_config not in ret[smpls[i]]:
                    ret[smpls[i]][server_config] = 0
                ret[smpls[i]][server_config] += int(seamoney_row[i+1])
        return ret

    seamoney_map = {
        "us": _get_seamoney_map(seamoney_sheet_us),
        "others": _get_seamoney_map(seamoney_sheet_others)
    }

    for loc, smp in seamoney_map.items():
        smpls = []
        smpl_capex_weights = []
        smpl_power_weights = []
        seamoney_ret[loc] = {}
        for smpl, pl_config_map in smp.items():
            pl_sm_capex, pl_sm_power, pl_sm_count = _summary_prices(
                pl_config_map)
            pprint.pprint(pl_config_map)
            seamoney_ret[loc][smpl] = {
                "server_capex": pl_sm_capex,
                "server_power": pl_sm_power,
                "server_count": pl_sm_count,
            }
            smpls.append(smpl)
            smpl_capex_weights.append(pl_sm_capex)
            smpl_power_weights.append(pl_sm_power)
        smpl_capex_weights = normalize_weights(smpl_capex_weights)
        smpl_power_weights = normalize_weights(smpl_power_weights)
        smpl_capex_frac = dict(zip(smpls, smpl_capex_weights))
        smpl_power_frac = dict(zip(smpls, smpl_power_weights))

        for smpl, smr in seamoney_ret[loc].items():
            smr["projected_capex"] = ret[loc]["seamoney"]["projected_capex"] * \
                smpl_capex_frac[smpl] / 10000
            smr["projected_opex"] = ret[loc]["seamoney"]["projected_opex"] * \
                smpl_power_frac[smpl] / 10000

    # del ret[loc]["AZ-Baremetal"]

    return ret, bare_metal_ret, seamoney_ret


def get_pl_usage(platform_sheets):
    '''
    each sheets has 6 cols
    header line
    rows <business, platform, indicator, budget, quota, usage>
    return {
        platform: {
            product_line: {
                budget,
                quota,
                usage,
                percentage // max(quota, usage)/sum(max)
            }
        }
    }

    az, storage is quite special
    az-baremetal need use special format
    '''
    ret = {}
    indicators = {}
    for platform_name, platform_sheet in platform_sheets.items():
        properties = platform_sheet["properties"]
        title = properties["title"]

        rows = spreadsheet.get_one_sheet_content(BILL_SHEET_ID, title)
        assert platform_name not in ret, F"duplicated {platform_name}?"
        ret[platform_name] = {}
        indicators[platform_name] = rows[1][1]  # set the indicators

        for row in rows[1:]:
            try:
                product_line, _, _, budget, quota, usage = row
            except ValueError:
                print(
                    colored(F"illegal row found in platform {platform_name} usage, line is {row}"))
                exit(-1)
            try:
                ret[platform_name][product_line] = {
                    "budget": float(budget),
                    "quota": float(quota),
                    "usage": float(usage),
                    "maxqu": max(float(quota), float(usage)),
                    "percentage": 0.0,
                }
            except ValueError as e:
                print(colored(
                    F"please check the format error in platform {platform_name} usage: {e}", "red"))
                exit(-1)

        qu_weights = []
        pls = []
        for pl, _ in ret[platform_name].items():
            qu_weights.append(ret[platform_name][pl]["maxqu"])
            pls.append(pl)
        qu_weights = normalize_weights(qu_weights, 1000000)
        for i in range(0, len(pls)):
            ret[platform_name][pls[i]]["percentage"] = qu_weights[i]

    return ret


def get_pl_bill(product_line_map, platform_cost, pl_usage, bare_metal_info, seamoney_info):
    ret = {
        "others": {},
        "us": {},
    }
    loc = "others"  # don't consider us yet
    cksum_capex = {"us": {}, "others": {}}
    cksum_opex = {"us": {}, "others": {}}
    for pl, _ in product_line_map.items():
        ret[loc][pl] = {}

    for platform, product_lines_usages in pl_usage.items():
        for pl, _ in product_lines_usages.items():
            assert pl in ret[loc], F"dummy productline {pl} in platform {platform}"
        for pl, _ in bare_metal_info[loc].items():
            assert pl in ret[loc], F"dummy productline {pl} in baremetal"
        for pl, _ in seamoney_info[loc].items():
            assert pl in ret[loc], F"dummy productline {pl} in seamoney"

    for platform, product_lines_usages in pl_usage.items():
        for pl, pl_usage in product_lines_usages.items():
            assert platform not in ret[loc][pl]
            ret[loc][pl][platform] = pl_usage.copy()
            ret[loc][pl][platform]["capex"] = platform_cost[loc][platform]["projected_capex"] * \
                pl_usage["percentage"] / 1000000
            ret[loc][pl][platform]["opex"] = platform_cost[loc][platform]["projected_opex"] * \
                pl_usage["percentage"] / 1000000

            if platform not in cksum_capex[loc]:
                cksum_capex[loc][platform] = Decimal()
            if platform not in cksum_opex[loc]:
                cksum_opex[loc][platform] = Decimal()
            cksum_capex[loc][platform] += ret[loc][pl][platform]["capex"]
            cksum_opex[loc][platform] += ret[loc][pl][platform]["opex"]

    for pl in bare_metal_info[loc]:
        ret[loc][pl]["AZ-Baremetal"] = {
            "capex": bare_metal_info[loc][pl]["projected_capex"],
            "opex": bare_metal_info[loc][pl]["projected_opex"]
        }

    for pl in seamoney_info[loc]:
        ret[loc][pl]["seamoney"] = {
            "capex": seamoney_info[loc][pl]["projected_capex"],
            "opex": seamoney_info[loc][pl]["projected_opex"],
        }

    # checkpoint 2, check the sum of productline cost similar to platform projected capex
        for platform, pv in platform_cost[loc].items():
            if platform not in cksum_capex[loc]:
                # print(colored(F"unknow checksum for {platform}, continue", "red"))
                continue
            capex_error = pv["projected_capex"] - cksum_capex[loc][platform]
            opex_error = pv["projected_opex"] - cksum_opex[loc][platform]
            print(colored(
                F"checking for platform {platform} cost.. capex_error: {capex_error}, opex_error: {opex_error}, Done ", "green"))
            assert capex_error < 10, "platform {platform} capex error {capex_error} > 10"
            assert opex_error < 10, "platform {platform} capex error {opex_error} > 10"

    return ret


def get_pl_map(product_line_sheet):
    properties = product_line_sheet["properties"]
    title = properties["title"]

    ret = {}
    rows = spreadsheet.get_one_sheet_content(BILL_SHEET_ID, title)
    for row in rows[1:]:
        row = row + [''] * max(0, 6 - len(row))  # padding
        division, l0, l1, mapping, cpo_office_link, quota_link = row
        if len(mapping) == 0 or mapping == '-':
            continue
        for m in mapping.split("\n"):
            ret[m.strip()] = (l0, l1, cpo_office_link, quota_link)

    return ret


'''
generation parts
> overviews
> final bills
> cpo office update
'''


def generate_overviews(pl_map, platform_cost, pl_bills):
    body = {
        "properties": {
            "title": F"App Platform Overviews - {BILL_MONTH}"
        },
        "sheets": [

        ]
    }
    # get a sequencial platforms
    platform_set = set()
    platforms = []
    for pl, platform_bills in pl_bills["others"].items():
        for platform in platform_bills.keys():
            if platform not in platform_set:
                platforms.append(platform)
                platform_set.add(platform)
    platforms.sort()

    # platform costs sheets
    for loc, pcm in platform_cost.items():
        row_data = [
            spreadsheet.get_cell_value("Platforms"),
            spreadsheet.get_cell_value("#Server"),
            spreadsheet.get_cell_value("Total Power"),
            spreadsheet.get_cell_value("Total Capex"),
            spreadsheet.get_cell_value("Projected Capex"),
            spreadsheet.get_cell_value("Projected Opex"),
            spreadsheet.get_cell_value("Projected Server Capex"),
            spreadsheet.get_cell_value("Projected Network Capex"),
            spreadsheet.get_cell_value("Projected Power Opex"),
            spreadsheet.get_cell_value("Projected Conn Opex"),
        ]
        rows_data = [{"values": row_data}]

        for platform in platforms + ["nonlive"]:
            if platform not in pcm:
                continue
            row_data = [
                spreadsheet.get_cell_value(platform),
                spreadsheet.get_cell_value(
                    pcm[platform]["server_count"], try_use_number=True),
                spreadsheet.get_cell_value(
                    pcm[platform]["server_power"], try_use_number=True),
                spreadsheet.get_cell_value(
                    Decimal(pcm[platform]["server_capex"]) / 100, try_use_number=True),
                spreadsheet.get_cell_value(
                    pcm[platform]["projected_capex"], try_use_number=True),
                spreadsheet.get_cell_value(
                    pcm[platform]["projected_opex"], try_use_number=True),
                spreadsheet.get_cell_value(
                    pcm[platform]["projected_server_capex"], try_use_number=True),
                spreadsheet.get_cell_value(
                    pcm[platform]["projected_network_capex"], try_use_number=True),
                spreadsheet.get_cell_value(
                    pcm[platform]["projected_power_opex"], try_use_number=True),
                spreadsheet.get_cell_value(
                    pcm[platform]["projected_conn_opex"], try_use_number=True),
            ]
            rows_data.append({"values": row_data})

        body["sheets"].append(
            {
                "properties": {
                    "title": F"{BILL_MONTH} Platform Costs Overview - {loc}",
                },
                "data": {
                    "rowData": rows_data
                },
            },
        )

    rows_data = []
    header1 = [
        spreadsheet.get_cell_value("L0"),
        spreadsheet.get_cell_value("L1"),
        spreadsheet.get_cell_value("CMDB mapping"),
        spreadsheet.get_cell_value("BizSum"),
        spreadsheet.get_cell_value(""),
    ]
    header2 = [
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value("Capex"),
        spreadsheet.get_cell_value("Opex"),
    ]
    for platform in platforms:
        header1.append(spreadsheet.get_cell_value(platform))
        header1.append(spreadsheet.get_cell_value(''))
        header2.append(spreadsheet.get_cell_value('Capex'))
        header2.append(spreadsheet.get_cell_value('Opex'))
    rows_data.append({"values": header1})
    rows_data.append({"values": header2})

    for pl, pv in pl_map.items():
        l0, l1, cpo_office_link, quota_link = pv
        if pl not in pl_bills["others"]:
            continue
        platform_bills = pl_bills["others"][pl]
        row_data = [
            spreadsheet.get_cell_value(l0),
            spreadsheet.get_cell_value(l1),
            spreadsheet.get_cell_value(pl),
        ]
        capex_sum = Decimal()
        opex_sum = Decimal()
        for platform in platforms:
            if platform in platform_bills:
                capex_sum += platform_bills[platform]["capex"]
                opex_sum += platform_bills[platform]["opex"]
                row_data.append(spreadsheet.get_cell_value(
                    float(platform_bills[platform]["capex"]), try_use_number=True))
                row_data.append(spreadsheet.get_cell_value(
                    float(platform_bills[platform]["opex"]), try_use_number=True))
            else:
                row_data.append(spreadsheet.get_cell_value(''))
                row_data.append(spreadsheet.get_cell_value(''))
        row_data[3:3] = [
            spreadsheet.get_cell_value(float(capex_sum), try_use_number=True),
            spreadsheet.get_cell_value(float(opex_sum), try_use_number=True)
        ]
        rows_data.append({"values": row_data})

    body["sheets"].append(
        {
            "properties": {
                "title": F"{BILL_MONTH} Product Line Bills Overview",
            },
            "data": {
                "rowData": rows_data
            },
        },
    )

    overview_ss = spreadsheet.create_spreadsheet_file(body)
    doc_id = overview_ss["spreadsheetId"]
    drive.move_doc_to_folder(doc_id, OUTPUT_FOLDER)

    merge_req = []

    for sheet in overview_ss["sheets"]:
        merge_req.append(spreadsheet.get_autosize(
            sheet["properties"]["sheetId"]))

    overview_bill_sheet_id = overview_ss["sheets"][2]["properties"]["sheetId"]
    # merge every two columns
    merge_req.append(spreadsheet.get_merge_cells_cmd(
        overview_bill_sheet_id, 0, 1, 3, 5),  # sum col
    )
    for i, platform in enumerate(platforms):
        merge_req.append(spreadsheet.get_merge_cells_cmd(
            overview_bill_sheet_id, 0, 1, 2 * i + 5, 2 * i + 7))

    spreadsheet.batch_update(doc_id, merge_req)

    return doc_id


if __name__ == "__main__":
    meta = spreadsheet.get_spreadsheet_meta(
        BILL_SHEET_ID, fields=HIDDEN_BY_USER_FIELD)
    cpo_office_overall_sheet, server_qty_sheet, pricing_sheet, \
        product_line_sheet, additional_storage_sheet, seamoney_sheet_us, seamoney_sheet_others, bare_metal_sheet, platform_sheets = get_key_sheets(
            meta)
    product_line_map = get_pl_map(product_line_sheet)

    cpo_overall = get_cpo_office_overall_bill(cpo_office_overall_sheet)
    server_qty = get_platform_servers(
        server_qty_sheet, additional_storage_sheet)
    server_unit_price = get_price_unit(pricing_sheet)

    platform_cost, bare_metal_cost, seamoney_cost = calculate_platform_cost(
        cpo_overall, server_qty, server_unit_price, bare_metal_sheet, seamoney_sheet_us, seamoney_sheet_others)

    pl_usage = get_pl_usage(platform_sheets)
    pl_bill = get_pl_bill(product_line_map, platform_cost, pl_usage,
                          bare_metal_cost, seamoney_cost)

    # generate the bill
    product_line_map = get_pl_map(product_line_sheet)
    generate_overviews(product_line_map, platform_cost, pl_bill)
