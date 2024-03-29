#!/usr/bin/env python

from decimal import Decimal
import spreadsheet
import drive
import sys
import logging
import pprint

BILL_YEAR = 2023
# INPUT A: Change month to bill month, must in MONTHS List
# BILL_MONTH = "June"
BILL_MONTH = "May"
# INPUT B: Input sheets ID
# BILL_SHEET_ID = "1UBk_y84Ekje_Dlqs3S3aXmzUgfEtt18tr5-XWr3Xntc"
# MAY BILL
BILL_SHEET_ID = "1VXbMo0fFjNANF02lxPIqGJxhdeRja7SVs3VtksdKfU8"

# INPUT C: output folder ID
OUTPUT_FOLDER = "1ixZ-VtoPV2i6-SQ_q0lDsmx_9svrOvH1"
# OUTPUT_FOLDER = "1eCsCABMYJYw8M68-MKPPKYAD3l8U1C7b"
# INPUT D: (Only used for R2)
OVERVIEW_SHEET_ID = '17eJHldqKpH8njqSmnfnvrZBOWQTJcrOA-FIHLb2p7H8'

MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December']
MONTH_INDEX = MONTHS.index(BILL_MONTH)
assert MONTH_INDEX != -1, "invalid BILL MONTH!"
LOG_LEVEL = logging.INFO

# change these vars if needs
CPO_OFFICE_OVERALL_SHEET_NAME = "CPO Office Bill"
SERVER_MAP_SHEET_NAME = "ServerMap"
SERVER_PRICE_SHEET_NAME = "Server Pricing"
ADDITIONAL_STORAGE_MAP_SHEET_NAME = "Additional-SM-Storage"
ADDITIONAL_AZ_BAREMETAL_SHEET_NAME = "Additional-AZ-Baremetal"
ADDITIONAL_SEAMONEY_US_SHEET_NAME = "Additional-Seamoney-US"
ADDITIONAL_SEAMONEY_OTHERS_SHEET_NAME = "Additional-Seamoney-Others"
PRODUCT_LINE_MAP_SHEET_NAME = "Productline Mapping"
R1_SHARE_SHEET_NAME = "R1-Share"
STANDARD_SERVER_CONFIG = "s1_v2"
NON_LIVE_DC = "DC West"
NON_BANK_FILTER = "Exclude Bank"
CATEGORY_FILTER = "APP"
EI_L0_NAME = "Engineering Infrastructure"

log = logging.getLogger('bunnyapple')
log.setLevel(level=LOG_LEVEL)


HIDDEN_BY_USER_FIELD = 'sheets(data(columnMetadata(hiddenByUser))),sheets(data(rowMetadata(hiddenByUser))),sheets(properties)'
MERGE_PLATFORMS = {
    "MMDB": "DB",
    "Data Transmission Service": "DB",
    "Video Network": "AZ",
}
BAREMETAL_PL_MERGE = {
    "DEV Efficiency": "engineering_infra.dev_efficiency",
}


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


def _convert_num(num):
    if len(num) == 0:
        return Decimal()
    if num == "#VALUE!" or num.upper() == "NA" or num.upper() == "N/A" or num.upper() == "#REF!":
        return Decimal()
    return Decimal(float(num.replace(",", "").replace("M", "")))


def _generate_column_sequence(num_columns):
    columns = []
    base = ord('A')
    while num_columns > 0:
        num_columns, remainder = divmod(num_columns - 1, 26)
        columns.append(chr(base + remainder))
    return ''.join(reversed(columns))


MAX_COL_NUM = 100
SPREADSHEET_COLS = [_generate_column_sequence(
    i) for i in range(1, MAX_COL_NUM + 1)]


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


def get_key_r2_sheets(meta):
    R1_SHARE = None

    for sheet_meta in meta["sheets"]:
        properties = sheet_meta["properties"]
        title = properties["title"]
        if title == R1_SHARE_SHEET_NAME:
            assert R1_SHARE is None, F"there are two sheets name called {R1_SHARE_SHEET_NAME}"
            R1_SHARE = sheet_meta

    return R1_SHARE


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
    bank_ret = {
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

        assert team == "Application", "team should be 'Application'"
        if region.upper() == "SG":
            region = "others"
        elif region.upper() == "US":
            region = "us"
        elif region == "Others":
            region = "others"
        else:
            log.warning(
                F"there is a empty region in cpo office bill {row}, will ignore")
            continue

        if bu != NON_BANK_FILTER:
            r = bank_ret[region]
        else:
            r = ret[region]
        # clean up the data
        r["power_opex"] += Decimal(power_opex)
        r["conn_opex"] += Decimal(conn_opex)
        r["mw"] += Decimal(mw)
        r["server_capex"] += Decimal(server_capex)
        r["network_capex"] += Decimal(network_capex)
        r["server_count"] += Decimal(server_count)

    return ret, bank_ret


def get_platform_servers(server_qty_sheet, storage_addtional_sheet, bare_metal_sheet):
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

        if category != CATEGORY_FILTER:  # ignore DI/AI
            continue
        if BU != NON_BANK_FILTER:  # ignore bank
            continue

        if location.lower() == "us":
            loc = "us"
        else:
            loc = "others"

        bu = bu.lower()
        if idc == NON_LIVE_DC and bu == "shopee":
            platform = "nonlive"
        if bu == "seamoney":
            platform = "seamoney"

        server_config = server_config.lower()

        if platform in MERGE_PLATFORMS:
            platform = MERGE_PLATFORMS[platform]

        if platform not in ret[loc]:
            ret[loc][platform] = {}
        if server_config not in ret[loc][platform]:
            ret[loc][platform][server_config] = 0
        ret[loc][platform][server_config] += int(qty)

    # get bare metal productline map
    bare_metal_map = {"us": {}, "others": {}}
    bare_metal_rows = spreadsheet.get_one_sheet_content(
        BILL_SHEET_ID, bare_metal_sheet["properties"]["title"])
    for bare_metal_row in bare_metal_rows[1:]:
        try:
            product_line, location, server_config, qty = bare_metal_row
            server_config = server_config.lower()
        except ValueError:
            print(F"abormal row for baremetal map: {bare_metal_row}")
            exit(-1)
        if location.lower() == "us":
            loc = "us"
        else:
            loc = "others"
        if product_line not in bare_metal_map[loc]:
            bare_metal_map[loc][product_line] = {}
        if server_config not in bare_metal_map[loc][product_line]:
            bare_metal_map[loc][product_line][server_config] = 0

        bare_metal_map[loc][product_line][server_config] += int(qty)

    # we move some platforms server to AZ+bare_metal
    for loc, platforms_map in ret.items():
        delete_platforms = []
        for platform, scs in platforms_map.items():
            if platform in BAREMETAL_PL_MERGE:
                # we remove this platform, add add them as a whole to barematal map
                product_line = BAREMETAL_PL_MERGE[platform]
                delete_platforms.append(platform)
                if product_line not in bare_metal_map[loc]:
                    bare_metal_map[loc][product_line] = {}
                for sc, qty in scs.items():
                    if sc not in bare_metal_map[loc][product_line]:
                        bare_metal_map[loc][product_line][sc] = 0
                    bare_metal_map[loc][product_line][sc] += qty
                    if sc not in ret[loc]["AZ"]:
                        ret[loc]["AZ"][sc] = 0
                    ret[loc]["AZ"][sc] += qty
        for p in delete_platforms:
            del platforms_map[p]

    # validate CPO office's storage platform and split map
    for c, q in ret["others"]["Storage"].items():
        assert ret["others"]["Storage-USS"].get(c, 0) + ret["others"]["Storage-Ceph"].get(c, 0) == q, \
            F"additional uss, ceph serverconfig {c} mismatch count with CPO Office's bill {q}: "\
            F'additional value is uss: {ret["others"]["Storage-USS"].get(c, 0)}, ceph {ret["others"]["Storage-Ceph"].get(c, 0)}'

    del ret["others"]["Storage"]  # splited straoge
    return ret, bare_metal_map


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


def get_ei_cost_share(r1_share_sheet):
    properties = r1_share_sheet["properties"]
    title = properties["title"]
    rows = spreadsheet.get_one_sheet_content(BILL_SHEET_ID, title)
    ret = {
        "us": {"capex": Decimal(), "opex": Decimal()},
        "others": {"capex": Decimal(), "opex": Decimal()},
    }

    for row in rows[1:]:
        fro, division, l0, l1, others_capex, others_opex, us_capex, us_opex = row
        assert fro == "DI", "we don't support other platform yet, just support DI"
        assert l0 == EI_L0_NAME, F"we just support share for {EI_L0_NAME}"
        ret["others"]["capex"] += Decimal(others_capex)
        ret["others"]["opex"] += Decimal(others_opex)
        ret["us"]["capex"] += Decimal(us_capex)
        ret["us"]["opex"] += Decimal(us_opex)

    log.debug(F"EI shared: {ret}")
    return ret


def calculate_platform_cost(cpo_bill, server_qty, server_unit_price, bare_metal_map, seamoney_sheet_us, seamoney_sheet_others):
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
                ret[smpls[i]][server_config] += int(seamoney_row[i + 1])
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
    for loc, smp in seamoney_ret.items():
        qty = 0
        for smpl, q in smp.items():
            qty += q["server_count"]
        if qty != ret[loc]["seamoney"]["server_count"]:
            log.warning(
                F'unblance seamoney count {qty} with platform count {ret[loc]["seamoney"]["server_count"]}')

    for loc, lc in ret.items():
        capex_sum = Decimal()
        opex_sum = Decimal()
        for platform, pm in lc.items():
            capex_sum += pm["projected_capex"]
            opex_sum += pm["projected_opex"]
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
    ret = {"us": {}, "others": {}}
    loc = "others"
    indicators = {}

    for platform_name, platform_sheet in platform_sheets.items():
        properties = platform_sheet["properties"]
        title = properties["title"]

        rows = spreadsheet.get_one_sheet_content(BILL_SHEET_ID, title)
        assert platform_name not in ret[loc], F"duplicated {platform_name}?"
        ret[loc][platform_name] = {}
        indicators[platform_name] = rows[1][1]  # set the indicators

        if MONTH_INDEX <= 4:  # before MAY, the format is a bit difference
            rows = rows[1:]
        else:
            rows = rows[2:]
        for row in rows:
            try:
                if MONTH_INDEX <= 4:
                    product_line, _, _, budget, quota, usage = row
                else:
                    row = row + [''] * max(0, 11 - len(row))
                    _, _, product_line, _, _, budget, quota, _, _, usage, _ = row
            except ValueError:
                print(
                    F"illegal row found in platform {platform_name} usage, line is {row}")
                exit(-1)
            if product_line == "":
                continue
            if product_line == "bank":
                continue
            try:
                budget = _convert_num(budget)
                quota = _convert_num(quota)
                usage = _convert_num(usage)
                ret[loc][platform_name][product_line] = {
                    "budget": budget,
                    "quota": quota,
                    "usage": usage,
                    "maxqu": max(quota, usage),
                    "percentage": 0.0,
                }
            except ValueError as e:
                print(
                    F"please check the format error in platform {platform_name} usage: {e}, {row}")
                exit(-1)

        qu_weights = []
        pls = []
        for pl, _ in ret[loc][platform_name].items():
            qu_weights.append(ret[loc][platform_name][pl]["maxqu"])
            pls.append(pl)
        qu_weights = normalize_weights(qu_weights, 1000000)
        for i in range(0, len(pls)):
            ret[loc][platform_name][pls[i]]["percentage"] = qu_weights[i]

    return ret


def get_pl_r1_bill(product_line_map, platform_cost, pl_usages, bare_metal_info, seamoney_info):
    ret = {
        "others": {},
        "us": {},
    }
    loc = "others"  # don't consider us yet

    cksum_capex = {"us": {}, "others": {}}
    cksum_opex = {"us": {}, "others": {}}
    for pl, _ in product_line_map.items():
        ret[loc][pl] = {}

    for platform, product_lines_usages in pl_usages[loc].items():
        for pl, _ in product_lines_usages.items():
            assert pl in ret[loc], F"dummy productline {pl} in platform {platform}"
        for pl, _ in bare_metal_info[loc].items():
            assert pl in ret[loc], F"dummy productline {pl} in baremetal"
        for pl, _ in seamoney_info[loc].items():
            assert pl in ret[loc], F"dummy productline {pl} in seamoney"

    for platform, product_lines_usages in pl_usages[loc].items():
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
        assert capex_error < 10, "platform {platform} capex error {capex_error} > 10"
        assert opex_error < 10, "platform {platform} capex error {opex_error} > 10"

    return ret


def get_pl_r2_bill(department_sheets, r1_shares):
    for loc, department_sheet in department_sheets.items():
        capex_sum = Decimal()
        opex_sum = Decimal()

        properties = department_sheet["properties"]
        title = properties["title"]

        dp_rows = spreadsheet.get_one_sheet_content(OVERVIEW_SHEET_ID, title)
        output_rows = []
        capex_weights = []
        opex_weights = []

        for i, dp_row in enumerate(dp_rows[2:]):
            division, l0, l1, capex, opex = dp_row
            capex = Decimal(capex)
            opex = Decimal(opex)

            if l0 == EI_L0_NAME:
                capex_sum += capex
                opex += opex
            else:
                capex_weights.append(capex)
                opex_weights.append(opex)
                output_rows.append(dp_row.copy())

        log.debug(F"EI original sum: Capex {capex_sum}, Opex {opex_sum}")
        for share in r1_shares[loc]:
            capex_sum += share["capex"]
            opex_sum += share["opex"]
        log.debug(F"EI R2 sum: Capex {capex_sum}, Opex {opex_sum}")

        capex_fracs = normalize_weights(capex_weights)
        opex_fracs = normalize_weights(opex_weights)

        for output_row, capex_frac, opex_frac in zip(output_rows, capex_fracs, opex_fracs):
            output_row[3] = Decimal(output_row[3]) * capex_frac / 10000
            output_row[4] = Decimal(output_row[4]) * opex_frac / 10000
    return output_row


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
            m = m.strip()
            if m in ret:
                assert ret[m][
                    1] == l0, F"we don't support L1 share cost yet!, duplicated {m}"
                ret[m][2].append(l1)
            else:
                ret[m] = (division, l0, [l1], cpo_office_link, quota_link)

    return ret


'''
generation parts
> overviews
> final bills
> cpo office update
'''


def generate_overviews(cpo_bill, pl_map, platform_cost, pl_r1_bills, bank_info):
    body = {
        "properties": {
            "title": F"App Platform Overviews V2 - {BILL_MONTH}"
        },
        "sheets": [

        ]
    }
    # get a sequencial platforms
    platform_set = set()
    platforms = []
    for pl, platform_bills in pl_r1_bills["others"].items():
        for platform in platform_bills.keys():
            if platform not in platform_set:
                platforms.append(platform)
                platform_set.add(platform)
    platforms.sort()
    last_diff_rows = []
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

        last_row = len(rows_data)
        # generate the check point
        rows_data.append({"values": [
            spreadsheet.get_cell_value("Total"),
            spreadsheet.get_cell_value(
                F"=SUM({SPREADSHEET_COLS[1]}2:{SPREADSHEET_COLS[1]}{last_row})", formula=True),
            spreadsheet.get_cell_value(""),
            spreadsheet.get_cell_value(""),
            spreadsheet.get_cell_value(
                F"=SUM({SPREADSHEET_COLS[4]}2:{SPREADSHEET_COLS[4]}{last_row})", formula=True),
            spreadsheet.get_cell_value(
                F"=SUM({SPREADSHEET_COLS[5]}2:{SPREADSHEET_COLS[5]}{last_row})", formula=True),
            spreadsheet.get_cell_value(
                F"=SUM({SPREADSHEET_COLS[6]}2:{SPREADSHEET_COLS[6]}{last_row})", formula=True),
            spreadsheet.get_cell_value(
                F"=SUM({SPREADSHEET_COLS[7]}2:{SPREADSHEET_COLS[7]}{last_row})", formula=True),
            spreadsheet.get_cell_value(
                F"=SUM({SPREADSHEET_COLS[8]}2:{SPREADSHEET_COLS[8]}{last_row})", formula=True),
            spreadsheet.get_cell_value(
                F"=SUM({SPREADSHEET_COLS[9]}2:{SPREADSHEET_COLS[9]}{last_row})", formula=True),
        ]})
        rows_data.append({"values": [
            spreadsheet.get_cell_value("CPO Office bill"),
            spreadsheet.get_cell_value(
                F"{cpo_bill[loc]['server_count']}", try_use_number=True),
            spreadsheet.get_cell_value(""),
            spreadsheet.get_cell_value(""),
            spreadsheet.get_cell_value(
                F"{cpo_bill[loc]['server_capex'] + cpo_bill[loc]['network_capex']}", try_use_number=True),
            spreadsheet.get_cell_value(
                F"{cpo_bill[loc]['power_opex'] + cpo_bill[loc]['conn_opex']}", try_use_number=True),
            spreadsheet.get_cell_value(
                F"{cpo_bill[loc]['server_capex']}", try_use_number=True),
            spreadsheet.get_cell_value(
                F"{cpo_bill[loc]['network_capex']}", try_use_number=True),
            spreadsheet.get_cell_value(
                F"{cpo_bill[loc]['power_opex']}", try_use_number=True),
            spreadsheet.get_cell_value(
                F"{cpo_bill[loc]['conn_opex']}", try_use_number=True),
        ]})
        rows_data.append({"values": [
            spreadsheet.get_cell_value("Diff"),
            spreadsheet.get_cell_value(
                F"=ABS(ROUND({SPREADSHEET_COLS[1]}{last_row+2}-{SPREADSHEET_COLS[1]}{last_row+2}, 2))", formula=True),
            spreadsheet.get_cell_value(""),
            spreadsheet.get_cell_value(""),
            spreadsheet.get_cell_value(
                F"=ABS(ROUND({SPREADSHEET_COLS[4]}{last_row+2}-{SPREADSHEET_COLS[4]}{last_row+2}, 2))", formula=True),
            spreadsheet.get_cell_value(
                F"=ABS(ROUND({SPREADSHEET_COLS[5]}{last_row+2}-{SPREADSHEET_COLS[5]}{last_row+2}, 2))", formula=True),
            spreadsheet.get_cell_value(
                F"=ABS(ROUND({SPREADSHEET_COLS[6]}{last_row+2}-{SPREADSHEET_COLS[6]}{last_row+2}, 2))", formula=True),
            spreadsheet.get_cell_value(
                F"=ABS(ROUND({SPREADSHEET_COLS[7]}{last_row+2}-{SPREADSHEET_COLS[7]}{last_row+2}, 2))", formula=True),
            spreadsheet.get_cell_value(
                F"=ABS(ROUND({SPREADSHEET_COLS[8]}{last_row+2}-{SPREADSHEET_COLS[8]}{last_row+2}, 2))", formula=True),
            spreadsheet.get_cell_value(
                F"=ABS(ROUND({SPREADSHEET_COLS[9]}{last_row+2}-{SPREADSHEET_COLS[9]}{last_row+2}, 2))", formula=True),
        ]

        })
        body["sheets"].append(
            {
                "properties": {
                    "title": F"[{BILL_MONTH}] Platform R1 Costs Overview - {loc}",
                },
                "data": {
                    "rowData": rows_data
                },
            },
        )
        last_diff_rows.append(last_row)

    dp_rows_data = []
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
    dp_header1 = [
        spreadsheet.get_cell_value("Department (Shopee App)"),
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value(F"{BILL_YEAR}-{BILL_MONTH}"),
        spreadsheet.get_cell_value(""),
    ]

    dp_header2 = [
        spreadsheet.get_cell_value("Division"),
        spreadsheet.get_cell_value("L0"),
        spreadsheet.get_cell_value("L1"),
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
    dp_rows_data.append({"values": dp_header1})
    dp_rows_data.append({"values": dp_header2})
    bank_row_index = -1
    dp_bank_row_index = -1

    departments = {}
    for pl, pv in pl_map.items():
        division, l0, l1, cpo_office_link, quota_link = pv
        if pl not in pl_r1_bills["others"]:
            continue
        if (division, l0, l1[0]) not in departments or pl == "bank":  # bank is quite special
            departments[(division, l0, l1[0])] = {
                "capex": Decimal(),
                "opex": Decimal(),
                "l1": l1,
            }

        platform_bills = pl_r1_bills["others"][pl]
        row_data = [
            spreadsheet.get_cell_value(l0),
            spreadsheet.get_cell_value("\n".join(l1)),
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
        if pl == "bank":
            capex_sum = bank_info["others"]["network_capex"] + \
                bank_info["others"]["server_capex"]
            opex_sum = bank_info["others"]["conn_opex"] + \
                bank_info["others"]["power_opex"]
            bank_row_index = len(rows_data) + 1
        departments[(division, l0, l1[0])]["capex"] += capex_sum
        departments[(division, l0, l1[0])]["opex"] += opex_sum
        row_data[3:3] = [
            spreadsheet.get_cell_value(float(capex_sum), try_use_number=True),
            spreadsheet.get_cell_value(float(opex_sum), try_use_number=True)
        ]
        rows_data.append({"values": row_data})

    last_row = len(rows_data)
    summary_row = [spreadsheet.get_cell_value("")] * 2
    summary_row.append(spreadsheet.get_cell_value("Total"))
    for i in range(3, (len(platforms) + 2) * 2, 2):
        summary_row.append(
            spreadsheet.get_cell_value(
                F"=SUM({SPREADSHEET_COLS[i]}2:{SPREADSHEET_COLS[i]}{last_row})-{SPREADSHEET_COLS[i]}{bank_row_index}",
                formula=True)
        ),
        summary_row.append(
            spreadsheet.get_cell_value(
                F"=SUM({SPREADSHEET_COLS[i+1]}2:{SPREADSHEET_COLS[i+1]}{last_row})-{SPREADSHEET_COLS[i+1]}{bank_row_index}",
                formula=True))
    rows_data.append({"values": summary_row})
    platform_row = [
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value("Platform Bills"),
        spreadsheet.get_cell_value(
            F'{cpo_bill["others"]["server_capex"]+cpo_bill["others"]["network_capex"]}', try_use_number=True),
        spreadsheet.get_cell_value(
            F'{cpo_bill["others"]["conn_opex"] + cpo_bill["others"]["power_opex"]}', try_use_number=True),
    ]
    for platform in platforms:
        platform_row.append(spreadsheet.get_cell_value(
            platform_cost[loc][platform]["projected_capex"], try_use_number=True))
        platform_row.append(spreadsheet.get_cell_value(
            platform_cost[loc][platform]["projected_opex"], try_use_number=True))
    rows_data.append({"values": platform_row})
    last_row = len(rows_data)
    diff_row = [
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value("Diff Check"),
    ]
    for i in range(3, (len(platforms) + 2) * 2+1):
        diff_row.append(spreadsheet.get_cell_value(
            F"=ABS(ROUND({SPREADSHEET_COLS[i]}{last_row} - {SPREADSHEET_COLS[i]}{last_row-1}, 2))", formula=True,
        ))
    rows_data.append({"values": diff_row})
    last_diff_rows.append(len(rows_data))

    body["sheets"].append(
        {
            "properties": {
                "title": F"{BILL_YEAR} - {BILL_MONTH} Product Line Bills Overview - others",
            },
            "data": {
                "rowData": rows_data
            },
        },
    )
    for (division, l0, _), dc in departments.items():
        def _i(l1_, capex, opex):
            return {"values": [
                spreadsheet.get_cell_value(division),
                spreadsheet.get_cell_value(l0),
                spreadsheet.get_cell_value(l1_),
                spreadsheet.get_cell_value(capex, try_use_number=True),
                spreadsheet.get_cell_value(opex, try_use_number=True),
            ]}

        l1 = dc["l1"]
        if len(l1) > 1:
            frac = Decimal(1) / len(l1)
            for _l1 in l1:
                dp_rows_data.append(
                    _i(_l1, dc["capex"] * frac, dc["opex"] * frac))
        else:
            dp_rows_data.append(_i(l1[0], dc["capex"], dc["opex"]))
        if l0 == "Digital Bank":
            dp_bank_row_index = len(dp_rows_data)

    assert dp_bank_row_index != -1, F"didn't find Digital bank in all departments {departments.keys()}"
    dp_rows_data.append({"values": [
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value("Total"),
        spreadsheet.get_cell_value(
            F"=ABS(ROUND(SUM({SPREADSHEET_COLS[3]}3:{SPREADSHEET_COLS[3]}{len(dp_rows_data)}) - {SPREADSHEET_COLS[3]}{dp_bank_row_index}, 2))", formula=True),
        spreadsheet.get_cell_value(
            F"=ABS(ROUND(SUM({SPREADSHEET_COLS[4]}3:{SPREADSHEET_COLS[4]}{len(dp_rows_data)}) - {SPREADSHEET_COLS[4]}{dp_bank_row_index}, 2))", formula=True),
    ]})
    dp_rows_data.append({"values": [
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value("CPO BILL"),
        spreadsheet.get_cell_value(
            F'{cpo_bill["others"]["server_capex"]+cpo_bill["others"]["network_capex"]}', try_use_number=True),
        spreadsheet.get_cell_value(
            F'{cpo_bill["others"]["conn_opex"] + cpo_bill["others"]["power_opex"]}', try_use_number=True),
    ]})
    dp_rows_data.append({"values": [
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value(""),
        spreadsheet.get_cell_value("DIFF Check"),
        spreadsheet.get_cell_value(
            F"=ABS(ROUND({SPREADSHEET_COLS[3]}{len(dp_rows_data)}-{SPREADSHEET_COLS[3]}{len(dp_rows_data)-1}, 2))", formula=True),
        spreadsheet.get_cell_value(
            F"=ABS(ROUND({SPREADSHEET_COLS[4]}{len(dp_rows_data)}-{SPREADSHEET_COLS[4]}{len(dp_rows_data)-1}, 2))", formula=True),
    ]
    })
    last_diff_rows.append(len(dp_rows_data))

    body["sheets"].append(
        {
            "properties": {
                "title": F"{BILL_YEAR} - {BILL_MONTH} Application Overview (Department) - others",
            },
            "data": {
                "rowData": dp_rows_data,
            },
        })

    overview_ss = spreadsheet.create_spreadsheet_file(body)
    doc_id = overview_ss["spreadsheetId"]
    drive.move_doc_to_folder(doc_id, OUTPUT_FOLDER)

    merge_req = []

    for sheet in overview_ss["sheets"]:
        merge_req.append(spreadsheet.get_autosize(
            sheet["properties"]["sheetId"]))

    def _diff_highlight_row(sheetId, row, startCol):
        return [{
            "addConditionalFormatRule": {
                "rule": spreadsheet.get_ge_rule(
                    sheetId,
                    "0.01",
                    condition="NUMBER_GREATER_THAN_EQ",
                    startRow=row,
                    endRow=row + 1,
                    startCol=startCol),
                "index": 0,
            }, },
            {
            "addConditionalFormatRule": {
                "rule": spreadsheet.get_ge_rule(
                    sheetId,
                    "0.01",
                    foreground_rgb="006400",
                    condition="NUMBER_LESS",
                    startRow=row,
                    endRow=row + 1,
                    startCol=startCol),
                "index": 0,
            },
        }]

    merge_req.extend(_diff_highlight_row(
        overview_ss["sheets"][0]["properties"]["sheetId"], last_diff_rows[0] + 2, 1))
    merge_req.extend(_diff_highlight_row(
        overview_ss["sheets"][1]["properties"]["sheetId"], last_diff_rows[1] + 2, 1))
    merge_req.extend(_diff_highlight_row(
        overview_ss["sheets"][2]["properties"]["sheetId"], last_diff_rows[2] - 1, 3))
    merge_req.extend(_diff_highlight_row(
        overview_ss["sheets"][3]["properties"]["sheetId"], last_diff_rows[3] - 1, 3))

    overview_bill_sheet_id = overview_ss["sheets"][2]["properties"]["sheetId"]
    # merge every two columns
    merge_req.append(spreadsheet.get_merge_cells_cmd(
        overview_bill_sheet_id, 0, 1, 3, 5),  # sum col
    )
    for i, platform in enumerate(platforms):
        merge_req.append(spreadsheet.get_merge_cells_cmd(
            overview_bill_sheet_id, 0, 1, 2 * i + 5, 2 * i + 7))

    depart_overview_bill_sheet_id = overview_ss["sheets"][3]["properties"]["sheetId"]
    merge_req.append(spreadsheet.get_merge_cells_cmd(
        depart_overview_bill_sheet_id, 0, 1, 0, 3
    ))
    merge_req.append(spreadsheet.get_merge_cells_cmd(
        depart_overview_bill_sheet_id, 0, 1, 3, 5
    ))
    spreadsheet.batch_update(doc_id, merge_req)

    return doc_id


def usage():
    print('''\033[0;35mpython3 bunnyapple.py [phrase]
    phrase=r1: generate the r1 report
    phase=r2: generate the r2 report \033[0;0m''')
    exit(0)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        usage()
    mode = sys.argv[1]

    if mode in ("r1", "r2"):
        meta = spreadsheet.get_spreadsheet_meta(
            BILL_SHEET_ID, fields=HIDDEN_BY_USER_FIELD)
        if mode == "r1":
            cpo_office_overall_sheet, server_qty_sheet, pricing_sheet, \
                product_line_sheet, additional_storage_sheet, seamoney_sheet_us, seamoney_sheet_others, bare_metal_sheet, platform_sheets = get_key_sheets(
                    meta)
            product_line_map = get_pl_map(product_line_sheet)

            cpo_overall, bank_overall = get_cpo_office_overall_bill(
                cpo_office_overall_sheet)
            server_qty, bare_metal_map = get_platform_servers(
                server_qty_sheet, additional_storage_sheet, bare_metal_sheet)
            server_unit_price = get_price_unit(pricing_sheet)

            platform_cost, bare_metal_cost, seamoney_cost = calculate_platform_cost(
                cpo_overall, server_qty,
                server_unit_price, bare_metal_map,
                seamoney_sheet_us, seamoney_sheet_others)

            pl_usage = get_pl_usage(platform_sheets)
            for loc, u in pl_usage.items():
                usage_keys = set(u.keys())
                pc_keys = set(platform_cost[loc].keys())
                usage_keys.update(['nonlive', 'AZ-Baremetal', 'seamoney'])
                diff1 = usage_keys.difference(pc_keys)
                diff2 = pc_keys.difference(usage_keys)

                if diff1 or diff2:
                    log.warning(
                        F'''Location '{loc}' platforms have difference, diff (usage-pc)={diff1}, (pc-usage)={diff2}''')

            pl_r1_bill = get_pl_r1_bill(product_line_map, platform_cost, pl_usage,
                                        bare_metal_cost, seamoney_cost)

            # generate the bill
            product_line_map = get_pl_map(product_line_sheet)
            generate_overviews(cpo_overall, product_line_map, platform_cost,
                               pl_r1_bill, bank_overall)
        elif mode == "r2":
            r1_di_sheet, _ = get_key_r2_sheets(meta)
            shared_costs = get_ei_cost_share(r1_di_sheet)
            get_pl_r2_bill()
    else:
        usage()
