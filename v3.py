#!/usr/bin/env python

from drive import create_file_in_folder, GOOGLE_SHEET_MIME
import spreadsheet
import drive
import pprint
from termcolor import colored

BILL_SHEET_ID = "1VXbMo0fFjNANF02lxPIqGJxhdeRja7SVs3VtksdKfU8"
CPO_OFFICE_OVERALL_SHEET_NAME = "CPO Office Bill"
SERVER_MAP_SHEET_NAME = "ServerMap"
SERVER_PRICE_SHEET_NAME = "Server Pricing"
STANDARD_SERVER_CONFIG = "s1_v2"

NON_BANK_FILTER = "Exclude Bank"

HIDDEN_BY_USER_FIELD = 'sheets(data(columnMetadata(hiddenByUser))),sheets(data(rowMetadata(hiddenByUser))),sheets(properties)'
MERGE_PLATFORMS = {
    "MMDB": "DB",
    "Data Transmission Service": "DB",
}


def get_key_sheets(meta):
    CPO_OVERALL_BILL_SHEET = None
    SERVER_QTY_SHEET = None
    SERVER_PRICE_SHEET = None

    for sheet_meta in meta["sheets"]:
        properties = sheet_meta["properties"]
        if properties["title"] == CPO_OFFICE_OVERALL_SHEET_NAME:
            assert CPO_OVERALL_BILL_SHEET is None, F"there are two sheets name called {CPO_OFFICE_OVERALL_SHEET_NAME}"
            CPO_OVERALL_BILL_SHEET = sheet_meta
        elif properties["title"] == SERVER_MAP_SHEET_NAME:
            assert SERVER_QTY_SHEET is None, F"there are two sheets name called {SERVER_QTY_SHEET}"
            SERVER_QTY_SHEET = sheet_meta
        elif properties["title"] == SERVER_PRICE_SHEET_NAME:
            assert SERVER_PRICE_SHEET is None, F"there are two sheets name called {SERVER_PRICE_SHEET}"
            SERVER_PRICE_SHEET = sheet_meta

    return CPO_OVERALL_BILL_SHEET, SERVER_QTY_SHEET, SERVER_PRICE_SHEET


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
        return {"idc_cost": 0, "conn_cost": 0, "mw": 0, "server_cost": 0, "network_cost": 0, "server_count": 0}
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
        team, bu, region, idc_cost, conn_cost, mw, server_cost, network_cost, server_count = row

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
        ret[region]["idc_cost"] += float(idc_cost)
        ret[region]["conn_cost"] += float(conn_cost)
        ret[region]["mw"] += float(mw)
        ret[region]["server_cost"] += float(server_cost)
        ret[region]["network_cost"] += float(network_cost)
        ret[region]["server_count"] += float(server_count)

    return ret


def get_platform_servers(server_qty_sheet):
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
        "others": {}
    }
    properties = server_qty_sheet.get("properties")
    title = properties.get("title")

    rows = spreadsheet.get_one_sheet_content(BILL_SHEET_ID, title)
    for row in rows[1:]:
        try:
            date, bu, platform, region, idc, server_config, qty, category, location, BU = row[
                :10]
        except ValueError:
            print(F"illegal row in server_quantity sheet {str(row)}")
            exit(0)

        if category != "APP":  # ignore DI/AI
            continue
        if BU != NON_BANK_FILTER:  # ignore bank
            continue

        if idc == "DC West":
            platform = "nonlive"
        elif bu == "shopee" or bu == "seamoney":
            if location.lower() == "us":
                loc = "us"
            else:
                loc = "others"
            
            if bu == "seamoney":
                platform = "seamoney"
        else:
            continue  # ignore others, like seamoney etc.

        if platform in MERGE_PLATFORMS:
            platform = MERGE_PLATFORMS[platform]

        server_config = server_config.lower()

        if platform not in ret[loc]:
            ret[loc][platform] = {}
        if server_config not in ret[loc][platform]:
            ret[loc][platform][server_config] = 0
        ret[loc][platform][server_config] += int(qty)

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
            "price": float(price),
            "power": float(power),
        }

    return ret


def calculate_platform_cost(cpo_bill, server_qty, server_unit_price):
    '''
    output format is 
    platform, server_count, total_capex, total_server_power, projected_server_capex, projected_network_device_capex, 
    projected_server_opex, projected_connectivity_opex, allocated_capex, allocated_opex
    '''
    ret = {
        "us": {},
        "others": {}
    }
    total_server_count = 0
    total_server_capex = 0.0
    total_power = 0.0
    total_conn_opex = 0.0

    for loc, platforms_qty in server_qty.items():
        for platform, server_config_map in platforms_qty.items():
            if platform not in ret[loc]:
                ret[loc][platform] = {}
            capex = 0
            server_power = 0
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

            ret[loc][platform] = {
                "capex": capex,
                "server_power": server_power,
                "count": server_count,
                "projected_server_capex": 0.0,
                "projected_network_capex": 0.0,
                "projected_power_opex": 0.0,
                "projected_conn_opex": 0.0,
            }
            if platform != "nonlive":  #non live no need join total calculation 
                total_server_count += server_count
                total_server_capex += capex
                total_power += server_power
    
    #print(F"{total_server_count}, {total_server_capex}, {total_power}")

    for loc, loc_details in ret.items():
        for platform, platform_cost in loc_details.items():
            ret[loc][platform]["server_share"] = float(
                platform_cost["count"]) / total_server_count
            details = cpo_bill[loc]

            ret[loc][platform]["projected_server_capex"] += float(
                platform_cost["capex"]) / total_server_capex * details["server_cost"]
            ret[loc][platform]["projected_network_capex"] += float(
                platform_cost["count"])/total_server_count * details["network_cost"]
            ret[loc][platform]["projected_power_opex"] += float(
                platform_cost["server_power"] / total_power * details["idc_cost"]
            )
            ret[loc][platform]["projected_conn_opex"] += float(
                platform_cost["count"]) / total_server_count * details["conn_cost"]

    return ret


if __name__ == "__main__":
    meta = spreadsheet.get_spreadsheet_meta(
        BILL_SHEET_ID, fields=HIDDEN_BY_USER_FIELD)
    cpo_office_overall_sheet, server_qty_sheet, pricing_sheet = get_key_sheets(
        meta)
    cpo_overall = get_cpo_office_overall_bill(cpo_office_overall_sheet)
    server_qty = get_platform_servers(server_qty_sheet)
    server_unit_price = get_price_unit(pricing_sheet)
    
    ret = calculate_platform_cost(cpo_overall, server_qty, server_unit_price)
    pprint.pprint(ret)