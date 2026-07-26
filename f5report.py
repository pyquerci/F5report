# Author: Andrea Querci
# version: 1.0.0
# project: https://github.com/pyquerci/f5report
# license: GPLv2

import re
import sys
import argparse
from pathlib import Path
import json
from io import BytesIO
import xlsxwriter
import datetime
import ipaddress

DEFAULT_JSON_VIPS = "vips.json"
DEFAULT_JSON_POOLS = "pools.json"
DEFAULT_JSON_NODES = "nodes.json"
DEFAULT_JSON_IFS = "ifs.json"
DEFAULT_JSON_ROUTES = "routes.json"
DEFAULT_JSON_DCS = "dcs.json"
DEFAULT_JSON_VSTATS = "vstats.json"
DEFAULT_JSON_PSTATS = "pstats.json"

NONE = "*none"
ALL = "*all"
ERROR = "*error"

class LTMConfig:
    def __init__(self,
        vips_config: dict,
        pools_config: dict,
        nodes_config: dict,
        ifs_config: dict,
        routes_config: dict,
        dcs_config: dict,
        vstats_config=dict,
        pstats_config=dict
    ) -> None:
        
        self.vips_config = vips_config
        self.pools_config = pools_config
        self.nodes_config = nodes_config
        self.ifs_config = ifs_config
        self.routes_config = routes_config
        self.dcs_config = dcs_config
        self.vstats_config = vstats_config
        self.pstats_config = pstats_config
        
        self.vips = self._improve_vips_config()
        self.pools = self._improve_pools_config()
        self.nodes = self._improve_nodes_config()
        self.ifs = self._improve_ifs_config()
        self.routes = self._improve_routes_config()
        self.dcs = self._improve_dcs_config()
        self.vstats = self._improve_vstats_config()
        self.pstats = self._improve_pstats_config()
    
    def mask_to_prefix(self, mask: str) -> int:
        if mask == "any":
            return "0"
        else:
            return sum(bin(int(octet)).count("1") for octet in mask.split("."))

    def _improve_vips_config(self):
        result = {}
        
        for i, dic in enumerate(self.vips_config["items"]):
            partition = dic.get("partition", NONE)
            p = f"/{partition}/"
            vip = dic.get("name", NONE)
            src = dic.get("source", NONE)
            raw = dic.get("destination", NONE)
            dst = raw.replace(p, "")
            raw = dic.get("mask")
            prefix = self.mask_to_prefix(raw) if raw else NONE
            sni = dic.get("serversslUseSni", NONE)
            snat = dic.get("sourceAddressTranslation", {}).get("type") or NONE
            dnat = dic.get("translateAddress", NONE)
            snatp = dic.get("sourcePort", NONE)
            dnatp = dic.get("translatePort", NONE)
            protocol = dic.get("ipProtocol", NONE)

            if "enabled" in dic:
                status = "enabled"
            elif "disabled" in dic:
                status = "disabled"
            else:
                status = ERROR
            
            raw = re.search(r'%(\d+)', dst)
            rd = raw.group(1) if raw else "0"
            
            des = dic.get("description", NONE)
            persist = [d.get("name", NONE) for d in dic.get("persist", {})] or NONE
            policy = [d["name"] for d in dic.get("policiesReference", {}).get("items", [])] or NONE
            profile = [d["name"] for d in dic.get("profilesReference", {}).get("items", [])] or NONE
            vlan = [r.replace(p, "") for r in dic.get("vlans", [])] or ALL
            rule = [r.replace(p, "") for r in dic.get("rules", [])] or NONE
            raw = dic.get("pool")
            pool = raw.replace(p, "") if raw else NONE
            
            result[i] = {
                "partition": partition,
                "rd": rd,
                "vip": vip,
                "src": src,
                "dst": dst,
                "prefix": prefix,
                "sni": sni,
                "persist": persist,
                "protocol": protocol,
                "snat": snat,
                "dnat": dnat,
                "snatp": snatp,
                "dnatp": dnatp,
                "pool": pool,
                "rule": rule,
                "profile": profile,
                "policy": policy,
                "vlan": vlan,
                "status": status,
                "des": des
            }
        
        return result

    def _improve_pools_config(self):
        result = {}
        
        for i, dic in enumerate(self.pools_config["items"]):
            partition = dic.get("partition", NONE)
            p = f"/{partition}/"
            pool = dic.get("name", NONE)
            mode = dic.get("loadBalancingMode", NONE)
            member = [d["name"] for d in dic.get("membersReference", {}).get("items", [])] or NONE
            node = [d["address"] for d in dic.get("membersReference", {}).get("items", [])] or NONE
            
            raw = dic.get("monitor")
            monitor = raw.replace(p, "") if raw else NONE

            raw = re.search(r'%(\d+)', node[0])
            rd = raw.group(1) if raw else "0"

            result[i] = {
                "partition": partition,
                "pool": pool,
                "rd": rd,
                "member": member,
                "node": node,
                "mode": mode,
                "monitor": monitor
            }

        return result

    def _improve_nodes_config(self):
        result = {}
        
        for i, dic in enumerate(self.nodes_config["items"]):
            partition = dic.get("partition", NONE)
            p = f"/{partition}/"
            node = dic.get("name", NONE)
            address = dic.get("address", NONE)
            
            raw = dic.get("monitor")
            monitor = raw.replace(p, "") if raw else NONE

            raw = re.search(r'%(\d+)', address)
            rd = raw.group(1) if raw else "0"

            result[i] = {
                "partition": partition,
                "rd": rd,
                "node": node,
                "address": address,
                "monitor": monitor
            }

        return result

    def _improve_ifs_config(self):
        result = {}

        for i, dic in enumerate(self.ifs_config["items"]):
            partition = dic.get("partition", NONE)
            p = f"/{partition}/"
            name = dic.get("name", NONE)
            address = dic.get("address", NONE)
            raw = re.search(r'%(\d+)', address)
            rd = raw.group(1) if raw else "0"
            floating = dic.get("floating", NONE)
            raw = dic.get("trafficGroup", NONE)
            tg = raw.replace(p, "") if raw else NONE
            itg = dic.get("inheritedTrafficGroup", NONE)
            raw = dic.get("vlan")
            vlan = raw.replace(p, "") if raw else NONE
            services = dic.get("allowService") or NONE

            try:
                raw = re.sub(r"%\d+", "", address)
                net = str(ipaddress.ip_network(raw, strict=False))
            except ValueError:
                net = ERROR

            result[i] = {
                "partition": partition,
                "rd": rd,
                "name": name,
                "address": address,
                "floating": floating,
                "tg": tg,
                "itg": itg,
                "net": net,
                "vlan": vlan,
                "services": services
            }

        return result

    def _ifs_lookup(self, ip: str) -> bool:
        try:
            for dic in self.ifs.values():
                ip = ipaddress.ip_address(ip)
                raw = re.sub(r"%\d+", "", dic["address"])
                net = ipaddress.ip_network(raw, strict=False)
                
                if ip in net:
                    con = str(net)
                    vlan = dic["vlan"]
                    return con, vlan
        
            return ERROR, ERROR
        
        except ValueError:
            return ERROR, ERROR

    def _improve_routes_config(self):
        result = {}

        for i, dic in enumerate(self.routes_config["items"]):
            partition = dic.get("partition", NONE)
            name = dic.get("name", NONE)
            net = dic.get("network", NONE)
            raw = re.search(r'%(\d+)', net)
            rd = raw.group(1) if raw else "0"
            gw = dic.get("gw", NONE)
            
            con, vlan = self._ifs_lookup(ip = gw.split("%")[0])

            result[i] = {
                "partition": partition,
                "rd": rd,
                "name": name,
                "net": net,
                "gw": gw,
                "con": con,
                "vlan": vlan 
            }
        
        return result

    def _improve_dcs_config(self):
        result = {}
        
        for i, dic in enumerate(self.dcs_config["items"]):
            partition = dic.get("partition", NONE)
            dc = dic.get("name", NONE)
            subject = dic.get("subject", NONE)
            version = dic.get("version", NONE)
            ks = dic.get("certificateKeySize", NONE)
            creation = dic.get("createTime", NONE)
            ctype = "bundle" if dic.get("isBundle") == "true" else "normal"
            issuer = dic.get("issuer", NONE)
            ktype = dic.get("keyType", NONE)
            sn = dic.get("serialNumber", NONE)
            checksum = dic.get("checksum", NONE)
            fingerprint = dic.get("fingerprint", NONE)

            raw = dic.get("expirationDate", NONE)
            expiration = datetime.datetime.fromtimestamp(raw, datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

            result[i] = {
                "partition": partition,
                "dc": dc,
                "subject": subject,
                "version": version,
                "ks": ks,
                "creation": creation,
                "expiration": expiration,
                "ctype": ctype,
                "issuer": issuer,
                "ktype": ktype,
                "sn": sn,
                "checksum": checksum,
                "fingerprint": fingerprint
            }

        return result

    def _fix_any(self, value: str) -> str:
        # example: any%1:any
        if "any" in value:
            first, last = value.split(":")
            if first.startswith("any"):
                first = first.replace("any", "0.0.0.0")
            if last.endswith("any"):
                last = "0"
            return f"{first}:{last}"
        else:
            return value

    def _improve_vstats_config(self):
        result = {}
        
        for i, mdic in enumerate(self.vstats_config["entries"].values()):
            dic = mdic.get("nestedStats", {}).get("entries", {})
            raw = dic.get("destination", {}).get("description", NONE)
            dst = self._fix_any(value=raw)
            raw = dic.get("tmName", {}).get("description", NONE)
            partition = raw.split("/")[1]
            p = f"/{partition}/"
            vip = raw.replace(p, "")
            raw = re.search(r'%(\d+)', dst)
            rd = raw.group(1) if raw else "0"
            sas = dic.get("status.availabilityState", {}).get("description", NONE)
            ses = dic.get("status.enabledState", {}).get("description", NONE)
            ccc = dic.get("clientside.curConns", {}).get("value", NONE)
            ctc = dic.get("clientside.totConns", {}).get("value", NONE)
            cbi = dic.get("clientside.bitsIn", {}).get("value", NONE)
            cbo = dic.get("clientside.bitsOut", {}).get("value", NONE)

            status_reason = dic.get("status.statusReason", {}).get("description", NONE)

            result[i] = {
                "partition": partition,
                "rd": rd,
                "vip": vip,
                "dst": dst,
                "sas": sas,
                "ses": ses,
                "ccc": str(ccc),
                "ctc": str(ctc),
                "cbi": str(cbi),
                "cbo": str(cbo),
                "status_reason": status_reason
            }

        return result

    def _improve_pstats_config(self):
        result = {}
        
        for i, mdic in enumerate(self.pstats_config["entries"].values()):
            dic = mdic.get("nestedStats", {}).get("entries", {})
            
            # pool
            raw = dic.get("tmName", {}).get("description", NONE)
            partition = raw.split("/")[1]
            p = f"/{partition}/"
            pool = raw.replace(p, "")
            am = dic.get("activeMemberCnt", {}).get("value", NONE)
            amc = dic.get("availableMemberCnt", {}).get("value", NONE)
            raw = dic.get("monitorRule", {}).get("description", NONE)
            monitor = raw.replace(p, "")
            ses = dic.get("status.enabledState", {}).get("description", NONE)
            sas = dic.get("status.availabilityState", {}).get("description", NONE)
            scc = dic.get("serverside.curConns", {}).get("value", NONE)
            stc = dic.get("serverside.totConns", {}).get("value", NONE)
            sbi = dic.get("serverside.pktsIn", {}).get("value", NONE)
            sbo = dic.get("serverside.pktsOut", {}).get("value", NONE)
            status_reason = dic.get("status.statusReason", {}).get("description", NONE)
            
            # members
            members_name = []
            members_dst = []
            members_monitor = []
            members_ses = []
            members_sas = []
            members_check = []
            members_scc = []
            members_stc = []
            members_sbi = []
            members_sbo = []
            members_status_reason = []
            
            members_cmd = next(
                (k for k in dic if k.startswith("https")),
                None
            )
            if members_cmd:
                dic2=dic[members_cmd]["nestedStats"]["entries"]
                
                for key in dic2:
                    dic3=dic2[key]["nestedStats"]["entries"]
                    
                    raw = dic3.get("nodeName", {}).get("description", NONE)
                    members_name.append(raw.replace(p, ""))
                    port = f":{str(dic3.get("port", {}).get("value", NONE))}"
                    members_dst.append(dic3.get("addr", {}).get("description", NONE) + port)
                    raw = dic3.get("monitorRule", {}).get("description", NONE)
                    members_monitor.append(raw.replace(p, ""))
                    raw = dic3.get("status.enabledState", {}).get("description", NONE)
                    members_ses.append(raw)
                    raw = dic3.get("status.availabilityState", {}).get("description", NONE)
                    members_sas.append(raw)
                    raw = dic3.get("monitorStatus", {}).get("description", NONE)
                    members_check.append(raw)
                    raw = dic3.get("serverside.curConns", {}).get("value", NONE)
                    members_scc.append(str(raw))
                    raw = dic3.get("serverside.totConns", {}).get("value", NONE)
                    members_stc.append(str(raw))
                    raw = dic3.get("serverside.pktsIn", {}).get("value", NONE)
                    members_sbi.append(str(raw))
                    raw = dic3.get("serverside.pktsOut", {}).get("value", NONE)
                    members_sbo.append(str(raw))
                    members_status_reason.append(dic3.get("status.statusReason", {}).get("description", NONE))
            else:
                members_name = ERROR
                members_dst = ERROR
                members_monitor = ERROR
                members_ses = ERROR
                members_sas = ERROR
                members_check = ERROR
                members_scc = ERROR
                members_stc = ERROR
                members_sbi = ERROR
                members_sbo = ERROR
                members_status_reason = ERROR

            # finish
            result[i] = {
                "partition": partition,
                "pool": pool,
                "monitor": monitor,
                "am": str(am),
                "amc": str(amc),
                "sas": sas,
                "ses": ses,
                "scc": str(scc),
                "stc": str(stc),
                "sbi": str(sbi),
                "sbo": str(sbo),
                "status_reason": status_reason,
                "members_name":members_name,
                "members_dst": members_dst,
                "members_monitor": members_monitor,
                "members_ses": members_ses,
                "members_sas": members_sas,
                "members_check": members_check,
                "members_scc": members_scc,
                "members_stc": members_stc,
                "members_sbi": members_sbi,
                "members_sbo": members_sbo,
                "members_status_reason": members_status_reason               
            }

        return result

    def export(self, file_path: str) -> None:
        output = BytesIO()
        wb = xlsxwriter.Workbook(output)
        
        # vips
        ws_1 = wb.add_worksheet("VIPs")

        header = {
            "partition": "PARTITION",
            "rd": "RD",
            "vip": "VIP NAME",
            "src": "SRC IP + RD",
            "dst": "DST IP + RD + PORT",
            "prefix": "PREFIX",
            "sni": "SNI",
            "persist": "PERSIST",
            "protocol": "PROTOCOL",
            "snat": "SNAT IP",
            "dnat": "DNAT IP",
            "snatp": "SNAT PORT",
            "dnatp": "DNAT PORT",
            "pool": "POOL NAME",
            "rule": "RULE",
            "profile": "PROFILE",
            "policy": "POLICY",
            "vlan": "VLAN",
            "status": "STATUS",
            "des": "DESCRIPTION"
        }
        
        ls = list(self.vips.values())
        create_table(wb, ws_1, header, ls)
        
        # pools
        ws_2 = wb.add_worksheet("POOLs")

        header = {
            "partition": "PARTITION",
            "rd": "RD",
            "pool": "POOL NAME",
            "member": "NODE NAME + PORT",
            "node": "NODE IP + RD",
            "mode": "LB MODE",
            "monitor": "MONITOR"
        }

        ls = list(self.pools.values())
        create_table(wb, ws_2, header, ls)

        # nodes
        ws_3 = wb.add_worksheet("NODEs")

        header = {
            "partition": "PARTITION",
            "rd": "RD",
            "node": "NODE NAME",
            "address": "ADDRESS + PORT",
            "monitor": "MONITOR"
        }

        ls = list(self.nodes.values())
        create_table(wb, ws_3, header, ls)

        # dcs
        ws_4 = wb.add_worksheet("DCs")

        header = {
            "partition": "PARTITION",
            "dc": "NAME",
            "subject": "SUBJECT",
            "ctype": "TYPE",
            "version": "VERSION",
            "ks": "KEY STRING",
            "ktype": "KEY TYPE",
            "creation": "START TIME",
            "expiration": "END TIME",
            "issuer": "ISSUER",
            "sn": "S/N",
            "checksum": "CHECKSUM",
            "fingerprint": "FINGERPRINT"
        }

        ls = list(self.dcs.values())
        create_table(wb, ws_4, header, ls)

        # ifs
        ws_5 = wb.add_worksheet("IFs")

        header = {
            "partition": "PARTITION",
            "rd": "RD",
            "name": "IF NAME",
            "address": "IP + RD + PREFIX",
            "floating": "FLOATING",
            "tg": "TRAFFIG GROUP",
            "itg": "INHERITED TG",
            "net": "NETWORK",
            "vlan": "VLAN",
            "services": "SERVICES"
        }

        ls = list(self.ifs.values())
        create_table(wb, ws_5, header, ls)

        # routes
        ws_6 = wb.add_worksheet("ROUTEs")

        header = {
            "partition": "PARTITION",
            "rd": "RD",
            "name": "STATIC NAME",
            "net": "NETWORK",
            "gw": "NEXT-HOP",
            "con": "CONNECTED",
            "vlan": "VLAN",
        }

        ls = list(self.routes.values())
        create_table(wb, ws_6, header, ls)

        # vstats
        ws_7 = wb.add_worksheet("VSTATs")

        header = {
            "partition": "PARTITION",
            "rd": "RD",
            "vip": "VIP NAME",
            "dst": "DST IP + RD + PORT",
            "ses": "CONFIG",
            "sas": "STATUS",
            "ccc": "CUR CONN",
            "ctc": "TOT CONN",
            "cbi": "BITS IN",
            "cbo": "BITS OUT",
            "status_reason": "NOTE"
        }

        ls = list(self.vstats.values())
        create_table(wb, ws_7, header, ls)                

        # pstats
        ws_8 = wb.add_worksheet("PSTATs")

        header = {
            "partition": "PARTITION",
            "pool": "POOL NAME",
            "monitor": "POOL MONITOR",
            "am": "ACTIVE MEMEBERS",
            "amc": "AVAILABLE MEMBERS",
            "ses": "POOL CONFIG",
            "sas": "POOL STATUS",
            "scc": "POOL CUR CONN",
            "stc": "POOL TOT CONN",
            "sbi": "POOL BITS IN",
            "sbo": "POOL BITS OUT",
            "status_reason": "POOL INFO",
            "members_name": "MEMBER NAME",
            "members_dst": "NODE IP + RD + PORT",
            "members_monitor": "MEMBER MONITOR",
            "members_ses": "MEMBER CONFIG",
            "members_sas": "MEMBER STATUS",
            "members_check": "MEMBER CHECK",
            "members_scc": "MEMBER CUR CONN",
            "members_stc": "MEMBER TOT CONN",
            "members_sbi": "MEMBER BITS IN",
            "members_sbo": "MEMBER BITS OUT",
            "members_status_reason": "MEMBER INFO"
        }

        ls = list(self.pstats.values())
        create_table(wb, ws_8, header, ls)   

        # export
        wb.close()

        try:
            with open(file_path, "wb") as f:
                f.write(output.getvalue())
        except PermissionError:
            raise SystemExit(f"error: permission denied")


def create_table(
    wb: xlsxwriter.Workbook,
    ws: xlsxwriter.worksheet.Worksheet,
    header: dict,
    ls: list[dict]
    ) -> None:
    
    ls_header = []
    int_header = len(header) - 1
    
    for key in header:
        ls_header.append(header[key])
    ws.freeze_panes(1, 0)
    
    cell_format = wb.add_format({"bold": True, "align": "top", "bg_color": "#e6b8b7", "num_format": '@'})
    for column_id in range(len(ls_header)):
        ws.write(0, column_id, ls_header[column_id], cell_format)
    ws.autofilter(0, 0, 0, int_header)

    cell_format = wb.add_format({"align": "top", "text_wrap": True, "num_format": '@'})
    for row_id, obj in enumerate(ls):
        row_id += 1
        for column_id, key in enumerate(header):
            cell = obj.get(key)
            if cell is None:
                cell = ERROR
            elif isinstance(cell, list):
                cell = [item if item else ERROR for item in cell]
                cell = "\n".join(cell)
            ws.write(row_id, column_id, cell, cell_format)


def load_json(path: str) -> dict:
    file_path = Path(path)

    if not file_path.is_file():
        raise SystemExit(f"error: file not found: {path}")

    try:
        with file_path.open("r", encoding="utf-8") as f:
            return json.load(f)

    except PermissionError:
        raise SystemExit(f"error: permission denied: {path}")

    except json.JSONDecodeError:
        raise SystemExit(f"error: invalid json format: {path}")


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        usage="%(prog)s [-h] [-a] [-c VIPS POOLS NODES IFS ROUTES DCS VSTATS PSTATS] -e [FILE]",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog=(
            "description:\n"
            "  generate a LTM report from F5 BIG-IP JSON files\n"
            "  they must be generated via API:\n\n"
            "  - vips.json:\t\t"
            "https://IP_ADDRESS/mgmt/tm/ltm/virtual?expandSubcollections=true\n"
            "  - pools.json:\t\t"
            "https://IP_ADDRESS/mgmt/tm/ltm/pool?expandSubcollections=true\n"
            "  - nodes.json\t\t"
            "https://IP_ADDRESS/mgmt/tm/ltm/node?expandSubcollections=true\n"
            "  - ifs.json:\t\t"
            "https://IP_ADDRESS/mgmt/tm/net/self\n"
            "  - routes.json:\t"
            "https://IP_ADDRESS/mgmt/tm/net/route\n"
            "  - dcs.json:\t\t"
            "https://IP_ADDRESS/mgmt/tm/sys/file/ssl-cert\n"
            "  - vstats.json:\t"
            "https://IP_ADDRESS/mgmt/tm/ltm/virtual/stats\n"
            "  - pstats.json:\t"
            "https://IP_ADDRESS/mgmt/tm/ltm/pool/members/stats\n\n"
        ),
    )

    parser.add_argument(
        "-a",
        "--about",
        action="store_true",
        help="show author, version, project URL and license information",
    )

    parser.add_argument(
        "-c",
        "--config",
        nargs=8,
        metavar=("VIPS", "POOLS", "NODES", "IFS", "ROUTES", "DCS", "VSTATS", "PSTATS"),
        help=(
            "JSON input files (default file names in description)"
        ),
    )

    parser.add_argument(
        "-e",
        "--export",
        nargs="?",
        const="report.xlsx",
        default=None,
        metavar="NAME",
        help="export the LTM report to an xlsx file (default: report.xlsx)",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_arguments()

    if args.about:
        print(
        "author: Andrea Querci\n"
        "version: 1.0\n"
        "project: https://github.com/pyquerci/f5report\n"
        "license: GPLv2"
        )
        return

    if not args.export:
        raise SystemExit("error: argument -e/--export is required")

    if args.config:
        vips_file, pools_file, nodes_file, ifs_file, routes_file, dcs_file, vstats_file, pstats_file = args.config
    else:
        vips_file = DEFAULT_JSON_VIPS
        pools_file = DEFAULT_JSON_POOLS
        nodes_file = DEFAULT_JSON_NODES
        ifs_file = DEFAULT_JSON_IFS
        routes_file = DEFAULT_JSON_ROUTES
        dcs_file = DEFAULT_JSON_DCS
        vstats_file = DEFAULT_JSON_VSTATS
        pstats_file = DEFAULT_JSON_PSTATS

    vips = load_json(vips_file)
    pools = load_json(pools_file)
    nodes = load_json(nodes_file)
    ifs = load_json(ifs_file)
    routes = load_json(routes_file)
    dcs = load_json(dcs_file)
    vstats = load_json(vstats_file)
    pstats = load_json(pstats_file)
    
    config = LTMConfig(
        vips_config=vips,
        pools_config=pools,
        nodes_config=nodes,
        ifs_config=ifs,
        routes_config=routes,
        dcs_config=dcs,
        vstats_config=vstats,
        pstats_config=pstats
    )

    config.export(args.export)


if __name__ == "__main__":
    main()