# f5report

A command-line tool to export the main information from an F5 BIG-IP device to an Excel report.

---

## Overview

`f5report` reads a set of JSON files exported from an F5 BIG-IP device's iControl REST API (virtual servers, pools, nodes, self-IPs, routes, SSL certificates and LTM statistics) and builds a single, multi-sheet `.xlsx` report.

It is particularly useful for **documentation, assessment, audits and migration activities**, where having the full LTM configuration and its live status laid out in one workbook makes it much easier to review than digging through the API output, CLI or GUI.

---

## Features

- Parses 8 different F5 iControl REST API JSON exports in a single run: virtual servers, pools, nodes, self-IPs, routes, SSL certificates, virtual server stats and pool member stats
- Produces a single `.xlsx` workbook with one sheet per object type (`VIPs`, `POOLs`, `NODEs`, `DCs`, `IFs`, `ROUTEs`, `VSTATs`, `PSTATs`)
- Strips partition prefixes and resolves route domains (`%n`) for every object
- Converts dotted-decimal netmasks to CIDR prefixes
- Resolves each static route's gateway to its connected self-IP network and VLAN
- Merges live statistics (connections, bits/packets in/out, availability status) with the corresponding VIP/pool/member configuration
- Frozen header row and autofilter enabled on every sheet for quick sorting/filtering
- Missing or malformed fields are clearly flagged (`*none`, `*all`, `*error`) instead of failing silently
- Custom input file paths via CLI argument

---

## Requirements

- Python 3.10+
- Uses the Python standard library (`argparse`, `re`, `json`, `pathlib`, `ipaddress`, `datetime`) plus the third-party package `xlsxwriter`.

---

## Installation

```
git clone https://github.com/pyquerci/f5report.git
cd f5report
pip install xlsxwriter
```

### Windows

A pre-compiled Windows executable can be built with PyInstaller using the command:

```
pyinstaller --onefile f5report.py
```

No Python installation is needed if you use the compiled `f5report.exe`. For convenience, you can add it to a folder in your system `PATH` to invoke it from any directory; for example, I keep mine in `C:\Tools\f5report`.

---

## Data Sources

[#data-sources](#data-sources)

`f5report` does not talk to the F5 device itself — it only consumes JSON files that you must generate beforehand. These files come straight from the BIG-IP iControl REST API:

| File (default name) | Endpoint |
| --- | --- |
| `vips.json` | `https://IP_ADDRESS/mgmt/tm/ltm/virtual?expandSubcollections=true` |
| `pools.json` | `https://IP_ADDRESS/mgmt/tm/ltm/pool?expandSubcollections=true` |
| `nodes.json` | `https://IP_ADDRESS/mgmt/tm/ltm/node?expandSubcollections=true` |
| `ifs.json` | `https://IP_ADDRESS/mgmt/tm/net/self` |
| `routes.json` | `https://IP_ADDRESS/mgmt/tm/net/route` |
| `dcs.json` | `https://IP_ADDRESS/mgmt/tm/sys/file/ssl-cert` |
| `vstats.json` | `https://IP_ADDRESS/mgmt/tm/ltm/virtual/stats` |
| `pstats.json` | `https://IP_ADDRESS/mgmt/tm/ltm/pool/members/stats` |

You have two ways to get them:

- **Manually**, hitting each URL with a browser or a tool like `curl`/`Postman` (HTTP Basic Auth against the device), saving each response as the corresponding local JSON file.
- **With [`sfetchapi`](https://github.com/pyquerci/sfetchapi)**, another tool of mine built exactly for this: it reads a simple YAML file listing `filename: url` pairs and downloads all of them in one run, using a single set of credentials, and pretty-prints the JSON automatically. Its own README uses these very 8 endpoints as a worked example, so it's a drop-in way to produce all the input files `f5report` expects without manually calling the API eight times. Much faster and less error-prone than doing it by hand.

---

## Usage

```
f5report.py [-h]
            [-a]
            [-c VIPS POOLS NODES IFS ROUTES DCS VSTATS PSTATS]
            -e [FILE]
```

### Arguments

| Argument | Description |
| --- | --- |
| `-h, --help` | Show the help message and exit. |
| `-a, --about` | Show author, version, project URL and license information. |
| `-c, --config VIPS POOLS NODES IFS ROUTES DCS VSTATS PSTATS` | Custom JSON input files, in this exact order. Accepts plain filenames (looked up in the current directory) or full/relative paths. Defaults to `vips.json`, `pools.json`, `nodes.json`, `ifs.json`, `routes.json`, `dcs.json`, `vstats.json` and `pstats.json` in the current directory. |
| `-e, --export [NAME]` | Export the LTM report to an xlsx file. Required. Defaults to `report.xlsx` if no name is given. |

If a JSON file is not found, cannot be read due to permission issues, or is not valid JSON, `f5report` exits with a clear error message.

### Examples

```
# Generate report.xlsx using the default JSON filenames in the current directory
f5report.py -e

# Generate a custom named report
f5report.py -e ltm_report.xlsx

# Use custom input file paths
f5report.py -c D:\f5\vips.json D:\f5\pools.json D:\f5\nodes.json D:\f5\ifs.json D:\f5\routes.json D:\f5\dcs.json D:\f5\vstats.json D:\f5\pstats.json -e ltm_report.xlsx

# Show author and version information
f5report.py -a
```

---

## How It Works

[#how-it-works](#how-it-works)

For each of the 8 JSON files, `f5report` walks the `items` (or `entries`, for the two stats files) collection and rebuilds every object into a flat, readable record:

- **Partitions** are stripped from every reference (`/Common/my_vip` → `my_vip`), and the **route domain** (`%n`) is extracted separately when present in an address.
- **Netmasks** on virtual servers are converted from dotted-decimal to CIDR prefix length.
- **Self-IP networks** (`ifs.json`) are parsed with `ipaddress` to compute their actual network address, and each **static route**'s next-hop is looked up against them to resolve the connected network and VLAN.
- **SSL certificate expiration timestamps** are converted from epoch to ISO 8601 UTC.
- **Statistics** (`vstats.json`, `pstats.json`) are unnested from F5's `nestedStats.entries` structure; for pools, each member's stats are extracted individually and merged into the parent pool row as newline-separated lists, aligned by position with the member name/monitor/status columns.
- Any field that cannot be found is written as `*none`, an empty list of VLANs is written as `*all`, and any value that cannot be parsed at all is written as `*error` — so gaps in the source data are visible in the spreadsheet rather than causing a crash.

Each object type is written to its own worksheet with a frozen header row, an autofilter, and word-wrap enabled, so multi-value fields (members, pools, VLANs, rules...) stay readable.

---

## License

This project is licensed under the **GNU General Public License v2.0 (GPLv2)**. You are free to use, modify, and distribute this software under the terms of that license. See the [LICENSE](https://github.com/pyquerci/f5report/blob/main/LICENSE) file for the full license text.

---

## Donations

If you value the work and want to help support its development, feel free to make a donation. Your support will be greatly appreciated:

- PayPal: <https://paypal.me/pyquerci>
- Buy Me a Coffee: <https://buymeacoffee.com/pyquerci>
