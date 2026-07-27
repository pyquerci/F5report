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
- Merges live statistics with the corresponding configuration
- Frozen header row and autofilter enabled on every sheet for quick sorting/filtering
- Missing or malformed fields are clearly flagged (`*none`, `*all`, `*error`) instead of failing silently

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

A pre-compiled Windows executable is included in the repository, built with PyInstaller 6.21.0 using the command:

```
pyinstaller --onefile f5report.py
```

No Python installation is needed, just download and run f5report.exe. For convenience, you can add it to a folder in your system PATH to invoke it from any directory; for example, I keep mine in C:\Tools\f5report.

---

## Data Sources

`f5report` does not talk to the F5 device itself, it only consumes JSON files that you must generate beforehand. These files come straight from the BIG-IP iControl REST API:

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
- **With [`sfetchapi`](https://github.com/pyquerci/sfetchapi)**, another tool of mine built exactly for this: it reads a simple YAML file listing `filename: url` pairs and downloads all of them in one run, using a single set of credentials, and pretty-prints the JSON automatically.

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
| `-e, --export [NAME]` | Export the LTM report to an xlsx file. This argument is required. Defaults to `report.xlsx` if no name is given. |

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

For each of the 8 JSON files, `f5report` extracts the main information and writes it to a dedicated sheet in the Excel report: `VIPs`, `POOLs`, `NODEs`, `DCs`, `IFs`, `ROUTEs`, `VSTATs` and `PSTATs`. Along the way, a few values are cleaned up to make them easier to read, partition prefixes are stripped from names, netmasks are converted to CIDR prefixes, and statistics are matched to the corresponding VIP, pool or member.

Each sheet has a frozen header row and an autofilter enabled, so you can sort and filter the data directly in Excel.

### Example

The GIF below shows the export in action: `f5report` generating the xlsx report from the JSON input files. Before opening it in Excel, the file was also processed with [`xlfit`](https://github.com/pyquerci/xlfit), another tool of mine that auto-fits rows and columns, to keep the sheets readable without having to resize everything by hand.

![f5report demo](f5report.gif)

---

## Production Testing and Compatibility

`f5report` has been tested against a production F5 BIG-IP configuration extracted from a system running version **17.5.1.5-0.0.6**, virtualized on **F5OS 1.8.3-23493**, on **F5-R5800** hardware. No issues were encountered during its use. The provisioned modules in use were `Local Traffic (LTM)`, `Access Policy (APM)`, `iRules Language Extensions (iRulesLX)`, and `Application Visibility and Reporting (AVR)`. Keep in mind that other versions may rename fields, change nesting, or restructure the API response, which could cause parsing errors or silently wrong output.

---

## License

This project is licensed under the **GNU General Public License v2.0 (GPLv2)**. You are free to use, modify, and distribute this software under the terms of that license. See the [LICENSE](https://github.com/pyquerci/f5report/blob/main/LICENSE) file for the full license text.

---

## Donations

If you value the work and want to help support its development, feel free to make a donation. Your support will be greatly appreciated:

- PayPal: <https://paypal.me/pyquerci>
- Buy Me a Coffee: <https://buymeacoffee.com/pyquerci>
