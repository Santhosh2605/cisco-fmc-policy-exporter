# Cisco FMC Access Control Policy Exporter

A Python utility for exporting Cisco Firepower Management Center (FMC/FMCv) Access Control Policies (ACPs) and rules to a formatted Excel workbook. The exporter makes read-only configuration requests; it does not modify policies or deploy configuration.

## What the workbook contains

For each ACP, the workbook includes:

- A **rule worksheet** with rule actions, source and destination zones, networks and ports, applications, users, URLs, comments, intrusion policy, file policy, and logging details.
- A separate **Index worksheet** with total, enabled, and disabled rule counts; action counts; IPS and file policy status; logging; and time-range counts.
- An **Export Summary** listing each policy, rule count, and rule worksheet name.
- An **Export Errors** worksheet if the exporter encounters retrieval or object-resolution errors.

Referenced network and port groups are expanded where the FMC API returns the required data and the API account has permission.

## Requirements

- Network access from the computer running the exporter to the FMC management interface, normally HTTPS/TCP 443.
- FMC REST API access and an API account with read permissions for ACPs, rules, and referenced objects.
- Python 3.9 or newer recommended.
- The exporter file: `FMC_ACP_Policy_Exporter_With_Index.py`.
- Python packages: `requests` and `openpyxl`.

Use a dedicated FMC API account with the least read access needed. Do not place the FMC password or API token in the script or repository.

## Windows setup and run

These steps use PowerShell. Replace the folder or filename only if you have chosen a different location.

1. Install Python 3 from [python.org](https://www.python.org/downloads/windows/) or your organization’s approved software source. Select **Add python.exe to PATH** during installation.
2. Open PowerShell and confirm Python is available:

   ```powershell
   py --version
   ```

3. Create a working folder and copy `FMC_ACP_Policy_Exporter_With_Index.py` into it:

   ```powershell
   New-Item -ItemType Directory -Force -Path "$HOME\Desktop\FMC_Export"
   Set-Location "$HOME\Desktop\FMC_Export"
   ```

4. Create and activate a virtual environment, then install the required packages:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\python.exe -m pip install --upgrade pip
   .\.venv\Scripts\python.exe -m pip install requests openpyxl
   ```

5. Run the exporter from the working folder:

   ```powershell
   .\.venv\Scripts\python.exe .\FMC_ACP_Policy_Exporter_With_Index.py
   ```

6. Enter the requested information:
   - FMC hostname or IP address, without `https://`.
   - FMC API username.
   - FMC API password. The password prompt does not display typed characters.
   - TLS certificate verification choice. Choose **Y** when FMC presents a certificate trusted by the computer. Disable verification only when required by your environment and permitted by your security policy.

PowerShell may block virtual-environment activation under your organization’s policy. The commands above call the virtual environment’s Python directly and do not require activation.

## Linux setup and run

The following commands create a virtual environment so packages are installed for this exporter only. Run them from the directory where you saved the script.

### Ubuntu

Install Python, `venv`, and `pip` if they are not already installed:

```bash
sudo apt update
sudo apt install -y python3 python3-venv python3-pip
```

### CentOS Stream, RHEL, and Oracle Linux

On current dnf-based releases (typically version 8 or newer), install Python and pip:

```bash
sudo dnf install -y python3 python3-pip
```

On older yum-based CentOS/RHEL/Oracle Linux releases, use the approved package repository and package manager for that host. Ensure the installed Python version is supported by your organization and can create virtual environments before continuing. Older enterprise releases may not provide a sufficiently recent Python by default.

### Create the environment and install packages

Run this on Ubuntu or an EL-based distribution after installing the operating-system packages:

```bash
python3 --version
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install requests openpyxl
```

If the `venv` command is unavailable on Ubuntu, install `python3-venv` as shown above. On RHEL-compatible systems, use the Python packages provided or approved for your release; do not install Python packages into the system interpreter with `sudo pip`.

### Run the exporter

From the directory containing the script:

```bash
.venv/bin/python FMC_ACP_Policy_Exporter_With_Index.py
```

Enter the FMC hostname/IP, API username, password, and TLS verification preference when prompted. The password is hidden while typing.

## Output and count interpretation

The Excel workbook is created in the current working directory. Its filename begins with `FMC_ACP_Rules_By_Policy_` and includes a timestamp. Run the script from a dedicated export folder to keep the workbook there.

The index action counts cover all rules in the policy, including disabled rules; enabled and disabled totals are reported separately. FMC action values are mapped as follows:

| FMC action value | Index action |
| --- | --- |
| `ALLOW` | Allow |
| `BLOCK` | Block |
| `TRUST` | Trust |
| `MONITOR` | Monitor |
| `BLOCK_WITH_RESET`, `BLOCK_RESET` | Block with reset |
| `INTERACTIVE_BLOCK`, `BLOCK_INTERACTIVE` | Interactive block |
| `INTERACTIVE_BLOCK_WITH_RESET`, `BLOCK_RESET_INTERACTIVE` | Interactive block with reset |

Unmatched or missing action values appear in **Unrecognized / missing action** for review. The requested action categories show plain numeric counts.

For IPS/file policy and time-range counts, the exporter counts only values it can confirm from the rule or a read-only resolved FMC object. If FMC omits a field or the referenced object cannot be resolved, the corresponding count may show `0`; it does not distinguish an explicit “not enabled” value from missing/unavailable data in the index. Time ranges are counted as available or expired only when FMC returns an explicit status. Check these counts in FMC when the policy is expected to contain such settings.

Logging enabled is counted when a rule has `logBegin` or `logEnd` set to true. Review the rule worksheet for the exported logging fields.

## Troubleshooting

| Symptom | Checks |
| --- | --- |
| `py` or `python3` is not found | Install Python 3 from an approved source, reopen the terminal, and confirm the command is on `PATH`. |
| `No module named requests` or `openpyxl` | Install the packages with the virtual-environment Python shown above. |
| Connection timeout | Check the FMC hostname/IP, DNS, routing, firewall rules, and HTTPS/TCP 443 connectivity. |
| HTTP 401 or authentication failure | Verify credentials, account status, REST API access, and the FMC domain access. |
| HTTP 403 or forbidden | Ask the FMC administrator to review read permissions for policies, rules, and referenced objects. |
| TLS/certificate error | Install or trust the correct FMC certificate. Do not disable certificate verification unless your security policy permits it. |
| Workbook cannot be saved | Close the workbook in Excel or another office application and check write permission in the current folder. |
| Group members are not expanded | Expansion depends on the object type, API response, and account permissions. Review **Export Errors** and compare with FMC. |
| An action appears under Unrecognized / missing action | Check the rule’s exact `action` value in FMC and compare it with the mappings above. |

## Security and validation

- Configuration retrieval uses read-only API requests. Authentication may use the FMC token endpoint; the exporter does not create, edit, or deploy policy configuration.
- Use a dedicated API account with least-privilege read access.
- Never share the FMC password, access token, or exported workbook publicly.
- Treat the workbook as sensitive because it may include internal IP addresses, network names, and security policy details.
- Compare the exported policy names and rule totals with FMC. Review **Export Errors** and confirm expected policy status counts in FMC.
- This export is not a replacement for a supported FMC backup.
