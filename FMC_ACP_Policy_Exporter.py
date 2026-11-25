#!/usr/bin/env python3
"""
Cisco FMC / FMCv Access Control Policy exporter.
Creates ONE Excel workbook with ONE worksheet per Access Control Policy.
Network groups and port groups are expanded recursively where FMC API access permits.

Read-only: all FMC configuration requests are GET requests.
Authentication uses FMC REST API token endpoint.

Install:
    py -m pip install requests openpyxl

Run:
    py fmc_acp_policy_sheets.py
"""

import getpass
import ipaddress
import json
import re
import sys
from datetime import datetime
from urllib.parse import urljoin

import requests
import urllib3
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

VERIFY_TLS = True
PAGE_SIZE = 1000

HEADERS = [
    "Name",
    "Action",
    "Source Zones",
    "Destination Zones",
    "Source Network",
    "Source Port",
    "Destination Network",
    "Destination Port",
    "Dynamic Attributes",
    "Applications",
    "Users",
    "URLs",
    "Comments",
    "Intrusion policy",
    "File policy",
    "Logging",
]

# Known FMC object type names -> FMC object collection endpoints.
OBJECT_COLLECTIONS = {
    "Network": "networks",
    "NetworkGroup": "networkgroups",
    "Host": "hosts",
    "Range": "ranges",
    "FQDN": "fqdns",
    "IPv4Address": "ipv4addresses",
    "IPv6Address": "ipv6addresses",
    "ProtocolPortObject": "protocolportobjects",
    "PortObjectGroup": "portobjectgroups",
    "SecurityZone": "securityzones",
    "VlanTag": "vlantags",
    "URL": "urls",
    "URLCategory": "urlcategories",
    "URLGroup": "urlgroups",
    "Application": "applications",
    "ApplicationFilter": "applicationfilters",
    "User": "users",
    "UserGroup": "usergroups",
    "DynamicObject": "dynamicobjects",
    "DynamicObjectGroup": "dynamicobjectgroups",
    "VariableSet": "variablesets",
}

NETWORK_GROUP_TYPES = {"NetworkGroup"}
PORT_GROUP_TYPES = {"PortObjectGroup"}


def json_text(value):
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def clean_cell(value):
    """Avoid Excel illegal control characters and keep values readable."""
    text = json_text(value)
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", text)


def get_name(ref):
    if isinstance(ref, dict):
        return ref.get("name") or ref.get("value") or ref.get("id") or ""
    return str(ref) if ref is not None else ""


def make_session():
    session = requests.Session()
    retry = Retry(
        total=4,
        connect=4,
        read=3,
        backoff_factor=1,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET"]),
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.headers.update({"Accept": "application/json"})
    return session


class FMC:
    def __init__(self, host, username, password, verify_tls=True):
        self.base = "https://" + host.strip().replace("https://", "").replace("http://", "").strip("/")
        self.verify_tls = verify_tls
        self.session = make_session()
        self.username = username
        self.password = password
        self.token = None
        self.refresh_token = None
        self.domain_uuid = None
        self.cache = {}
        self.authenticate()

    def authenticate(self):
        url = self.base + "/api/fmc_platform/v1/auth/generatetoken"
        response = self.session.post(
            url,
            auth=(self.username, self.password),
            data="",
            timeout=45,
            verify=self.verify_tls,
        )
        if response.status_code not in (200, 204):
            raise RuntimeError(
                f"Authentication failed: HTTP {response.status_code}\n{response.text[:1600]}"
            )

        self.token = response.headers.get("X-auth-access-token")
        self.refresh_token = response.headers.get("X-auth-refresh-token")
        self.domain_uuid = (
            response.headers.get("DOMAIN_UUID")
            or response.headers.get("X-auth-domain-uuid")
            or response.headers.get("global")
        )

        if not self.token:
            raise RuntimeError(
                "FMC did not return X-auth-access-token. "
                f"Headers received: {', '.join(response.headers.keys())}"
            )
        self.session.headers["X-auth-access-token"] = self.token

        # Fallback for releases that expose domain information separately.
        if not self.domain_uuid:
            info = self.request_json(self.base + "/api/fmc_platform/v1/info/domain")
            if isinstance(info, dict):
                self.domain_uuid = info.get("uuid") or info.get("id")
                domains = info.get("items") or info.get("domains") or []
                if not self.domain_uuid and domains:
                    self.domain_uuid = domains[0].get("uuid") or domains[0].get("id")

        if not self.domain_uuid:
            raise RuntimeError(
                "FMC returned an access token but no domain UUID. "
                "Check FMC REST API version/domain access."
            )

    def refresh(self):
        if not self.refresh_token:
            self.authenticate()
            return
        response = self.session.post(
            self.base + "/api/fmc_platform/v1/auth/refreshtoken",
            headers={
                "X-auth-access-token": self.token,
                "X-auth-refresh-token": self.refresh_token,
            },
            timeout=45,
            verify=self.verify_tls,
        )
        if response.status_code not in (200, 204):
            self.authenticate()
            return
        self.token = response.headers.get("X-auth-access-token", self.token)
        self.refresh_token = response.headers.get("X-auth-refresh-token", self.refresh_token)
        self.session.headers["X-auth-access-token"] = self.token

    def request_json(self, url, params=None):
        response = self.session.get(
            url, params=params, timeout=90, verify=self.verify_tls
        )
        if response.status_code == 401:
            self.refresh()
            response = self.session.get(
                url, params=params, timeout=90, verify=self.verify_tls
            )
        if response.status_code == 200:
            try:
                return response.json()
            except ValueError:
                return None
        return None

    def get_all(self, url, params=None):
        result = []
        offset = 0
        while True:
            q = dict(params or {})
            q.update({"offset": offset, "limit": PAGE_SIZE})
            response = self.session.get(
                url, params=q, timeout=90, verify=self.verify_tls
            )
            if response.status_code == 401:
                self.refresh()
                response = self.session.get(
                    url, params=q, timeout=90, verify=self.verify_tls
                )
            if response.status_code != 200:
                raise RuntimeError(
                    f"GET {response.url} -> HTTP {response.status_code}: "
                    f"{response.text[:1200]}"
                )
            payload = response.json()
            items = payload.get("items", [])
            result.extend(items)
            paging = payload.get("paging") or {}
            count = paging.get("count")
            page_limit = paging.get("limit", PAGE_SIZE)
            page_offset = paging.get("offset", offset)
            if not items:
                break
            if count is not None and page_offset + len(items) >= count:
                break
            if len(items) < page_limit:
                break
            offset = page_offset + len(items)
        return result

    @property
    def root(self):
        return f"{self.base}/api/fmc_config/v1/domain/{self.domain_uuid}"

    def resolve(self, ref):
        """Resolve one FMC reference to its full object JSON, with cache."""
        if not isinstance(ref, dict):
            return None
        obj_id = ref.get("id")
        obj_type = ref.get("type")
        if not obj_id or not obj_type:
            return None
        key = (obj_type, obj_id)
        if key in self.cache:
            return self.cache[key]

        # FMC self links are the most accurate endpoint when supplied.
        self_link = (ref.get("links") or {}).get("self")
        if self_link:
            url = urljoin(self.base + "/", self_link.lstrip("/"))
            data = self.request_json(url)
            if data:
                self.cache[key] = data
                return data

        collection = OBJECT_COLLECTIONS.get(obj_type)
        if collection:
            url = f"{self.root}/object/{collection}/{obj_id}"
            data = self.request_json(url)
            if data:
                self.cache[key] = data
                return data

        self.cache[key] = None
        return None


def network_value_from_object(obj):
    """Extract a network leaf's actual address/value without repeating its object name."""
    value = (
        obj.get("value")
        or obj.get("ipAddress")
        or obj.get("address")
        or obj.get("network")
        or obj.get("fqdn")
    )
    if value:
        return str(value)
    if obj.get("start") or obj.get("end"):
        return f"{obj.get('start', '')}-{obj.get('end', '')}".strip("-")
    if obj.get("ip") and obj.get("netmask"):
        return f"{obj['ip']}/{obj['netmask']}"
    return obj.get("name") or ""


def compact_network_values(values):
    """
    Collapse only exact address coverage. This never broadens the matched address set.
    A single host stays as a single IP; adjacent hosts/subnets may become a CIDR.
    """
    ip_networks = []
    non_ip = []
    for value in values:
        value = str(value).strip()
        if not value:
            continue
        try:
            if "/" in value:
                ip_networks.append(ipaddress.ip_network(value, strict=False))
            else:
                address = ipaddress.ip_address(value)
                ip_networks.append(ipaddress.ip_network(f"{address}/{address.max_prefixlen}"))
        except ValueError:
            non_ip.append(value)

    compacted = []
    if ip_networks:
        # Separate address families; collapse_addresses removes contained ranges
        # and combines adjacent ranges only when their union is exact.
        for version in (4, 6):
            family = [net for net in ip_networks if net.version == version]
            for net in ipaddress.collapse_addresses(family):
                if net.prefixlen == net.max_prefixlen:
                    compacted.append(str(net.network_address))
                else:
                    compacted.append(str(net))
    compacted.extend(non_ip)
    return list(dict.fromkeys(compacted))


def collect_network_group_values(fmc, ref, visited=None, depth=0):
    """Recursively collect leaf IP/FQDN/range values from nested network groups."""
    if visited is None:
        visited = set()
    if depth > 12 or not isinstance(ref, dict):
        return [str(ref)] if ref is not None else []

    obj_type = ref.get("type", "")
    obj_id = ref.get("id", "")
    key = (obj_type, obj_id) if obj_id else None
    if key and key in visited:
        return []
    if key:
        visited = visited | {key}

    full = fmc.resolve(ref)
    obj = full or ref
    typ = obj.get("type") or obj_type

    members = []
    for field in ("objects", "literals", "members", "networkObjects"):
        value = obj.get(field)
        if isinstance(value, list):
            members.extend(value)
        elif value:
            members.append(value)

    if members or typ in NETWORK_GROUP_TYPES or "Group" in typ:
        values = []
        for member in members:
            if isinstance(member, dict):
                values.extend(collect_network_group_values(fmc, member, visited, depth + 1))
            else:
                values.append(str(member))
        return values

    value = network_value_from_object(obj)
    return [value] if value else []


def flatten_object_tree(fmc, ref, mode, visited=None, depth=0):
    """
    Render object values like FMC's rule table.
    Single host objects display only their IP (not 'object name: same IP').
    Network groups show the group name followed by compacted member addresses.
    """
    if visited is None:
        visited = set()
    if depth > 12:
        return ["[maximum nested group depth reached]"]
    if not isinstance(ref, dict):
        return [str(ref)] if ref is not None else []

    label = get_name(ref)
    obj_type = ref.get("type", "")
    obj_id = ref.get("id", "")
    key = (obj_type, obj_id) if obj_id else None
    if key and key in visited:
        return [f"{label} [duplicate/circular reference]"]
    if key:
        visited.add(key)

    full = fmc.resolve(ref)
    obj = full or ref
    name = obj.get("name") or label
    typ = obj.get("type") or obj_type

    if mode == "network":
        is_group = typ in NETWORK_GROUP_TYPES or "Group" in typ
        if is_group:
            raw_values = collect_network_group_values(fmc, ref, set(), depth)
            compacted = compact_network_values(raw_values)
            if compacted:
                return [f"{name} (expanded):"] + [f"  - {value}" for value in compacted]
            return [f"{name} (members could not be resolved)"]
        value = network_value_from_object(obj)
        if value:
            # Avoid duplicate output such as "5.2.25.1: 5.2.25.1".
            return [str(value)]
        return [name or json_text(obj)]

    if mode == "port":
        is_group = typ in PORT_GROUP_TYPES or "PortObjectGroup" in typ
        if not is_group:
            pieces = []
            protocol = obj.get("protocol")
            port = obj.get("port")
            port_range = obj.get("portRange") or obj.get("range")
            if protocol:
                pieces.append(f"protocol={protocol}")
            if port is not None:
                pieces.append(f"port={port}")
            if port_range is not None:
                pieces.append(f"range={json_text(port_range)}")
            if obj.get("icmpType") is not None:
                pieces.append(f"ICMP type={obj.get('icmpType')}")
            if obj.get("icmpCode") is not None:
                pieces.append(f"ICMP code={obj.get('icmpCode')}")
            if pieces:
                return [", ".join(pieces)]
            return [name or json_text(obj)]

    members = []
    for field in ("objects", "literals", "members", "portObjects"):
        value = obj.get(field)
        if isinstance(value, list):
            members.extend(value)
        elif value:
            members.append(value)

    if members:
        output = [f"{name} (expanded):"]
        for member in members:
            if isinstance(member, dict):
                output.extend(
                    "  - " + item
                    for item in flatten_object_tree(fmc, member, mode, visited.copy(), depth + 1)
                )
            else:
                output.append(f"  - {member}")
        return output

    return [name or json_text(ref)]

def condition_items(value):
    """
    FMC rule conditions commonly wrap references as:
      {"objects": [...], "literals": [...]}
    Return the actual members, not the wrapper dictionary.
    """
    if value is None:
        return []
    if isinstance(value, dict):
        if "objects" in value or "literals" in value:
            items = []
            objects = value.get("objects") or []
            literals = value.get("literals") or []
            if isinstance(objects, list):
                items.extend(objects)
            elif objects:
                items.append(objects)
            if isinstance(literals, list):
                items.extend(literals)
            elif literals:
                items.append(literals)
            return items
        return [value]
    if isinstance(value, list):
        return value
    return [value]


def literal_display(item, mode=None):
    """Render FMC literal entries as the concise values shown in the FMC UI."""
    if not isinstance(item, dict):
        return str(item)
    if mode == "network":
        value = (
            item.get("value") or item.get("ipAddress") or item.get("address")
            or item.get("network") or item.get("fqdn")
        )
        if value:
            return str(value)
        # Some FMC literal records use start/end or address/netmask.
        if item.get("start") or item.get("end"):
            return f"{item.get('start', '')}-{item.get('end', '')}".strip("-")
        if item.get("ip") and item.get("netmask"):
            return f"{item['ip']}/{item['netmask']}"
    if mode == "port":
        protocol = item.get("protocol") or item.get("protocolName") or ""
        port = item.get("port") or item.get("portValue")
        start = item.get(" 시작")  # harmless fallback for unusual localized payloads
        end = item.get("end")
        if port is not None:
            return f"{protocol + ' ' if protocol else ''}{port}".strip()
        if item.get("begin") is not None or end is not None:
            return f"{protocol + ' ' if protocol else ''}{item.get('begin', '')}-{end or ''}".strip()
        if item.get("portRange"):
            return f"{protocol + ' ' if protocol else ''}{item['portRange']}".strip()
    return item.get("name") or item.get("value") or item.get("id") or json_text(item)


def render_refs(fmc, value, mode=None, any_if_empty=True):
    items = condition_items(value)
    if not items:
        return "Any" if any_if_empty else ""

    rendered = []
    for ref in items:
        if isinstance(ref, dict):
            # Literal entries are often not API object references (no id/type).
            if not (ref.get("id") and ref.get("type")):
                shown = literal_display(ref, mode)
                if shown:
                    rendered.append(shown)
                continue

            if mode in ("network", "port"):
                rendered.extend(flatten_object_tree(fmc, ref, mode))
            else:
                full = fmc.resolve(ref)
                name = (full or {}).get("name") or get_name(ref)
                rendered.append(name or str(ref.get("id", "")))
        else:
            rendered.append(str(ref))

    unique = list(dict.fromkeys(x for x in rendered if x))
    return "\n".join(unique) if unique else ("Any" if any_if_empty else "")

def first_present(rule, *keys):
    for key in keys:
        if key in rule:
            return rule.get(key)
    return None


def get_zone_fields(fmc, rule):
    src = first_present(rule, "sourceZones", "sourceZone")
    dst = first_present(rule, "destinationZones", "destinationZone")
    src_text = render_refs(fmc, src, any_if_empty=True)
    dst_text = render_refs(fmc, dst, any_if_empty=True)
    combined = f"Source: {src_text} | Destination: {dst_text}"
    return combined, src_text, dst_text


def get_policy_name(fmc, ref):
    if not isinstance(ref, dict):
        return ""
    full = fmc.resolve(ref)
    return (full or {}).get("name") or ref.get("name") or ""


def get_logging(rule):
    fields = []
    for key, label in (
        ("logBegin", "Log at beginning"),
        ("logEnd", "Log at end"),
        ("sendEventsToFMC", "Send events to FMC"),
        ("syslogConfig", "Syslog"),
        ("snmpConfig", "SNMP"),
    ):
        if key in rule and rule.get(key) is not None:
            value = rule.get(key)
            if isinstance(value, dict):
                value = value.get("name") or json_text(value)
            fields.append(f"{label}: {value}")
    return "\n".join(fields)


def build_rule_row(fmc, rule):
    _zones, src_zones, dst_zones = get_zone_fields(fmc, rule)

    src_networks = first_present(rule, "sourceNetworks", "sourceNetwork")
    dst_networks = first_present(rule, "destinationNetworks", "destinationNetwork")
    src_ports = first_present(rule, "sourcePorts", "sourcePort")
    dst_ports = first_present(rule, "destinationPorts", "destinationPort")

    dynamic_attrs = first_present(
        rule, "dynamicAttributes", "dynamicAttribute", "dynamicObjects"
    )
    applications = first_present(rule, "applications", "applicationFilters")
    users = first_present(rule, "users", "userGroups")
    urls = first_present(rule, "urls", "urlCategories", "urlGroups")

    ips_ref = first_present(rule, "ipsPolicy", "intrusionPolicy")
    file_ref = first_present(rule, "filePolicy")

    return {
        "Name": rule.get("name", ""),
        "Action": rule.get("action", ""),
        "Source Zones": src_zones,
        "Destination Zones": dst_zones,
        "Source Network": render_refs(fmc, src_networks, "network"),
        "Source Port": render_refs(fmc, src_ports, "port"),
        "Destination Network": render_refs(fmc, dst_networks, "network"),
        "Destination Port": render_refs(fmc, dst_ports, "port"),
        "Dynamic Attributes": render_refs(fmc, dynamic_attrs),
        "Applications": render_refs(fmc, applications),
        "Users": render_refs(fmc, users),
        "URLs": render_refs(fmc, urls),
        "Comments": rule.get("comment") or rule.get("comments") or rule.get("description") or "",
        "Intrusion policy": get_policy_name(fmc, ips_ref),
        "File policy": get_policy_name(fmc, file_ref),
        "Logging": get_logging(rule),
    }


def excel_sheet_name(name, used):
    name = re.sub(r"[\[\]:*?/\\]", "_", str(name)).strip("'")
    if not name:
        name = "Access Control Policy"
    base = name[:31]
    candidate = base
    n = 1
    while candidate in used:
        suffix = f"_{n}"
        candidate = base[:31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def style_sheet(ws):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 32
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for col_idx, cells in enumerate(ws.columns, 1):
        header = str(ws.cell(1, col_idx).value or "")
        max_width = max(14, min(len(header) + 3, 32))
        for cell in list(cells)[1:]:
            if cell.value is not None:
                max_width = max(max_width, min(max((len(line) for line in str(cell.value).splitlines()), default=0) + 2, 55))
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_width, 55)



ACTION_LABELS = {
    "ALLOW": "Allow",
    "BLOCK": "Block",
    "TRUST": "Trust",
    "MONITOR": "Monitor",
    "BLOCK_WITH_RESET": "Block with reset",
    "BLOCK_RESET": "Block with reset",
    "INTERACTIVE_BLOCK": "Interactive block",
    "BLOCK_INTERACTIVE": "Interactive block",
    "INTERACTIVE_BLOCK_WITH_RESET": "Interactive block with reset",
    "BLOCK_RESET_INTERACTIVE": "Interactive block with reset",
}


def explicit_bool(value):
    """Return only explicit booleans; never infer false from missing/other values."""
    return value if isinstance(value, bool) else None


def policy_ref_enabled(fmc, rule, keys):
    ref = first_present(rule, *keys)
    if not isinstance(ref, dict):
        return None
    obj = fmc.resolve(ref) or ref
    return explicit_bool(obj.get("enabled"))


def time_range_status(fmc, rule):
    """Use explicit status fields from the rule or its read-only resolved time range."""
    key = next((k for k in ("timeRange", "timeRangeObject") if k in rule), None)
    if key is None:
        return "unknown"
    ref = rule.get(key)
    if ref in (None, {}, []):
        return "none"
    obj = (fmc.resolve(ref) or ref) if isinstance(ref, dict) else {}
    if not isinstance(obj, dict):
        return "unknown"
    # These are authoritative only when explicitly returned by FMC. We do not
    # infer expiration from guessed date fields or from the mere presence of a ref.
    expired = explicit_bool(obj.get("expired"))
    if expired is True:
        return "expired"
    if expired is False:
        return "available"
    status = str(obj.get("status", "")).strip().lower()
    if status in ("expired", "inactive"):
        return "expired"
    if status in ("available", "active"):
        return "available"
    return "unknown"


def counted_value(known_count, unknown_count):
    """Show a plain numeric count; fields without a reported value contribute zero."""
    return known_count


def summarize_policy(fmc, rules):
    total = len(rules)
    enabled = disabled = enabled_unknown = 0
    actions = {label: 0 for label in ACTION_LABELS.values()}
    unknown_actions = 0
    ips_yes = ips_unknown = file_yes = file_unknown = 0
    logging_yes = logging_no = logging_unknown = 0
    time_available = time_expired = time_unknown = 0

    for rule in rules:
        state = explicit_bool(rule.get("enabled")) if "enabled" in rule else None
        if state is True:
            enabled += 1
        elif state is False:
            disabled += 1
        else:
            enabled_unknown += 1

        action = str(rule.get("action", "")).strip().upper().replace(" ", "_")
        label = ACTION_LABELS.get(action)
        if label:
            actions[label] += 1
        else:
            unknown_actions += 1

        ips = policy_ref_enabled(fmc, rule, ("ipsPolicy", "intrusionPolicy"))
        if ips is True:
            ips_yes += 1
        elif ips is None:
            ips_unknown += 1
        file_state = policy_ref_enabled(fmc, rule, ("filePolicy",))
        if file_state is True:
            file_yes += 1
        elif file_state is None:
            file_unknown += 1

        log_keys = ("logBegin", "logEnd")
        log_values = [explicit_bool(rule.get(k)) for k in log_keys if k in rule]
        if any(v is True for v in log_values):
            logging_yes += 1
        elif len(log_values) == len(log_keys) and all(v is False for v in log_values):
            logging_no += 1
        else:
            logging_unknown += 1

        tr = time_range_status(fmc, rule)
        if tr == "available":
            time_available += 1
        elif tr == "expired":
            time_expired += 1
        elif tr == "unknown":
            time_unknown += 1

    rows = [
        ("Rules", "Total access control rules", total),
        ("Rules", "Enabled rules", counted_value(enabled, enabled_unknown)),
        ("Rules", "Disabled rules", counted_value(disabled, enabled_unknown)),
        ("Actions", "Allow", counted_value(actions["Allow"], unknown_actions)),
        ("Actions", "Block", counted_value(actions["Block"], unknown_actions)),
        ("Actions", "Trust", counted_value(actions["Trust"], unknown_actions)),
        ("Actions", "Monitor", counted_value(actions["Monitor"], unknown_actions)),
        ("Actions", "Block with reset", counted_value(actions["Block with reset"], unknown_actions)),
        ("Actions", "Interactive block", counted_value(actions["Interactive block"], unknown_actions)),
        ("Actions", "Interactive block with reset", counted_value(actions["Interactive block with reset"], unknown_actions)),
        ("Actions", "Unrecognized / missing action", unknown_actions),
        ("Security", "IPS policy enabled", counted_value(ips_yes, ips_unknown)),
        ("Security", "File policy enabled", counted_value(file_yes, file_unknown)),
        ("Logging", "Logging enabled", counted_value(logging_yes, logging_unknown)),
        ("Time range", "Time range available", counted_value(time_available, time_unknown)),
        ("Time range", "Time range expired", counted_value(time_expired, time_unknown)),
    ]
    return rows


def create_index_sheet(wb, policy_name, used_names, rows=None, issue=None):
    ws = wb.create_sheet(excel_sheet_name(f"{policy_name} Index", used_names))
    ws.append(["Access Control Policy", policy_name])
    ws.append(["Category", "Metric", "Count / status"])
    if rows is None:
        ws.append(["Status", "Policy counts", "Unknown - rules could not be retrieved"] )
        if issue:
            ws.append(["Status", "Reason", clean_cell(issue)])
    else:
        for row in rows:
            ws.append(list(row))
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:C{ws.max_row}"
    ws.row_dimensions[1].height = 28
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", size=12)
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for cell in ws[2]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=3):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 38
    return ws


def main():
    global VERIFY_TLS
    print("\nCisco FMC - policy rule worksheets and per-policy index sheets")
    print("Expanded network groups and port groups where FMC API permits.")
    print("Read-only: no configuration changes are made.\n")

    host = input("FMC hostname/IP (example: 192.0.2.10): ").strip()
    username = input("FMC API username: ").strip()
    password = getpass.getpass("FMC API password: ")
    tls = input("Verify FMC TLS certificate? [Y/n]: ").strip().lower()
    VERIFY_TLS = tls not in ("n", "no")
    if not VERIFY_TLS:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    try:
        fmc = FMC(host, username, password, VERIFY_TLS)
    except Exception as exc:
        sys.exit(f"\nFMC connection/authentication failed:\n{exc}")

    policies_url = f"{fmc.root}/policy/accesspolicies"
    print(f"\nConnected. Domain UUID: {fmc.domain_uuid}")
    print("Retrieving access control policies...")
    try:
        policies = fmc.get_all(policies_url, {"expanded": "true"})
    except Exception as exc:
        sys.exit(f"Could not retrieve policies: {exc}")
    if not policies:
        sys.exit("No access control policies returned. Check API permissions.")

    wb = Workbook()
    default = wb.active
    wb.remove(default)
    used_names = set()
    summary = wb.create_sheet("Export Summary")
    summary.append(["Policy Name", "Rules exported", "Excel worksheet"])
    errors = []

    for p_index, policy in enumerate(policies, 1):
        policy_name = policy.get("name") or f"Policy_{p_index}"
        policy_id = policy.get("id")
        sheet_name = excel_sheet_name(policy_name, used_names)
        ws = wb.create_sheet(sheet_name)
        ws.append(HEADERS)

        if not policy_id:
            errors.append(f"{policy_name}: policy response did not contain an ID")
            create_index_sheet(wb, policy_name, used_names, issue="Policy response did not contain an ID")
            summary.append([policy_name, 0, sheet_name])
            style_sheet(ws)
            continue

        print(f"[{p_index}/{len(policies)}] {policy_name}")
        rules_url = f"{policies_url}/{policy_id}/accessrules"
        try:
            rules = fmc.get_all(rules_url, {"expanded": "true"})
        except Exception as exc:
            errors.append(f"{policy_name}: {exc}")
            create_index_sheet(wb, policy_name, used_names, issue=str(exc))
            summary.append([policy_name, 0, sheet_name])
            style_sheet(ws)
            continue

        for rule in rules:
            try:
                row_data = build_rule_row(fmc, rule)
                ws.append([clean_cell(row_data.get(col, "")) for col in HEADERS])
            except Exception as exc:
                errors.append(f"{policy_name} / {rule.get('name', 'unknown rule')}: {exc}")
                # Keep a row for the rule even if one condition could not be expanded.
                fallback = {h: "" for h in HEADERS}
                fallback["Name"] = rule.get("name", "")
                fallback["Action"] = rule.get("action", "")
                fallback["Comments"] = "Export warning: " + str(exc)
                ws.append([clean_cell(fallback.get(col, "")) for col in HEADERS])

        summary.append([policy_name, len(rules), sheet_name])
        create_index_sheet(wb, policy_name, used_names, summarize_policy(fmc, rules))
        style_sheet(ws)
        print(f"    {len(rules)} rules exported to worksheet '{sheet_name}'")

    style_sheet(summary)
    if errors:
        err_ws = wb.create_sheet("Export Errors")
        err_ws.append(["Error"])
        for err in errors:
            err_ws.append([clean_cell(err)])
        style_sheet(err_ws)

    filename = "FMC_ACP_Rules_By_Policy_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
    try:
        wb.save(filename)
    except PermissionError:
        sys.exit(f"\nCannot save {filename}. Close the workbook if it is already open and run again.")
    print("\nEXPORT COMPLETE")
    print(f"Policies: {len(policies)}")
    print(f"Workbook: {filename}")
    print("Each access control policy has its own rule worksheet and index worksheet.")
    print("Review Export Summary and Export Errors (if present).")
    print("Index values include unknown counts when FMC omits required status fields. Time ranges are classified only from explicit FMC status fields.")
    print("Compare rule totals with FMC. Group expansion depends on object API permissions and object types.")


if __name__ == "__main__":
    main()
