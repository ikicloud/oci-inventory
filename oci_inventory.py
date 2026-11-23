#!/usr/bin/env python3
"""Generate Excel and HTML inventory reports for an OCI tenancy.

Examples:
  python3 oci_inventory.py --tenancy-id <TENANCY_OCID>
  python3 oci_inventory.py --profile CLIENT --region eu-milan-1
  python3 oci_inventory.py --input export.zip --customer CLIENT

For offline NSG rules, use <NSG_OCID>/rules.json or JSON with
"nsg-id" and "data" fields.
"""
import argparse
import collections
import configparser
from functools import lru_cache
import hashlib
import html
from itertools import chain, groupby
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from xml.etree import ElementTree as ET

GLOBAL = {'Tenancy', 'Compartment', 'User', 'Group', 'Policy', 'DynamicGroup',
          'TagDefault', 'TagNamespace', 'Domain', 'IdentityDomain'}
DETAILS = {
    'Instance': ('compute instance get', 'instance-id'),
    'Vnic': ('network vnic get', 'vnic-id'),
    'PrivateIp': ('network private-ip get', 'private-ip-id'),
    'PublicIp': ('network public-ip get', 'public-ip-id'),
    'Vcn': ('network vcn get', 'vcn-id'),
    'Subnet': ('network subnet get', 'subnet-id'),
    'RouteTable': ('network route-table get', 'rt-id'),
    'SecurityList': ('network security-list get', 'security-list-id'),
    'NetworkSecurityGroup': ('network nsg get', 'nsg-id'),
    'DHCPOptions': ('network dhcp-options get', 'dhcp-id'),
    'InternetGateway': ('network internet-gateway get', 'ig-id'),
    'NatGateway': ('network nat-gateway get', 'nat-gateway-id'),
    'ServiceGateway': ('network service-gateway get', 'service-gateway-id'),
    'DrgAttachment': ('network drg-attachment get', 'drg-attachment-id'),
    'LocalPeeringGateway': ('network local-peering-gateway get', 'local-peering-gateway-id'),
    'RemotePeeringConnection': ('network remote-peering-connection get', 'remote-peering-connection-id'),
    'LoadBalancer': ('lb load-balancer get', 'load-balancer-id'),
    'AutonomousDatabase': ('db autonomous-database get', 'autonomous-database-id'),
    'DbSystem': ('db system get', 'db-system-id'),
    'BootVolume': ('bv boot-volume get', 'boot-volume-id'),
    'Volume': ('bv volume get', 'volume-id'),
    'VolumeBackupPolicy': ('bv volume-backup-policy get', 'policy-id'),
    'ResourceSchedule': ('resource-scheduler schedule get', 'schedule-id'),
}
# Exclude instance user data and client configuration.
FIELDS = set('''id identifier resource-type compartment-id display-name name lifecycle-state
availability-domain vcn-id subnet-id subnet-ids nsg-ids security-list-ids route-table-id
dhcp-options-id private-ip public-ip ip-address ip-addresses cidr-block cidr-blocks ipv6-cidr-blocks
is-primary is-private is-enabled prohibit-public-ip-on-vnic block-traffic nat-ip lifetime
assigned-entity-id assigned-entity-type private-ip-id vnic-id hostname-label dns-label
vcn-domain-name drg-id peer-id network-id network-details drg-route-table-id shape shape-config
image-id source-details boot-volume-id volume-id instance-id size-in-gbs unique-size-in-gbs
vpus-per-gb time-created expiration-time source-type type asset-id policy-id schedules
policy-name route-rules ingress-security-rules egress-security-rules options services
direction protocol source source-type destination destination-type is-stateless tcp-options
udp-options icmp-options description url home-region domain-type statements time-last-run
last-run-status action resources topic-id target-resource-id scope view-id attached-vcn-id
default-view-id attached-views zone-type peer-region-name network-entity-id
db-name db-workload private-endpoint-ip private-endpoint-label cpu-core-count
compute-count compute-model db-version hostname backend-sets listeners'''.split())
LINKS = {
    'vcn-id':'VCN', 'subnet-id':'Subnet', 'subnet-ids':'Subnet',
    'attached-vcn-id':'VCN del resolver', 'default-view-id':'View predefinita',
    'view-id':'View DNS', 'nsg-ids':'Network Security Group',
    'security-list-ids':'Security List', 'dhcp-options-id':'DHCP Options',
    'route-table-id':'Route table', 'drg-id':'DRG', 'drg-route-table-id':'DRG route table',
    'network-id':'Rete collegata', 'peer-id':'Peering', 'vnic-id':'VNIC',
    'private-ip-id':'IP privato', 'assigned-entity-id':'Entità assegnata',
    'instance-id':'Istanza', 'volume-id':'Volume di origine', 'boot-volume-id':'Boot volume di origine',
    'image-id':'Immagine', 'asset-id':'Volume protetto', 'policy-id':'Policy di backup',
    'target-resource-id':'Target', 'source-boot-volume-backup-id':'Backup sorgente',
    'source-volume-backup-id':'Backup sorgente',
}
FIELDS.update(LINKS)
NETWORK_TYPES = set('''Vcn Subnet Vnic PrivateIp PublicIp RouteTable InternetGateway NatGateway
ServiceGateway Drg DrgAttachment DrgRouteTable RemotePeeringConnection LocalPeeringGateway
DHCPOptions NetworkSecurityGroup SecurityList CustomerDnsZone DnsZone DnsResolver DnsView
IpSecConnection VirtualCircuit LoadBalancer NetworkLoadBalancer'''.split())
GLOBAL_PREFIXES = {'policy','compartment','user','group','domain','dynamicgroup','tagdefault','tagnamespace','tenancy'}
SHEETS = ('Inventario', 'Compute e Storage', 'Networking', 'Security', 'Backup', 'IAM')
BACKUP_ASSIGNMENT_OP = 'bv volume-backup-policy-assignment get-volume-backup-policy-asset-assignment'
NSG_RULES_OP = 'network nsg rules list'

def response_list(document):
    rows = document.get('data') if isinstance(document, dict) else document
    if isinstance(rows, list):
        return rows
    raise ValueError('Expected a list of results')

def backup_assignment_rows(document, asset_id):
    rows = response_list(document)
    if (len(rows) > 1 or isinstance(document, dict) and document.get('opc-next-page')):
        raise ValueError('Ambiguous or incomplete backup policy assignments')
    if rows:
        row = rows[0]
        if (not isinstance(row, dict) or row.get('asset-id') != asset_id
                or not all(isinstance(row.get(k), str) and row[k].strip() for k in ('id', 'policy-id'))):
            raise ValueError('Incomplete assignment or mismatched volume')
    return rows

def is_nsg_rule(row):
    return (isinstance(row, dict) and isinstance(row.get('id'), str) and bool(row['id'])
            and row.get('direction') in ('INGRESS', 'EGRESS') and row.get('protocol') is not None
            and not row.get('resource-type'))

def compact(value):
    if value is None:
        return ''
    if isinstance(value, (list, dict)):
        return json.dumps(value, ensure_ascii=False, separators=(', ', ': '))
    return str(value)

def many(value):
    if isinstance(value, list):
        return value
    return [] if value is None else [value]

def payload(document):
    data = document.get('data', document) if isinstance(document, dict) else document
    if isinstance(data, dict) and isinstance(data.get('items'), list):
        return data['items']
    return many(data)

def region_of(identifier):
    parts = identifier.split('.')
    return parts[3] if len(parts) > 3 and parts[0] == 'ocid1' else ''

def global_kind(kind, oid):
    return kind in GLOBAL or (oid.startswith('ocid1.') and oid.split('.')[1] in GLOBAL_PREFIXES)

def clean_name(text):
    return re.sub(r'[^\w.-]+', '_', text, flags=re.UNICODE).strip('._') or 'tenancy'

def id_of(row):
    return row.get('identifier') or row.get('id') or row.get('topic-id')

class Inventory:
    def __init__(self):
        self.resources = {}
        self.edges = set()
        self.issues = []
        self.rules = []
        self.members = []
        self.assignment_results = {}
        self.empty_nsgs = set()
        self.unassigned_rules = {}
        self.regions = set()
        self.region_keys = {}
        self.roots = set()
        self.origin = ''
        self.customer = ''
        self.checked = collections.Counter()

    def add(self, kind, raw, region, source):
        oid = id_of(raw)
        if not isinstance(oid, str) or not oid:
            return None
        encoded_region = region_of(oid)
        if (kind in ('Instance', 'BootVolumeAttachment')
                and oid.startswith('ocid1.instance.')
                and raw.get('instance-id') == oid
                and raw.get('boot-volume-id')):
            kind = 'BootVolumeAttachment'
            native_id = raw.get('id') or ''
            if native_id.startswith('ocid1.bootvolumeattachment.'):
                oid = native_id
            else:
                oid = f"boot-attachment|{oid}|{raw['boot-volume-id']}"
        canonical = self.region_keys.get(encoded_region.lower(), encoded_region)
        if self.origin == 'OCI CLI' and region and canonical and '-' not in canonical:
            canonical = region
        scope = 'Globale' if global_kind(kind, oid) else canonical or region or ''
        key = oid if oid.startswith('ocid1.') else '|'.join((scope, raw.get('compartment-id') or '', kind, oid))
        res = self.resources.setdefault(key, {
            'id': oid, 'kind': kind, 'region': scope, 'data': {}, 'sources': set(),
        })
        res['sources'].add(source)
        extra = raw.get('additional-details') or {}
        for segment in (extra if isinstance(extra, dict) else {}, raw):
            for k, v in segment.items():
                if k in FIELDS and v is not None:
                    res['data'][k] = v
        if scope and scope != 'Globale':
            self.regions.add(scope)
        if kind == 'Tenancy':
            self.roots.add(oid)
        compartment = raw.get('compartment-id') or ''
        if compartment.startswith('ocid1.tenancy.'):
            self.roots.add(compartment)
        return key

    def link(self, a, b, relation, detail=''):
        if a and isinstance(b, str) and b and a != b:
            self.edges.add((a, b, relation, str(detail or '')))

    def name(self, key):
        res = self.resources.get(key, {})
        row = res.get('data', {})
        return row.get('display-name') or row.get('name') or res.get('id') or key or ''

    def compartment(self, key):
        if not key:
            return ''
        path, seen = [], set()
        while key and key not in seen:
            seen.add(key)
            if key in self.roots:
                path.append(self.name(key) if key in self.resources else (self.customer or key))
                break
            path.append(self.name(key))
            key = self.resources.get(key, {}).get('data', {}).get('compartment-id')
        return ' / '.join(reversed(path))

    def build_links(self):
        for key, res in self.resources.items():
            row = res['data']
            self.link(key, row.get('compartment-id'), 'Compartment')
            for field, relation in LINKS.items():
                for target in many(row.get(field)):
                    self.link(key, target, relation)
            for item in row.get('attached-views') or []:
                self.link(key, item.get('view-id'), 'View DNS associata')
            network = row.get('network-details') or {}
            self.link(key, network.get('id'), 'Rete collegata')
            for route in row.get('route-rules') or []:
                dest = route.get('destination') or route.get('cidr-block')
                self.link(key, route.get('network-entity-id'), 'Next hop', dest)
            for target in row.get('resources') or []:
                if isinstance(target, dict):
                    self.link(key, target.get('id'), 'Target schedule')
            if res['kind'] == 'Instance':
                source = row.get('source-details') or {}
                self.link(key, source.get('image-id'), 'Immagine')
            for field, direction in [('ingress-security-rules','INGRESS'),('egress-security-rules','EGRESS')]:
                for rule in row.get(field) or []:
                    self.rules.append((key, direction, rule))
            if row.get('lifecycle-state') == 'ATTACHED':
                vm = row.get('instance-id')
                if res['kind'] == 'VnicAttachment':
                    self.link(vm, row.get('vnic-id'), 'Interfaccia collegata', res['id'])
                if res['kind'] == 'BootVolumeAttachment':
                    self.link(vm, row.get('boot-volume-id'), 'Boot volume collegato', res['id'])
                if res['kind'] == 'VolumeAttachment':
                    self.link(vm, row.get('volume-id'), 'Block volume collegato', res['id'])
        for key, member in self.members:
            self.link(member, key, 'Membro del gruppo')
        for key, direction, rule in self.rules:
            peer = rule.get('source') if direction == 'INGRESS' else rule.get('destination')
            if isinstance(peer, str) and peer.startswith('ocid1.'):
                self.link(key, peer, 'Riferimento in regola '+direction, rule.get('id'))

class MarioAnimation:
    """Terminal progress animation."""
    SPRITE = (
        '  RRRR  ',
        ' RRRRRR ',
        ' KKSSK  ',
        ' KSSSSK ',
        '  RSSR  ',
        ' SRBBRS ',
        '  BBBB  ',
        '  K  K  ',
    )
    COLORS = {'R': 196, 'S': 223, 'K': 94, 'B': 27, 'G': 220,
              'P': 34, 'D': 130, 'W': 255}
    WIDTH = 28
    HEIGHT = 12

    LOGO_WIDE = (
        "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW",
        "WWWWWWWWWWLMMSSMLWWWWWWWWWWWWWWWWW",
        "WWWWWWWWWHSSHHHHSMWWWWWWWWWWWWWWWW",
        "WWWWWWWWHBHWWWWWLMHLHMHLWWWWWWWWWW",
        "WWWWWWWLDHWWWWWWWWMSMMMSBLWWWWWWWW",
        "WWWWWWWHMWWWWWWWWWLHWWWWMBLWWWWWWW",
        "WWWWWLLSHWWWWWWWWWWWWWWWWBMWWWWWWW",
        "WWWLMMMSHWWWWWWWWWWWWWWWWHMLLWWWWW",
        "WWLSHLWWWWWWWWWWWWWWWWWWWHSSBBHWWW",
        "WWSMWWWWWWWWWWWWWWWWWWWWWWLWLHDMWW",
        "WHDWWWWWWWWWWWWWWWWWWWWWWWWWWWLDHW",
        "WMMWWWWWWWWWWWWWWWWWWWWWWWWWWWWMBW",
        "WSHWWWWWWWWWWWWWWWWWWWWWWWWWWWWHDW",
        "WMBWWWWWWWWWWWWWWWWWWWWWWWWWWWWHSW",
        "WLDHWWWWWWWWWWWWWWWWWWWWWWWWWWWBHW",
        "WWHDHWWWWWWWWWWWWWWWWWWWWWWWWLSBWW",
        "WWWHBSHHHHHHHHHHHHHHHHHHHHHHMBSLWW",
        "WWWWLHSBSSSSSSSSSSSSSSSSSSBBSHWWWW",
    )
    LOGO_COMPACT = (
        "WWWWWWLHLWWWWWWWWWW",
        "WWWWWHHLHHLWLWWWWWW",
        "WWWWHHWWWWMHHMLWWWW",
        "WWWLMWWWWWWWWLMWWWW",
        "WLMHLWWWWWWWWWHMHWW",
        "LMWWWWWWWWWWWWWLHMW",
        "HHWWWWWWWWWWWWWWWHH",
        "LMWWWWWWWWWWWWWWWHH",
        "WMHWWWWWWWWWWWWWLMW",
        "WWHMMMMMMMMMMMMMMLW",
    )
    LOGO_WORDMARK = (
        "   ▘▟   ▝         ▌           ▐   ",
        "   ▌█▄▞▘▐▌▄▄▄▖▞▀▀▘▌▟▀▀▙▐  ▐▐▀▀▜   ",
        "   ▌▜ ▝▄▐▘    ▜▄▄▘▌▝▄▄▘▝▄▖▜▝▄▄▜   ",
    )
    LOGO_COLORS = {'W': 231, 'L': 252, 'H': 247, 'M': 244,
                   'S': 241, 'B': 60, 'D': 237}
    LOGO_GAP = 4

    def __init__(self, enabled, operation, region, checked, resources, started):
        self.stream = sys.stderr
        size = self.terminal_size()
        self.enabled = (enabled and self.stream.isatty()
                        and os.environ.get('TERM', 'dumb') != 'dumb'
                        and size.columns >= 60 and size.lines >= 16)
        try:
            ('▀' + ''.join(self.LOGO_WORDMARK)).encode(getattr(self.stream, 'encoding', None) or 'ascii')
            self.half_blocks = True
        except (LookupError, UnicodeEncodeError):
            self.half_blocks = False
        self.operation = operation
        self.region = region or 'default'
        self.checked = checked
        self.resources = resources
        self.started = started
        self.stop = threading.Event()
        self.thread = None
        self.drawn = 0

    def terminal_size(self):
        try:
            return os.get_terminal_size(self.stream.fileno())
        except (AttributeError, OSError, ValueError):
            return shutil.get_terminal_size((80, 24))

    @classmethod
    @lru_cache(maxsize=4)
    def logo_lines(cls, raster, half_blocks):
        lines = []
        for upper, lower in zip(raster[::2], raster[1::2]):
            parts = []
            for (top, bottom), cells in groupby(zip(upper, lower)):
                count = sum(1 for _ in cells)
                if top == bottom or not half_blocks:
                    ink = bottom if top == 'W' else top
                    style = f'48;5;{cls.LOGO_COLORS[ink]}'
                    glyph = ' '
                else:
                    style = f'38;5;{cls.LOGO_COLORS[top]};48;5;{cls.LOGO_COLORS[bottom]}'
                    glyph = '▀'
                parts.append(f'\x1b[{style}m' + glyph * count)
            lines.append(''.join(parts) + '\x1b[0m')
        if raster == cls.LOGO_WIDE and half_blocks:
            lettering = cls.LOGO_WORDMARK
        else:
            lettering = ('iki-cloud'.center(len(raster[0])),)
        for row in lettering:
            lines.append('\x1b[1;38;5;23;48;5;231m' + row + '\x1b[0m')
        return tuple(lines)

    def frame(self):
        size = self.terminal_size()
        if size.columns < 60 or size.lines < 16:
            return []
        cols = size.columns
        available = cols - 1 - self.LOGO_GAP
        raster = self.LOGO_COMPACT
        if available >= self.WIDTH * 2 + len(self.LOGO_WIDE[0]):
            raster = self.LOGO_WIDE
        width = min(self.WIDTH, (available - len(raster[0])) // 2)
        logo = self.logo_lines(raster, self.half_blocks)
        logo_top = (self.HEIGHT - len(logo)) // 2
        elapsed = time.monotonic() - self.started
        tick = int(elapsed * 9) % 24
        jump = (0, 0, 0, 0, 0, 0, 1, 2, 3, 3, 2, 1,
                0, 0, 0, 0, 0, 1, 2, 3, 2, 1, 0, 0)[tick]
        travel = width - len(self.SPRITE[0]) - 3
        x = 1 + tick * travel // 23
        y = 3 - jump
        board = [[' '] * width for _ in range(self.HEIGHT)]
        def put(xx, yy, color):
            if 0 <= xx < width and 0 <= yy < self.HEIGHT:
                board[yy][xx] = color
        for xx in range(width):
            put(xx, 11, 'D' if xx % 2 else 'P')
        for full_width_x in (10, 18, 25):
            cx = round(full_width_x * width / self.WIDTH)
            if abs((x + 4) - cx) <= 3 and jump >= 2:
                continue
            put(cx, 0, 'G'); put(cx, 1, 'G'); put(cx, 2, 'G')
            if tick % 4 < 2:
                put(cx + 1, 1, 'G')
        for dy, row in enumerate(self.SPRITE):
            for dx, color in enumerate(row):
                if color != ' ':
                    put(x + dx, y + dy, color)
        def pixels(row):
            return ''.join('  ' if c == ' ' else
                           '\x1b[48;5;'+str(self.COLORS[c])+'m  \x1b[0m'
                           for c in row)
        mins, sec = divmod(int(elapsed), 60)
        status = f'Resources: {self.resources} | Successful API calls: {self.checked} | {mins:02d}:{sec:02d}'
        lines = [f'OCI INVENTORY  -  {self.region}']
        for index, row in enumerate(board):
            line = pixels(row)
            logo_row = index - logo_top
            if 0 <= logo_row < len(logo):
                line += ' ' * self.LOGO_GAP + logo[logo_row]
            lines.append(line)
        lines += [status[:max(1, cols-1)], ('Running: '+self.operation)[:max(1, cols-1)]]
        lines[0] = lines[0][:max(1, cols-1)]
        return lines

    def draw(self):
        lines = self.frame()
        if self.drawn:
            self.stream.write(f'\x1b[{self.drawn}A')
            if not lines:
                self.stream.write('\r\x1b[0m\x1b[J')
        self.stream.write(''.join('\r\x1b[2K'+line+'\n' for line in lines))
        self.stream.flush()
        self.drawn = len(lines)

    def animate(self):
        while not self.stop.wait(0.11):
            try:
                self.draw()
            except OSError:
                return

    def __enter__(self):
        if self.enabled:
            self.draw()
            self.thread = threading.Thread(target=self.animate, daemon=True)
            self.thread.start()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.stop.set()
        if self.thread is not None:
            self.thread.join()
        if self.drawn:
            self.stream.write(f'\x1b[{self.drawn}A' +
                              '\r\x1b[0m\x1b[J')
            self.stream.flush()
        return False


class Cli:
    def __init__(self, args, inv):
        self.base = ['oci', '--output', 'json']
        if args.profile:
            self.base += ['--profile', args.profile]
        if args.config_file:
            self.base += ['--config-file', str(args.config_file)]
        if args.auth:
            self.base += ['--auth', args.auth]
        self.inv = inv
        self.animation = not getattr(args, "no_animation", False)
        self.started = time.monotonic()

    def run(self, operation, region='', required=False, **options):
        args = self.base + (['--region', region] if region else []) + operation.split()
        for key, value in options.items():
            if value is True:
                args.append('--'+key.replace('_','-'))
            elif value is not None and value is not False:
                args.extend(['--'+key.replace('_','-'), str(value)])
        with MarioAnimation(self.animation, operation, region,
                            sum(self.inv.checked.values()), len(self.inv.resources), self.started):
            job = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if job.returncode:
            match = re.search(r'"code"\s*:\s*"([^"]+)"', job.stderr)
            lines = job.stderr.strip().splitlines()
            detail = match.group(1) if match else (lines[-1][:180] if lines else f'Exit {job.returncode}')
            scope = ', '.join(str(v) for k,v in options.items() if k.endswith('_id'))
            self.inv.issues.append((region, operation, scope, detail))
            if required:
                raise ValueError(f'{operation}: {detail}')
            return None
        # Empty CLI output does not confirm a complete response.
        scope = ', '.join(str(v) for k, v in options.items() if k.endswith('_id'))
        output = job.stdout.strip()
        list_operation = operation.split()[-1] in ('list', 'list-users')
        if not output and list_operation:
            self.inv.issues.append((region, operation, scope,
                'Exit 0, empty stdout: no records returned; verify if unexpected'))
            if required:
                raise ValueError(f'{operation}: empty stdout, coverage not verified')
            return None
        try:
            response = json.loads(output)
            if not isinstance(response, (dict, list)):
                raise ValueError('Unexpected JSON structure')
            if operation == BACKUP_ASSIGNMENT_OP:
                backup_assignment_rows(response, options['asset_id'])
            elif operation == NSG_RULES_OP:
                if not all(is_nsg_rule(row) for row in response_list(response)):
                    raise ValueError('Incomplete NSG rule')
                if isinstance(response, dict) and response.get('opc-next-page'):
                    raise ValueError('Incomplete NSG rule list after --all')
            elif operation == 'search resource structured-search':
                data = response.get('data', response) if isinstance(response, dict) else None
                items = data.get('items') if isinstance(data, dict) else data
                if not isinstance(items, list):
                    raise ValueError('Missing Resource Search data.items field')
                if not all(isinstance(row, dict) and isinstance(id_of(row), str) and id_of(row) for row in items):
                    raise ValueError('Resource Search item without an identifier')
        except ValueError as error:
            detail = ('Exit 0, empty stdout' if not output else
                      f'Exit 0, invalid or unexpected JSON ({len(output)} characters): {error}')
            self.inv.issues.append((region, operation, scope, detail))
            print(f'[WARN] {region} {operation} [{scope}]: {detail}',
                  file=sys.stderr, flush=True)
            if required:
                raise ValueError(f'{operation}: {detail}') from error
            return None
        self.inv.checked[(region,operation)] += 1
        return response

    def rows(self, op, region='', **kw):
        result = self.run(op, region, **kw)
        return payload(result) if result is not None else []

def configured_tenancy(args):
    if args.tenancy_id:
        return args.tenancy_id
    for key in ('OCI_CLI_TENANCY', 'OCI_TENANCY'):
        value = os.environ.get(key,'')
        if value.startswith('ocid1.tenancy.'):
            return value
    config = configparser.ConfigParser(interpolation=None)
    path = args.config_file or os.environ.get('OCI_CLI_CONFIG_FILE','~/.oci/config')
    config.read(os.path.expanduser(path))
    profile = args.profile or os.environ.get('OCI_CLI_PROFILE','DEFAULT')
    return config.get(profile,'tenancy',fallback='')

def collect_live(args, inv):
    inv.origin = 'OCI CLI'
    if not shutil.which('oci'):
        raise ValueError('OCI CLI not found. Use OCI Cloud Shell or --input.')
    tenancy = configured_tenancy(args)
    if not tenancy.startswith('ocid1.tenancy.'):
        raise ValueError('Set --tenancy-id to the tenancy OCID.')
    cli = Cli(args, inv)
    initial = args.region[0] if args.region else ''
    tenancy_rows = cli.rows('iam tenancy get', initial, tenancy_id=tenancy)
    if tenancy_rows and tenancy_rows[0].get('id') != tenancy:
        raise ValueError('IAM response does not match the requested tenancy.')
    inv.add('Tenancy', tenancy_rows[0] if tenancy_rows else {'id':tenancy}, '', 'IAM')
    subscriptions = cli.rows('iam region-subscription list', initial, tenancy_id=tenancy, all=True)
    inv.region_keys.update({r['region-key'].lower(): r['region-name'] for r in subscriptions
                            if r.get('region-key') and r.get('region-name')})
    regions = list(dict.fromkeys(args.region or sorted(r['region-name'] for r in subscriptions if r.get('region-name') and r.get('status','READY')=='READY')))
    if not regions:
        raise ValueError('No regions found. Specify --region.')
    home = next((r['region-name'] for r in subscriptions if r.get('is-home-region')), regions[0])
    compartments = cli.rows('iam compartment list', home, compartment_id=tenancy, compartment_id_in_subtree='true', access_level='ACCESSIBLE', all=True)
    for c in compartments:
        inv.add('Compartment', c, '', 'IAM')
    scopes = {tenancy} | {c['id'] for c in compartments if c.get('lifecycle-state')=='ACTIVE'}
    for region in regions:
        print(f'[{region}] Resource Search', file=sys.stderr, flush=True)
        inv.regions.add(region)
        page, tokens = None, set()
        while True:
            reply = cli.run('search resource structured-search', region, query_text='query all resources', tenant_id=tenancy, page=page)
            if reply is None:
                break
            for row in payload(reply):
                kind = row.get('resource-type') or 'Resource'
                inv.add(kind, row, region, 'Resource Search')
                if row.get('compartment-id'):
                    scopes.add(row['compartment-id'])
            page = reply.get('opc-next-page')
            if not page:
                break
            if page in tokens:
                inv.issues.append((region,'Resource Search','','Token di paginazione ripetuto'))
                break
            tokens.add(page)
    print(f'[{home}] IAM: users, groups, domains and policies', file=sys.stderr, flush=True)
    for kind, op in [('User','iam user list'),('Group','iam group list'),('DynamicGroup','iam dynamic-group list')]:
        for row in cli.rows(op, home, compartment_id=tenancy, all=True):
            inv.add(kind, row, '', op)
    for compartment in sorted(scopes):
        for kind, op in [('Domain','iam domain list'),('Policy','iam policy list')]:
            for row in cli.rows(op, home, compartment_id=compartment, all=True):
                inv.add(kind, row, '', op)
    for oid, res in list(inv.resources.items()):
        if res['kind']=='Policy' and 'statements' not in res['data']:
            for policy in cli.rows('iam policy get',home,policy_id=oid):
                inv.add('Policy',policy,'','iam policy get')
        if res['kind']=='Group':
            for user in cli.rows('iam group list-users',home,group_id=oid,all=True):
                key = inv.add('User',user,'','IAM group list-users')
                if key:
                    inv.members.append((oid,key))
    for region in regions:
        print(f'[{region}] compute, networking and backups', file=sys.stderr, flush=True)
        for compartment in sorted(scopes):
            for kind, op in [('BootVolumeBackup','bv boot-volume-backup list'),('VolumeBackup','bv backup list')]:
                for row in cli.rows(op,region,compartment_id=compartment,all=True):
                    inv.add(kind, row, region, op)
        targets = [(k,v) for k,v in inv.resources.items() if v['region']==region]
        for key, res in targets:
            kind = res['kind']
            if kind in DETAILS and key.startswith('ocid1.'):
                op, arg = DETAILS[kind]
                for row in cli.rows(op,region,**{arg:key}):
                    inv.add(kind,row,region,op)
            if kind=='NetworkSecurityGroup':
                result = cli.run(NSG_RULES_OP, region, nsg_id=key, all=True)
                if result is not None:
                    rules = response_list(result)
                    if not rules:
                        inv.empty_nsgs.add(key)
                    inv.rules.extend((key, rule['direction'], rule) for rule in rules)
            if kind in ('Volume','BootVolume'):
                result = cli.run(BACKUP_ASSIGNMENT_OP, region, asset_id=key)
                if result is None:
                    inv.assignment_results[key] = 'Da verificare'
                else:
                    rows = response_list(result)
                    inv.assignment_results[key] = rows[0]['policy-id'] if rows else 'Non assegnata'
                    for row in rows:
                        inv.add('VolumeBackupPolicyAssignment', row, region, BACKUP_ASSIGNMENT_OP)
        vm_domains = collections.defaultdict(set)
        for _, resource in targets:
            row = resource['data']
            if resource['kind'] == 'Instance' and row.get('compartment-id'):
                domains = vm_domains[row['compartment-id']]
                if row.get('availability-domain'):
                    domains.add(row['availability-domain'])
        for compartment, domains in sorted(vm_domains.items()):
            for kind,op in [('VnicAttachment','compute vnic-attachment list'),('VolumeAttachment','compute volume-attachment list')]:
                for row in cli.rows(op,region,compartment_id=compartment,all=True):
                    inv.add(kind,row,region,op)
            for ad in sorted(domains):
                op='compute boot-volume-attachment list'
                for row in cli.rows(op,region,compartment_id=compartment,availability_domain=ad,all=True):
                    inv.add('BootVolumeAttachment',row,region,op)
        nic_ids = {v['data']['vnic-id'] for v in inv.resources.values() if v['region']==region and v['kind']=='VnicAttachment' and v['data'].get('vnic-id')}
        nic_ids.update(k for k,v in targets if v['kind']=='Vnic')
        policy_ids = {r['data']['policy-id'] for r in inv.resources.values() if r['region']==region
                      and r['kind']=='VolumeBackupPolicyAssignment' and r['data'].get('policy-id')}
        for policy_id in sorted(policy_ids - inv.resources.keys()):
            for row in cli.rows('bv volume-backup-policy get',region,policy_id=policy_id):
                inv.add('VolumeBackupPolicy',row,region,'bv volume-backup-policy get')
        for nic in sorted(nic_ids):
            if nic not in inv.resources or 'network vnic get' not in inv.resources[nic]['sources']:
                for row in cli.rows('network vnic get',region,vnic_id=nic):
                    inv.add('Vnic',row,region,'network vnic get')
            for row in cli.rows('network private-ip list',region,vnic_id=nic,all=True):
                inv.add('PrivateIp',row,region,'network private-ip list')
    if not any(op=='search resource structured-search' for _,op in inv.checked):
        inv.issues.append(('', 'Copertura Resource Search', '', 'Nessuna region interrogata con successo; dati limitati alle API di servizio riuscite.'))
    if len(inv.resources) == 1 and not any(op == 'search resource structured-search' for _,op in inv.checked):
        raise ValueError('No inventory collected. Check the OCI profile, tenancy, regions and permissions.')

TYPE_BY_OCID = {
    'instance':'Instance', 'vnic':'Vnic','privateip':'PrivateIp','publicip':'PublicIp',
    'vcn':'Vcn','subnet':'Subnet','routetable':'RouteTable','securitylist':'SecurityList',
    'networksecuritygroup':'NetworkSecurityGroup','dhcpoptions':'DHCPOptions',
    'internetgateway':'InternetGateway','natgateway':'NatGateway','servicegateway':'ServiceGateway',
    'drg':'Drg','drgattachment':'DrgAttachment','localpeeringgateway':'LocalPeeringGateway',
    'remotepeeringconnection':'RemotePeeringConnection','volume':'Volume','bootvolume':'BootVolume',
    'bootvolumebackup':'BootVolumeBackup','volumebackup':'VolumeBackup',
    'volumebackuppolicy':'VolumeBackupPolicy','volumebackuppolicyassign':'VolumeBackupPolicyAssignment',
    'bootvolumeattachment':'BootVolumeAttachment','volumeattachment':'VolumeAttachment','vnicattachment':'VnicAttachment',
    'compartment':'Compartment','tenancy':'Tenancy','domain':'Domain','policy':'Policy',
    'group':'Group','user':'User','dynamicgroup':'DynamicGroup','onstopic':'OnsTopic',
    'resourceschedule':'ResourceSchedule','cloudguardtarget':'CloudGuardTarget',
    'dns-zone':'CustomerDnsZone','dnsresolver':'DnsResolver','dnsview':'DnsView',
    'autonomousdatabase':'AutonomousDatabase','dbsystem':'DbSystem','loadbalancer':'LoadBalancer',
}

def load_exports(path, inv, region):
    imported_rules = {}

    def read_json(name, content):
        if not content.strip():
            inv.issues.append(('', 'Import JSON', name, 'File vuoto'))
            return
        try:
            doc = json.loads(content)
        except (ValueError, UnicodeError):
            inv.issues.append(('', 'Import JSON', name, 'JSON illeggibile'))
            return
        # Rule IDs are unique only within their NSG.
        envelope = doc if isinstance(doc, dict) else {}
        file_path = Path(name)
        parent = envelope.get('nsg-id') or envelope.get('network-security-group-id')
        if not parent and file_path.name.lower() == 'rules.json':
            parent = file_path.parent.name
        if not isinstance(parent, str) or not parent.startswith('ocid1.networksecuritygroup.'):
            parent = None
        if parent and envelope.get('opc-next-page'):
            inv.issues.append((region, 'Import regole NSG', name, 'Elenco paginato: verificare la copertura'))
        elif parent and (doc == [] or envelope.get('data') == []):
            inv.add('NetworkSecurityGroup', {'id': parent}, region, name)
            inv.empty_nsgs.add(parent)
        ignored = 0
        missing_owner = 0
        for row in payload(doc):
            if is_nsg_rule(row):
                owner = row.get('nsg-id') or row.get('network-security-group-id') or parent
                if isinstance(owner, str) and owner.startswith('ocid1.networksecuritygroup.'):
                    inv.add('NetworkSecurityGroup', {'id': owner}, region, name)
                    imported_rules[(owner, row['id'])] = row
                else:
                    inv.unassigned_rules[(name, region, row['id'])] = row
                    missing_owner += 1
                continue
            oid = id_of(row) if isinstance(row, dict) else None
            if not oid:
                ignored += 1
                continue
            ocid_type = oid.split('.')[1] if isinstance(oid,str) and oid.startswith('ocid1.') else ''
            kind = row.get('resource-type') or TYPE_BY_OCID.get(ocid_type, ocid_type or 'Resource')
            inv.add(kind,row,region,name)
        if missing_owner:
            inv.issues.append((region, 'Import regole NSG', name,
                f'{missing_owner} regole con NSG di appartenenza non identificato'))
        if ignored:
            inv.issues.append(('', 'Import JSON', name, f'{ignored} righe senza identificatore di risorsa'))
    if path.is_dir():
        for source in sorted(path.rglob('*.json')):
            read_json(str(source.relative_to(path)),source.read_bytes())
    elif zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as archive:
            for item in archive.infolist():
                if item.filename.lower().endswith('.json') and not item.is_dir():
                    read_json(item.filename,archive.read(item))
    else:
        name = str(Path(path.parent.name) / path.name) if path.parent.name.startswith('ocid1.networksecuritygroup.') else path.name
        read_json(name, path.read_bytes())
    inv.rules.extend((owner, rule['direction'], rule) for (owner, _), rule in imported_rules.items())
    if not inv.resources and not inv.unassigned_rules:
        raise ValueError('No recognizable resources found in the selected exports.')
    if len(inv.roots)>1:
        raise ValueError('Exports contain multiple tenancies. Process one tenancy at a time.')
    inv.origin = 'Export OCI JSON'
    inv.issues.append(('', 'Copertura offline', '', 'Solo i dati presenti nei JSON sono disponibili. Un campo vuoto indica dettaglio non esportato.'))

def rule_fields(direction, row):
    peer = row.get('source') if direction=='INGRESS' else row.get('destination')
    options = row.get('tcp-options') or row.get('udp-options') or {}
    def port(key):
        limits = options.get(key) or {}
        low, high = limits.get('min'),limits.get('max')
        if low == high:
            return compact(low)
        return f'{low}–{high}'
    icmp = row.get('icmp-options') or {}
    stateful = {True:'No',False:'Sì'}.get(row.get('is-stateless'),'')
    return [direction,row.get('protocol'),peer,port('source-port-range'),port('destination-port-range'),
            compact(icmp) if icmp else '',stateful,row.get('description'),row.get('id')]

HEADERS = {
    'Inventario':['Nome risorsa','Tipo OCI','Compartment','Region / ambito','Stato','IP privato','IP pubblico','CIDR','VCN','Subnet','Risorse associate','OCID / identificatore','Fonte'],
    'Compute e Storage':['Tipo','Nome','Compartment','Region','AD','Shape','OCPU','RAM (GB)','Dimensione (GB)','VPUs/GB','IP privato','IP pubblico','Relazioni','OCID / identificatore'],
    'Networking':['Tipo','Nome','Compartment','Region','VCN','Subnet','CIDR / IP','Associazione / destinazione','Next hop / valore','OCID collegato','OCID risorsa'],
    'Security':['Tipo','Nome','Compartment','Region','Direzione','Protocollo OCI','Sorgente / destinazione','Porta sorgente','Porta destinazione','ICMP','Stateful','Descrizione','ID regola','OCID risorsa'],
    'Backup':['Tipo','Nome','Compartment','Region','Volume / asset','Policy','Esecuzione','Tipo backup','Creato','Scadenza','GB','Unica (GB)','Stato','Pianificazione','OCID risorsa'],
    'IAM':['Tipo','Nome','Compartment','Ambito','Stato','Gruppo / endpoint','Statement / dettaglio','OCID risorsa'],
}
def spreadsheet_data(inv):
    compartment_path = lru_cache(maxsize=None)(inv.compartment)
    sheets = {sheet:[] for sheet in SHEETS}
    source_rows = collections.defaultdict(list)
    outgoing = collections.defaultdict(list)
    incoming = collections.defaultdict(list)
    for a,b,t,d in sorted(inv.edges):
        outgoing[a].append((t,b,d))
        incoming[b].append((t,a,d))
    def emit(sheet,key,values):
        if len(values)!=len(HEADERS[sheet]):
            raise ValueError('Numero di colonne errato: '+sheet)
        sheets[sheet].append(values)
        source_rows[key].append((sheet, len(sheets[sheet])+1))
    grouped_rules = collections.defaultdict(list)
    for key,direction,rule in inv.rules:
        grouped_rules[key].append(rule_fields(direction,rule))
    assignments = collections.defaultdict(list)
    vm_nics = collections.defaultdict(list)
    for key,res in inv.resources.items():
        row = res['data']
        if res['kind']=='VolumeBackupPolicyAssignment':
            assignments[row.get('asset-id')].append(row.get('policy-id'))
        if res['kind']=='VnicAttachment' and row.get('lifecycle-state')=='ATTACHED' and row.get('vnic-id'):
            vm_nics[row.get('instance-id')].append(row['vnic-id'])
    keys = sorted(inv.resources,key=lambda k:(inv.resources[k]['region'],compartment_path(inv.resources[k]['data'].get('compartment-id')),inv.resources[k]['kind'],inv.name(k).casefold(),k))
    def unique(values):
        return ', '.join(dict.fromkeys(str(v) for v in values if v))
    for key in keys:
        res = inv.resources[key]; row = res['data']
        kind,name,region,oid = res['kind'],inv.name(key),res['region'],res['id']
        comp = compartment_path(row.get('compartment-id'))
        state = row.get('lifecycle-state')
        edges = outgoing[key]
        assoc = unique(f'{t} → {inv.name(b)}' for t,b,d in edges if t!='Compartment')
        nics = [inv.resources[n]['data'] for n in vm_nics[key] if n in inv.resources]
        private = unique([row.get('private-ip'), row.get('ip-address') if kind=='PrivateIp' else '',row.get('private-endpoint-ip')]+[n.get('private-ip') for n in nics])
        public = unique([row.get('public-ip'), row.get('ip-address') if kind=='PublicIp' else '',row.get('nat-ip')]+[n.get('public-ip') for n in nics])
        cidr = unique(many(row.get('cidr-block'))+many(row.get('cidr-blocks'))+many(row.get('ipv6-cidr-blocks')))
        subnet_ids = many(row.get('subnet-id'))+many(row.get('subnet-ids'))+[n.get('subnet-id') for n in nics]
        vcn_ids = many(row.get('vcn-id'))+[inv.resources.get(s,{}).get('data',{}).get('vcn-id') for s in subnet_ids]
        subnet = unique(inv.name(s) for s in subnet_ids if s)
        vcn = unique(inv.name(v) for v in vcn_ids if v)
        emit('Inventario',key,[name,kind,comp,region,state,private,public,cidr,vcn,subnet,assoc,oid,', '.join(sorted(res['sources']))])
        if kind in ('Instance','BootVolume','Volume','BootVolumeAttachment','VolumeAttachment','VnicAttachment','Image'):
            shape = row.get('shape-config') or {}
            related = unique([assoc]+[f'{t} ← {inv.name(a)}' for t,a,d in incoming[key] if t!='Compartment'])
            emit('Compute e Storage',key,[kind,name,comp,region,row.get('availability-domain'),row.get('shape'),shape.get('ocpus'),shape.get('memory-in-gbs'),row.get('size-in-gbs'),row.get('vpus-per-gb'),private,public,related,oid])
        if kind in NETWORK_TYPES:
            base = [kind,name,comp,region,vcn,subnet,cidr or unique([private,public])]
            emit('Networking',key,base+['','','',oid])
            for t,b,d in edges:
                if t!='Compartment':
                    emit('Networking',key,base+[d if t=='Next hop' else t,inv.name(b),b,oid])
            for field in ('options','services'):
                for obj in row.get(field) or []:
                    emit('Networking',key,base+[field,compact(obj),'',oid])
        if kind in ('NetworkSecurityGroup','SecurityList') or kind.startswith(('CloudGuard','Vault','Certificate','SecurityZone','Waf','WebAppFirewall')):
            rules = grouped_rules.get(key)
            if not rules:
                description = None
                if kind == 'NetworkSecurityGroup':
                    description = 'Nessuna regola (elenco verificato)' if key in inv.empty_nsgs else 'Elenco regole da verificare'
                rules = [[None] * 7 + [description, None]]
            for values in rules:
                emit('Security',key,[kind,name,comp,region]+values+[oid])
        if kind in ('BootVolumeBackup','VolumeBackup','VolumeBackupPolicy','VolumeBackupPolicyAssignment','VolumeGroupBackup') or key in inv.assignment_results or key in assignments:
            policy = row.get('policy-id')
            policy_text = unique(inv.name(x) for x in assignments.get(key, ()) if x) or inv.name(policy) or row.get('policy-name','')
            if key in inv.assignment_results:
                policy_text = inv.name(inv.assignment_results[key])
            asset = row.get('asset-id') or row.get('volume-id') or row.get('boot-volume-id')
            schedules = row.get('schedules') or ['']
            for schedule in schedules:
                emit('Backup',key,[kind,name,comp,region,inv.name(asset),policy_text,row.get('source-type'),row.get('type'),row.get('time-created'),row.get('expiration-time'),row.get('size-in-gbs'),row.get('unique-size-in-gbs'),state,compact(schedule),oid])
        if kind in GLOBAL:
            groups = unique(inv.name(b) for t,b,d in edges if t=='Membro del gruppo')
            for statement in row.get('statements') or ['']:
                emit('IAM',key,[kind,name,comp,region,state,groups or row.get('url'),statement,oid])
    for (source, region, _), rule in sorted(inv.unassigned_rules.items()):
        values = rule_fields(rule['direction'], rule)
        sheets['Security'].append(['Regola NSG', 'NSG non identificato · ' + source, '', region, *values, ''])
    for region,op,scope,msg in inv.issues:
        sheets['Inventario'].append(['Raccolta dati','Nota di copertura',scope,region,'','','','','','',msg,'',op])
    return sheets,source_rows,outgoing,incoming

# Store strings as text to prevent spreadsheet formula execution.
XLS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CONTENT = 'http://schemas.openxmlformats.org/package/2006/content-types'
def xmlbytes(root):
    return ET.tostring(root,encoding='utf-8',xml_declaration=True)
def cell_column(n):
    result=''
    while n:
        n, rem=divmod(n-1,26)
        result=chr(65+rem)+result
    return result

def write_xlsx(path, sheets):
    workbook=ET.Element('workbook',{'xmlns':XLS,'xmlns:r':REL})
    book_sheets=ET.SubElement(workbook,'sheets')
    relationships=ET.Element('Relationships',xmlns=PKG)
    types=ET.Element('Types',xmlns=CONTENT)
    ET.SubElement(types,'Default',Extension='rels',ContentType='application/vnd.openxmlformats-package.relationships+xml')
    ET.SubElement(types,'Default',Extension='xml',ContentType='application/xml')
    prefix='application/vnd.openxmlformats-officedocument.spreadsheetml.'
    ET.SubElement(types,'Override',PartName='/xl/workbook.xml',ContentType=prefix+'sheet.main+xml')
    ET.SubElement(types,'Override',PartName='/xl/styles.xml',ContentType=prefix+'styles+xml')
    rootrels=ET.Element('Relationships',xmlns=PKG)
    ET.SubElement(rootrels,'Relationship',Id='rId1',Type=REL+'/officeDocument',Target='xl/workbook.xml')
    styles=f'''<styleSheet xmlns="{XLS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF17435A"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''
    with zipfile.ZipFile(path,'w',compression=zipfile.ZIP_DEFLATED,allowZip64=True) as z:
        z.writestr('_rels/.rels',xmlbytes(rootrels))
        z.writestr('xl/styles.xml',styles)
        ET.SubElement(relationships,'Relationship',Id='rs',Type=REL+'/styles',Target='styles.xml')
        for num,(name,rows) in enumerate(sheets.items(),1):
            header=HEADERS[name]
            if len(rows)>=1048576:
                raise ValueError(f'Il foglio {name} supera il limite di righe Excel.')
            ET.SubElement(book_sheets,'sheet',{'name':name,'sheetId':str(num),'r:id':f'rId{num}'})
            sheetpath=f'worksheets/sheet{num}.xml'
            ET.SubElement(relationships,'Relationship',Id=f'rId{num}',Type=REL+'/worksheet',Target=sheetpath)
            ET.SubElement(types,'Override',PartName='/xl/'+sheetpath,ContentType=prefix+'worksheet+xml')
            sheet=ET.Element('worksheet',xmlns=XLS)
            area=f'A1:{cell_column(len(header))}{len(rows)+1}'
            ET.SubElement(sheet,'dimension',ref=area)
            view=ET.SubElement(ET.SubElement(sheet,'sheetViews'),'sheetView',workbookViewId='0')
            ET.SubElement(view,'pane',ySplit='1',topLeftCell='A2',activePane='bottomLeft',state='frozen')
            cols=ET.SubElement(sheet,'cols')
            for j,title in enumerate(header,1):
                width=min(72,max(15,len(title)+4,max((len(compact(r[j-1]))+2 for r in rows[:250]),default=0)))
                ET.SubElement(cols,'col',min=str(j),max=str(j),width=str(width),customWidth='1')
            data=ET.SubElement(sheet,'sheetData')
            for i,values in enumerate(chain([header],rows),1):
                row=ET.SubElement(data,'row',r=str(i))
                for j,val in enumerate(values,1):
                    if val in (None,''):
                        continue
                    c=ET.SubElement(row,'c',r=f'{cell_column(j)}{i}')
                    if i==1:
                        c.set('s','1')
                    if type(val) in (int,float):
                        ET.SubElement(c,'v').text=str(val)
                    else:
                        c.set('t','inlineStr')
                        text=ET.SubElement(ET.SubElement(c,'is'),'t',{'xml:space':'preserve'})
                        value=re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', compact(val))
                        text.text=value if len(value)<=32767 else value[:32730]+' […] Dettaglio completo in HTML'
            ET.SubElement(sheet,'autoFilter',ref=area)
            z.writestr('xl/'+sheetpath,xmlbytes(sheet))
        z.writestr('xl/workbook.xml',xmlbytes(workbook))
        z.writestr('xl/_rels/workbook.xml.rels',xmlbytes(relationships))
        z.writestr('[Content_Types].xml',xmlbytes(types))

def esc(value):
    return html.escape(compact(value),quote=True)

def table(headers, rows):
    """Render an HTML table."""
    heads=''.join(f'<th>{esc(cell)}</th>' for cell in headers)
    body=''.join('<tr>'+''.join(f'<td>{cell}</td>' for cell in row)+'</tr>' for row in rows)
    if not body:
        body=f'<tr><td colspan="{len(headers)}">Nessuna riga disponibile.</td></tr>'
    return f'<div class="scroll"><table><thead><tr>{heads}</tr></thead><tbody>{body}</tbody></table></div>'

def html_report(inv, sheets, sources, outgoing, incoming):
    def anchor(key):
        return 'r'+hashlib.sha256(key.encode('utf-8')).hexdigest()[:24]
    def link(key):
        if key in inv.resources:
            return f'<a href="#{anchor(key)}" class="res">{esc(inv.name(key))}</a>'
        return f'<span>{esc(key)}</span>'
    labels = {
        'Instance': 'Istanze Compute', 'BootVolume': 'Volumi di avvio',
        'Volume': 'Volumi a blocchi', 'BootVolumeAttachment': 'Associazioni volumi di avvio',
        'VolumeAttachment': 'Associazioni volumi a blocchi', 'VnicAttachment': 'Associazioni VNIC',
        'Vcn': 'Reti virtuali (VCN)', 'Vnic': 'Interfacce di rete (VNIC)',
        'PrivateIp': 'IP privati', 'PublicIp': 'IP pubblici', 'RouteTable': 'Tabelle di routing',
        'SecurityList': 'Liste di sicurezza', 'NetworkSecurityGroup': 'Gruppi di sicurezza (NSG)',
    }

    def options(values, names=None):
        names = names or {}
        return ''.join(
            f'<option value="{esc(v)}">{esc(names.get(v, v))}</option>'
            for v in sorted(set(values) - {''}, key=lambda v: names.get(v, v).casefold())
            if v
        )
    def excel_table(sheet, pairs):
        def rows():
            for index,row in pairs:
                cells=[esc(v) if v not in (None,'') else '—' for v in row]
                yield cells+[str(index)]
        return table(HEADERS[sheet]+['Riga Excel'],rows())
    def selected(sheet, predicate):
        return [(i,r) for i,r in enumerate(sheets[sheet],2) if predicate(r)]
    def details(key, res):
        row=res['data']
        name, kind, region = inv.name(key),res['kind'],res['region'] or 'Non indicata'
        places=sources[key]
        main=sheets['Inventario'][places[0][1]-2]
        comp=main[2]
        related=[(link(key),esc(t),link(target),esc(note)) for t,target,note in outgoing[key]]
        related += [(link(target),esc(t),link(key),esc(note)) for t,target,note in incoming[key]]
        operational=any(t!='Compartment' for t,_,_ in outgoing[key]+incoming[key])
        status='Con relazioni documentate' if operational else 'Sola appartenenza o nessuna relazione'
        search=' '.join(compact(v) for v in [name,kind,labels.get(kind),region,res['id'],comp,*main[5:10]]).lower()
        state = row.get('lifecycle-state') or ''
        status_names = {'RUNNING': 'In esecuzione', 'STOPPED': 'Arrestata',
                        'STARTING': 'In avvio', 'STOPPING': 'In arresto', 'ACTIVE': 'Attivo'}
        badge = ('positive' if state in ('RUNNING', 'ACTIVE')
                 else 'paused' if state in ('STOPPED', 'STARTING', 'STOPPING') else 'neutral')
        parts=[f'<details class="item" id="{anchor(key)}" data-q="{esc(search)}" data-type="{esc(kind)}" data-region="{esc(region)}" data-status="{esc(status)}">',
               f'<summary><span class="resource-name"><strong>{esc(name)}</strong>'
               f'<small>{esc(labels.get(kind, kind))}</small></span>'
               f'<span class="state {badge}">{esc(status_names.get(state, state) or "—")}</span>'
               f'<small class="private-ip">{esc(main[5]) or "—"}</small>'
               f'<small class="region">{esc(region)}</small>'
               f'<small class="compartment">{esc(comp) or "—"}</small></summary><section>',
               f'<p>{esc(region)} &nbsp; / &nbsp; {esc(row.get("lifecycle-state")) or "Stato non indicato"}</p>',
               f'<code>{esc(res["id"])}</code>',
               table(['Attributo','Valore'],[(esc(k),esc(v)) for k,v in row.items() if k not in ('identifier','id')]),
               '<h3>Relazioni documentate</h3>',table(['Origine','Relazione','Destinazione','Dettaglio'],related),
               '<h3>Righe Excel</h3>']
        by_sheet=collections.defaultdict(list)
        for sheet,index in places:
            by_sheet[sheet].append((index,sheets[sheet][index-2]))
        for sheet,pairs in by_sheet.items():
            parts.append('<details class="raw"><summary>'+esc(sheet)+' · '+', '.join(str(i) for i,r in pairs)+'</summary>'+excel_table(sheet,pairs)+'</details>')
        parts.append('</section></details>')
        return ''.join(parts)
    def button(i,name):
        return f'<button type="button" data-tab="{i}" aria-current="{str(i==0).lower()}">{esc(name)}</button>'
    css = """
*{box-sizing:border-box}body{margin:0;background:#f5f8f9;color:#173445;font:16px/1.45 Arial,sans-serif}
main,header{max-width:1400px;margin:auto;padding:25px 32px}header{border-bottom:2px solid #adc3cb}
h1{font-size:36px;margin:6px 0 12px}h2{font-size:24px}h3{font-size:19px;margin:18px 0 6px}
.cap{color:#975337;font-weight:bold}.filters,nav{display:flex;flex-wrap:wrap;gap:9px;margin:16px 0}
button,select,input{font:inherit;background:white;color:#173445;border:1px solid #a9bac1;border-radius:4px;padding:7px 10px}
label{display:grid;gap:3px;font-size:14px}button,summary{cursor:pointer}button[aria-current=true]{background:#173e51;color:white}
.item{margin:5px 0;border:1px solid #cad8de;border-radius:5px;background:white}
.item>summary{display:grid;grid-template-columns:2fr 1fr 1fr;gap:10px;padding:11px 15px}
.item>summary strong:before{content:"＋ ";color:#316983}.item[open]>summary strong:before{content:"− "}
small{color:#506975}.item section{padding:8px 16px 18px}.raw{padding:5px 0;border-bottom:1px solid #dbe3e7}
code,td,summary{overflow-wrap:anywhere}code{background:#e8eff1;padding:3px 6px}.scroll{overflow-x:auto}
table{width:100%;border-collapse:collapse;margin:8px 0;background:white;font-size:14px}
th,td{text-align:left;vertical-align:top;padding:6px 9px;border-bottom:1px solid #d4dfe3}th{background:#e7eff2}a{color:#14617e}
[hidden]{display:none!important}@media(max-width:700px){main,header{padding:15px}.item>summary{grid-template-columns:1fr}}
"""
    css += '\n:root{--ink:#162637;--muted:#506377;--line:#dce5ee;--accent:#006f77;\n--grid:minmax(0,2.15fr) minmax(0,.85fr) minmax(0,.95fr) minmax(0,1fr) minmax(0,1.35fr)}\nbody{background:#f1f5f9;color:var(--ink);font:15px/1.45 Inter,system-ui,sans-serif}\nheader{max-width:none;background:#102638;color:#fff;border:0;padding-block:28px 0;\npadding-inline:max(24px,calc((100% - 1336px)/2))}\nheader h1{font-size:clamp(25px,4vw,38px);margin:4px 0 6px;letter-spacing:-.035em}\nheader p{margin:0;color:#c7d8e7}.cap{color:#80e0da;font-size:.86rem}\nnav{margin:18px 0 0;gap:6px}nav button{border-color:#3b566a;background:#19344a;\ncolor:#e9f2fa;border-radius:8px 8px 0 0;padding:9px 14px}\nnav button[aria-current=true]{color:#102638;background:#eef9fb;border-color:#eef9fb}\nmain{padding-top:22px}h2{margin-top:0;letter-spacing:-.02em}\n.filters{background:white;border:1px solid var(--line);border-radius:12px;padding:14px;\ndisplay:grid;grid-template-columns:minmax(220px,1.6fr) repeat(3,minmax(0,1fr))}\nlabel{min-width:0;color:var(--muted);font-weight:650}input,select{width:100%;min-width:0;border-color:#c5d3df;border-radius:8px}\n:focus-visible{outline:2px solid #00878d;outline-offset:2px}\n#found{color:var(--muted);margin:14px 0 8px}\n.columns,.item>summary{display:grid;grid-template-columns:var(--grid);align-items:center;\ngap:12px;padding:10px 16px}\n.columns{color:var(--muted);font-size:.8rem;font-weight:700}\n.item{border-color:var(--line);border-radius:10px;margin:6px 0}\n.item>summary:hover,.item[open]>summary{background:#eaf5f6;border-radius:inherit}\n.resource-name small{display:block;padding-left:1.24em;margin-top:2px}\n.item>summary strong{color:#075d66}\n.item>summary strong:before{color:var(--accent)}\n.state{justify-self:start;padding:3px 8px;border-radius:99px;font-size:.78rem;\nfont-weight:650;background:#eef2f6;color:#415769}\n.state.positive{background:#dff4e9;color:#15543c}\n.state.paused{background:#fff0d9;color:#754700}\n.private-ip,.region{font-variant-numeric:tabular-nums}\n.item section{border-top:1px solid var(--line);padding-top:10px}\nth{background:#e8f0f6}a.res{color:#08656f}\n@media(max-width:820px){\n  main{padding:18px 16px}.filters{grid-template-columns:repeat(2,minmax(0,1fr))}\n  .columns{display:none}.item>summary{grid-template-columns:minmax(0,2fr) minmax(0,1fr);gap:6px 10px}\n  .private-ip:before{content:"IP: "}.region:before{content:"Region: "}\n  .compartment{grid-column:1/-1}\n}\n@media(max-width:480px){.filters{grid-template-columns:1fr}}\n'
    js = """
const cards=Array.from(document.querySelectorAll('.item'));
const q=document.getElementById('q'), kind=document.getElementById('kind'), region=document.getElementById('region'), statusFilter=document.getElementById('status');
function filter(){let count=0;const term=q.value.trim().toLowerCase();
  for(const row of cards){row.hidden=!(row.dataset.q.includes(term)&&(!kind.value||row.dataset.type===kind.value)&&(!region.value||row.dataset.region===region.value)&&(!statusFilter.value||row.dataset.status===statusFilter.value));if(!row.hidden)count++;}
  document.getElementById('found').textContent=count+' di '+cards.length+' risorse';}
function showTab(index){for(const tab of document.querySelectorAll('[data-view]'))tab.hidden=tab.dataset.view!==String(index);
  for(const b of document.querySelectorAll('button[data-tab]'))b.setAttribute('aria-current',String(b.dataset.tab===String(index)));}
function openResource(){const obj=document.getElementById(location.hash.slice(1));if(!obj||!obj.classList.contains('item'))return;
  showTab(0);q.value=kind.value=region.value=statusFilter.value='';filter();obj.open=true;obj.scrollIntoView();}
for(const b of document.querySelectorAll('button[data-tab]'))b.onclick=()=>showTab(b.dataset.tab);
q.oninput=kind.onchange=region.onchange=statusFilter.onchange=filter;
document.addEventListener('click',e=>{const a=e.target.closest('a.res');if(a&&location.hash===a.hash){e.preventDefault();openResource();}});
window.addEventListener('hashchange',openResource);filter();openResource();
"""
    keys=sorted(inv.resources,key=lambda k:(inv.resources[k]['kind'],inv.name(k).casefold(),k))
    tabs=['Risorse','Regole di sicurezza','Statement IAM','Route e backup','Copertura e limiti']
    out=['<!doctype html><html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">',
         '<title>'+esc(inv.customer)+' | Registro tecnico OCI</title><style>'+css+'</style></head><body><header>',
         '<div class="cap">OCI / '+esc(inv.customer)+'</div><h1>Inventario delle risorse OCI</h1>',
         '<p>'+str(len(keys))+' oggetti · '+esc(inv.origin)+' · '+esc(', '.join(sorted(inv.regions)) or 'Region non indicata')+'</p>',
         '<nav>'+''.join(button(i,t) for i,t in enumerate(tabs))+'</nav></header><main>',
         '<div data-view="0"><h2>Risorse censite</h2><div class="filters"><label>Cerca risorsa<input id="q" type="search" placeholder="Nome, IP, CIDR, compartment o OCID"></label>',
         '<label>Tipologia<select id="kind"><option value="">Tutte</option>'+options((v['kind'] for v in inv.resources.values()), labels)+'</select></label>',
         '<label>Region<select id="region"><option value="">Tutte</option>'+options(v['region'] or 'Non indicata' for v in inv.resources.values())+'</select></label>',
         '<label>Relazioni<select id="status"><option value="">Tutte</option><option>Con relazioni documentate</option><option>Sola appartenenza o nessuna relazione</option></select></label></div><p id="found" aria-live="polite"></p>',
         '<div class="columns" aria-hidden="true"><span>Nome e tipologia</span><span>Stato</span>'
         '<span>IP privato</span><span>Region</span><span>Compartment</span></div>']
    out.extend(details(k,inv.resources[k]) for k in keys)
    rules=selected('Security',lambda r:bool(r[4]))
    policies=selected('IAM',lambda r:r[0]=='Policy' and bool(r[6]))
    routes=selected('Networking',lambda r:r[0]=='RouteTable' and bool(r[9]) and r[7] not in ('VCN','Compartment'))
    coverage = selected('Security', lambda r: r[0] == 'NetworkSecurityGroup' and not r[4])
    out+=['</div><div hidden data-view="1"><h2>Regole di sicurezza documentate · '+str(len(rules))+'</h2>',excel_table('Security',rules),
          '<h3>NSG senza regole documentate</h3>' + excel_table('Security', coverage) if coverage else '',
          '</div><div hidden data-view="2"><h2>Statement IAM · '+str(len(policies))+'</h2>',excel_table('IAM',policies),
          '</div><div hidden data-view="3"><h2>Route configurate · '+str(len(routes))+'</h2>',excel_table('Networking',routes),
          '<h2>Backup e assegnazioni</h2>',excel_table('Backup',enumerate(sheets['Backup'],2)),
          '</div><div hidden data-view="4"><h2>Copertura</h2>',
          '<p>Le relazioni usano gli identificatori esposti nelle risposte OCI. Gli oggetti disponibili solo da Resource Search possono avere una scheda con i soli dati di base. Per IAM, gli utenti e i gruppi approfonditi sono quelli visibili tramite le API IAM classiche; i domini sono censiti separatamente.</p>',
          table(['Tipo OCI','Oggetti'],[(esc(k),str(v)) for k,v in sorted(collections.Counter(r['kind'] for r in inv.resources.values()).items())]),
          '<h2>Note di raccolta e richieste non riuscite</h2>',table(['Region','Operazione','Ambito','Esito'],[[esc(v) for v in issue] for issue in inv.issues]),
          '<h2>Chiamate OCI riuscite</h2>',table(['Region','Operazione','Chiamate'],[(esc(r),esc(op),str(count)) for (r,op),count in sorted(inv.checked.items())]),
          '</div></main><script>'+js+'</script></body></html>']
    return ''.join(out)

def main():
    parser=argparse.ArgumentParser(description=__doc__,formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--input',type=Path,metavar='ZIP|DIRECTORY|JSON',help='Read OCI CLI JSON exports instead of querying OCI')
    parser.add_argument('--tenancy-id',help='Tenancy OCID (defaults to OCI CLI configuration or environment)')
    parser.add_argument('--customer', dest='cliente', metavar='NAME',help='Name used in report filenames and titles (defaults to tenancy name)')
    parser.add_argument('--cliente', dest='cliente', help=argparse.SUPPRESS)
    parser.add_argument('--region',action='append',default=[],help='Region to scan; repeat for multiple regions (default: all subscribed regions)')
    parser.add_argument('--profile',help='OCI CLI profile (default: current profile)')
    parser.add_argument('--config-file',type=Path,help='OCI CLI configuration file')
    parser.add_argument('--auth',help='OCI CLI authentication method (e.g. instance_principal or security_token)')
    parser.add_argument('--output-dir',type=Path,default=Path('.'),help='Directory for the generated Excel and HTML reports')
    parser.add_argument('--no-animation', action='store_true', help='Disable terminal animation during OCI CLI calls')
    args=parser.parse_args()
    inv=Inventory()
    if args.input:
        if len(args.region)>1:
            parser.error('Offline mode accepts only one fallback region for OCIDs without a region.')
        load_exports(args.input,inv,args.region[0] if args.region else '')
    else:
        collect_live(args,inv)
    if args.tenancy_id and inv.roots and inv.roots!={args.tenancy_id}:
        raise ValueError('The specified tenancy does not match the inventory data.')
    root_name=next((inv.name(k) for k in sorted(inv.roots) if k in inv.resources and inv.name(k)!=k),None)
    fallback = 'OCI_'+next(iter(inv.roots))[-8:] if inv.roots else 'OCI'
    inv.customer=args.cliente or root_name or fallback
    inv.build_links()
    sheets,refs,forward,reverse=spreadsheet_data(inv)
    tag=clean_name(inv.customer)
    args.output_dir.mkdir(parents=True,exist_ok=True)
    xlsx=args.output_dir/f'Inventario_{tag}.xlsx'
    report=args.output_dir/f'Registro_tecnico_{tag}.html'
    # Replace reports only after both have been generated.
    with tempfile.TemporaryDirectory(dir=args.output_dir) as tmp:
        directory=Path(tmp)
        write_xlsx(directory/xlsx.name,sheets)
        (directory/report.name).write_text(html_report(inv,sheets,refs,forward,reverse),encoding='utf-8')
        os.replace(directory/xlsx.name,xlsx)
        os.replace(directory/report.name,report)
    summary = f'{len(inv.resources)} resources · {len(inv.edges)} relations · {len(inv.issues)} collection notes'
    try:
        print(summary)
    except UnicodeEncodeError:
        print(summary.replace('·', '|'))
    print(xlsx.resolve())
    print(report.resolve())

if __name__=='__main__':
    try:
        main()
    except (OSError, ValueError, zipfile.BadZipFile, configparser.Error) as error:
        sys.exit('Error: '+str(error))
