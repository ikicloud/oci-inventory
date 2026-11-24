# OCI Resource Inventory

**A read-only inventory and reporting tool for Oracle Cloud Infrastructure (OCI), developed by Iki-Cloud.**

OCI Resource Inventory collects information about resources in an OCI tenancy and turns it into two complementary reports: an interactive HTML technical register and a structured Excel workbook. It is designed to make infrastructure assessments, documentation, and migration planning easier by bringing resource details and their relationships into a single view.

The script uses the **OCI CLI** for live data collection. It can also process previously exported OCI CLI JSON files without connecting to OCI.

## Features

- **Tenancy-wide discovery:** use OCI Resource Search and additional service API calls to inventory accessible resources across subscribed regions, or restrict collection to selected regions.
- **Resource relationships:** connect compute instances, VNICs, subnets, VCNs, boot and block volumes, route tables, security controls, and other resources where OCI exposes the relevant identifiers.
- **Networking and security:** document routing information, Security Lists, Network Security Groups (NSGs), and their available ingress/egress rules.
- **Backup visibility:** include volume backups, backup policies, and policy assignments where available.
- **IAM visibility:** collect accessible compartments, users, groups, policies, policy statements, and identity domains, subject to API and permission coverage.
- **Two report formats:** browse a searchable, self-contained HTML report or analyze the data in an Excel workbook with dedicated worksheets.
- **Coverage reporting:** record unsuccessful requests and collection limitations instead of silently treating missing information as complete.
- **Offline mode:** generate reports from OCI CLI JSON exports supplied as a file, directory, or ZIP archive.

## Requirements

- **Python 3.8 or later**
- **OCI CLI** installed and configured for live collection (already available in OCI Cloud Shell)
- An OCI identity or session with sufficient **read access** to the resources and APIs being inventoried

The Python script uses only the **standard library**: no `pip install` or additional Python packages are required. Offline mode does not require the OCI CLI.

## Quick start

### 1. Run in OCI Cloud Shell

From Cloud Shell, clone the repository and run the inventory using the current OCI CLI session.

```bash
git clone https://github.com/ikicloud/oci-inventory.git
cd oci-inventory
bash run.sh
```

The launcher runs the Python script and saves both reports under `reports/` inside the checkout. It accepts the same options as the script. For example, `bash run.sh --region eu-milan-1` restricts collection to Milan. If the tenancy cannot be detected from the environment or profile, pass its OCID with `--tenancy-id`.

To update an existing checkout, run `git pull --ff-only` from its directory. To invoke Python directly, use `python3 oci_inventory.py --output-dir reports`.

### 2. Run with a specific OCI CLI profile

```bash
python3 oci_inventory.py --profile CUSTOMER --region eu-milan-1
```

You can specify multiple regions by repeating `--region`:

```bash
python3 oci_inventory.py \
  --region eu-milan-1 \
  --region eu-frankfurt-1 \
  --customer "Example Customer" \
  --output-dir ./reports
```

The script also supports `--config-file` and `--auth` to pass the corresponding OCI CLI configuration and authentication options.

### 3. Generate reports from JSON exports (offline)

```bash
bash run.sh --input sample-inventory.json --customer Demo
```

`--input` accepts a **JSON file**, a **directory containing JSON files**, or a **ZIP archive** of OCI CLI exports. In offline mode, the reports can only contain information present in the supplied exports.

For exported NSG rules, preserve their owning NSG by storing each rule list at:

```text
<NSG_OCID>/rules.json
```

Alternatively, a JSON object with an `nsg-id` field and a `data` array is supported. This matters because individual NSG rule IDs do not identify the owning NSG by themselves.

## Generated reports

When running Python directly, unless `--output-dir` is provided, both files are written to the current working directory. The customer/tenancy name is sanitized for use in filenames.

| File | Description |
| --- | --- |
| `Inventory_<customer>.xlsx` | Structured Excel inventory with filters and frozen header rows. |
| `Technical_register_<customer>.html` | Standalone interactive HTML technical register, viewable in a browser. |

The Excel workbook contains six worksheets:

1. **Inventory** — consolidated inventory, resource identifiers, location, network attributes, and relationships.
2. **Compute and Storage** — compute instances, volumes, attachments, shapes, and associated details.
3. **Networking** — network resources, associations, routing, and related identifiers.
4. **Security** — security resources and available rules.
5. **Backup** — backups, policies, assignments, and schedules.
6. **IAM** — tenancy identity resources and policy information.

The HTML report provides resource search and filters, expandable resource details, links between related resources, and dedicated views for security rules, IAM statements, routes, backups, and collection coverage.

## Command-line options

```bash
python3 oci_inventory.py --help
```

| Option | Purpose |
| --- | --- |
| `--tenancy-id` | Specify the tenancy OCID. |
| `--profile` | Select an OCI CLI configuration profile. |
| `--config-file` | Specify a custom OCI CLI configuration file. |
| `--auth` | Select an OCI CLI authentication mechanism. |
| `--region` | Restrict live inventory to a region; repeat to add more regions. |
| `--customer` | Set the customer name displayed in reports and used in filenames. |
| `--output-dir` | Select the output directory. |
| `--input` | Read OCI CLI JSON exports instead of querying OCI. |
| `--no-animation` | Disable the optional terminal animation during live collection. |

## Scope and limitations

- **Read-only in OCI:** the script uses OCI CLI search, list, and get operations. It does not create, update, or delete OCI resources.
- **Permissions and API coverage matter:** Resource Search is used for broad discovery; additional details are collected for supported resource types. Some resources may therefore appear with basic metadata only.
- **Not a completeness guarantee:** unavailable permissions, failed calls, API limitations, or incomplete offline exports can result in missing details. Review the **Coverage and limitations** section of the HTML report and the collection notes in the Excel inventory.
- **IAM coverage varies:** classic IAM users/groups and identity domains are not interchangeable; identity-domain details are not exhaustively expanded.
- **Local output files are replaced** if a report with the same name already exists in the selected output directory.

## Security and data handling

Inventory reports can contain **OCIDs, private/public IP addresses, CIDR ranges, routing information, security rules, and IAM policy statements**. Treat generated reports as potentially sensitive infrastructure documentation.

Before sharing the repository or publishing report examples:

- Do not commit OCI configuration files, API private keys, authentication tokens, or customer-specific exports.
- Do not publish real customer inventory reports without authorization and appropriate sanitization.
- Review access permissions before distributing generated HTML or Excel files.

## Project

Developed by **Iki-Cloud** to support OCI infrastructure discovery, technical documentation, and assessment activities.
