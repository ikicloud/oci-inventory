#!/usr/bin/env bash
set -eu

# Resolve the checkout independently of the caller's working directory.
repo_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
exec python3 "$repo_dir/oci_inventory.py" --output-dir "$repo_dir/reports" "$@"
