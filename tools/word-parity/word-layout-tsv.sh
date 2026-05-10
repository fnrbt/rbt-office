#!/usr/bin/env bash
set -euo pipefail

script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
exec "$script_dir/run-word-vbs-on-vm.sh" "$script_dir/extract-word-layout.vbs" "$@"
