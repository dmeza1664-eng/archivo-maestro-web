#!/usr/bin/env bash
# Per-boot startup: bring up MariaDB and make sure the database is ready.
# The frontend and backend dev servers are launched as persistent terminals.
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

bash "$REPO_ROOT/.cursor/init-db.sh"

echo "==> start.sh completed; dev servers run in the 'backend' and 'frontend' terminals"
