#!/usr/bin/env bash
# Idempotent repository bootstrap for the Archivo Maestro Web dev environment.
# Installs system packages, JS dependencies, and initializes the MySQL database.
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

echo "==> Installing system packages (MariaDB)"
export DEBIAN_FRONTEND=noninteractive
sudo apt-get update -qq
sudo apt-get install -y -qq mariadb-server mariadb-client

echo "==> Installing frontend dependencies"
npm ci

echo "==> Installing backend dependencies"
(cd backend && npm ci)

echo "==> Ensuring backend/.env exists"
if [ ! -f backend/.env ]; then
  cp backend/.env.example backend/.env
fi

echo "==> Initializing MariaDB (database, schema, credentials)"
bash "$REPO_ROOT/.cursor/init-db.sh"

echo "==> install.sh completed"
