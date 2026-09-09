#!/usr/bin/env bash
# Ensures the MariaDB server is running and that the archivo_maestro database,
# schema, and root TCP credentials exist. Safe to run repeatedly.
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

DB_NAME="${DB_NAME:-archivo_maestro}"

echo "==> Preparing MariaDB runtime directory"
sudo mkdir -p /run/mysqld
sudo chown mysql:mysql /run/mysqld

if ! sudo mysqladmin ping --silent >/dev/null 2>&1; then
  echo "==> Starting MariaDB server"
  sudo -b bash -c 'mariadbd-safe --datadir=/var/lib/mysql >/tmp/mariadb.log 2>&1'
  echo "==> Waiting for MariaDB to accept connections"
  for _ in $(seq 1 60); do
    if sudo mysqladmin ping --silent >/dev/null 2>&1; then
      break
    fi
    sleep 1
  done
fi

if ! sudo mysqladmin ping --silent >/dev/null 2>&1; then
  echo "ERROR: MariaDB did not start. Last log lines:" >&2
  sudo tail -n 40 /tmp/mariadb.log >&2 || true
  exit 1
fi

echo "==> Creating database and enabling password-less TCP access for root"
# The backend (mysql2) connects over TCP as root with an empty password, matching
# backend/.env. Name resolution maps 127.0.0.1/::1 to 'localhost', so switching
# root@localhost to native empty-password auth serves both socket and TCP clients.
sudo mysql <<SQL
CREATE DATABASE IF NOT EXISTS \`${DB_NAME}\` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
ALTER USER 'root'@'localhost' IDENTIFIED BY '';
FLUSH PRIVILEGES;
SQL

echo "==> Applying schema"
sudo mysql "${DB_NAME}" < "$REPO_ROOT/backend/schema.sql"

echo "==> Database ready: ${DB_NAME}"
