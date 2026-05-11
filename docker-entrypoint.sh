#!/bin/bash
export PGPASSWORD="${POSTGRES_PASSWORD}"

until pg_isready -U postgres -h db; do
    sleep 1
done

psql -U postgres -h db -d crm <<SQL
INSERT INTO _users (username, password, role)
VALUES ('admin', '${ADMIN_PASSWORD}', 'admin')
ON CONFLICT (username) DO NOTHING;
SQL

echo "Admin user created."
