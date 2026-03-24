#!/usr/bin/env bash
# Bootstrap the development environment.
#
# Usage:
#   ./scripts/bootstrap.sh
#
# Installs web part dependencies (npm install) and creates .env from
# .env.example if .env does not already exist. Run once after cloning.
#
# After bootstrapping, start a dev server:
#   ./scripts/dev-webpart.sh    # SPFx web part (hosted workbench)
#   ./scripts/dev-function.sh   # Azure Function (requires az login)

set -euo pipefail

# Always run from the repository root so paths resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

echo "Installing web part dependencies..."
npm install

ENV_FILE=".env"
if [[ ! -f "${ENV_FILE}" ]]; then
  cp .env.example "${ENV_FILE}"
  echo ""
  echo "Created .env from .env.example."
  echo "  → Edit .env and set SPFX_SERVE_TENANT_DOMAIN=<your-tenant>.sharepoint.com"
  echo "  (or export SPFX_SERVE_TENANT_DOMAIN on your host OS — see .devcontainer/devcontainer.json)"
else
  echo ""
  echo ".env already exists — skipped."
fi

echo ""
echo "Bootstrap complete. Next steps:"
echo "  ./scripts/dev-webpart.sh    # start the SPFx dev server"
echo "  ./scripts/dev-function.sh   # start the Azure Function locally"
