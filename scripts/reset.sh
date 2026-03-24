#!/usr/bin/env bash
# Reset the development environment to a clean state.
#
# Usage:
#   scripts/reset.sh
#
# Removes all build outputs, caches, and node_modules for both the SPFx
# web part and the Azure Function, then re-installs dependencies via
# scripts/bootstrap.sh. Useful after branch switches with major dependency
# changes or when builds produce unexpected results.

set -euo pipefail

# Always run from the repository root so paths resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

echo "Cleaning web part build outputs..."
npm run clean

echo ""
echo "Removing web part node_modules..."
rm -rf node_modules

echo ""
echo "Cleaning Azure Function build output and node_modules..."
rm -rf azure-function/dist azure-function/node_modules

echo ""
echo "Re-installing dependencies..."
./scripts/bootstrap.sh
