#!/usr/bin/env bash
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>

set -euo pipefail

# Always run from repository root so relative paths resolve reliably.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

# shellcheck source=scripts/colors.sh
source "$(dirname "${BASH_SOURCE[0]}")/colors.sh"

if ! command -v az >/dev/null 2>&1; then
  echo "${C_RED}✗${C_RST} Azure CLI not found. Install Azure CLI to run Bicep lint."
  exit 1
fi

if ! az bicep version >/dev/null 2>&1; then
  echo "${C_RED}✗${C_RST} Bicep CLI not available via Azure CLI."
  next_steps "Run ${C_BLD}az bicep install${C_RST} and retry ${C_BLD}npm run lint:bicep${C_RST}."
  exit 1
fi

if [[ $# -gt 0 ]]; then
  # lint-staged appends the staged file list to the command. Keep only .bicep
  # files so accidental extra arguments do not break az bicep lint.
  BICEP_FILES=()
  for file in "$@"; do
    if [[ "$file" == *.bicep ]]; then
      BICEP_FILES+=("$file")
    fi
  done
else
  mapfile -t BICEP_FILES < <(find azure-function/infra -type f -name '*.bicep' | sort)
fi

if [[ ${#BICEP_FILES[@]} -eq 0 ]]; then
  echo "${C_RED}✗${C_RST} No Bicep files found under azure-function/infra."
  exit 1
fi

echo "${C_BLD}Linting ${#BICEP_FILES[@]} Bicep file(s)...${C_RST}"
for file in "${BICEP_FILES[@]}"; do
  echo "  ${C_CYN}•${C_RST} ${file}"
  az bicep lint --file "${file}"
done

echo "${C_GRN}✓${C_RST} Bicep lint passed for all infra templates."
