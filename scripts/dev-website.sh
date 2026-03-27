#!/usr/bin/env bash
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>
#
# Start the Jekyll website locally for development (live-reload).
#
# Usage:
#   scripts/dev-website.sh
#
# The site is served at http://localhost:4000 with incremental rebuilds
# and live-reload enabled. Changes to files under website/ are picked up
# automatically — no manual restart needed.
#
# Prerequisites:
#   - Ruby and Bundler — pre-installed in the dev container
#   - Jekyll gems installed via `cd website && bundle install`
#     (done automatically by post-create.sh)

set -euo pipefail

# Always run from the repository root so paths resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

# shellcheck source=scripts/colors.sh
source "$(dirname "${BASH_SOURCE[0]}")/colors.sh"

WEBSITE_DIR="website"

# --- GEM_HOME ---
# Keep gems in user space so no sudo is needed. post-create.sh persists
# this in .bashrc, but non-interactive shells (e.g. VS Code tasks) may
# not source .bashrc — set it explicitly here.
export GEM_HOME="${HOME}/.gems"
export PATH="${GEM_HOME}/bin:${PATH}"

# --- Preflight checks ---

if ! command -v ruby &>/dev/null; then
  echo "${C_RED}ERROR:${C_RST} Ruby not found."
  echo "  The dev container should have Ruby pre-installed."
  echo "  Try rebuilding the container."
  exit 1
fi

if ! command -v bundle &>/dev/null; then
  echo "${C_RED}ERROR:${C_RST} Bundler not found."
  echo "  Install it with: gem install bundler"
  exit 1
fi

if ! command -v jekyll &>/dev/null; then
  echo "${C_RED}ERROR:${C_RST} Jekyll not found."
  echo "  Run: cd ${WEBSITE_DIR} && bundle install"
  exit 1
fi

# --- Dependencies ---

if [[ ! -d "${WEBSITE_DIR}/vendor" && ! -f "${WEBSITE_DIR}/Gemfile.lock" ]]; then
  echo "Installing Jekyll gems..."
  (cd "${WEBSITE_DIR}" && bundle install --jobs 4 --retry 2)
  echo ""
fi

# --- Start ---

echo "Starting Jekyll development server..."
hint "Website: ${C_BLD}http://localhost:4000${C_RST}" \
  "" \
  "Live-reload is enabled — changes are applied automatically." \
  "Press Ctrl+C to stop."
echo ""

cd "${WEBSITE_DIR}"
bundle exec jekyll serve \
  --livereload \
  --incremental \
  --host 0.0.0.0 \
  --port 4000
