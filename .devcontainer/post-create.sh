#!/usr/bin/env bash
set -euo pipefail

workspace_dir="${containerWorkspaceFolder:-$(pwd)}"

git config --global --add safe.directory "${workspace_dir}"
npm config set save-exact true --location=user

# Enable npm tab-completion in bash.
if ! grep -q "npm completion" "${HOME}/.bashrc" 2>/dev/null; then
  npm completion >> "${HOME}/.bashrc"
fi

# Configure git identity from host gitconfig.
bash .devcontainer/setup-git.sh

# Pre-install git-cliff so ./scripts/release-notes.sh works without a download delay.
# The pinned version is read from the script itself (single source of truth).
GIT_CLIFF_VERSION="$(grep '^GIT_CLIFF_VERSION=' scripts/release-notes.sh | sed 's/.*"\(.*\)".*/\1/')"
INSTALL_DIR="${HOME}/.local/bin"
mkdir -p "${INSTALL_DIR}"
if [[ ! -x "${INSTALL_DIR}/git-cliff" ]]; then
  TRIPLE="x86_64-unknown-linux-musl"
  TARBALL="git-cliff-${GIT_CLIFF_VERSION}-${TRIPLE}.tar.gz"
  curl -fsSL \
    "https://github.com/orhun/git-cliff/releases/download/v${GIT_CLIFF_VERSION}/${TARBALL}" \
    | tar -xz -C "${INSTALL_DIR}" \
        --strip-components=1 \
        "git-cliff-${GIT_CLIFF_VERSION}/git-cliff"
  chmod +x "${INSTALL_DIR}/git-cliff"
fi
# Ensure ~/.local/bin is on PATH for interactive shells.
if ! grep -q '\.local/bin' "${HOME}/.bashrc" 2>/dev/null; then
  echo 'export PATH="${HOME}/.local/bin:${PATH}"' >> "${HOME}/.bashrc"
fi

# Install project dependencies and set up git hooks (husky prepare script).
npm install

# Ensure Husky git hooks are properly initialized in dev-container.
# Note: npm install already triggers the prepare script (which runs 'husky'),
# but we call it explicitly here as a belt-and-suspenders measure.
npm run prepare

# Hard-enforce the correct hooksPath. 'husky' silently skips when HUSKY=0 is
# set in the environment, which would leave core.hooksPath pointing to whatever
# it was before (possibly a stale or incorrect value). Explicitly setting it
# to the canonical value ensures git hooks are always active in the container.
git config core.hooksPath .husky/_

# Install/upgrade the Bicep CLI through the Azure CLI.
# 'az bicep install' is idempotent and always fetches the latest release,
# so the version is up to date on every container rebuild — no stale binary
# baked into the image. Azure CLI is installed as a devcontainer feature and
# therefore not available during the Dockerfile build; this is the right hook.
az bicep install

echo "Node version: $(node --version)"
echo "npm version: $(npm --version)"
echo "Yeoman version: $(yo --version)"
echo "SPFx generator version: $(npm view @microsoft/generator-sharepoint version)"
echo "Bicep version: $(az bicep version)"

