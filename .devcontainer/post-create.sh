#!/usr/bin/env bash
set -euo pipefail

# shellcheck source=scripts/colors.sh
source scripts/colors.sh

# Collect non-fatal warnings so they can be repeated at the end of output,
# where they are easier to spot than inline among scrolling install logs.
WARNINGS=()
warn() {
  local msg="$1"
  local hint="${2:-}" # optional second line with recovery instructions
  WARNINGS+=("${msg}")
  echo "${C_YLW}⚠${C_RST} ${msg}" >&2
  [[ -n "${hint}" ]] && echo "  ${C_DIM}${hint}${C_RST}" >&2
}

workspace_dir="${containerWorkspaceFolder:-$(pwd)}"

git config --global --add safe.directory "${workspace_dir}"
git config --global pull.rebase true
npm config set save-exact true --location=user

# Enable npm tab-completion in bash.
if ! grep -q "npm completion" "${HOME}/.bashrc" 2>/dev/null; then
  npm completion >>"${HOME}/.bashrc"
fi

# Configure git identity from host gitconfig.
bash .devcontainer/setup-git.sh

# Run a network-dependent command with a fast connectivity guard and retry.
#
# First probes <host> with a 3-second curl HEAD request.  If the host is
# completely unreachable (no DNS, no TCP), returns 1 immediately — no further
# waiting.  When the host is reachable but the command fails (transient error),
# retries up to 3 times with short backoff (2 s → 4 s) — fast enough for
# container builds without hanging indefinitely.
#
# Usage: try_net <host> <command> [args…]
try_net() {
  local host="$1"
  shift
  # 3-second timeout — treat DNS failure or TCP refusal as "no network".
  # Do NOT use -f: a 4xx/5xx HTTP response (curl exit 22) still proves the
  # host is reachable.  Only true network errors (exit 6 = DNS, 7 = refused,
  # 28 = timeout) should be treated as "no network".
  local probe_exit
  curl -sS --max-time 3 --head "https://${host}" >/dev/null 2>&1 || probe_exit=$?
  case "${probe_exit:-0}" in
    0 | 22) ;; # 0 = success, 22 = HTTP 4xx/5xx — host IS reachable
    *) return 1 ;;
  esac
  local max=3 attempt=1 delay=2
  until "$@"; do
    if [[ "${attempt}" -ge "${max}" ]]; then
      return 1
    fi
    echo "${C_YLW}⚠${C_RST} Attempt ${attempt}/${max} failed — retrying in ${delay}s…" >&2
    sleep "${delay}"
    attempt=$((attempt + 1))
    delay=$((delay * 2)) # exponential backoff: 2 s → 4 s
  done
}

# Pre-install git-cliff so ./scripts/release-notes.sh works without a download delay.
# The pinned version is read from the script itself (single source of truth).
GIT_CLIFF_VERSION="$(grep '^GIT_CLIFF_VERSION=' scripts/release-notes.sh | sed 's/.*"\(.*\)".*/\1/')"
INSTALL_DIR="${HOME}/.local/bin"
mkdir -p "${INSTALL_DIR}"
if [[ ! -x "${INSTALL_DIR}/git-cliff" ]]; then
  TRIPLE="x86_64-unknown-linux-musl"
  TARBALL="git-cliff-${GIT_CLIFF_VERSION}-${TRIPLE}.tar.gz"
  # Download to a temp file so the network step can be retried independently
  # from the tar extraction, and the temp file is always cleaned up.
  TARBALL_TMP="$(mktemp)"
  if try_net github.com curl -fsSL \
    "https://github.com/orhun/git-cliff/releases/download/v${GIT_CLIFF_VERSION}/${TARBALL}" \
    -o "${TARBALL_TMP}"; then
    tar -xz -C "${INSTALL_DIR}" \
      --strip-components=1 \
      "git-cliff-${GIT_CLIFF_VERSION}/git-cliff" <"${TARBALL_TMP}"
    rm -f "${TARBALL_TMP}"
    chmod +x "${INSTALL_DIR}/git-cliff"
  else
    rm -f "${TARBALL_TMP}"
    warn "git-cliff installation skipped — network unavailable?" \
      "Run 'bash .devcontainer/post-create.sh' once the network is available."
  fi
fi
# Ensure ~/.local/bin is on PATH for interactive shells.
if ! grep -q '\.local/bin' "${HOME}/.bashrc" 2>/dev/null; then
  # SC2016: single quotes intentional — $HOME must expand at shell startup, not now.
  # shellcheck disable=SC2016
  echo 'export PATH="${HOME}/.local/bin:${PATH}"' >>"${HOME}/.bashrc"
fi

# Install project dependencies and create .env from .env.example if absent.
# bootstrap.sh is the single source of truth for both steps.
# _CALLOUT_SUPPRESS tells bootstrap.sh to skip its own next_steps box — we
# print a combined summary at the very end of this script instead.
_CALLOUT_SUPPRESS=1 bash scripts/bootstrap.sh

# Install Jekyll gems for local website development.
# GEM_HOME keeps gems in user space so no sudo is needed.
export GEM_HOME="${HOME}/.gems"
export PATH="${GEM_HOME}/bin:${PATH}"
if [[ -f website/Gemfile ]]; then
  echo "Installing Jekyll gems for local website development..."
  # The apt package ruby-dev only ships 'bundle3.1', not the bare 'bundle'
  # command. Install bundler into GEM_HOME so 'bundle' is available on PATH.
  gem install bundler --no-document --quiet
  if (cd website && bundle install --jobs 4 --retry 2 2>&1); then
    echo "${C_GRN}✓${C_RST} Jekyll gems installed."
  else
    warn "Jekyll gem installation failed." \
      "Run 'cd website && bundle install' to retry."
  fi
fi
# Persist GEM_HOME on PATH for interactive shells.
if ! grep -q 'GEM_HOME' "${HOME}/.bashrc" 2>/dev/null; then
  # SC2016: single quotes intentional — variables must expand at shell startup.
  # shellcheck disable=SC2016
  {
    echo 'export GEM_HOME="${HOME}/.gems"'
    echo 'export PATH="${GEM_HOME}/bin:${PATH}"'
  } >>"${HOME}/.bashrc"
fi

# Ensure Husky git hooks are properly initialized in dev-container.
# set in the environment, which would leave core.hooksPath pointing to whatever
# it was before (possibly a stale or incorrect value). Explicitly setting it
# to the canonical value ensures git hooks are always active in the container.
git config core.hooksPath .husky/_

# Configure delta as the git diff pager for syntax-highlighted diffs.
# Side-by-side mode is off by default (too wide for most terminals) but can
# be toggled with `delta --side-by-side` or by setting GIT_PAGER at runtime.
if command -v delta &>/dev/null; then
  git config --global core.pager delta
  git config --global interactive.diffFilter "delta --color-only"
  git config --global delta.navigate true
  git config --global delta.line-numbers true
fi

# Install/upgrade the Bicep CLI through the Azure CLI.
# 'az bicep install' is idempotent and always fetches the latest release,
# so the version is up to date on every container rebuild — no stale binary
# baked into the image. Azure CLI is installed as a devcontainer feature and
# therefore not available during the Dockerfile build; this is the right hook.
# az bicep resolves downloads through aka.ms
# Redirect stdout so the noisy "already installed" line does not interleave
# with the final summary block that follows.
if ! try_net aka.ms az bicep install >/dev/null; then
  warn "Bicep CLI installation skipped — network unavailable?" \
    "Run 'az bicep install' once the network is available."
fi

# Pre-fetch external Bicep extensions used by the infra template so editor
# diagnostics work immediately in a fresh container.
if ! try_net mcr.microsoft.com az bicep restore --file azure-function/infra/main.bicep >/dev/null; then
  warn "Bicep extension restore skipped — network unavailable?" \
    "Run 'az bicep restore --file azure-function/infra/main.bicep' once the network is available."
fi

# Install repo-defined APM packages so every contributor gets the same shared
# skills and agent context immediately after container creation.
if [[ -f apm.yml ]]; then
  echo "Installing APM-managed agent dependencies..."
  if ! command -v apm >/dev/null 2>&1; then
    warn "APM CLI is not available in PATH." \
      "Rebuild the dev container to pick up the Dockerfile changes."
  elif ! try_net github.com apm install; then
    warn "APM dependency installation skipped — network unavailable?" \
      "Run 'apm install' once the network is available."
  fi
fi

echo ""
echo "${C_BLD}━━━ Dev Container Ready ━━━${C_RST}"
echo ""
echo "  Node     $(node --version)"
echo "  npm      $(npm --version)"
echo "  Yeoman   $(yo --version)"
echo "  SPFx     $(npm view @microsoft/generator-sharepoint version)"
echo "  APM      $(apm --version 2>/dev/null || echo "(not installed)")"
echo "  Ruby     $(ruby --version 2>/dev/null | awk '{print $2}' || echo "(not found)")"
echo "  Jekyll   $(jekyll --version 2>/dev/null | awk '{print $2}' || echo "(not installed)")"
echo "  Bicep    $(az bicep version 2>/dev/null || echo "(not installed)")"

# Repeat collected warnings so they are visible after the version summary.
if [[ ${#WARNINGS[@]} -gt 0 ]]; then
  # Build warning lines array for the important() box.
  WARN_LINES=()
  for w in "${WARNINGS[@]}"; do
    WARN_LINES+=("${C_YLW}⚠${C_RST} ${w}")
  done
  WARN_LINES+=("")
  WARN_LINES+=("Re-run ${C_BLD}bash .devcontainer/post-create.sh${C_RST} to retry.")
  important "${WARN_LINES[@]}"
else
  echo ""
  echo "${C_GRN}✓${C_RST} All setup steps completed successfully."
fi

next_steps "${C_BLD}./scripts/dev-webpart.sh${C_RST}    # start the SPFx dev server" \
  "${C_BLD}./scripts/dev-function.sh${C_RST}   # start the Azure Function locally" \
  "${C_BLD}./scripts/dev-website.sh${C_RST}    # start the Jekyll website locally"
