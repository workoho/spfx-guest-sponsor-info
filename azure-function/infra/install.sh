#!/usr/bin/env bash
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>
#
# POSIX bootstrapper for macOS and Linux. It installs PowerShell when missing,
# downloads install.ps1, and lets the PowerShell installer perform the actual
# Guest Sponsor Info deployment workflow.
#
# Usage:
#   curl -fsSL https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.sh | bash
#   curl -fsSL https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.sh | bash -s -- -Version v1.2.0

set -euo pipefail

INSTALL_PS1_URL="${GSI_INSTALL_PS1_URL:-https://raw.githubusercontent.com/workoho/spfx-guest-sponsor-info/main/azure-function/infra/install.ps1}"
TMP_DIR=''

cleanup() {
  if [[ -n "${TMP_DIR}" && -d "${TMP_DIR}" ]]; then
    rm -rf "${TMP_DIR}"
  fi
}
trap cleanup EXIT

die() {
  printf 'ERROR: %s\n' "$*" >&2
  exit 1
}

info() {
  printf '  %s\n' "$*"
}

have() {
  command -v "$1" >/dev/null 2>&1
}

prompt_yes_no() {
  local prompt="$1"
  local answer=''

  if [[ "${GSI_INSTALL_ASSUME_YES:-}" == "1" ]]; then
    return 0
  fi

  if [[ ! -r /dev/tty ]]; then
    die "${prompt} Set GSI_INSTALL_ASSUME_YES=1 to approve non-interactively."
  fi

  printf '%s ' "${prompt}" >/dev/tty
  read -r answer </dev/tty || answer=''
  [[ -z "${answer}" || "${answer}" =~ ^[Yy] ]]
}

make_temp_dir() {
  TMP_DIR="$(mktemp -d 2>/dev/null || mktemp -d -t gsi-install)"
}

download_to() {
  local url="$1"
  local target="$2"

  if ! curl -fsSL -o "${target}" "${url}"; then
    die "Download failed: ${url}"
  fi
}

run_sudo() {
  if [[ "$(id -u)" -eq 0 ]]; then
    "$@"
    return
  fi

  have sudo || die 'sudo is required to install PowerShell on this system.'
  sudo "$@"
}

add_path_front() {
  local path_entry="$1"

  if [[ -d "${path_entry}" && ":${PATH}:" != *":${path_entry}:"* ]]; then
    export PATH="${path_entry}:${PATH}"
  fi
}

refresh_common_paths() {
  add_path_front '/opt/homebrew/bin'
  add_path_front '/usr/local/bin'
}

get_homebrew() {
  refresh_common_paths
  if have brew; then
    command -v brew
    return
  fi

  for candidate in /opt/homebrew/bin/brew /usr/local/bin/brew; do
    if [[ -x "${candidate}" ]]; then
      printf '%s\n' "${candidate}"
      return
    fi
  done
}

install_homebrew_if_needed() {
  local brew_path
  brew_path="$(get_homebrew || true)"
  if [[ -n "${brew_path}" ]]; then
    return
  fi

  prompt_yes_no 'Homebrew is required to install PowerShell on macOS. Install Homebrew now? [Y/n]' ||
    die 'Homebrew is required to continue.'

  info 'Installing Homebrew...'
  local installer="${TMP_DIR}/homebrew-install.sh"
  download_to 'https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh' "${installer}"

  if [[ "${GSI_INSTALL_ASSUME_YES:-}" == "1" ]]; then
    NONINTERACTIVE=1 /bin/bash "${installer}"
  else
    /bin/bash "${installer}"
  fi

  refresh_common_paths
  brew_path="$(get_homebrew || true)"
  [[ -n "${brew_path}" ]] ||
    die 'Homebrew was installed, but brew is not available in this shell. Open a new terminal and re-run the command.'
}

install_powershell_macos() {
  install_homebrew_if_needed

  local brew_path
  brew_path="$(get_homebrew || true)"
  [[ -n "${brew_path}" ]] || die 'Homebrew is required to install PowerShell on macOS.'

  prompt_yes_no 'PowerShell is required. Install it now via Homebrew? [Y/n]' ||
    die 'PowerShell is required to continue.'

  info 'Installing PowerShell via Homebrew...'
  "${brew_path}" update
  "${brew_path}" install --cask powershell
  refresh_common_paths
}

install_powershell_debian_or_ubuntu() {
  [[ -r /etc/os-release ]] || die 'Unsupported Linux distribution: /etc/os-release not found.'

  # shellcheck source=/dev/null
  . /etc/os-release

  case "${ID:-}" in
    debian | ubuntu) ;;
    *)
      die 'Automatic PowerShell installation currently supports only Debian and Ubuntu Linux.'
      ;;
  esac

  [[ -n "${VERSION_ID:-}" ]] || die 'Cannot detect Linux VERSION_ID from /etc/os-release.'
  have apt-get || die 'apt-get is required to install PowerShell on this Linux distribution.'
  have dpkg || die 'dpkg is required to install PowerShell on this Linux distribution.'

  prompt_yes_no 'PowerShell is required. Install it now via the Microsoft package repository? [Y/n]' ||
    die 'PowerShell is required to continue.'

  info 'Installing PowerShell prerequisites...'
  local packages=(curl ca-certificates apt-transport-https)
  if [[ "${ID}" == 'ubuntu' ]]; then
    packages+=(software-properties-common)
  fi

  run_sudo apt-get update
  run_sudo apt-get install -y "${packages[@]}"

  local repo_package="${TMP_DIR}/packages-microsoft-prod.deb"
  local repo_url="https://packages.microsoft.com/config/${ID}/${VERSION_ID}/packages-microsoft-prod.deb"

  info 'Registering Microsoft package repository...'
  download_to "${repo_url}" "${repo_package}"
  run_sudo dpkg -i "${repo_package}"

  info 'Installing PowerShell...'
  run_sudo apt-get update
  run_sudo apt-get install -y powershell
}

install_powershell_if_needed() {
  refresh_common_paths
  if have pwsh; then
    return
  fi

  info 'PowerShell (pwsh) is not installed.'
  case "$(uname -s)" in
    Darwin)
      install_powershell_macos
      ;;
    Linux)
      install_powershell_debian_or_ubuntu
      ;;
    *)
      die 'Unsupported OS. Install PowerShell 7+ manually and run install.ps1.'
      ;;
  esac

  refresh_common_paths
  have pwsh || die 'PowerShell was installed, but pwsh is not available in this shell. Open a new terminal and re-run the command.'
}

run_powershell_installer() {
  local install_ps1="${TMP_DIR}/install.ps1"
  download_to "${INSTALL_PS1_URL}" "${install_ps1}"

  info 'Starting Guest Sponsor Info PowerShell installer...'
  pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File "${install_ps1}" "$@"
}

main() {
  make_temp_dir

  printf '\n'
  printf '  Guest Sponsor Info  -  Bootstrap Installer\n'
  printf '  ----------------------------------------------------------\n'

  have curl || die 'curl is required to download the installer.'
  install_powershell_if_needed
  run_powershell_installer "$@"
}

main "$@"
