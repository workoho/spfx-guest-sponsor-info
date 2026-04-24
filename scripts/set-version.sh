#!/usr/bin/env bash
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>
#
# Stamp a SemVer tag into package.json and config/package-solution.json.
#
# Usage:
#   scripts/set-version.sh              # interactive mode
#   scripts/set-version.sh --help       # show this help
#   scripts/set-version.sh v1.2.3           # stamp only (for CI)
#   scripts/set-version.sh v1.2.3 --commit  # stamp + git commit + git tag
#   scripts/set-version.sh v1.2.3 --commit --push  # stamp + commit + tag + push
#
# Both forms are accepted; a leading "v" is stripped automatically.
# SPFx requires a four-part version (major.minor.patch.build), so ".0" is
# appended for package-solution.json.
#
# Recommended release workflow:
#   ./scripts/set-version.sh v1.2.3 --commit --push
# The pushed tag triggers the release GitHub Actions workflow automatically.

set -euo pipefail

# Always run from the repository root so paths resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

# --------------------------------------------------------------------------- #
# Terminal formatting
# --------------------------------------------------------------------------- #
# shellcheck source=scripts/colors.sh
source "$(dirname "${BASH_SOURCE[0]}")/colors.sh"

# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

show_help() {
  cat <<'EOF'
Usage: scripts/set-version.sh [OPTIONS] [<version>] [--commit]

Stamp a SemVer version into package.json, azure-function/package.json,
and config/package-solution.json.

Arguments:
  <version>   Target version, e.g. v1.2.3 or 1.2.3. A leading "v" is stripped
              automatically. When omitted, interactive mode starts.
  --commit    After stamping, create a git commit and annotated tag.
              In interactive mode you will be asked.
  --push      After committing and tagging, push branch and tags to origin.
              Implies --commit. In interactive mode you will be asked.

Options:
  -h, --help  Show this help and exit.

Interactive mode (no arguments):
  Detects the current version from the last git tag (falling back to
  package.json) and suggests next patch, minor, and major versions.

Examples:
  scripts/set-version.sh                         # interactive
  scripts/set-version.sh v1.2.3                  # stamp only (CI)
  scripts/set-version.sh v1.2.3 --commit         # stamp + commit + tag
  scripts/set-version.sh v1.2.3 --commit --push  # stamp + commit + tag + push
EOF
}

# Prints "X.Y.Z (from git tag vX.Y.Z)" or "X.Y.Z (from package.json)"
get_current_label() {
  local tag
  tag=$(git describe --tags --match "v*" --abbrev=0 2>/dev/null || true)
  if [[ -n "$tag" ]]; then
    printf '%s (from git tag %s)' "${tag#v}" "$tag"
    return
  fi
  local ver
  ver=$(node -p "require('./package.json').version" 2>/dev/null || true)
  if [[ -n "$ver" ]]; then
    printf '%s (from package.json)' "$ver"
    return
  fi
  echo "unknown"
}

# Returns just the bare semver string (no label)
get_current_semver() {
  local tag
  tag=$(git describe --tags --match "v*" --abbrev=0 2>/dev/null || true)
  if [[ -n "$tag" ]]; then
    echo "${tag#v}"
    return
  fi
  node -p "require('./package.json').version" 2>/dev/null || echo "0.0.0"
}

# bump_version <x.y.z> <patch|minor|major>
bump_version() {
  local version="$1" bump="$2"
  local major minor patch
  IFS='.' read -r major minor patch <<<"$version"
  patch="${patch%%[-+]*}" # strip any pre-release/build suffix
  case "$bump" in
    major) echo "$((major + 1)).0.0" ;;
    minor) echo "${major}.$((minor + 1)).0" ;;
    patch) echo "${major}.${minor}.$((patch + 1))" ;;
  esac
}

# suggest_bump [<base_tag>]
# Analyses Conventional Commits since <base_tag> (or the last git tag) and
# returns "major", "minor", or "patch".
suggest_bump() {
  local base_tag="${1:-}"
  if [[ -z "$base_tag" ]]; then
    base_tag=$(git describe --tags --match "v*" --abbrev=0 2>/dev/null || true)
  fi

  local range="HEAD"
  if [[ -n "$base_tag" ]]; then
    range="${base_tag}..HEAD"
  fi

  # Collect all one-line subjects (safe in a variable — no NUL bytes needed)
  local subjects
  subjects=$(git log "$range" --format="%s" 2>/dev/null || true)

  if [[ -z "$subjects" ]]; then
    echo "patch"
    return
  fi

  # Breaking change in subject line: type!: or type(scope)!:
  # Each grep runs inside an if-condition, so set -e does not apply.
  if echo "$subjects" | grep -qE '^[a-zA-Z]+(\([^)]*\))?!:'; then
    echo "major"
    return
  fi

  # Breaking change token in any commit body / footer
  if git log "$range" --format="%b" 2>/dev/null |
    grep -qiE '^(BREAKING CHANGE|BREAKING-CHANGE):'; then
    echo "major"
    return
  fi

  # Feature commit → minor
  if echo "$subjects" | grep -qE '^feat(\([^)]*\))?:'; then
    echo "minor"
    return
  fi

  echo "patch"
}

# --------------------------------------------------------------------------- #
# Argument parsing
# --------------------------------------------------------------------------- #

TAG=""
DO_COMMIT=false
DO_PUSH=false

for arg in "$@"; do
  case "$arg" in
    -h | --help)
      show_help
      exit 0
      ;;
    --commit)
      DO_COMMIT=true
      ;;
    --push)
      DO_COMMIT=true
      DO_PUSH=true
      ;;
    -*)
      echo "${C_RED}Unknown option:${C_RST} $arg" >&2
      gha_error "Unknown option: $arg"
      echo "Run '$0 --help' for usage." >&2
      exit 1
      ;;
    *)
      if [[ -z "$TAG" ]]; then
        TAG="$arg"
      else
        echo "${C_RED}Unexpected argument:${C_RST} $arg" >&2
        gha_error "Unexpected argument: $arg"
        echo "Run '$0 --help' for usage." >&2
        exit 1
      fi
      ;;
  esac
done

# --------------------------------------------------------------------------- #
# Interactive mode (no version argument)
# --------------------------------------------------------------------------- #

if [[ -z "$TAG" ]]; then
  if [[ ! -t 0 ]]; then
    echo "${C_RED}Error:${C_RST} interactive mode requires a TTY. Pass a version explicitly." >&2
    echo "Run '$0 --help' for usage." >&2
    exit 1
  fi

  CURRENT_LABEL=$(get_current_label)
  CURRENT_SEMVER=$(get_current_semver)
  NEXT_PATCH=$(bump_version "$CURRENT_SEMVER" patch)
  NEXT_MINOR=$(bump_version "$CURRENT_SEMVER" minor)
  NEXT_MAJOR=$(bump_version "$CURRENT_SEMVER" major)

  # Determine recommended bump from Conventional Commits since last tag
  LAST_TAG=$(git describe --tags --match "v*" --abbrev=0 2>/dev/null || true)
  RECOMMENDED=$(suggest_bump "$LAST_TAG")

  # Map recommendation to menu choice number and default version
  case "$RECOMMENDED" in
    major)
      DEFAULT_CHOICE=3
      DEFAULT_VER="$NEXT_MAJOR"
      ;;
    minor)
      DEFAULT_CHOICE=2
      DEFAULT_VER="$NEXT_MINOR"
      ;;
    *)
      DEFAULT_CHOICE=1
      DEFAULT_VER="$NEXT_PATCH"
      ;;
  esac

  # Build menu labels — mark recommended entry with ★
  LABEL_PATCH="patch  →  ${C_CYN}${NEXT_PATCH}${C_RST}"
  LABEL_MINOR="minor  →  ${C_CYN}${NEXT_MINOR}${C_RST}"
  LABEL_MAJOR="major  →  ${C_CYN}${NEXT_MAJOR}${C_RST}"
  case "$RECOMMENDED" in
    major) LABEL_MAJOR="${LABEL_MAJOR}  ${C_YLW}★ recommended${C_RST}" ;;
    minor) LABEL_MINOR="${LABEL_MINOR}  ${C_YLW}★ recommended${C_RST}" ;;
    *) LABEL_PATCH="${LABEL_PATCH}  ${C_YLW}★ recommended${C_RST}" ;;
  esac

  echo ""
  echo "${C_BLD}Current version:${C_RST} ${C_CYN}${CURRENT_LABEL}${C_RST}"
  if [[ -n "$LAST_TAG" ]]; then
    COMMIT_COUNT=$(git rev-list "${LAST_TAG}..HEAD" --count 2>/dev/null || echo "?")
    echo "${C_DIM}Commits since ${LAST_TAG}: ${COMMIT_COUNT}${C_RST}"
  fi
  echo ""
  echo "${C_BLD}Suggested next versions:${C_RST}"
  echo "  ${C_BLD}${C_YLW}1)${C_RST} ${LABEL_PATCH}"
  echo "  ${C_BLD}${C_YLW}2)${C_RST} ${LABEL_MINOR}"
  echo "  ${C_BLD}${C_YLW}3)${C_RST} ${LABEL_MAJOR}"
  echo "  ${C_BLD}${C_YLW}4)${C_RST} Enter a custom version"
  echo ""

  while true; do
    read -rp "Select [1-4] or press Enter for recommended (${DEFAULT_VER}): " CHOICE
    CHOICE="${CHOICE:-${DEFAULT_CHOICE}}"
    case "$CHOICE" in
      1)
        TAG="$NEXT_PATCH"
        break
        ;;
      2)
        TAG="$NEXT_MINOR"
        break
        ;;
      3)
        TAG="$NEXT_MAJOR"
        break
        ;;
      4)
        read -rp "Enter version (e.g. 1.2.3 or v1.2.3): " CUSTOM
        if [[ -z "$CUSTOM" ]]; then
          echo "${C_YLW}No version entered, please try again.${C_RST}" >&2
          continue
        fi
        TAG="$CUSTOM"
        break
        ;;
      *)
        echo "${C_YLW}Invalid choice — enter 1, 2, 3, or 4.${C_RST}" >&2
        ;;
    esac
  done

  echo ""
  read -rp "Create git commit and tag? [y/N]: " COMMIT_ANSWER
  if [[ "${COMMIT_ANSWER,,}" == "y" || "${COMMIT_ANSWER,,}" == "yes" ]]; then
    DO_COMMIT=true
    read -rp "Also push to origin? [y/N]: " PUSH_ANSWER
    if [[ "${PUSH_ANSWER,,}" == "y" || "${PUSH_ANSWER,,}" == "yes" ]]; then
      DO_PUSH=true
    fi
  fi
  echo ""
fi

# --------------------------------------------------------------------------- #
# Validate and normalise
# --------------------------------------------------------------------------- #

SEMVER="${TAG#v}"      # strip leading "v" if present
VTAG="v${SEMVER}"      # ensure "v" prefix for the git tag
SPFX_VER="${SEMVER}.0" # SPFx requires four-part version (major.minor.patch.build)

if ! [[ "$SEMVER" =~ ^[0-9]+\.[0-9]+\.[0-9]+([.-][a-zA-Z0-9.]+)?$ ]]; then
  echo "${C_RED}Error:${C_RST} '${C_BLD}${SEMVER}${C_RST}' is not a valid SemVer string (expected e.g. 1.2.3)." >&2
  gha_error "Invalid SemVer string: ${SEMVER}"
  exit 1
fi

echo "${C_BLD}Stamping version:${C_RST} ${C_CYN}${SEMVER}${C_RST}  ${C_DIM}(SPFx: ${SPFX_VER})${C_RST}"
gha_notice "Stamping version ${SEMVER} (SPFx ${SPFX_VER})."

# --------------------------------------------------------------------------- #
# Stamp files
# --------------------------------------------------------------------------- #

npm version "$SEMVER" --no-git-tag-version --allow-same-version

# Stamp azure-function/package.json if it exists
if [[ -f "azure-function/package.json" ]]; then
  npm version "$SEMVER" --no-git-tag-version --allow-same-version --prefix azure-function
  echo "${C_GRN}✓${C_RST} azure-function/package.json → ${C_CYN}${SEMVER}${C_RST}"
fi

SPFX_VER="$SPFX_VER" node -e "
const fs  = require('fs');
const ver = process.env.SPFX_VER;
const p   = 'config/package-solution.json';
const obj = JSON.parse(fs.readFileSync(p, 'utf8'));
obj.solution.version = ver;
obj.solution.features.forEach(f => { f.version = ver; });
fs.writeFileSync(p, JSON.stringify(obj, null, 2) + '\n');
"
echo "${C_GRN}✓${C_RST} config/package-solution.json → ${C_CYN}${SPFX_VER}${C_RST}"

# --------------------------------------------------------------------------- #
# Commit and tag
# --------------------------------------------------------------------------- #

if [[ "${DO_COMMIT}" == "true" ]]; then
  git add package.json package-lock.json config/package-solution.json
  if [[ -f "azure-function/package.json" ]]; then
    git add azure-function/package.json azure-function/package-lock.json 2>/dev/null || true
  fi

  RETAG=false
  if git diff --cached --quiet; then
    # Nothing staged — version files are already at SEMVER (re-tagging scenario).
    # Move the tag to HEAD without creating an empty commit.
    echo "${C_YLW}⚠${C_RST}  No file changes (already at ${C_CYN}${SEMVER}${C_RST}) — moving tag ${C_BLD}${VTAG}${C_RST} to HEAD."
    git tag -f -a "${VTAG}" -m "Release ${VTAG}"
    RETAG=true
  else
    git commit -m "chore: release ${VTAG}"
    git tag -a "${VTAG}" -m "Release ${VTAG}"
  fi

  echo ""
  echo "${C_GRN}✓${C_RST} Tag ${C_BLD}${VTAG}${C_RST} set at HEAD."
  if [[ "${DO_PUSH}" == "true" ]]; then
    echo "${C_DIM}Pushing to origin…${C_RST}"
    git push
    if [[ "$RETAG" == "true" ]]; then
      # The tag already exists on the remote — force-push it to its new position.
      git push --force origin "refs/tags/${VTAG}"
    else
      git push --tags
    fi
    echo "${C_GRN}✓${C_RST} Pushed. The release workflow will start automatically."
  else
    if [[ "$RETAG" == "true" ]]; then
      next_steps "Push with:" \
        "  ${C_BLD}git push && git push --force origin refs/tags/${VTAG}${C_RST}"
    else
      next_steps "Push with:" \
        "  ${C_BLD}git push && git push --tags${C_RST}"
    fi
  fi
fi

if [[ "${DO_COMMIT}" == "true" ]]; then
  gha_notice "Version ${VTAG} prepared successfully."
fi
