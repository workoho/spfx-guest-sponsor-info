#!/usr/bin/env bash
# Stamp a SemVer tag into package.json and config/package-solution.json.
#
# Usage:
#   scripts/set-version.sh              # interactive mode
#   scripts/set-version.sh --help       # show this help
#   scripts/set-version.sh v1.2.3           # stamp only (for CI)
#   scripts/set-version.sh v1.2.3 --commit  # stamp + git commit + git tag
#
# Both forms are accepted; a leading "v" is stripped automatically.
# SPFx requires a four-part version (major.minor.patch.build), so ".0" is
# appended for package-solution.json.
#
# Recommended release workflow (run locally, then push):
#   ./scripts/set-version.sh v1.2.3 --commit
#   git push && git push --tags
# The pushed tag triggers the release GitHub Actions workflow automatically.

set -euo pipefail

# Always run from the repository root so paths resolve correctly.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

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
  --commit    After stamping, create a git commit and annotated tag, then
              print the push command. In interactive mode you will be asked.

Options:
  -h, --help  Show this help and exit.

Interactive mode (no arguments):
  Detects the current version from the last git tag (falling back to
  package.json) and suggests next patch, minor, and major versions.

Examples:
  scripts/set-version.sh                    # interactive
  scripts/set-version.sh v1.2.3             # stamp only (CI)
  scripts/set-version.sh v1.2.3 --commit    # stamp + commit + tag
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
    # Only analyse commits that are not reachable from the base tag
    range="${base_tag}..HEAD"
  fi

  # Check whether there are any commits to analyse at all
  if ! git log "$range" --format="%s" 2>/dev/null | grep -q .; then
    echo "patch"
    return
  fi

  local bump="patch"

  # Patterns stored in variables — required for complex ERE in bash [[ =~ ]]
  local re_breaking='^[a-zA-Z]+(\([^)]*\))?!:'
  local re_feat='^feat(\([^)]*\))?:'

  # Pipe directly into the loop — bash variables cannot hold NUL bytes, so
  # storing the output in a variable via $() would silently drop all \0
  # delimiters and break the read -d $'\0' loop.
  while IFS= read -r -d $'\0' subject && IFS= read -r -d $'\0' body; do
    [[ -z "$subject" ]] && continue

    # Breaking change: exclamation mark after type/scope, e.g. feat!: or feat(x)!:
    if [[ "$subject" =~ $re_breaking ]]; then
      echo "major"
      return
    fi

    # Breaking change in footer or body (BREAKING CHANGE: or BREAKING-CHANGE:)
    if echo "$body" | grep -qiE '^(BREAKING CHANGE|BREAKING-CHANGE):'; then
      echo "major"
      return
    fi

    # Feature commit → at least minor
    if [[ "$subject" =~ $re_feat ]] && [[ "$bump" != "major" ]]; then
      bump="minor"
    fi
  done < <(git log "$range" --format="%s%x00%b%x00" 2>/dev/null)

  echo "$bump"
}

# --------------------------------------------------------------------------- #
# Argument parsing
# --------------------------------------------------------------------------- #

TAG=""
DO_COMMIT=false

for arg in "$@"; do
  case "$arg" in
    -h | --help)
      show_help
      exit 0
      ;;
    --commit)
      DO_COMMIT=true
      ;;
    -*)
      echo "Unknown option: $arg" >&2
      echo "Run '$0 --help' for usage." >&2
      exit 1
      ;;
    *)
      if [[ -z "$TAG" ]]; then
        TAG="$arg"
      else
        echo "Unexpected argument: $arg" >&2
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
    echo "Error: interactive mode requires a TTY. Pass a version explicitly." >&2
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
  LABEL_PATCH="patch  →  ${NEXT_PATCH}"
  LABEL_MINOR="minor  →  ${NEXT_MINOR}"
  LABEL_MAJOR="major  →  ${NEXT_MAJOR}"
  case "$RECOMMENDED" in
    major) LABEL_MAJOR="${LABEL_MAJOR}  ★ recommended" ;;
    minor) LABEL_MINOR="${LABEL_MINOR}  ★ recommended" ;;
    *) LABEL_PATCH="${LABEL_PATCH}  ★ recommended" ;;
  esac

  echo ""
  echo "Current version: ${CURRENT_LABEL}"
  if [[ -n "$LAST_TAG" ]]; then
    COMMIT_COUNT=$(git rev-list "${LAST_TAG}..HEAD" --count 2>/dev/null || echo "?")
    echo "Commits since ${LAST_TAG}: ${COMMIT_COUNT}"
  fi
  echo ""
  echo "Suggested next versions:"
  echo "  1) ${LABEL_PATCH}"
  echo "  2) ${LABEL_MINOR}"
  echo "  3) ${LABEL_MAJOR}"
  echo "  4) Enter a custom version"
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
          echo "No version entered, please try again." >&2
          continue
        fi
        TAG="$CUSTOM"
        break
        ;;
      *)
        echo "Invalid choice — enter 1, 2, 3, or 4." >&2
        ;;
    esac
  done

  echo ""
  read -rp "Create git commit and tag? [y/N]: " COMMIT_ANSWER
  if [[ "${COMMIT_ANSWER,,}" == "y" || "${COMMIT_ANSWER,,}" == "yes" ]]; then
    DO_COMMIT=true
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
  echo "Error: '$SEMVER' is not a valid SemVer string (expected e.g. 1.2.3)." >&2
  exit 1
fi

echo "Stamping version: semver=${SEMVER}  spfx=${SPFX_VER}"

# --------------------------------------------------------------------------- #
# Stamp files
# --------------------------------------------------------------------------- #

npm version "$SEMVER" --no-git-tag-version --allow-same-version

# Stamp azure-function/package.json if it exists
if [[ -f "azure-function/package.json" ]]; then
  npm version "$SEMVER" --no-git-tag-version --allow-same-version --prefix azure-function
  echo "azure-function/package.json → ${SEMVER}"
fi

SPFX_VER="$SPFX_VER" node -e "
const fs  = require('fs');
const ver = process.env.SPFX_VER;
const p   = 'config/package-solution.json';
const obj = JSON.parse(fs.readFileSync(p, 'utf8'));
obj.solution.version = ver;
obj.solution.features.forEach(f => { f.version = ver; });
fs.writeFileSync(p, JSON.stringify(obj, null, 2) + '\n');
console.log('config/package-solution.json → ' + ver);
"

# Compile azuredeploy.json from main.bicep so it is part of the release commit.
# The tag must point to a commit that already contains the correct ARM template —
# otherwise the Deploy-to-Azure button would serve a stale version until CI commits back.
BICEP_COMPILED=false
if command -v az &>/dev/null && [[ -f "azure-function/infra/main.bicep" ]]; then
  echo "Compiling azure-function/infra/main.bicep → azuredeploy.json via az bicep"
  az bicep build --file azure-function/infra/main.bicep --outfile azure-function/infra/azuredeploy.json
  BICEP_COMPILED=true
else
  echo "⚠ az CLI not found (or main.bicep missing) — azuredeploy.json will be regenerated by CI if needed"
fi

# --------------------------------------------------------------------------- #
# Commit and tag
# --------------------------------------------------------------------------- #

if [[ "${DO_COMMIT}" == "true" ]]; then
  git add package.json package-lock.json config/package-solution.json
  if [[ -f "azure-function/package.json" ]]; then
    git add azure-function/package.json azure-function/package-lock.json 2>/dev/null || true
  fi
  if [[ "$BICEP_COMPILED" == "true" ]]; then
    git add azure-function/infra/azuredeploy.json
  fi
  git commit -m "chore: release ${VTAG}"
  git tag -a "${VTAG}" -m "Release ${VTAG}"
  echo ""
  echo "Created commit and tag ${VTAG}."
  echo "Push with:  git push && git push --tags"
fi
