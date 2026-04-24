FROM mcr.microsoft.com/devcontainers/base:trixie

ARG NODE_VERSION=22.14.0
ARG TARGETARCH

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update \
  && apt-get install -y --no-install-recommends \
  bat \
  build-essential \
  ca-certificates \
  curl \
  fd-find \
  fzf \
  git \
  jq \
  python-is-python3 \
  python3 \
  python3-venv \
  ripgrep \
  ruby-dev \
  shellcheck \
  unzip \
  xz-utils \
  && rm -rf /var/lib/apt/lists/* \
  # Debian installs bat as "batcat" and fd-find as "fdfind" to avoid name
  # conflicts. Create canonical symlinks so scripts/agents can use bat and fd.
  && ln -s /usr/bin/batcat /usr/local/bin/bat \
  && ln -s /usr/bin/fdfind /usr/local/bin/fd

# yq (mikefarah/yq v4) — the apt package is a Python wrapper with different
# syntax. Install the official Go binary directly for `yq eval` compatibility.
ARG YQ_VERSION=v4.44.6
RUN arch="${TARGETARCH:-$(dpkg --print-architecture)}" \
  && case "${arch}" in \
  amd64|x86_64) yq_arch="amd64" ;; \
  arm64|aarch64) yq_arch="arm64" ;; \
  *) echo "Unsupported architecture: ${arch}" >&2; exit 1 ;; \
  esac \
  && curl -fsSLo /usr/local/bin/yq \
  --retry 3 --retry-delay 2 --retry-all-errors \
  "https://github.com/mikefarah/yq/releases/download/${YQ_VERSION}/yq_linux_${yq_arch}" \
  && chmod +x /usr/local/bin/yq

RUN arch="${TARGETARCH:-$(dpkg --print-architecture)}" \
  && case "${arch}" in \
  amd64|x86_64) node_arch="x64" ;; \
  arm64|aarch64) node_arch="arm64" ;; \
  *) echo "Unsupported architecture: ${arch}" >&2; exit 1 ;; \
  esac \
  && curl -fsSLO --retry 3 --retry-delay 2 --retry-all-errors \
  "https://nodejs.org/dist/v${NODE_VERSION}/node-v${NODE_VERSION}-linux-${node_arch}.tar.xz" \
  && curl -fsSLO --retry 3 --retry-delay 2 --retry-all-errors \
  "https://nodejs.org/dist/v${NODE_VERSION}/SHASUMS256.txt" \
  && grep " node-v${NODE_VERSION}-linux-${node_arch}.tar.xz$" SHASUMS256.txt | sha256sum -c - \
  && tar -xJf "node-v${NODE_VERSION}-linux-${node_arch}.tar.xz" -C /usr/local --strip-components=1 --no-same-owner \
  && rm "node-v${NODE_VERSION}-linux-${node_arch}.tar.xz" SHASUMS256.txt \
  && corepack enable

# Keep the global SPFx generator version in sync with package.json.
COPY package.json /tmp/package.json

RUN SPFX_VERSION="$(jq -r '.config.spfxVersion // .dependencies["@microsoft/sp-core-library"]' /tmp/package.json)" \
  && npm install --global @rushstack/heft yo "@microsoft/generator-sharepoint@${SPFX_VERSION}" \
  && rm -f /tmp/package.json \
  && npm cache clean --force

# Azure Functions Core Tools v4 — enables `func start` for local function development.
RUN npm install --global azure-functions-core-tools@4 \
  && npm cache clean --force

# APM CLI — reproducible agent package management for shared skills,
# instructions, and MCP dependencies across similar repositories.
ARG APM_VERSION=0.9.2
RUN python3 -m venv /opt/apm \
  && /opt/apm/bin/pip install --no-cache-dir --upgrade pip \
  && /opt/apm/bin/pip install --no-cache-dir "apm-cli==${APM_VERSION}" \
  && ln -s /opt/apm/bin/apm /usr/local/bin/apm

# Playwright browser runtime — installs Chromium plus the Linux libraries
# needed by Playwright/Playwright MCP in this container.
RUN npx -y playwright@latest install --with-deps chromium \
  && npm cache clean --force

# delta — syntax-highlighting git diff pager (https://github.com/dandavison/delta).
# Installed as a static binary; makes `git diff` and `git show` output readable
# with colour and line numbers in the terminal — useful for both developers and agents.
ARG DELTA_VERSION=0.19.1
RUN arch="${TARGETARCH:-$(dpkg --print-architecture)}" \
  && case "${arch}" in \
  # x86_64 has a musl (fully static) build; arm64 ships gnu only
  amd64|x86_64) delta_arch="x86_64" ; delta_libc="musl" ;; \
  arm64|aarch64) delta_arch="aarch64" ; delta_libc="gnu" ;; \
  *) echo "Unsupported architecture: ${arch}" >&2; exit 1 ;; \
  esac \
  && curl -fsSLo /tmp/delta.tar.gz \
  --retry 3 --retry-delay 2 --retry-all-errors \
  "https://github.com/dandavison/delta/releases/download/${DELTA_VERSION}/delta-${DELTA_VERSION}-${delta_arch}-unknown-linux-${delta_libc}.tar.gz" \
  && tar -xzf /tmp/delta.tar.gz -C /usr/local/bin \
  --strip-components=1 \
  "delta-${DELTA_VERSION}-${delta_arch}-unknown-linux-${delta_libc}/delta" \
  && rm /tmp/delta.tar.gz

# shfmt — shell script formatter (https://github.com/mvdan/sh).
# Installed as a static binary; no Go toolchain required at runtime.
ARG SHFMT_VERSION=v3.10.0
RUN arch="${TARGETARCH:-$(dpkg --print-architecture)}" \
  && case "${arch}" in \
  amd64|x86_64) shfmt_arch="amd64" ;; \
  arm64|aarch64) shfmt_arch="arm64" ;; \
  *) echo "Unsupported architecture: ${arch}" >&2; exit 1 ;; \
  esac \
  && curl -fsSLo /usr/local/bin/shfmt \
  --retry 3 --retry-delay 2 --retry-all-errors \
  "https://github.com/mvdan/sh/releases/download/${SHFMT_VERSION}/shfmt_${SHFMT_VERSION}_linux_${shfmt_arch}" \
  && chmod +x /usr/local/bin/shfmt

# actionlint — GitHub Actions workflow linter (https://github.com/rhysd/actionlint).
# Validates action names, expression types, job/step schema, and permission scopes —
# catches semantic errors in workflow files that YAML formatting tools cannot detect.
ARG ACTIONLINT_VERSION=1.7.7
RUN arch="${TARGETARCH:-$(dpkg --print-architecture)}" \
  && case "${arch}" in \
  amd64|x86_64) al_arch="amd64" ;; \
  arm64|aarch64) al_arch="arm64" ;; \
  *) echo "Unsupported architecture: ${arch}" >&2; exit 1 ;; \
  esac \
  && curl -fsSLo /tmp/actionlint.tar.gz \
  --retry 3 --retry-delay 2 --retry-all-errors \
  "https://github.com/rhysd/actionlint/releases/download/v${ACTIONLINT_VERSION}/actionlint_${ACTIONLINT_VERSION}_linux_${al_arch}.tar.gz" \
  && tar -xzf /tmp/actionlint.tar.gz -C /usr/local/bin actionlint \
  && rm /tmp/actionlint.tar.gz


