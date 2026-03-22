FROM mcr.microsoft.com/devcontainers/base:bookworm

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
  ripgrep \
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
  "https://github.com/mikefarah/yq/releases/download/${YQ_VERSION}/yq_linux_${yq_arch}" \
  && chmod +x /usr/local/bin/yq

RUN arch="${TARGETARCH:-$(dpkg --print-architecture)}" \
  && case "${arch}" in \
  amd64|x86_64) node_arch="x64" ;; \
  arm64|aarch64) node_arch="arm64" ;; \
  *) echo "Unsupported architecture: ${arch}" >&2; exit 1 ;; \
  esac \
  && curl -fsSLO "https://nodejs.org/dist/v${NODE_VERSION}/node-v${NODE_VERSION}-linux-${node_arch}.tar.xz" \
  && curl -fsSLO "https://nodejs.org/dist/v${NODE_VERSION}/SHASUMS256.txt" \
  && grep " node-v${NODE_VERSION}-linux-${node_arch}.tar.xz$" SHASUMS256.txt | sha256sum -c - \
  && tar -xJf "node-v${NODE_VERSION}-linux-${node_arch}.tar.xz" -C /usr/local --strip-components=1 --no-same-owner \
  && rm "node-v${NODE_VERSION}-linux-${node_arch}.tar.xz" SHASUMS256.txt \
  && corepack enable \
  && npm install --global yo @microsoft/generator-sharepoint@1.22.2 \
  && npm cache clean --force

# Azure Functions Core Tools v4 — enables `func start` for local function development.
RUN npm install --global azure-functions-core-tools@4 \
  && npm cache clean --force


