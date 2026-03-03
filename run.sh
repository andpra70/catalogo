#!/usr/bin/env bash

set -euo pipefail

REGISTRY="docker.io"
DOCKER_USER="andpra70"
IMAGE_NAME="catalogo-opere"
CONTAINER_NAME="catalogo-opere"
IMAGE_REF="${REGISTRY}/${DOCKER_USER}/${IMAGE_NAME}"
VERSION_TAG="${1:-latest}"

docker pull "${IMAGE_REF}:${VERSION_TAG}"

if docker ps -a --format '{{.Names}}' | grep -Fxq "${CONTAINER_NAME}"; then
  docker rm -f "${CONTAINER_NAME}"
fi

docker run -d \
  --name "${CONTAINER_NAME}" \
  --restart unless-stopped \
  -p 6063:6063 \
  "${IMAGE_REF}:${VERSION_TAG}"
