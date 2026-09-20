#!/usr/bin/env bash

set -euo pipefail

REGISTRY="${REGISTRY:-docker.io/andpra70}"
IMAGE_NAME="${IMAGE_NAME:-catalogo-opere}"
CONTAINER_NAME="${CONTAINER_NAME:-catalogo-opere}"
IMAGE_REF="${IMAGE:-${REGISTRY}/${IMAGE_NAME}}"
VERSION_TAG="${TAG:-${1:-latest}}"
HOST_PORT="${HOST_PORT:-6063}"
CONTAINER_PORT="${CONTAINER_PORT:-8080}"

docker pull "${IMAGE_REF}:${VERSION_TAG}"

if docker ps -a --format '{{.Names}}' | grep -Fxq "${CONTAINER_NAME}"; then
  docker rm -f "${CONTAINER_NAME}"
fi

docker run -d \
  --name "${CONTAINER_NAME}" \
  --restart unless-stopped \
  -p "${HOST_PORT}:${CONTAINER_PORT}" \
  "${IMAGE_REF}:${VERSION_TAG}"

echo "Catalogo disponibile su http://localhost:${HOST_PORT}"
