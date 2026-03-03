#!/usr/bin/env bash

set -euo pipefail

REGISTRY="docker.io"
DOCKER_USER="andpra70"
IMAGE_NAME="catalogo-opere"
IMAGE_REF="${REGISTRY}/${DOCKER_USER}/${IMAGE_NAME}"
VERSION_TAG="${1:-latest}"

docker build -t "${IMAGE_REF}:${VERSION_TAG}" -t "${IMAGE_REF}:latest" .
docker push "${IMAGE_REF}:${VERSION_TAG}"

if [[ "${VERSION_TAG}" != "latest" ]]; then
  docker push "${IMAGE_REF}:latest"
fi
