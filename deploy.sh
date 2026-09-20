#!/usr/bin/env bash

set -euo pipefail

cd "$(dirname "$0")"
REGISTRY="${REGISTRY:-docker.io/andpra70}"
IMAGE_NAME="${IMAGE_NAME:-catalogo-opere}"
IMAGE_REF="${REGISTRY}/${IMAGE_NAME}"
VERSION_TAG="${TAG:-${1:-latest}}"

docker build -t "${IMAGE_REF}:${VERSION_TAG}" -t "${IMAGE_REF}:latest" .
docker push "${IMAGE_REF}:${VERSION_TAG}"

if [[ "${VERSION_TAG}" != "latest" ]]; then
  docker push "${IMAGE_REF}:latest"
fi
