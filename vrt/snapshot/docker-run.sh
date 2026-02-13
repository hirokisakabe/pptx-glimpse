#!/bin/bash
set -euo pipefail

cd /workspace
npm ci
exec "$@"
