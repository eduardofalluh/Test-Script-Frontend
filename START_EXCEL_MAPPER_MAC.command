#!/bin/bash

set -e

cd "$(dirname "$0")"

if ! command -v node >/dev/null 2>&1; then
  echo "Node.js is required to run Excel Mapper."
  echo "Install the LTS version from https://nodejs.org, then run this file again."
  read -r -p "Press Enter to close."
  exit 1
fi

if ! command -v npm >/dev/null 2>&1; then
  echo "npm is required to run Excel Mapper. It is included with Node.js."
  read -r -p "Press Enter to close."
  exit 1
fi

if [ ! -d "node_modules" ]; then
  echo "Installing Excel Mapper dependencies. This can take a few minutes the first time."
  npm install
fi

(sleep 3 && open "http://localhost:3010") >/dev/null 2>&1 &
npm run dev:local
