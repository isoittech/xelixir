#!/usr/bin/env bash

# コンパイル用ツールスクリプト
# - /workspace/xelixir/java/src/xelixir/*.java を jars/* をクラスパスにしてコンパイル

set -eu

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
JAVA_DIR="${SCRIPT_DIR%/tools}"

cd "${JAVA_DIR}"

rm -rf dist
mkdir -p dist

javac -d dist -cp 'jars/*:src' src/xelixir/*.java
