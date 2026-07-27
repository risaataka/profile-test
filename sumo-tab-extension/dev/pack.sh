#!/bin/sh
# 配布用の zip を作る。
#   sh dev/pack.sh            → dist/chanko-tab-rikishi-<version>.zip
# 展開してできるフォルダを chrome://extensions から読み込めば動く。
set -e

cd "$(dirname "$0")/.."
root=$(pwd)
name=chanko-tab-rikishi
version=$(sed -n 's/.*"version"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' manifest.json | head -1)
out="$root/dist/$name-$version.zip"

rm -rf "$root/dist"
mkdir -p "$root/dist/$name"
cp -R manifest.json README.md src icons "$root/dist/$name/"

cd "$root/dist"
zip -qr "$out" "$name"
rm -rf "$root/dist/$name"
echo "$out"
