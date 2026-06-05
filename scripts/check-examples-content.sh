#!/usr/bin/env bash
#
# Compares the decompressed contents of every examples/*.xlsx against the same
# file at HEAD. The .xlsx ZIP wrapper carries a 2-second-precision DOS timestamp
# that changes on every run, so a binary diff is too strict. The actual XML
# payload is deterministic and is what we want to lock as a golden file.
#
# Exits non-zero (with a printed diff) if any example's content has changed.

set -euo pipefail

failed=0
tmp=$(mktemp -d)
trap 'rm -rf "$tmp"' EXIT

shopt -s nullglob
xlsx_files=(examples/*.xlsx)

if [ "${#xlsx_files[@]}" -eq 0 ]; then
	echo "No examples/*.xlsx found — nothing to check."
	exit 0
fi

for xlsx in "${xlsx_files[@]}"; do
	name=$(basename "$xlsx" .xlsx)
	expected="$tmp/$name.expected"
	actual="$tmp/$name.actual"
	mkdir -p "$expected" "$actual"

	# Extract the committed (HEAD) version.
	git show "HEAD:$xlsx" > "$tmp/$name.head.xlsx"
	unzip -q "$tmp/$name.head.xlsx" -d "$expected"

	# Extract the regenerated (working tree) version.
	unzip -q "$xlsx" -d "$actual"

	if ! diff -r "$expected" "$actual" > /dev/null 2>&1; then
		echo "❌ $xlsx content differs from HEAD"
		diff -r "$expected" "$actual" || true
		failed=1
	fi
done

if [ "$failed" -eq 0 ]; then
	echo "✅ All examples/*.xlsx contents match HEAD"
else
	exit 1
fi
