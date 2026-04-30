#!/bin/bash
# Intune Custom Compliance — Kernel version detection
# Strips distro suffix, e.g. "6.11.0-28-generic" → "6.11.0"

kernel_version=$(uname -r | grep -oP '^\d+\.\d+\.\d+')

cat <<EOF
{
  "KernelVersion": "$kernel_version"
}
EOF
