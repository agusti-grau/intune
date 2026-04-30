#!/bin/bash
# Intune Custom Compliance — Zscaler SIA detection
# Adjust package name if needed (e.g. zscaler-client-connector)

zs_installed="false"
zs_running="false"

if dpkg -l zscaler 2>/dev/null | grep -q "^ii"; then
    zs_installed="true"
fi

if systemctl is-active --quiet zscaler 2>/dev/null; then
    zs_running="true"
fi

cat <<EOF
{
  "ZscalerInstalled": $zs_installed,
  "ZscalerRunning": $zs_running
}
EOF
