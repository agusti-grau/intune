#!/bin/bash
# Intune Custom Compliance — CrowdStrike Falcon detection

cs_installed="false"
cs_running="false"

if dpkg -l falcon-sensor 2>/dev/null | grep -q "^ii"; then
    cs_installed="true"
fi

if systemctl is-active --quiet falcon-sensor 2>/dev/null; then
    cs_running="true"
fi

cat <<EOF
{
  "CrowdStrikeInstalled": $cs_installed,
  "CrowdStrikeRunning": $cs_running
}
EOF
