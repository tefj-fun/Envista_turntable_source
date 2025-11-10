#!/bin/bash
# WSL wrapper for build.cmd

CONFIG="${1:-Debug}"
PLATFORM="${2:-x64}"

cd "/mnt/c/Program Files/Solomon Technology Corp/SolVision6/Version_6.1.4/SampleCode/C_Sharp/DemoApp" || exit 1

cmd.exe /c build.cmd "$CONFIG" "$PLATFORM"
