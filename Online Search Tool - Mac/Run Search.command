#!/bin/bash
cd "$(dirname "$0")"

# Set window title
echo -n -e "\033]0;Online-Search-Tool V2\007"

echo "Checking prerequisites..."
if ! command -v node &> /dev/null
then
    echo ""
    echo "========================================================"
    echo " ERROR: Node.js is not installed!"
    echo " This tool requires Node.js to run."
    echo " Please download and install it from: https://nodejs.org/"
    echo "========================================================"
    echo ""
    echo "Press any key to exit..."
    read -n 1 -s
    exit
fi

echo "Setting up Online-Search-Tool V2..."
if [ ! -d "node_modules" ]; then
    echo "Installing dependencies..."
    npm install
    npx playwright install chromium
fi

echo "Starting CLI..."
node collect_input.js

echo ""
echo "Press any key to close..."
read -n 1 -s
