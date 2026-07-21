#!/bin/bash
set -e

echo "=== Building Python Dependencies ==="

# Check if python3 is available
if ! command -v python3 &> /dev/null; then
    echo "python3 is not installed. Attempting to install..."
    apt-get update && apt-get install -y python3 python3-distutils python3-apt
fi

# Check if pip is installed
if ! python3 -m pip --version &>/dev/null; then
    echo "pip not found. Downloading get-pip.py..."
    curl -sS https://bootstrap.pypa.io/get-pip.py -o get-pip.py
    python3 get-pip.py --no-warn-script-location || python3 get-pip.py --user --no-warn-script-location
    rm -f get-pip.py
fi

echo "Installing requirements.txt..."
python3 -m pip install -r requirements.txt --no-cache-dir || python3 -m pip install -r requirements.txt --user --no-cache-dir

echo "Python dependencies installed successfully."
