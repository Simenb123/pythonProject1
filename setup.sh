#!/bin/bash
# Simple setup script for the bilagsverktøy project
# Creates a virtual environment and installs all dependencies
set -e

# Default venv directory
VENV_DIR=".venv"

python3 -m venv "$VENV_DIR"
source "$VENV_DIR/bin/activate"
pip install --upgrade pip
pip install -r requirements.txt

echo "\nSetup complete. Activate the environment with 'source $VENV_DIR/bin/activate'"
