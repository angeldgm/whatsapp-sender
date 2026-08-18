#!/usr/bin/env bash

set -e

echo "========================================"
echo "  Node.js / Express Server Startup"
echo "========================================"

# Check if Node.js is installed
if command -v node >/dev/null 2>&1; then
    echo "Node.js is already installed."
    node --version
else
    echo "Node.js is not installed."
    echo "Installing Node.js LTS using nvm..."

    # Install nvm if it is not already installed
    if [ ! -d "$HOME/.nvm" ]; then
        echo "Installing nvm..."

        curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.40.3/install.sh | bash

        export NVM_DIR="$HOME/.nvm"

        # Load nvm
        if [ -s "$NVM_DIR/nvm.sh" ]; then
            . "$NVM_DIR/nvm.sh"
        else
            echo "ERROR: nvm installation failed."
            exit 1
        fi
    else
        export NVM_DIR="$HOME/.nvm"

        if [ -s "$NVM_DIR/nvm.sh" ]; then
            . "$NVM_DIR/nvm.sh"
        fi
    fi

    # Install and use the latest LTS Node.js
    nvm install --lts
    nvm use --lts

    echo "Node.js installed successfully."
    node --version
fi

echo
echo "Installing npm dependencies..."
npm install

echo
echo "Starting server..."
echo

node server.js
