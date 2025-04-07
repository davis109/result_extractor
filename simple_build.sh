#!/bin/bash

# Exit on error
set -e

# Install Python dependencies
echo "Installing Python dependencies..."
pip install -r simple_requirements.txt

# Create necessary directories (just in case)
echo "Creating template directories if they don't exist..."
mkdir -p templates
mkdir -p static

echo "Build script completed successfully." 