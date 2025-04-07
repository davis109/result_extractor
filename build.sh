#!/bin/bash

# Exit on error
set -e

# Install Python dependencies
echo "Installing Python dependencies..."
pip install -r requirements.txt

# Make Chrome installer executable
echo "Making Chrome installer executable..."
chmod +x chrome-installer.sh

# Install Chrome
echo "Installing Chrome..."
./chrome-installer.sh

# Print Chrome version
echo "Chrome version:"
google-chrome --version || echo "Chrome version check failed, but continuing..."

# Print ChromeDriver version
echo "ChromeDriver info (from webdriver-manager):"
python -c "from webdriver_manager.chrome import ChromeDriverManager; print(ChromeDriverManager().driver_version)" || echo "ChromeDriver version check failed, but continuing..."

# Create necessary directories
echo "Creating template directories if they don't exist..."
mkdir -p templates
mkdir -p static

echo "Build script completed successfully." 