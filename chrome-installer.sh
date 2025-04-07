#!/bin/bash

# Check if Chrome is already installed
if [ -f "/opt/google/chrome/chrome" ]; then
    echo "Chrome is already installed. Skipping installation."
    exit 0
fi

# Install dependencies
apt-get update
apt-get install -y wget gnupg2 unzip

# Download and install Chrome
wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | apt-key add -
echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google-chrome.list
apt-get update
apt-get install -y google-chrome-stable

# Verify installation
google-chrome --version

echo "Chrome installation completed successfully." 