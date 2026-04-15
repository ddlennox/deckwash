#!/bin/bash
# DeckWash Launcher
# Double-click this file to start DeckWash in your browser.

cd "$(dirname "$0")"

# Check Python is available
if ! command -v python3 &>/dev/null; then
  osascript -e 'display alert "Python 3 not found" message "Please install Python 3 from python.org and try again."'
  exit 1
fi

# Install dependencies if needed (silently)
python3 -m pip install flask lxml --break-system-packages -q 2>/dev/null || \
python3 -m pip install flask lxml -q 2>/dev/null

echo "🌊 Starting DeckWash..."
python3 deckwash.py
