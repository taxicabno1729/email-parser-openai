#!/bin/bash

# Cleanup script to remove non-Streamlit components and set up the clean repository

echo "Cleaning up repository to keep only Streamlit components..."

# Create a backup of important files
echo "Creating backup..."
mkdir -p backup_old_app
cp -r app backup_old_app/
cp -r exports backup_old_app/ 2>/dev/null || true
cp .env backup_old_app/ 2>/dev/null || true

# Remove old app structure
echo "Removing old app structure..."
rm -rf app
rm -rf src
rm -rf public
rm -rf venv
rm -rf .venv
rm -rf __pycache__
rm -rf .git 2>/dev/null || true
rm app_streamlit.py 2>/dev/null || true
rm requirements.txt
rm tailwind.config.js 2>/dev/null || true
rm package.json 2>/dev/null || true
rm test_imports.py 2>/dev/null || true
rm README_STREAMLIT.md 2>/dev/null || true

# Rename files
echo "Setting up new structure..."
mv streamlit_app.py app.py
mv requirements_streamlit.txt requirements.txt

# Create exports directory if it doesn't exist
mkdir -p exports

echo "Cleanup complete! Your Streamlit app is ready."
echo "Run 'streamlit run app.py' to start the application." 