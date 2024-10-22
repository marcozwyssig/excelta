#!/bin/bash

# Define the virtual environment directory
VENV_DIR="myenv"

# Function to reset the virtual environment by deleting it
reset_venv() {
  if [ -d "$VENV_DIR" ]; then
    echo "Resetting virtual environment directory..."
    rm -rf "$VENV_DIR"
    echo "Virtual environment directory reset."
  else
    echo "Virtual environment directory does not exist."
  fi
}

# Check if the --reset or -r argument is provided
if [ "$1" == "--reset" ] || [ "$1" == "-r" ]; then
  reset_venv
fi

# Create the virtual environment if it doesn't exist
if [ ! -d "$VENV_DIR" ]; then
  python3 -m venv "$VENV_DIR"
fi

# Activate the virtual environment
source "$VENV_DIR/bin/activate"

# Install required packages
pip3 install --upgrade pip
pip3 install -r requirements.txt

# Keep the virtual environment activated (for demonstration purposes, otherwise you can deactivate)
exec $SHELL
