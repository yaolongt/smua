#!/bin/bash

# Change to the source directory
cd src

# Echo a message to indicate that Python dependencies are being installed
echo "Installing Python dependencies..."

# Install Python dependencies using pip3
pip3 install -r ./requirements.txt

# Check if the pip3 command exited successfully
if [[ $? -ne 0 ]]; then
    # If the pip3 command failed, echo an error message and exit the script
    echo "Failed to install dependencies. Exiting..."
    exit $?
fi

# Echo a message to indicate that Python dependencies were installed successfully
echo "Dependencies installed successfully..."

# Echo a message to indicate that the Python code is being run
echo "Running Python code..."

# Run the Python code
python ./compare_files.py

# Check if the Python script exited successfully
if [[ $? -ne 0 ]]; then
    # If the Python script failed, echo an error message and exit the script
    echo "Python script execution failed."
    exit $?
else
    # If the Python script was successful, echo a message and exit the script
    echo "Execution complete..."
    exit 0
fi
