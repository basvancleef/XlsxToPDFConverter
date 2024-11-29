#!/bin/bash

# Define paths relative to the current script location
base_folder="$(cd "$(dirname "$0")" && pwd)"
input_folder="$base_folder/input"
output_folder="$base_folder/output"

mkdir -p "$output_folder"

# Find the most recent .xlsx file in the input folder
latest_file=$(find "$input_folder" -type f -name "*.xlsx" -exec stat -f "%m %N" {} + 2>/dev/null | sort -n | tail -1 | cut -d' ' -f2-)

# Check if a file was found
if [[ -z "$latest_file" ]]; then
  echo "No .xlsx files found in $input_folder."
  exit 1
fi

filename=$(basename "$latest_file" .xlsx)

output_pdf="$output_folder/$filename.pdf"

# Use AppleScript to open Excel and export the file as a PDF
osascript <<EOF
tell application "Microsoft Excel"
  try
    open POSIX file "$latest_file"
    set activeSheet to active sheet of active workbook
    set outputFilePath to POSIX file "$output_pdf"
    save activeSheet in outputFilePath as PDF file format
    close active workbook without saving
  on error errMsg
    display dialog "Error: " & errMsg
    error errMsg
  end try
end tell
EOF

if [[ -f "$output_pdf" ]]; then
  echo "Successfully saved $latest_file as PDF in $output_folder."
else
  echo "Failed to save $latest_file as PDF."
fi
