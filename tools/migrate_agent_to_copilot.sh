#!/bin/bash

# migrate_agent_to_copilot.sh
# Copies Antigravity .agent/ configurations into GitHub Copilot .github/ configurations.

set -e

AGENT_DIR=".agent"
GITHUB_DIR=".github"
COPILOT_FILE="$GITHUB_DIR/copilot-instructions.md"

echo "Starting migration from $AGENT_DIR to $GITHUB_DIR..."

# Ensure source directory exists
if [ ! -d "$AGENT_DIR" ]; then
    echo "Error: Source directory $AGENT_DIR does not exist."
    exit 1
fi

# 1. Create .github structure
echo "Preparing $GITHUB_DIR directory..."
mkdir -p "$GITHUB_DIR/skills"
mkdir -p "$GITHUB_DIR/prompts"

# 2. Concatenate rules into copilot-instructions.md
echo "Compiling rules into $COPILOT_FILE..."
echo "# GitHub Copilot Instructions" > "$COPILOT_FILE"
echo "" >> "$COPILOT_FILE"
echo "The following instructions dictate the coding style, design principles, and general communication guidelines for this repository." >> "$COPILOT_FILE"
echo "" >> "$COPILOT_FILE"

if [ -d "$AGENT_DIR/rules" ]; then
    for rule_file in "$AGENT_DIR"/rules/*.md; do
        if [ -f "$rule_file" ]; then
            echo "-> Adding $(basename "$rule_file")"
            echo "## Extract from $(basename "$rule_file")" >> "$COPILOT_FILE"
            echo "" >> "$COPILOT_FILE"
            cat "$rule_file" >> "$COPILOT_FILE"
            echo "" >> "$COPILOT_FILE"
            echo "---" >> "$COPILOT_FILE"
            echo "" >> "$COPILOT_FILE"
        fi
    done
else
    echo "-> No rules directory found. Skipping."
fi

# Helper function to conditionally copy files
copy_if_newer() {
    local src_dir="$1"
    local dest_dir="$2"
    
    # Find all files recursively in the source directory
    find "$src_dir" -type f | while read -r src_file; do
        # Calculate the destination path
        local rel_path="${src_file#$src_dir/}"
        local dest_file="$dest_dir/$rel_path"
        
        # Ensure destination subdirectory exists
        mkdir -p "$(dirname "$dest_file")"
        
        # Check if destination exists and if source is newer
        if [ ! -f "$dest_file" ] || [ "$src_file" -nt "$dest_file" ]; then
            cp "$src_file" "$dest_file"
            echo "-> Updated: $dest_file"
        else
            echo "-> Skipped (destination is newer): $dest_file"
        fi
    done
}

# 3. Copy Skills
echo "Copying skills..."
if [ -d "$AGENT_DIR/skills" ]; then
    copy_if_newer "$AGENT_DIR/skills" "$GITHUB_DIR/skills"
else
    echo "-> No skills directory found. Skipping."
fi

# 4. Copy Workflows
echo "Copying workflows to prompts..."
if [ -d "$AGENT_DIR/workflows" ]; then
    copy_if_newer "$AGENT_DIR/workflows" "$GITHUB_DIR/prompts"
else
    echo "-> No workflows directory found. Skipping."
fi

echo "=========================================="
echo "Migration Complete!"
echo "Check the $GITHUB_DIR directory for your Copilot configuration."
echo "=========================================="
