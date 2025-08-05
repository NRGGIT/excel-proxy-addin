#!/usr/bin/env python3
"""
Setup script for WEF folder installation of Excel add-in.
This script helps you install the add-in using the WEF folder method.
"""

import os
import shutil
import subprocess
import sys

def get_wef_folder():
    """Get the WEF folder path for the current user."""
    home = os.path.expanduser("~")
    wef_path = os.path.join(home, "Library/Containers/com.microsoft.Excel/Data/Documents/wef")
    return wef_path

def copy_manifest():
    """Copy the manifest.xml to the WEF folder."""
    wef_path = get_wef_folder()
    
    if not os.path.exists(wef_path):
        print(f"❌ WEF folder not found at: {wef_path}")
        print("Please make sure Excel is installed and has been run at least once.")
        return False
    
    manifest_source = "manifest.xml"
    manifest_dest = os.path.join(wef_path, "manifest.xml")
    
    try:
        shutil.copy2(manifest_source, manifest_dest)
        print(f"✅ Manifest copied to: {manifest_dest}")
        return True
    except Exception as e:
        print(f"❌ Error copying manifest: {e}")
        return False

def start_server():
    """Start the local server."""
    print("🚀 Starting local server...")
    try:
        subprocess.Popen([sys.executable, "server.py"], 
                        stdout=subprocess.PIPE, 
                        stderr=subprocess.PIPE)
        print("✅ Server started at http://localhost:8000")
        return True
    except Exception as e:
        print(f"❌ Error starting server: {e}")
        return False

def main():
    """Main setup function."""
    print("🔧 Setting up Excel Add-in for WEF installation...\n")
    
    # Step 1: Copy manifest
    print("1. Copying manifest to WEF folder...")
    if not copy_manifest():
        return False
    
    # Step 2: Start server
    print("\n2. Starting local server...")
    if not start_server():
        return False
    
    print("\n🎉 Setup complete!")
    print("\n📋 Next steps:")
    print("1. Open Excel")
    print("2. Go to Insert > My Add-ins")
    print("3. Look for 'KMAPI Excel Add-in' in the list")
    print("4. If not visible, try restarting Excel")
    print("5. Configure your API settings using the Settings button")
    
    print("\n🔧 Troubleshooting:")
    print("- If the add-in doesn't appear, restart Excel")
    print("- Make sure the server is running: http://localhost:8000")
    print("- Check that manifest.xml is in the WEF folder")
    
    return True

if __name__ == "__main__":
    main() 