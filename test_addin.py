#!/usr/bin/env python3
"""
Simple test script to verify the Excel add-in installation.
"""

import os

def main():
    """Main test function."""
    print("🔍 Testing Excel Add-in Installation...\n")
    
    # Check if manifest is in WEF folder
    wef_path = os.path.expanduser("~/Library/Containers/com.microsoft.Excel/Data/Documents/wef")
    manifest_path = os.path.join(wef_path, "manifest.xml")
    
    if os.path.exists(manifest_path):
        print(f"✅ Manifest found in WEF folder: {manifest_path}")
        
        # Read and check manifest content
        with open(manifest_path, 'r') as f:
            content = f.read()
            if "file://" in content:
                print("✅ Manifest uses file:// URLs (good for WEF installation)")
            elif "localhost:8000" in content:
                print("⚠️  Manifest uses localhost URLs (requires server)")
            else:
                print("ℹ️  Manifest uses other URL format")
    else:
        print(f"❌ Manifest not found in WEF folder: {manifest_path}")
    
    # Check if project files exist
    project_files = [
        "taskpane.html",
        "taskpane.js", 
        "functions.js",
        "functions.json",
        "functions.html",
        "assets/icon-32.png"
    ]
    
    print("\n📁 Checking project files:")
    for file in project_files:
        if os.path.exists(file):
            print(f"✅ {file}")
        else:
            print(f"❌ {file}")
    
    print("\n📋 Next steps:")
    print("1. Open Excel")
    print("2. Go to Insert > My Add-ins")
    print("3. Look for 'KMAPI Excel Add-in' in the list")
    print("4. If not visible, try restarting Excel")
    print("5. Once loaded, configure your API settings")
    
    print("\n🔧 Troubleshooting:")
    print("- If add-in doesn't appear: restart Excel")
    print("- If functions don't work: check browser console")
    print("- Test with: =DEBUGLOG('Hello')")
    print("- Test API with: =TESTAPI('your_model_id', 'your_api_key')")

if __name__ == "__main__":
    main() 