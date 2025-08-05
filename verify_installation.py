#!/usr/bin/env python3
"""
Verification script to check if the Excel add-in is properly installed.
"""

import os
import requests
import time

def check_wef_folder():
    """Check if manifest is in WEF folder."""
    wef_path = os.path.expanduser("~/Library/Containers/com.microsoft.Excel/Data/Documents/wef")
    manifest_path = os.path.join(wef_path, "manifest.xml")
    
    if os.path.exists(manifest_path):
        print(f"✅ Manifest found in WEF folder: {manifest_path}")
        return True
    else:
        print(f"❌ Manifest not found in WEF folder: {manifest_path}")
        return False

def check_server():
    """Check if local server is running."""
    try:
        response = requests.get("http://localhost:8000/manifest.xml", timeout=5)
        if response.status_code == 200:
            print("✅ Local server is running and serving files")
            return True
        else:
            print(f"❌ Server responded with status: {response.status_code}")
            return False
    except requests.exceptions.RequestException as e:
        print(f"❌ Server not accessible: {e}")
        return False

def check_files():
    """Check if all required files are accessible via server."""
    files_to_check = [
        "taskpane.html",
        "taskpane.js", 
        "functions.js",
        "functions.json",
        "functions.html",
        "assets/icon-32.png"
    ]
    
    all_good = True
    for file in files_to_check:
        try:
            response = requests.get(f"http://localhost:8000/{file}", timeout=5)
            if response.status_code == 200:
                print(f"✅ {file} accessible")
            else:
                print(f"❌ {file} not accessible (status: {response.status_code})")
                all_good = False
        except requests.exceptions.RequestException as e:
            print(f"❌ {file} not accessible: {e}")
            all_good = False
    
    return all_good

def main():
    """Run all verification checks."""
    print("🔍 Verifying Excel Add-in Installation...\n")
    
    checks = [
        ("WEF Folder", check_wef_folder),
        ("Local Server", check_server),
        ("File Accessibility", check_files)
    ]
    
    passed = 0
    total = len(checks)
    
    for name, check_func in checks:
        print(f"Checking {name}...")
        if check_func():
            passed += 1
        print()
    
    print(f"📊 Verification Results: {passed}/{total} checks passed")
    
    if passed == total:
        print("🎉 All checks passed! Your add-in should be working.")
        print("\n📋 To use the add-in:")
        print("1. Open Excel")
        print("2. Go to Insert > My Add-ins")
        print("3. Look for 'KMAPI Excel Add-in'")
        print("4. Click 'Add' to install it")
        print("5. Configure your API settings")
    else:
        print("⚠️  Some checks failed. Please fix the issues above.")
        print("\n🔧 To fix:")
        print("1. Run: python setup_wef.py")
        print("2. Make sure Excel is closed")
        print("3. Restart Excel")
        print("4. Try again")

if __name__ == "__main__":
    main() 