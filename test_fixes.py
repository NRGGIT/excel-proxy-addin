#!/usr/bin/env python3
"""
Test script to verify the fixes for #BUSY and Trust Center issues.
"""

import os
import subprocess
import time

def check_server():
    """Check if server is running."""
    try:
        import requests
        response = requests.get("http://localhost:8000/taskpane.html", timeout=5)
        if response.status_code == 200:
            print("✅ Server is running at http://localhost:8000")
            return True
        else:
            print(f"❌ Server responded with status: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Server not accessible: {e}")
        return False

def start_server():
    """Start the server if not running."""
    if not check_server():
        print("🚀 Starting server...")
        try:
            subprocess.Popen(["python", "server.py"], 
                           stdout=subprocess.PIPE, 
                           stderr=subprocess.PIPE)
            time.sleep(2)  # Wait for server to start
            if check_server():
                print("✅ Server started successfully")
                return True
            else:
                print("❌ Failed to start server")
                return False
        except Exception as e:
            print(f"❌ Error starting server: {e}")
            return False
    return True

def main():
    """Main test function."""
    print("🔧 Testing Fixes for Excel Add-in Issues...\n")
    
    # Check and start server
    if not start_server():
        print("❌ Cannot start server. Please check the error above.")
        return
    
    # Check manifest
    wef_path = os.path.expanduser("~/Library/Containers/com.microsoft.Excel/Data/Documents/wef")
    manifest_path = os.path.join(wef_path, "manifest.xml")
    
    if os.path.exists(manifest_path):
        print("✅ Manifest is in WEF folder")
        
        # Check if manifest uses localhost URLs
        with open(manifest_path, 'r') as f:
            content = f.read()
            if "localhost:8000" in content:
                print("✅ Manifest uses localhost URLs (should fix Trust Center issue)")
            else:
                print("⚠️  Manifest doesn't use localhost URLs")
    else:
        print("❌ Manifest not found in WEF folder")
    
    print("\n📋 Testing Instructions:")
    print("1. Restart Excel completely")
    print("2. Go to Insert > My Add-ins")
    print("3. Look for 'KMAPI Excel Add-in'")
    print("4. Test these functions:")
    print("   =SIMPLETEST('Hello')     // Should work without #BUSY")
    print("   =DEBUGLOG('Hello')       // May still show #BUSY (Excel context issue)")
    print("   =TESTAPI('id', 'key')    // Test API connection")
    
    print("\n🔧 If Trust Center error persists:")
    print("1. Go to Excel > Preferences > Security & Privacy")
    print("2. Click 'Trust Center' > 'Trust Center Settings'")
    print("3. Go to 'Trusted Add-in Catalogs'")
    print("4. Add: http://localhost:8000")
    print("5. Check 'Show in Menu' and click 'OK'")
    print("6. Restart Excel")

if __name__ == "__main__":
    main() 