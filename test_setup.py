#!/usr/bin/env python3
"""
Test script to verify the Excel add-in setup.
"""

import os
import json
import xml.etree.ElementTree as ET

def test_files_exist():
    """Check if all required files exist."""
    required_files = [
        'manifest.xml',
        'taskpane.html',
        'taskpane.js',
        'functions.js',
        'functions.json',
        'functions.html',
        'server.py'
    ]
    
    missing_files = []
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print(f"❌ Missing files: {missing_files}")
        return False
    else:
        print("✅ All required files exist")
        return True

def test_manifest_xml():
    """Test if manifest.xml is valid."""
    try:
        tree = ET.parse('manifest.xml')
        root = tree.getroot()
        
        # Check for required elements
        id_elem = root.find('.//{http://schemas.microsoft.com/office/appforoffice/1.1}Id')
        if id_elem is None:
            print("❌ Manifest missing Id element")
            return False
            
        print("✅ Manifest XML is valid")
        return True
    except Exception as e:
        print(f"❌ Manifest XML error: {e}")
        return False

def test_functions_json():
    """Test if functions.json is valid."""
    try:
        with open('functions.json', 'r') as f:
            data = json.load(f)
        
        if 'functions' not in data:
            print("❌ functions.json missing 'functions' key")
            return False
            
        print("✅ functions.json is valid")
        return True
    except Exception as e:
        print(f"❌ functions.json error: {e}")
        return False

def test_server_import():
    """Test if server.py can be imported."""
    try:
        import server
        print("✅ server.py can be imported")
        return True
    except Exception as e:
        print(f"❌ server.py import error: {e}")
        return False

def main():
    """Run all tests."""
    print("🔍 Testing Excel Add-in Setup...\n")
    
    tests = [
        test_files_exist,
        test_manifest_xml,
        test_functions_json,
        test_server_import
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        if test():
            passed += 1
        print()
    
    print(f"📊 Test Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 All tests passed! Your add-in is ready to use.")
        print("\n📋 Next steps:")
        print("1. Run: python server.py")
        print("2. Open Excel and load the add-in")
        print("3. Configure your API settings")
    else:
        print("⚠️  Some tests failed. Please fix the issues above.")

if __name__ == "__main__":
    main() 