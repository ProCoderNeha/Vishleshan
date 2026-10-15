#!/usr/bin/env python
"""Test form submission and duplicate detection."""

import requests
import json
import time

BASE_URL = "http://127.0.0.1:5000"

def test_form_page():
    """Test if form page loads."""
    print("\n📝 Test 1: Form page load...")
    try:
        response = requests.get(f"{BASE_URL}/form", timeout=5)
        if response.status_code == 200:
            print("✅ Form page loads successfully (HTTP 200)")
            return True
        else:
            print(f"❌ Form page failed with status {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Error: {e}")
        return False


def test_valid_submission():
    """Test valid form submission."""
    print("\n📝 Test 2: Valid form submission...")
    payload = {
        "Year": "2025",
        "Name of Student": "Test Student One",
        "Gender": "Male",
        "Affiliation": "IIT Delhi",
        "State": "Delhi",
        "Branch": "B.Tech",
        "Branch Category": "Mechanical",
        "Field / Branch of Engg.": "Mechanical Engineering",
        "Division (DRDO)": "MEMS",
        "Scientist Name": "Dr. Sharma",
        "Designation": "Scientist-D",
        "Period of Attachment": "01 Jan 2025 to 28 Feb 2025",
        "Time in Weeks": "8",
        "Mode of Training": "Offline",
        "Training Period": "Winter 2025",
        "Topic Name": "Advanced Materials Research",
        "Certificate Issued": "Yes",
        "Certificate Issue Date": "2025-02-28",
    }
    
    try:
        response = requests.post(
            f"{BASE_URL}/submit-form",
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=5
        )
        
        if response.status_code == 201:
            data = response.json()
            print(f"✅ Entry submitted successfully (HTTP 201)")
            print(f"   Message: {data.get('message')}")
            print(f"   Assigned S.No: {data.get('sno')}")
            return data.get('sno')
        else:
            print(f"❌ Submission failed with status {response.status_code}")
            print(f"   Response: {response.json()}")
            return None
    except Exception as e:
        print(f"❌ Error: {e}")
        return False


def test_duplicate_detection():
    """Test duplicate detection with normalized values."""
    print("\n📝 Test 3: Duplicate detection with normalization...")
    
    # Submit first entry (exact same as test 2)
    payload1 = {
        "Year": "2025",
        "Name of Student": "Test Student One",
        "Gender": "Male",
        "Affiliation": "IIT Delhi",
        "State": "Delhi",
        "Branch": "B.Tech",
        "Branch Category": "Mechanical",
        "Field / Branch of Engg.": "Mechanical Engineering",
        "Division (DRDO)": "MEMS",
        "Scientist Name": "Dr. Sharma",
        "Designation": "Scientist-D",
        "Period of Attachment": "01 Jan 2025 to 28 Feb 2025",
        "Time in Weeks": "8",
        "Mode of Training": "Offline",
        "Training Period": "Winter 2025",
        "Topic Name": "Advanced Materials Research",
        "Certificate Issued": "Yes",
        "Certificate Issue Date": "2025-02-28",
    }
    
    # Submit second entry with normalized variations (should be detected as duplicate)
    payload2 = {
        "Year": "2025",
        "Name of Student": " test student one ",  # Extra spaces
        "Gender": "Male",
        "Affiliation": "iit delhi",  # Lowercase
        "State": "Delhi",
        "Branch": "B.Tech.",  # With period
        "Branch Category": "Mechanical",
        "Field / Branch of Engg.": "Mechanical Engineering",
        "Division (DRDO)": "mems",  # Lowercase
        "Scientist Name": "Dr. Sharma",
        "Designation": "Scientist-D",
        "Period of Attachment": "01 - Jan - 2025 to 28 - Feb - 2025",  # Different formatting
        "Time in Weeks": "8",
        "Mode of Training": "Offline",
        "Training Period": "Winter 2025",
        "Topic Name": "Different Topic",  # Different topic (should still be detected as dup)
        "Certificate Issued": "Yes",
        "Certificate Issue Date": "2025-02-28",
    }
    
    try:
        # First submission
        response1 = requests.post(
            f"{BASE_URL}/submit-form",
            json=payload1,
            headers={"Content-Type": "application/json"},
            timeout=5
        )
        
        if response1.status_code != 201:
            print(f"⚠️  First submission failed: {response1.json()}")
            return False
        
        print(f"✅ First entry submitted (S.No: {response1.json().get('sno')})")
        
        # Second submission (should be detected as duplicate)
        time.sleep(0.5)
        response2 = requests.post(
            f"{BASE_URL}/submit-form",
            json=payload2,
            headers={"Content-Type": "application/json"},
            timeout=5
        )
        
        if response2.status_code == 409:
            print(f"✅ Duplicate correctly detected (HTTP 409)")
            print(f"   Error: {response2.json().get('error')}")
            return True
        else:
            print(f"❌ Duplicate NOT detected (status: {response2.status_code})")
            print(f"   Response: {response2.json()}")
            return False
    except Exception as e:
        print(f"❌ Error: {e}")
        return False


def test_missing_fields():
    """Test missing required fields validation."""
    print("\n📝 Test 4: Missing fields validation...")
    
    # Minimal payload missing required fields
    payload = {
        "Year": "2025",
        "Name of Student": "Test",
    }
    
    try:
        response = requests.post(
            f"{BASE_URL}/submit-form",
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=5
        )
        
        if response.status_code == 400:
            print(f"✅ Missing fields rejected (HTTP 400)")
            data = response.json()
            print(f"   Error: {data.get('error')}")
            return True
        else:
            print(f"❌ Should have rejected missing fields (status: {response.status_code})")
            return False
    except Exception as e:
        print(f"❌ Error: {e}")
        return False


def main():
    """Run all tests."""
    print("=" * 60)
    print("  DRDO Dashboard - Form Testing")
    print("=" * 60)
    
    # Allow time for server startup
    time.sleep(1)
    
    results = {
        "Form page load": test_form_page(),
        "Valid submission": test_valid_submission() is not None,
        "Duplicate detection": test_duplicate_detection(),
        "Missing fields validation": test_missing_fields(),
    }
    
    print("\n" + "=" * 60)
    print("  TEST SUMMARY")
    print("=" * 60)
    for test_name, result in results.items():
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"{status}: {test_name}")
    
    passed = sum(1 for r in results.values() if r)
    total = len(results)
    print(f"\nTotal: {passed}/{total} tests passed")
    print("=" * 60)


if __name__ == "__main__":
    main()
