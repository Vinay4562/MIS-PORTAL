import requests
import sys
import json
from datetime import datetime, timedelta

class LineLosMISAPITester:
    def __init__(self, base_url="http://localhost:8000/api"):
        self.base_url = base_url
        self.token = None
        self.tests_run = 0
        self.tests_passed = 0
        self.user_id = None
        self.feeder_id = None
        self.entry_id = None

    def run_test(self, name, method, endpoint, expected_status, data=None, headers=None):
        """Run a single API test"""
        url = f"{self.base_url}/{endpoint}"
        test_headers = {'Content-Type': 'application/json'}
        
        if self.token:
            test_headers['Authorization'] = f'Bearer {self.token}'
        
        if headers:
            test_headers.update(headers)

        self.tests_run += 1
        print(f"\n🔍 Testing {name}...")
        print(f"   URL: {url}")
        
        try:
            if method == 'GET':
                response = requests.get(url, headers=test_headers, params=data)
            elif method == 'POST':
                response = requests.post(url, json=data, headers=test_headers)
            elif method == 'PUT':
                response = requests.put(url, json=data, headers=test_headers)
            elif method == 'DELETE':
                response = requests.delete(url, headers=test_headers)

            success = response.status_code == expected_status
            if success:
                self.tests_passed += 1
                print(f"✅ Passed - Status: {response.status_code}")
                try:
                    return success, response.json() if response.content else {}
                except:
                    return success, {}
            else:
                print(f"❌ Failed - Expected {expected_status}, got {response.status_code}")
                try:
                    print(f"   Response: {response.text}")
                except:
                    pass
                return False, {}

        except Exception as e:
            print(f"❌ Failed - Error: {str(e)}")
            return False, {}

    def test_user_registration(self):
        """Test user registration"""
        test_user_data = {
            "email": f"test_user_{datetime.now().strftime('%H%M%S')}@example.com",
            "password": "TestPass123!",
            "full_name": "Test User"
        }
        
        success, response = self.run_test(
            "User Registration",
            "POST",
            "auth/register",
            200,
            data=test_user_data
        )
        
        if success and 'access_token' in response:
            self.token = response['access_token']
            self.user_id = response['user']['id']
            print(f"   Token obtained: {self.token[:20]}...")
            return True, test_user_data
        return False, test_user_data

    def test_user_login(self, user_data):
        """Test user login"""
        login_data = {
            "email": user_data["email"],
            "password": user_data["password"]
        }
        
        success, response = self.run_test(
            "User Login",
            "POST",
            "auth/login",
            200,
            data=login_data
        )
        
        if success and 'access_token' in response:
            self.token = response['access_token']
            print(f"   Login token: {self.token[:20]}...")
            return True
        return False

    def test_get_current_user(self):
        """Test get current user endpoint"""
        success, response = self.run_test(
            "Get Current User",
            "GET",
            "auth/me",
            200
        )
        return success

    def test_initialize_feeders(self):
        """Test feeder initialization"""
        success, response = self.run_test(
            "Initialize Feeders",
            "POST",
            "init-feeders",
            200
        )
        
        if success:
            print(f"   Feeders initialized: {response.get('count', 0)}")
        return success

    def test_get_feeders(self):
        """Test get all feeders"""
        success, response = self.run_test(
            "Get All Feeders",
            "GET",
            "feeders",
            200
        )
        
        if success and response:
            self.feeder_id = response[0]['id']
            print(f"   Found {len(response)} feeders")
            print(f"   First feeder ID: {self.feeder_id}")
            return True
        return False

    def test_get_single_feeder(self):
        """Test get single feeder"""
        if not self.feeder_id:
            print("❌ No feeder ID available for testing")
            return False
            
        success, response = self.run_test(
            "Get Single Feeder",
            "GET",
            f"feeders/{self.feeder_id}",
            200
        )
        
        if success:
            print(f"   Feeder name: {response.get('name', 'Unknown')}")
        return success

    def test_create_entry(self):
        """Test creating a daily entry"""
        if not self.feeder_id:
            print("❌ No feeder ID available for testing")
            return False
            
        today = datetime.now().strftime("%Y-%m-%d")
        entry_data = {
            "feeder_id": self.feeder_id,
            "date": today,
            "end1_import_final": 100.50,
            "end1_export_final": 95.25,
            "end2_import_final": 200.75,
            "end2_export_final": 190.30
        }
        
        success, response = self.run_test(
            "Create Daily Entry",
            "POST",
            "entries",
            200,
            data=entry_data
        )
        
        if success and 'id' in response:
            self.entry_id = response['id']
            print(f"   Entry created with ID: {self.entry_id}")
            print(f"   Loss percentage: {response.get('loss_percent', 0):.2f}%")
            return True
        return False

    def test_get_entries(self):
        """Test getting entries with filters"""
        current_year = datetime.now().year
        current_month = datetime.now().month
        
        success, response = self.run_test(
            "Get Entries (Filtered)",
            "GET",
            "entries",
            200,
            data={
                "feeder_id": self.feeder_id,
                "year": current_year,
                "month": current_month
            }
        )
        
        if success:
            print(f"   Found {len(response)} entries")
        return success

    def test_update_entry(self):
        """Test updating an entry"""
        if not self.entry_id:
            print("❌ No entry ID available for testing")
            return False
            
        update_data = {
            "end1_import_final": 105.75,
            "end1_export_final": 98.50
        }
        
        success, response = self.run_test(
            "Update Entry",
            "PUT",
            f"entries/{self.entry_id}",
            200,
            data=update_data
        )
        
        if success:
            print(f"   Updated loss percentage: {response.get('loss_percent', 0):.2f}%")
        return success

    def test_export_data(self):
        """Test Excel export functionality"""
        if not self.feeder_id:
            print("❌ No feeder ID available for testing")
            return False
            
        current_year = datetime.now().year
        current_month = datetime.now().month
        
        success, response = self.run_test(
            "Export Excel Data",
            "GET",
            f"export/{self.feeder_id}/{current_year}/{current_month}",
            200
        )
        return success

    def test_delete_entry(self):
        """Test deleting an entry"""
        if not self.entry_id:
            print("❌ No entry ID available for testing")
            return False
            
        success, response = self.run_test(
            "Delete Entry",
            "DELETE",
            f"entries/{self.entry_id}",
            200
        )
        
        if success:
            print(f"   Entry deleted successfully")
        return success

def main():
    print("🚀 Starting Line Losses MIS Portal API Testing...")
    print("=" * 60)
    
    tester = LineLosMISAPITester()
    
    # Test user registration and authentication
    reg_success, user_data = tester.test_user_registration()
    if not reg_success:
        print("❌ Registration failed, stopping tests")
        return 1

    # Test login with the same credentials
    if not tester.test_user_login(user_data):
        print("❌ Login failed, stopping tests")
        return 1

    # Test get current user
    if not tester.test_get_current_user():
        print("❌ Get current user failed")

    # Test feeder operations
    if not tester.test_initialize_feeders():
        print("❌ Feeder initialization failed")
        return 1

    if not tester.test_get_feeders():
        print("❌ Get feeders failed, stopping tests")
        return 1

    if not tester.test_get_single_feeder():
        print("❌ Get single feeder failed")

    # Test entry operations
    if not tester.test_create_entry():
        print("❌ Create entry failed, stopping entry tests")
    else:
        # Only proceed with entry tests if creation succeeded
        if not tester.test_get_entries():
            print("❌ Get entries failed")

        if not tester.test_update_entry():
            print("❌ Update entry failed")

        if not tester.test_export_data():
            print("❌ Export data failed")

        if not tester.test_delete_entry():
            print("❌ Delete entry failed")

    # Print final results
    print("\n" + "=" * 60)
    print(f"📊 FINAL RESULTS:")
    print(f"   Tests Run: {tester.tests_run}")
    print(f"   Tests Passed: {tester.tests_passed}")
    print(f"   Success Rate: {(tester.tests_passed/tester.tests_run)*100:.1f}%")
    
    if tester.tests_passed == tester.tests_run:
        print("🎉 All tests passed!")
        return 0
    else:
        print(f"⚠️  {tester.tests_run - tester.tests_passed} tests failed")
        return 1

if __name__ == "__main__":
    sys.exit(main())