import requests
import pandas as pd
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

URL = "https://netenrich.opsramp.com/"
key = "cHrCgP3TWVtv3EwMzah3hfjH34eXUHM8"
secret = "c55PPRzMPg3BWp5tXZMwwT8Gzpq6GmUbBwWnAfdhJZjmVHfXB59ZMM5rZY3kA5wf"

def token_generation():
    try:
        token_url = URL + "auth/oauth/token"
        auth_data = {
            'client_secret': secret,
            'grant_type': 'client_credentials',
            'client_id': key
        }
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        
        response = requests.post(token_url, headers=headers, data=auth_data, verify=False)
        
        if response.status_code == 200:
            response = response.json()
            token = response.get("access_token")
            if token:
                return token
            else:
                print("Error: No access token returned in the response.")
        else:
            print(f"Failed to obtain access token: {response.status_code} - {response.text}")
    except Exception as e:
        print("Error during token generation:", str(e))
    return None

def process_correlation(access_token, tenantId, policyId, tenantname, policyname):
    #mode = "disable"
    try:
        correlation_url = f"{URL}api/v2/tenants/{tenantId}/policies/alertCorrelation/{policyId}"
        #print (f"Hitting this URL {correlation_url}")
        headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}
        results = requests.delete(correlation_url, headers=headers, verify=False)
        
        if results.status_code == 200:
           print(f"\033[92m[SUCCESS]\033[0m Disabled correlation policy for Tenant: {tenantname}, Policy: {policyname}")
        elif results.status_code in [407, 401] or "invalid_token" in results.text.lower():   
            print("Access token expired, regenerating the token...")
            access_token = token_generation()
            if access_token:
                print("Got the new token, retrying the request...")
                process_correlation(access_token, tenantId, policyId)
            else:
                print("Issue in regenerating the token")
        else:
            print(f"\033[91m[FAILED]\033[0m Could not disable policy for Tenant: {tenantname}, Policy: {policyname}, {results.status_code} - {results.text}")
    except Exception as e:
        print(f"Error disabling correlation policy for Tenant: {tenantname}, Policy: {policyname}. Error: {str(e)}")
        return None

def main():
    access_token = token_generation()
    if access_token:
        try:
            df = pd.read_excel("C:\\Users\\hari.boddu\\Downloads\\Correlation_Policies_OFF.xlsx")
            for index, row in df.iterrows():
                tenantname = row.get("Client Name")
                tenantId = row.get("Tenant ID")
                policyId = row.get("Policy ID")
                policyname = row.get("Policy Name")
                if tenantId and policyId:
                    process_correlation(access_token, tenantId, policyId, tenantname, policyname)
                else:
                    print(f"Skipping row {index + 1}: Missing Tenant ID or Policy ID.")
        except Exception as e:
            print("Error reading Excel file:", str(e))

if __name__ == "__main__":
    main()
