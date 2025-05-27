import json
import requests
import jwt
from exceptions import HermesMSGraphError

class HttpClient:
    def __init__(self, client_id, client_secret, tenant_id):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id

        self.access_token = self.__get_access_token()

    def __get_access_token(self):
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        payload = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        try:
            response = requests.post(url, data=payload)
        except Exception as e:
            raise RuntimeError("Unable to make post request to get access token") from e

        response_data = response.json()
        access_token = response_data["access_token"]
        return access_token
    

    def __headers(self):
        access_token = self.access_token
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        return headers

    def __verify_response_status_code(self, response):
        if response.status_code == 200:
            return 200
        elif response.status_code == 401:
            self.access_token = self.__get_access_token()
            return 401
        elif response.status_code == 403:
            return 403
        elif response.status_code == 404:
            return 404


    def __get_http(self, url, headers=None):
        if headers is None:
            headers = self.__headers()
        else:
            headers_raw = self.__headers()
            headers_raw.update(headers)
            headers = headers_raw
            print(headers_raw)
        response = requests.get(url, headers=headers)
        
        response_status_code = self.__verify_response_status_code(response)
        if response_status_code == 401:
            response = self.__get_http(url)

        return response

    def __post_http(self, url, payload):
        headers = self.__headers()
        data = json.dumps(payload)

        response = requests.post(url, headers=headers, data=data)

        response_status_code = self.__verify_response_status_code(response)
        
        if response_status_code == 401:
            response = self.__get_http(url)

        return response
    
    def post(self, url, payload):
        return self.__post_http(url, payload)
    
    def get(self, url, headers=None):
        return self.__get_http(url, headers)
    
    def list_msgraph_permisions(self):
        decoded_token = jwt.decode(self.access_token, options={"verify_signature": False})
        if "scp" in decoded_token:
            print("Delegated Permissions:", decoded_token["scp"])
        if "roles" in decoded_token:
            print("Application Permissions:", decoded_token["roles"])
    
    def get_json_response_by_url(self, url, get_value=True):
        response = self.get(url)

        if response.status_code == 200:
            try:
                json_response = response.json()
                return json_response.get("value", []) if get_value else json_response
            except ValueError:
                print(f"Invalid JSON response from {url}")
                return None
        else:
            raise HermesMSGraphError(f"Error fetching data from {url}: {response.status_code} - {response.text}")
