from exceptions import HermesMSGraphError
from tqdm import tqdm

class UsersService:
    def __init__(self, http):
        self.http = http

    def get_user_id_by_email(self, email_address: str) -> str:
        """
        Retrieve the user ID for a given email address.
        :param email_address: The email address of the user.
        :return: The user ID if found, None otherwise.
        :raises HermesMSGraphError: If the request fails or the user is not found.
        """
        url = f"https://graph.microsoft.com/v1.0/users/{email_address}"
        
        response = self.http.get(url)

        if response.status_code == 200:
            user_data = response.json()
            return user_data.get("id")
        else:
            error_message = f"Error fetching user ID: {response.status_code} - {response.text}"
            raise HermesMSGraphError(error_message)
    
    
    def __filter_data(self, users: list) -> list:
        users_filtered = []
        for user in users:
            users_filtered.append({
                "displayName": user.get("displayName"),
                "jobTitle": user.get("jobTitle"),
                "mail": user.get("mail"),
                "officeLocation": user.get("officeLocation"),
            })
        return users_filtered
    
    def __add_friendly_license_names(self, sku_list):
        friendly_names = {
            "POWER_BI_PRO": "Power BI Pro",
            "WINDOWS_STORE": "Windows Store",
            "FLOW_FREE": "Power Automate Free",
            "MICROSOFT_BUSINESS_CENTER": "Microsoft Business Center",
            "CCIBOTS_PRIVPREV_VIRAL": "Copilot (Preview)",
            "SPB": "Microsoft 365 Business Premium",
            "POWERAPPS_VIRAL": "Power Apps (Trial)",
            "EXCHANGESTANDARD": "Exchange Online (Plan 1)",
            "Microsoft_Teams_Exploratory_Dept": "Teams Exploratory",
            "O365_BUSINESS_PREMIUM": "Microsoft 365 Business Standard",
            "POWER_BI_STANDARD": "Power BI (Free)",
            "PBI_PREMIUM_PER_USER": "Power BI Premium (Per User)",
            "Power_Pages_vTrial_for_Makers": "Power Pages (Trial)",
            "RMSBASIC": "Rights Management (Basic)",
            "Teams_Premium_(for_Departments)": "Teams Premium (Departments)",
            "POWERAPPS_DEV": "Power Apps Developer",
            "PROJECT_PLAN3_DEPT": "Project Plan 3 (Departments)"
        }

        skuid_list = []

        for sku in sku_list:
            sku_id = sku.get("skuId")
            sku_name = sku.get("skuPartNumber")
            if sku_name in friendly_names:
                sku_friendly_name = friendly_names[sku_name]
            else:
                sku["friendlyName"] = friendly_names.get(sku_name, "Desconhecido")
            skuid_list.append({"skuId": sku_id, "skuName": sku_name, "friendlyName": sku_friendly_name})

        return skuid_list

    
    def get_tenant_licenses(self):
        response = self.http.get('https://graph.microsoft.com/v1.0/subscribedSkus')
        if response.status_code == 200:
            skus = response.json().get("value", [])
            skuid_list = self.__add_friendly_license_names(skus)
            return skuid_list
        else:
            error_message = f"Error fetching SKUs: {response.status_code} - {response.text}"
            raise HermesMSGraphError(error_message)


    def __get_license_details(self, users: list) -> list:
        users_all_info = []
        for user in tqdm(users, desc="Fetching user details", unit="user"):
            user_id = user['id']
            print(user_id)
            response = self.http.get(f"https://graph.microsoft.com/v1.0/users/{user_id}?$select=displayName,userPrincipalName,accountEnabled,assignedLicenses,assignedPlans")
        
            print(response)
            if response.status_code == 200:
                print(response.json())
                user['userDetails'] = response.json().get("value", [])
            else:
                user['userDetails'] = {}
            users_all_info.append(user)
        return users_all_info


    def get_all_users(self, data='all') -> list:
        """
        Retrieve all users' email addresses.
        :return: A list of email addresses.
        :raises HermesMSGraphError: If the request fails.
        """
        url = "https://graph.microsoft.com/v1.0/users?$top=999&$filter=userType eq 'Member'&$select=id,displayName,mail,officeLocation,userPrincipalName,accountEnabled,assignedLicenses,assignedPlans"

        users = []

        while url:
            response = self.http.get(url)
            if response.status_code != 200:
                raise HermesMSGraphError(f"Error fetching users' emails: {response.status_code} - {response.text}")
                
            response_data = response.json()
            users.extend(response_data.get("value", []))
            url = response_data.get("@odata.nextLink")
        
        match data:
            case 'simple':
                return self.__filter_data(users)
            case 'all':
                #users_all_info = self.__get_license_details(users)
                return users
            case _:
                raise ValueError(f"Invalid data type: {data}")
            
    def search_from_mailboxes(self, query: str) -> list:
        """
        Search for users by email address or display name using $search.
        Args:
            query (str): Search term (e.g., part of an email or name).
        Returns:
            list: List of matching users.
        """
        url = f"https://graph.microsoft.com/v1.0/users?$search=\"displayName:{query}\""

        headers = {
            "ConsistencyLevel": "eventual"
        }

        response = self.http.get(url, headers=headers)

        if response.status_code == 200:
            return response.json().get("value", [])
        else:
            error_message = f"Error searching users by email: {response.status_code} - {response.text}"
            raise HermesMSGraphError(error_message)

    #def add_user_to_shared_mailbox(self, user_email_address, shared_mailbox_address, access_type="ReadWrite"):
     



"""
ProductName;LicensePartNumber;LicenseSKUID
APP CONNECT IW;SPZA_IW;8f0c5670-4e56-4892-b06d-91c085d7004f
Microsoft 365 Audio Conferencing;MCOMEETADV;0c266dff-15dd-4b49-8397-2bb16070ed52
AZURE ACTIVE DIRECTORY BASIC;AAD_BASIC;2b9c8e7c-319c-43a2-a2a0-48c5c6161de7
AZURE ACTIVE DIRECTORY PREMIUM P1;AAD_PREMIUM;078d2b04-f1bd-4111-bbd4-b4b1b354cef4
AZURE ACTIVE DIRECTORY PREMIUM P2;AAD_PREMIUM_P2;84a661c4-e949-4bd2-a560-ed7766fcaf2b
AZURE INFORMATION PROTECTION PLAN 1;RIGHTSMANAGEMENT;c52ea49f-fe5d-4e95-93ba-1de91d380f89
DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION;DYN365_ENTERPRISE_PLAN1;ea126fc5-a19e-42e2-a731-da9d437bffcf
DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION;DYN365_ENTERPRISE_CUSTOMER_SERVICE;749742bf-0d37-4158-a120-33567104deeb
DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION;DYN365_FINANCIALS_BUSINESS_SKU;cc13a803-544e-4464-b4e4-6d6169a138fa
DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION;DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE;8edc2cf8-6438-4fa9-b6e3-aa1660c640cc
DYNAMICS 365 FOR SALES ENTERPRISE EDITION;DYN365_ENTERPRISE_SALES;1e1a282c-9c54-43a2-9310-98ef728faace
DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION;DYN365_ENTERPRISE_TEAM_MEMBERS;8e7a3d30-d97d-43ab-837c-d7701cef83dc
DYNAMICS 365 UNF OPS PLAN ENT EDITION;Dynamics_365_for_Operations;ccba3cfe-71ef-423a-bd87-b6df3dce59a9
ENTERPRISE MOBILITY + SECURITY E3;EMS;efccb6f7-5641-4e0e-bd10-b4976e1bf68e
ENTERPRISE MOBILITY + SECURITY E5;EMSPREMIUM;b05e124f-c7cc-45a0-a6aa-8cf78c946968
EXCHANGE ONLINE (PLAN 1);EXCHANGESTANDARD;4b9405b0-7788-4568-add1-99614e613b69
EXCHANGE ONLINE (PLAN 2);EXCHANGEENTERPRISE;19ec0d23-8335-4cbd-94ac-6050e30712fa
EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE;EXCHANGEARCHIVE_ADDON;ee02fd1b-340e-4a4b-b355-4a514e4c8943
EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER;EXCHANGEARCHIVE;90b5e015-709a-4b8b-b08e-3200f994494c
EXCHANGE ONLINE ESSENTIALS;EXCHANGEESSENTIALS;7fc0182e-d107-4556-8329-7caaa511197b
EXCHANGE ONLINE ESSENTIALS;EXCHANGE_S_ESSENTIALS;e8f81a67-bd96-4074-b108-cf193eb9433b
EXCHANGE ONLINE KIOSK;EXCHANGEDESKLESS;80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82
EXCHANGE ONLINE POP;EXCHANGETELCO;cb0a98a8-11bc-494c-83d9-c1b1ac65327e
INTUNE;INTUNE_A;061f9ace-7d42-4136-88ac-31dc755f143f
Microsoft 365 A1;M365EDU_A1;b17653a4-2443-4e8c-a550-18249dda78bb
Microsoft 365 A3 for faculty;M365EDU_A3_FACULTY;4b590615-0888-425a-a965-b3bf7789848d
Microsoft 365 A3 for students;M365EDU_A3_STUDENT;7cfd9a2b-e110-4c39-bf20-c6a3f36a3121
Microsoft 365 A5 for faculty;M365EDU_A5_FACULTY;e97c048c-37a4-45fb-ab50-922fbf07a370
Microsoft 365 A5 for students;M365EDU_A5_STUDENT;46c119d4-0379-4a9d-85e4-97c66d3f909e
MICROSOFT 365 APPS FOR BUSINESS;O365_BUSINESS;cdd28e44-67e3-425e-be4c-737fab2899d3
MICROSOFT 365 APPS FOR BUSINESS;SMB_BUSINESS;b214fe43-f5a3-4703-beeb-fa97188220fc
MICROSOFT 365 APPS FOR ENTERPRISE;OFFICESUBSCRIPTION;c2273bd0-dff7-4215-9ef5-2c7bcfb06425
MICROSOFT 365 BUSINESS BASIC;O365_BUSINESS_ESSENTIALS;3b555118-da6a-4418-894f-7df1e2096870
MICROSOFT 365 BUSINESS BASIC;SMB_BUSINESS_ESSENTIALS;dab7782a-93b1-4074-8bb1-0e61318bea0b
MICROSOFT 365 BUSINESS STANDARD;O365_BUSINESS_PREMIUM;f245ecc8-75af-4f8e-b61f-27d8114de5f3
MICROSOFT 365 BUSINESS STANDARD;SMB_BUSINESS_PREMIUM;ac5cef5d-921b-4f97-9ef3-c99076e5470f
MICROSOFT 365 BUSINESS PREMIUM;SPB;cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46
MICROSOFT 365 E3;SPE_E3;05e9a617-0261-4cee-bb44-138d3ef5d965
Microsoft 365 E5;SPE_E5;06ebc4ee-1bb5-47dd-8120-11324bc54e06
Microsoft 365 E3_USGOV_DOD;SPE_E3_USGOV_DOD;d61d61cc-f992-433f-a577-5bd016037eeb
Microsoft 365 E3_USGOV_GCCHIGH;SPE_E3_USGOV_GCCHIGH;ca9d1dd9-dfe9-4fef-b97c-9bc1ea3c3658
Microsoft 365 E5 Compliance;INFORMATION_PROTECTION_COMPLIANCE;184efa21-98c3-4e5d-95ab-d07053a96e67
Microsoft 365 E5 Security;IDENTITY_THREAT_PROTECTION;26124093-3d78-432b-b5dc-48bf992543d5
Microsoft 365 E5 Security for EMS E5;IDENTITY_THREAT_PROTECTION_FOR_EMS_E5;44ac31e7-2999-4304-ad94-c948886741d4
Microsoft 365 F1;M365_F1;44575883-256e-4a79-9da4-ebe9acabe2b2
Microsoft 365 F3;SPE_F1;66b55226-6b4f-492c-910c-a3b7a3c9d993
MICROSOFT FLOW FREE;FLOW_FREE;f30db892-07e9-47e9-837c-80727f46fd3d
MICROSOFT 365 PHONE SYSTEM;MCOEV;e43b5b99-8dfb-405f-9987-dc307f34bcbd
MICROSOFT 365 PHONE SYSTEM FOR DOD;MCOEV_DOD;d01d9287-694b-44f3-bcc5-ada78c8d953e
MICROSOFT 365 PHONE SYSTEM FOR FACULTY;MCOEV_FACULTY;d979703c-028d-4de5-acbf-7955566b69b9
MICROSOFT 365 PHONE SYSTEM FOR GCC;MCOEV_GOV;a460366a-ade7-4791-b581-9fbff1bdaa85
MICROSOFT 365 PHONE SYSTEM FOR GCCHIGH;MCOEV_GCCHIGH;7035277a-5e49-4abc-a24f-0ec49c501bb5
MICROSOFT 365 PHONE SYSTEM FOR SMALL AND MEDIUM BUSINESS;MCOEVSMB_1;aa6791d3-bb09-4bc2-afed-c30c3fe26032
MICROSOFT 365 PHONE SYSTEM FOR STUDENTS;MCOEV_STUDENT;1f338bbc-767e-4a1e-a2d4-b73207cc5b93
MICROSOFT 365 PHONE SYSTEM FOR TELSTRA;MCOEV_TELSTRA;ffaf2d68-1c95-4eb3-9ddd-59b81fba0f61
MICROSOFT 365 PHONE SYSTEM_USGOV_DOD;MCOEV_USGOV_DOD;b0e7de67-e503-4934-b729-53d595ba5cd1
MICROSOFT 365 PHONE SYSTEM_USGOV_GCCHIGH;MCOEV_USGOV_GCCHIGH;985fcb26-7b94-475b-b512-89356697be71
Microsoft Defender Advanced Threat Protection;WIN_DEF_ATP;111046dd-295b-4d6d-9724-d52ac90bd1f2
MICROSOFT DYNAMICS CRM ONLINE BASIC;CRMPLAN2;906af65a-2970-46d5-9b58-4e9aa50f0657
MICROSOFT DYNAMICS CRM ONLINE;CRMSTANDARD;d17b27af-3f49-4822-99f9-56a661538792
MS IMAGINE ACADEMY;IT_ACADEMY_AD;ba9a34de-4489-469d-879c-0f0f145321cd
MICROSOFT TEAM (FREE);TEAMS_FREE;16ddbbfc-09ea-4de2-b1d7-312db6112d70
Office 365 A5 for faculty;ENTERPRISEPREMIUM_FACULTY;a4585165-0533-458a-97e3-c400570268c4
Office 365 A5 for students;ENTERPRISEPREMIUM_STUDENT;ee656612-49fa-43e5-b67e-cb1fdf7699df
Office 365 Advanced Compliance;EQUIVIO_ANALYTICS;1b1b1f7a-8355-43b6-829f-336cfccb744c
Office 365 Advanced Threat Protection (Plan 1);ATP_ENTERPRISE;4ef96642-f096-40de-a3e9-d83fb2f90211
OFFICE 365 E1;STANDARDPACK;18181a46-0d4e-45cd-891e-60aabd171b4e
OFFICE 365 E2;STANDARDWOFFPACK;6634e0ce-1a9f-428c-a498-f84ec7b8aa2e
OFFICE 365 E3;ENTERPRISEPACK;6fd2c87f-b296-42f0-b197-1e91e994b900
OFFICE 365 E3 DEVELOPER;DEVELOPERPACK;189a915c-fe4f-4ffa-bde4-85b9628d07a0
Office 365 E3_USGOV_DOD;ENTERPRISEPACK_USGOV_DOD;b107e5a3-3e60-4c0d-a184-a7e4395eb44c
Office 365 E3_USGOV_GCCHIGH;ENTERPRISEPACK_USGOV_GCCHIGH;aea38a85-9bd5-4981-aa00-616b411205bf
OFFICE 365 E4;ENTERPRISEWITHSCAL;1392051d-0cb9-4b7a-88d5-621fee5e8711
OFFICE 365 E5;ENTERPRISEPREMIUM;c7df2760-2c81-4ef7-b578-5b5392b571df
OFFICE 365 E5 WITHOUT AUDIO CONFERENCING;ENTERPRISEPREMIUM_NOPSTNCONF;26d45bd9-adf1-46cd-a9e1-51e9a5524128
OFFICE 365 F1;DESKLESSPACK;4b585984-651b-448a-9e53-3b10f069cf7f
OFFICE 365 F3;DESKLESSPACK;4b585984-651b-448a-9e53-3b10f069cf7f
OFFICE 365 MIDSIZE BUSINESS;MIDSIZEPACK;04a7fb0d-32e0-4241-b4f5-3f7618cd1162
OFFICE 365 SMALL BUSINESS;LITEPACK;bd09678e-b83c-4d3f-aaba-3dad4abd128b
OFFICE 365 SMALL BUSINESS PREMIUM;LITEPACK_P2;fc14ec4a-4169-49a4-a51e-2c852931814b
ONEDRIVE FOR BUSINESS (PLAN 1);WACONEDRIVESTANDARD;e6778190-713e-4e4f-9119-8b8238de25df
ONEDRIVE FOR BUSINESS (PLAN 2);WACONEDRIVEENTERPRISE;ed01faf2-1d88-4947-ae91-45ca18703a96
POWER APPS PER USER PLAN;POWERAPPS_PER_USER;b30411f5-fea1-4a59-9ad9-3db7c7ead579
POWER BI (FREE);POWER_BI_STANDARD;a403ebcc-fae0-4ca2-8c8c-7a907fd6c235
POWER BI FOR OFFICE 365 ADD-ON;POWER_BI_ADDON;45bc2c81-6072-436a-9b0b-3b12eefbc402
POWER BI PRO;POWER_BI_PRO;f8a1db68-be16-40ed-86d5-cb42ce701560
PROJECT FOR OFFICE 365;PROJECTCLIENT;a10d5e58-74da-4312-95c8-76be4e5b75a0
PROJECT ONLINE ESSENTIALS;PROJECTESSENTIALS;776df282-9fc0-4862-99e2-70e561b9909e
PROJECT ONLINE PREMIUM;PROJECTPREMIUM;09015f9f-377f-4538-bbb5-f75ceb09358a
PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT;PROJECTONLINE_PLAN_1;2db84718-652c-47a7-860c-f10d8abbdae3
PROJECT ONLINE PROFESSIONAL;PROJECTPROFESSIONAL;53818b1b-4a27-454b-8896-0dba576410e6
PROJECT ONLINE WITH PROJECT FOR OFFICE 365;PROJECTONLINE_PLAN_2;f82a60b8-1ee3-4cfb-a4fe-1c6a53c2656c
SHAREPOINT ONLINE (PLAN 1);SHAREPOINTSTANDARD;1fc08a02-8b3d-43b9-831e-f76859e04e1a
SHAREPOINT ONLINE (PLAN 2);SHAREPOINTENTERPRISE;a9732ec9-17d9-494c-a51c-d6b45b384dcb
SKYPE FOR BUSINESS ONLINE (PLAN 1);MCOIMP;b8b749f8-a4ef-4887-9539-c95b1eaa5db7
SKYPE FOR BUSINESS ONLINE (PLAN 2);MCOSTANDARD;d42c793f-6c78-4f43-92ca-e8f6a02b035f
SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING;MCOPSTN2;d3b4fe1f-9992-4930-8acb-ca6ec609365e
SKYPE FOR BUSINESS PSTN DOMESTIC CALLING;MCOPSTN1;0dab259f-bf13-4952-b7f8-7db8f131b28d
SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes);MCOPSTN5;54a152dc-90de-4996-93d2-bc47e670fc06
VISIO ONLINE PLAN 1;VISIOONLINE_PLAN1;4b244418-9658-4451-a2b8-b5e2b364e9bd
VISIO Online Plan 2;VISIOCLIENT;c5928f49-12ba-48f7-ada3-0d743a3601d5
WINDOWS 10 ENTERPRISE E3;WIN10_PRO_ENT_SUB;cb10e6cd-9da4-4992-867b-67546b1db821
Windows 10 Enterprise E5;WIN10_VDA_E5;488ba24a-39a9-4473-8ee5-19291e71b002
WINDOWS STORE FOR BUSINESS;WINDOWS_STORE;6470687e-a428-4b7a-bef2-8a291ad947c9
"""