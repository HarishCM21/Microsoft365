# Microsoft365
### Purpase: This script help to fetch Microsoft 365 events using Graph API
### Creator: Harish C M
### Created on: 28-08-2023
### Version: 1.0
### Provide your feedback on a mail with subject "Microsoft 365" at "harishcm21@gmail.com" to update the script to help the users

Steps: The application requried to create  Azure AD application as shown explained below.
1. Login to portal.azure.com
2. Search for "Azure Active Directory" and click on it.
3. On the left panel click on "App registrations".
4. Click on the "New registration"
5. Provide any readble name, choose "Single Tenant"(we suggest), and redirect URI select platform "Web" entire url "http://localhost"
6. Click on "Register"
7. Once the application published open the created application available under "All Application"
8. Copy the "Application (client) ID" and "Directory (tenant) ID" (which used in the powershell application)
9. Click on the "Certificates & secrets" in left panel and click on "Client secrets" click on the "+ New client secret"
10. Provide readble name, choose duraction as per the need(suggest to use least possible duration), and click on "Add"
11. Copy the value available under "Value"(Note: Value available only once) (which used in the powershell application)
12. Assign the permission to the application, powershell application provide permission needed to fetch the events/logs from Microsoft 365 which need to be added as explained below steps
13. Click on the "API permissions", Click on the "+ Add a permission"
14. Under Microsoft APIs, click on "Microsoft Graph", and click on "Application permissions"
15. Check the permission mentioned in the Powershell application and click on "Add permissions"
16. click on the "Grant admin consent for <your org name> (repeat step 13 to 16 when you like to add permission to fetch different reports)

## Note: If "View Report" button enable after clicking, it means there is no events/logs available for 30days
## Note: There is no days limitation hard coded on the script.
