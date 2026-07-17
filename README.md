
# Purpose

Create a GAS (Google Apps Script) website that logs in to your Entra tenant and runs some simple tools using MS Graph calls to do things like type in a group name and see all the places that group is assigned, find devices with duplicate serial numbers, or provide a CSV of device serial numbers to delete those objects out of Intune.

# Requirements
* You'll need access in Google (or know someone who does) to create a GAS web page.
* You'll need access in Entra (or know someone who does) to create an App registration.
* A willingness to allow this site to access your Google Sheets (for logging/display) and run third-party web content (to log into MS).

# Setup

### A) In Google:
1. Head to the [GAS console](https://script.google.com/home) > New project.
2. Copy the `Code.gs` from this repo into the editor.
3. Click the + sign > HTML.
4. Name it `index` (it will add .html automatically, since, ya know, you told it it was an HTML file).
5. Copy the `index.html` from this repo into the editor.
6. Click Deploy > New deployment:
	* Select type > Web app
	* Give it a Description (because you're a good admin who always includes descriptions)
	* Execute as > Me
	* Who has access > whoever you want
	* You might get prompted to Authorize access here... do so.
7. Make note of the Web app URL (ends in `/exec`).
8. **[OPTIONAL:]** If you want to develop further, you can click Deploy again:
	* Test deployments
	* Make note of the Web app URL here (ends in `/dev`)
9. **[OPTIONAL:]** You can go to the website now (the one that ends in `/exec` or `/dev`), but it's not fully working yet. But you'll need to give it permissions eventually, so you can do that now. You'll be prompted to allow access.

### B) In Entra:
1. Head to Entra ID > App registrations > New registration:
	* Give it a name
	* Single tenant Only
	* Redirect URI
		* Platform: Web
		* URL: Put the `/exec` URL here; you can add the `/dev` URL later.
	* Register
2. Go to the Certificates & secrets blade by clicking it in the Resource menu:
	1. New Client Secret
	2. Enter a description, because you're awesome
	3. Choose an expiration length
	4. Add
3. **[IMPORTANT!]** Copy the secret now! If you don't, you'll lose the Value of the secret and you'll need to create a new one.
4. Go to the Overview of the app and copy the Application (client) ID.
5. Go to Entra ID Enterprise apps > find this app you just created.
6. Go to Properties and change "Assignment required?" to Yes.
7. Under Users and groups, add whichever users and groups you want to be able to use these tools.
8. Go back to App registrations > Find this new app.
9. Go to the API Permissions blade under Manage > API permissions > Add a permission:
	* Microsoft Graph
	* Delegated permissions
	* Add:
		* `DeviceManagementConfiguration.Read.All`
		* `DeviceManagementManagedDevices.ReadWrite.All`
			* Yes, there is a tool for bulk deleting device objects from a CSV.  Feel free to not grant this if you don't want to use this tool
		* `DeviceManagementRBAC.Read.All`
		* `Group.Read.All`
		* `User.Read`
10. Click "Grant admin consent for \<your tenant>" > Yes.
  
### C) Back to Google:
We need to create Script Properties, which we'll do manually for security purposes
1. Click the Project Settings gear
2. Scroll down to the Script Properties section and click Add script Property
	1. Property: ```CLIENT_ID```	Value: \<entra's application (client) ID>
	2. Add script property
	3. Property: ```CLIENT_SECRET``` Value:\<the Value of the Secret>
	4. Add script property
	5. Property: ```SHEET_ID``` Value:\<the ID of a Google sheet for logging>
	6. Add script property
	7. Property: ```TENANT_ID``` Value:\<your entra tenant ID found in Entra ID > overview>
	8. Add script property
	9. Property: ```redirectUri``` Value:\<this Script's URL ending in /exec>
3. In the `setSecretCredentials` function, set the specific Entra tenant ID (Entra ID > Overview), Client ID (from the Application (client) ID field of the newly registered app that you correctly followed instructions for), Client Secret (from the Value...), and this GAS' URL ending with `/exec`.
4. Search for the functions `logAction` and `triggerPermissions` and insert the ID of a Google Sheet you created for logging.
5. Next to the Debug button, make sure the dropdown box has the `setSecretCredentials` function selected.
6. Click Run.
7. **[OPTIONAL:]** You can check this worked not only in the run log, but also by clicking the Project settings gear and scrolling down to the Script Properties.
8. **[IMPORTANT!]** Delete those values in the `Code.gs` file now! Leaving those there will expose the secret to anyone savvy enough.
9. Click Save (or Ctrl+S).
10. Click Deploy > Manage Deployments > Edit pencil:
	* Change the version to New version
	* Add yet another dramatically descriptive description
11. Click Deploy > Done.
12. Go to the web page (make sure you're using an account that has access to the GAS as set in A.6) > Click the Sign In with Microsoft button.
13. Log in with your MS credentials.
14. Accept the perms.
15. Use dem tools.
