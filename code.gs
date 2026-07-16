/**
 * Code.gs - Backend Server Logic for Google Apps Script
 */

/**
 * 1. RUN THIS FUNCTION ONCE MANUALLY
 * This securely stores your Azure App Registration credentials so they 
 * aren't hardcoded in the source code.
 */
function setSecretCredentials() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties({
    'TENANT_ID': '<entra tenant ID>',
    'CLIENT_ID': '<entra app ID from the Overview page>',
    'CLIENT_SECRET': '<Value of the secret>',
    'redirectUri': '<URL of the deployed app, ending in /exec>',
    'SHEET_ID': '<your Google Sheets ID Here>'
  });
  Logger.log("Credentials securely saved!");
}

/**
 * Serves the webpage, acting as a secure gateway.
 */
function doGet(e) {
  const props = PropertiesService.getScriptProperties().getProperties();
  const tenantId = props['TENANT_ID'];
  const clientId = props['CLIENT_ID'];
  const clientSecret = props['CLIENT_SECRET'];
  const redirectUri = props['redirectUri'];
  
  //const redirectUri = "https://script.google.com/macros/s/AKfycbzB5D-Ot3tlrYaBbWb4dLPyI9E_CybdimFL_tTpfEUh7r5NPTUES5WXMwv5ejh6CKWL/exec";
  //const redirectUri = "https://script.google.com/macros/s/AKfycbwhn9brXG_mLkps6CZw8FH_4cnBuzoC7tt133ZMjhY/dev";

  // 1. Returning from Microsoft with an authorization code
  if (e.parameter.code) {
    const expectedState = CacheService.getUserCache().get("oauth_state");
    if (!e.parameter.state || e.parameter.state !== expectedState) {
      return HtmlService.createHtmlOutput(`
        <div style="font-family: sans-serif; text-align: center; margin-top: 50px;">
          <h2>Security Validation Failed</h2>
          <p>Invalid OAuth state. Please <a href="${redirectUri}" target="_top">click here to try again</a>.</p>
        </div>
      `);
    }
    CacheService.getUserCache().remove("oauth_state");

    try {
      // const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
      const tokenUrl = `https://login.microsoftonline.com/organizations/oauth2/v2.0/token`;

      // CHANGE: We now request the Graph API default scope to get the user's access token
      const payload = {
        client_id: clientId,
        scope: "https://graph.microsoft.com/.default offline_access openid profile email",
        code: e.parameter.code,
        redirect_uri: redirectUri,
        grant_type: "authorization_code",
        client_secret: clientSecret
      };
      
      const options = {
        method: "post",
        contentType: "application/x-www-form-urlencoded",
        payload: payload,
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(tokenUrl, options);
      const tokenResponse = JSON.parse(response.getContentText());
      
      if (!tokenResponse.access_token) {
        // Extract Microsoft's exact error details
        const msError = tokenResponse.error_description || JSON.stringify(tokenResponse);
        throw new Error("Microsoft rejected the token request: " + msError);
      } 
      
      const userEmail = tokenResponse.id_token ? decodeIdToken(tokenResponse.id_token) : "Unknown User";
      
      const sessionToken = Utilities.getUuid() + "-" + Utilities.getUuid();
      
      // CHANGE: Cache the user's email AND their specific Graph Access Token
      CacheService.getScriptCache().put(sessionToken, userEmail, 3600); 
      CacheService.getScriptCache().put(sessionToken + "_access", tokenResponse.access_token, 3500); // Store for ~1 hour
      
      const template = HtmlService.createTemplateFromFile('index');
      template.sessionToken = sessionToken;
      template.userEmail = userEmail;
      
      return template.evaluate()
        .setTitle('Intune Admin Tools')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
    } catch (err) {
      return HtmlService.createHtmlOutput("<h1>Authentication Error</h1><p>" + err.message + "</p>");
    }
  }

  // 2. User arrives with a valid session token in URL
  if (e.parameter.sessionToken) {
    const cachedEmail = CacheService.getScriptCache().get(e.parameter.sessionToken);
    if (cachedEmail) {
      const template = HtmlService.createTemplateFromFile('index');
      template.sessionToken = e.parameter.sessionToken;
      template.userEmail = cachedEmail;
      return template.evaluate()
        .setTitle('Intune Admin Tools')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
    }
  }

  // 3. Initial load - Needs Authentication
  const state = Utilities.getUuid();
  CacheService.getUserCache().put("oauth_state", state, 300);
  
  // CHANGE: The authorization URL must request Microsoft Graph scopes from the user
  const scopes = encodeURIComponent("https://graph.microsoft.com/.default offline_access openid profile email");
  // const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodeURIComponent(redirectUri)}&response_mode=query&scope=${scopes}&state=${state}`;
  const authUrl = `https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodeURIComponent(redirectUri)}&response_mode=query&scope=${scopes}&state=${state}`;
  
  const loginHtml = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; text-align: center; margin-top: 100px; color: #333;">
      <h2 style="color: #0f2b5b;">Intune Admin Tools</h2>
      <p>Authentication is required to access this portal.</p>
      <br>
      <a href="${authUrl}" target="_top" style="background-color: #0f2b5b; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px;">
        Sign In with Microsoft
      </a>
    </div>
  `;
  
  return HtmlService.createHtmlOutput(loginHtml);
}

/**
 * Decodes the Entra ID token to safely extract the user's email address
 */
function decodeIdToken(idToken) {
  const base64Url = idToken.split('.')[1];
  const decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(base64Url)).getDataAsString();
  const payload = JSON.parse(decoded);
  return payload.email || payload.upn || payload.preferred_username;
}

/**
 * Retrieves the logged-in user's Graph API Access Token from the cache
 */
function getGraphToken(sessionToken) {
  if (!sessionToken) {
    throw new Error("No session token provided.");
  }
  
  // Fetch the token stored during doGet
  const token = CacheService.getScriptCache().get(sessionToken + "_access");
  
  if (!token) {
    throw new Error("Session expired or invalid. Please refresh the page and sign in again.");
  }
  
  return token;
}

// ************************************************ Begin Tool scripts ***************************************
// ------------------------------------------------ Get duplicate devices ------------------------------------
/**
 * Called by the HTML button. Fetches all Windows devices from Intune,
 * finds duplicate serial numbers, and returns the sorted array to the browser.
 */
function getDuplicateSerials(sessionToken) {
  verifySession(sessionToken);
  const token = getGraphToken(sessionToken);
  // We only pull the 4 columns we actually care about to speed up the query drastically.
  let url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$select=id,deviceName,serialNumber,operatingSystem";
  const options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + token,
      'ConsistencyLevel': 'eventual'
    },
    'muteHttpExceptions': true
  };
  let allWindowsDevices = [];
  // Graph API uses pagination. We must loop until there's no '@odata.nextLink'
  while (url) {
    let response = UrlFetchApp.fetch(url, options);
    let json = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() !== 200) {
       throw new Error("Graph API Error: " + (json.error?.message || "Something went wrong fetching devices."));
    }

    // Filter out only valid Windows devices natively in JS
    let devices = json.value.filter(device => 
      device.operatingSystem === "Windows" &&
      device.serialNumber !== null && 
      device.serialNumber.trim() !== ""
    );
    
    allWindowsDevices = allWindowsDevices.concat(devices);
    url = json['@odata.nextLink']; // Get the next page URL
  }

  // Group devices by SerialNumber
  const groupedSerials = {};
  allWindowsDevices.forEach(device => {
    let serial = device.serialNumber;
    if (!groupedSerials[serial]) {
      groupedSerials[serial] = [];
    }
    groupedSerials[serial].push(device);
  });

  // Build the final array of just duplicates
  const finalOutput = [];
  
  // Iterate through the grouped Object
  for (const [serial, devices] of Object.entries(groupedSerials)) {
    // Only process groups more than 1 (Duplicates!)
    if (devices.length > 1) {
      
      // Add the Header Row for the Group
      finalOutput.push({
        SerialNumber: serial,
        DeviceName: "--",
        Id: "--"
      });

      // Add the individual duplicate devices underneath it
      devices.forEach(d => {
        finalOutput.push({
          SerialNumber: "",
          DeviceName: d.deviceName,
          Id: d.id
        });
      });
    }
  }

  return finalOutput;
}
// ------------------------------------------------ End Finding Serial Dupe scripts ---------------------------------------

// ------------------------------------------------ Begin Finding Group Assignment scripts ---------------------------------------

/**
 * 1. Finds the Entra ID Group and returns its ID and Name
 */
function getEntraGroup(groupName, sessionToken) {
  verifySession(sessionToken);
  const token = getGraphToken(sessionToken);
  const baseUrl = "https://graph.microsoft.com/beta";
  const groupQuery = `${baseUrl}/groups?$filter=displayName eq '${encodeURIComponent(groupName)}'`;
  const groupRes = UrlFetchApp.fetch(groupQuery, {
    method: "get",
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  });
  
  const groupData = JSON.parse(groupRes.getContentText());
  if (!groupData.value || groupData.value.length === 0) {
    throw new Error("Group not found: " + groupName);
  }
  
  return { 
    id: groupData.value[0].id, 
    displayName: groupData.value[0].displayName 
  };
}

/**
 * 2. Checks a SINGLE Intune endpoint for assignments matching the Group ID
 */
function checkIntuneEndpoint(groupId, endpoint, sessionToken) {
  verifySession(sessionToken);
  const token = getGraphToken(sessionToken);
  const baseUrl = "https://graph.microsoft.com/beta";
  
  let url = baseUrl + endpoint.path;
  let assignedItems = [];

  while (url) {
    let res = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });
    
    let data = JSON.parse(res.getContentText());
    if (res.getResponseCode() !== 200 || !data.value) break;

    // Filter items where an assignment matches the Group ID
    let matched = data.value.filter(item => {
      // Handle Role Assignments
      if (item.members && Array.isArray(item.members)) {
        return item.members.includes(groupId);
      }
      if (!item.assignments) return false;
      return item.assignments.some(assignment => {
        let target = assignment.target || assignment;
        return (target.groupId === groupId || target.id === groupId || assignment.id === groupId);
      });
    });

    // Extract the display name
    matched.forEach(item => {
      assignedItems.push(item.displayName || item.name || "Unknown Name");
    });

    url = data["@odata.nextLink"] || null;
  }

  return { 
    categoryName: endpoint.name, 
    items: assignedItems 
  };
}
// ------------------------------------------------ End Finding Group Assignment scripts ---------------------------------------
// ------------------------------------------------ Start Finding Keyword Scripts ---------------------------------------

function getAllPolicyMetadata(sessionToken) {
  verifySession(sessionToken);
  const token = getGraphToken(sessionToken);
  const headers = { "Authorization": "Bearer " + token };
  const options = { "method": "get", "headers": headers, "muteHttpExceptions": true };
  let allPolicies = [];

  // --- 1. Settings Catalog Policies ---
  let scUrl = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$select=id,name";
  while (scUrl) {
    let res = UrlFetchApp.fetch(scUrl, options);
    if (res.getResponseCode() === 200) {
      let data = JSON.parse(res.getContentText());
      if (data.value) {
        data.value.forEach(p => allPolicies.push({
          id: p.id, 
          name: p.name, 
          type: "SettingsCatalog", 
          ui: "Devices > Configuration"
        }));
      }
      scUrl = data['@odata.nextLink'] || null;
    } else break;
  }

  // --- 2. Legacy Templates (Device Configurations) ---
  let dcUrl = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?$select=id,displayName";
  while (dcUrl) {
    let res = UrlFetchApp.fetch(dcUrl, options);
    if (res.getResponseCode() === 200) {
      let data = JSON.parse(res.getContentText());
      if (data.value) {
        data.value.forEach(p => allPolicies.push({
          id: p.id, 
          name: p.displayName, 
          type: "Template", 
          ui: "Devices > Configuration"
        }));
      }
      dcUrl = data['@odata.nextLink'] || null;
    } else break;
  }

  // --- 3. Endpoint Security Policies (Intents) ---
  let intentUrl = "https://graph.microsoft.com/beta/deviceManagement/intents?$select=id,displayName";
  while (intentUrl) {
    let res = UrlFetchApp.fetch(intentUrl, options);
    if (res.getResponseCode() === 200) {
      let data = JSON.parse(res.getContentText());
      if (data.value) {
        data.value.forEach(p => allPolicies.push({
          id: p.id, 
          name: p.displayName, 
          type: "EndpointSecurity", 
          ui: "Endpoint Security"
        }));
      }
      intentUrl = data['@odata.nextLink'] || null;
    } else break;
  }

  return allPolicies;
}
// ------------------------------------------------ Begin Finding Group Assignment scripts --------------------------------

// --- CACHE HELPER ---
// Stores group names during the run so we don't query Graph API multiple times for the same group
const groupNameCache = {};

function getGroupNameCached(groupId, headers) {
  // If we've already looked up this ID, return the saved name immediately
  if (groupNameCache[groupId]) {
    return groupNameCache[groupId];
  }

  // Otherwise, query Graph API for it
  const url = `https://graph.microsoft.com/v1.0/groups/${groupId}?$select=displayName`;
  const res = UrlFetchApp.fetch(url, { headers: headers, muteHttpExceptions: true });
  const responseCode = res.getResponseCode();
  
  if (responseCode === 200) {
    const name = JSON.parse(res.getContentText()).displayName;
    groupNameCache[groupId] = name; // Save to cache
    return name;
  } else if (responseCode === 404) {
    // 404 means the group was deleted from Entra ID!
    groupNameCache[groupId] = "<em>&lt;Deleted Group&gt;</em> (ID: " + groupId + ")";
    return groupNameCache[groupId];
  } else {
    // If it fails for any other reason, save the ID
    groupNameCache[groupId] = "Group ID: " + groupId;
    return groupNameCache[groupId];
  }
}

// --- UPDATED BATCH SEARCH ---
function searchPoliciesBatch(policiesChunk, keyword, sessionToken) {
  verifySession(sessionToken);
  const token = getGraphToken(sessionToken);
  const headers = { "Authorization": "Bearer " + token };
  let lowerKeyword = keyword.toLowerCase();

  let requests = policiesChunk.map(p => {
    let url = "";
    if (p.type === "SettingsCatalog") url = `https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('${p.id}')/settings`;
    else if (p.type === "EndpointSecurity") url = `https://graph.microsoft.com/beta/deviceManagement/intents('${p.id}')/settings`;
    else url = `https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations('${p.id}')`;
    
    return { url: url, headers: headers, muteHttpExceptions: true };
  });

  let responses = UrlFetchApp.fetchAll(requests);
  let matchedResults = [];

  for (let i = 0; i < responses.length; i++) {
    let res = responses[i];
    let p = policiesChunk[i];
    let matched = false;

    // 1. Check Title First
    let pName = p.name ? p.name : "Unknown Policy";
    let titleMatch = pName.toLowerCase().indexOf(lowerKeyword) !== -1;
    let payloadMatch = false;
    let payloadSnippet = "";

    // 2. Check HTTP Response for JSON settings payload
    let code = res.getResponseCode();
    if (code === 200) {
      let content = res.getContentText();
      let matchIndex = content.toLowerCase().indexOf(lowerKeyword);
      if (matchIndex !== -1) {
        payloadMatch = true;
        let start = Math.max(0, matchIndex - 40);
        let end = Math.min(content.length, matchIndex + keyword.length + 60);
        payloadSnippet = "..." + content.substring(start, end).replace(/\r?\n|\r/g, " ").trim() + "...";
      }
    } else {
      payloadMatch = true;
      let errorText = res.getContentText().substring(0, 80);
      payloadSnippet = `⚠️ API Error ${code}: ${errorText}...`;
    }

    // 3. COMBINE AND FORMAT CLEANLY
    if (titleMatch || payloadMatch) {
      matched = true;
      let displaySnippets = [];
      
      // Add a bold indicator if the title matched
      if (titleMatch) {
        displaySnippets.push("<strong>✅ Matched Policy Name</strong>");
      }
      
      // Add the payload snippet on its own line if it matched the settings
      if (payloadMatch && payloadSnippet !== "") {
        // Wrap it in italics and a slightly muted color so it looks like code/context
        displaySnippets.push(`<span style="color: #555; font-style: italic;">${payloadSnippet}</span>`);
      }

      // Join them with a double line break so they stack neatly in the table
      p.snippet = displaySnippets.join("<br><br>");

      let assignUrl = "";
      if (p.type === "SettingsCatalog") assignUrl = `https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('${p.id}')/assignments`;
      else if (p.type === "EndpointSecurity") assignUrl = `https://graph.microsoft.com/beta/deviceManagement/intents('${p.id}')/assignments`;
      else assignUrl = `https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations('${p.id}')/assignments`;

      let assignRes = UrlFetchApp.fetch(assignUrl, { headers: headers, muteHttpExceptions: true });
      let assignList = [];

      if (assignRes.getResponseCode() === 200) {
        let aData = JSON.parse(assignRes.getContentText());
        if (aData.value && aData.value.length > 0) {
          aData.value.forEach(a => {
            let target = a.target || a;
            let type = target["@odata.type"] || "";
            
            if (type.includes("allDevices")) {
              assignList.push("All Devices");
            } else if (type.includes("allLicensedUsers")) {
              assignList.push("All Users");
            } else {
              // Safely extract the ID from whichever property Microsoft decided to use this time
              let groupId = target.groupId || target.id || a.id;
              
              if (groupId) {
                // Call our new caching function
                assignList.push(getGroupNameCached(groupId, headers));
              } else {
                assignList.push("Unknown Target");
              }
            }
          });
        }
      }

      // Format to show a max of 3
      if (assignList.length === 0) {
        p.assignments = "<em>Unassigned</em>";
      } else {
        let top3 = assignList.slice(0, 3);
        p.assignments = top3.join("<br>");
        if (assignList.length > 3) p.assignments += `<br><strong>...and ${assignList.length - 3} more</strong>`;
      }

      // THE FIX IS HERE: Push p.snippet instead of the empty snippets.join()
      matchedResults.push({ policy: p, snippet: p.snippet });
    }
  }
  
  return matchedResults;
}

/**
 * Appends a log entry to the designated Google Sheet.
 */
function logAction(userEmail, action, details, sessionToken) {
  verifySession(sessionToken);
  const sheetId = "<your Google Sheets ID Here>"; 
  
  try {
    // Check if they forgot to replace the placeholder
    if (!sheetId || sheetId.includes("your Google Sheets ID")) {
      throw new Error("Sheet ID placeholder was not replaced.");
    }
    
    const sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
    const timestamp = new Date(); 
    sheet.appendRow([timestamp, userEmail, action, details]);
    
  } catch (e) {
    // Throw this to the frontend
    throw new Error("Logging failed: " + e.message);
  }
}

function triggerPermissions() {
  SpreadsheetApp.openById("<your Google Sheets ID Here>");
}

/**
 * Bulk deletes Intune devices by Serial Number.
 * @param {Array} serials - Array of serial numbers.
 */
/**
 * STEP 1: Dry Run. Verifies serial numbers against Intune, handles duplicates, 
 * and returns the specific Intune Device IDs for deletion.
 */
function verifyBulkDelete(serials, sessionToken) {
  verifySession(sessionToken);
  const token = getGraphToken(sessionToken);
  const results = [];
  
  // Filter out empty or blank rows
  const validSerials = [];
  for (let i = 0; i < serials.length; i++) {
    const serial = serials[i];
    if (serial && serial.trim() !== "") {
      validSerials.push(serial.trim());
    }
  }
  
  if (validSerials.length === 0) {
    return results;
  }
  
  // Build batch request array for parallel execution
  const requests = validSerials.map(function(serial) {
    return {
      url: "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$filter=serialNumber eq '" + encodeURIComponent(serial) + "'&$select=id,deviceName",
      method: "get",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    };
  });
  
  // Execute all Graph API lookups concurrently
  const responses = UrlFetchApp.fetchAll(requests);
  
  // Process the batch results
  for (let i = 0; i < validSerials.length; i++) {
    const serial = validSerials[i];
    const res = responses[i];
    const code = res.getResponseCode();
    
    if (code !== 200) {
      results.push({ serial: serial, status: "Error", details: "Graph API Search Error (HTTP " + code + ")" });
      continue;
    }
    
    try {
      const searchData = JSON.parse(res.getContentText());
      if (!searchData.value || searchData.value.length === 0) {
        results.push({ serial: serial, status: "Not Found", details: "No device matches this SN" });
      } else if (searchData.value.length === 1) {
        // EXACTLY ONE MATCH - SAFE TO DELETE
        results.push({ 
          serial: serial, 
          status: "Ready", 
          id: searchData.value[0].id, 
          name: searchData.value[0].deviceName,
          details: "Safe to delete"
        });
      } else {
        // DUPLICATES FOUND - ABORT DELETION FOR THIS SN
        results.push({ 
          serial: serial, 
          status: "Duplicate", 
          details: "Found " + searchData.value.length + " devices with this SN. Skipping." 
        });
      }
    } catch (e) {
      results.push({ serial: serial, status: "Error", details: e.message });
    }
  }
  
  return results;
}

/**
 * STEP 2: Execution. Takes the validated array of objects and issues the DELETE command.
 */
function executeBulkDelete(devicesToProcess, sessionToken) {
  verifySession(sessionToken);
  const token = getGraphToken(sessionToken);
  const results = [];
  
  if (!devicesToProcess || devicesToProcess.length === 0) {
    return results;
  }
  
  // Build batch delete array
  const requests = devicesToProcess.map(function(target) {
    return {
      url: "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/" + target.id,
      method: "delete",
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    };
  });
  
  // Execute all deletions in parallel
  const responses = UrlFetchApp.fetchAll(requests);
  
  // Process delete outcomes
  for (let i = 0; i < devicesToProcess.length; i++) {
    const target = devicesToProcess[i];
    const res = responses[i];
    const delCode = res.getResponseCode();
    
    if (delCode === 200 || delCode === 204) {
      results.push({ serial: target.serial, status: "Deleted Successfully" });
    } else {
      results.push({ serial: target.serial, status: "Failed (HTTP " + delCode + ")" });
    }
  }
  
  return results;
}

function verifySession(sessionToken) {
  if (!sessionToken || CacheService.getUserCache().get(sessionToken)) {
    throw new Error("Unauthorized: Session expired or invalid. Please refresh the page to log in again.");
  }
}

/**
 * TEST FUNCTION: Fetch raw Role Assignments from Graph API
 */
function debugRoleAssignments(sessionToken) {
  // Use your existing token generation functions
  const token = getGraphToken(sessionToken); 
  const url = "https://graph.microsoft.com/v1.0/deviceManagement/roleAssignments";
  
  const options = {
    method: "get",
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  
  // Return the raw JSON back to the frontend
  return JSON.parse(response.getContentText());
}

/**
 * Retrieves purely Security Groups (filters out M365/Class groups).
 * Returns an array of group display names for the frontend dropdown.
 */
function getSecurityGroups(sessionToken) {
  verifySession(sessionToken);
  const token = getGraphToken(sessionToken);

  // Filtering for securityEnabled=true and mailEnabled=false neatly 
  // excludes all auto-generated Teams/Class Microsoft 365 groups.
  let url = "https://graph.microsoft.com/v1.0/groups?$filter=securityEnabled eq true and mailEnabled eq false&$select=id,displayName&$top=999";
  
  const options = {
    method: "get",
    headers: {
      "Authorization": "Bearer " + token,
      "ConsistencyLevel": "eventual" 
    },
    muteHttpExceptions: true
  };

  let securityGroups = [];

  // Handle Graph API pagination
  while (url) {
    let response = UrlFetchApp.fetch(url, options);
    let json = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
      throw new Error("Graph API Error fetching groups: " + (json.error?.message || response.getContentText()));
    }

    if (json.value) {
      json.value.forEach(g => {
        if (g.displayName) securityGroups.push(g.displayName);
      });
    }

    url = json['@odata.nextLink'] || null;
  }

  // Sort alphabetically so the dropdown is easy to read
  return securityGroups.sort((a, b) => a.localeCompare(b));
}
