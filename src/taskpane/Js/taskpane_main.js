// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const { createNestablePublicClientApplication } = require("@azure/msal-browser");

let pca = undefined;
let isPCAInitialized = false;
let requestId = generateUUID();
let logLines = [];


function log(message, error) {
  console.log(message, error);
  logLines.push({
    timestamp: Date.now().toString(),
    line: message + (error ? ' ' + JSON.stringify(error) : ''),
    level: error ? 'ERROR' : 'INFO',
  });
}

function initializePCA() {
  return new Office.Promise(function (resolve, reject) {
    if (isPCAInitialized) {
      resolve();
      return;
    }
    
    // Initialize the public client application
    createNestablePublicClientApplication({
      auth: {
        clientId: "3e201f82-64a7-469e-90d6-28722990edb5", // Replace with your actual client ID if not using env vars
        authority: 'https://login.microsoftonline.com/common',
      },
    }).then(function(localPca) {
      pca = localPca;
      isPCAInitialized = true;
      log('PCA initialized successfully');
      resolve();
    }).catch(function(error) {
      log('Error creating PCA', error);
      reject(error);
    });
  });
}

function save_user_settings_to_roaming_settings()
{
  Office.context.roamingSettings.saveAsync(function (asyncResult)
  {
	console.log("save_user_info_str_to_roaming_settings - " + JSON.stringify(asyncResult));
  });
}

function disable_client_signatures_if_necessary()
{
  if ($("#checkbox_sig").prop("checked") === true)
  {
	Office.context.mailbox.item.disableClientSignatureAsync(function (asyncResult)
	{
	  console.log("disable_client_signature_if_necessary - " + JSON.stringify(asyncResult));
	});
  }
}
function generateUUID() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
    var r = (Math.random() * 16) | 0,
      v = c == 'x' ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

function save_signature_settings()
{
  let user_info_str = localStorage.getItem('user_info');

  if (user_info_str)
  {
	if (!_user_info)
	{
	  _user_info = JSON.parse(user_info_str); 
	}

	Office.context.roamingSettings.set('user_info', user_info_str);
	// Office.context.roamingSettings.set('newMail', $("#new_mail option:selected").val());
	// Office.context.roamingSettings.set('reply', $("#reply option:selected").val());
	// Office.context.roamingSettings.set('forward', $("#forward option:selected").val());
	// Office.context.roamingSettings.set('override_olk_signature', $("#checkbox_sig").prop('checked'));

	save_user_settings_to_roaming_settings();

	disable_client_signatures_if_necessary();

	$("#message").show("slow");
  }
  else
  {
	// TBD display an error somewhere?
  }
}

function headers(token) {
  return {
    'Accept': 'application/json',
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + token,
    'X-Request-ID': requestId,
    'X-Client-Type': 'OutlookAddin',
  };
}

// Get request headers with token, similar to getRequestHeaders in the first code
function getRequestHeaders() {
  const tokenRequest = {
    scopes: ['User.Read', 'openid', 'profile', 'email'],
  };
  
  // Handle mobile platforms separately
  if (Office.context.diagnostics.platform === Office.PlatformType.Android || 
      Office.context.diagnostics.platform === Office.PlatformType.iOS) {
    return new Office.Promise(function (resolve, reject) {
      initializePCA().then(function() {
        pca.acquireTokenSilent(tokenRequest).then(function(userAccount) {
          log('Token acquired silently for mobile');
          resolve(headers(userAccount.idToken));
        }).catch(function(error) {
          log('Error acquiring token silently for mobile', error);
          reject(error);
        });
      });
    });
  } else {
    // For desktop/web platforms
    return new Office.Promise(function (resolve, reject) {
      // Try OfficeRuntime.auth first
      if (OfficeRuntime && OfficeRuntime.auth) {
        OfficeRuntime.auth.getAccessToken()
          .then(function (token) {
            log('Token acquired from OfficeRuntime');
            resolve(headers(token));
          }).catch(function (error) {
            log('Error acquiring token from OfficeRuntime, falling back to MSAL', error);
            // Fall back to MSAL if OfficeRuntime fails
            initializePCA().then(function() {
              pca.acquireTokenSilent(tokenRequest).then(function(userAccount) {
                log('Token acquired silently from MSAL fallback');
                resolve(headers(userAccount.idToken));
              }).catch(function(error) {
                log('Error acquiring token silently from MSAL fallback', error);
                reject(error);
              });
            }).catch(function(error) {
              log('Failed to initialize PCA for fallback', error);
              reject(error);
            });
          });
      } else {
        // If OfficeRuntime.auth is not available, use MSAL directly
        initializePCA().then(function() {
          pca.acquireTokenSilent(tokenRequest).then(function(userAccount) {
            log('Token acquired silently (OfficeRuntime unavailable)');
            resolve(headers(userAccount.idToken));
          }).catch(function(error) {
            log('Error acquiring token silently (OfficeRuntime unavailable)', error);
            reject(error);
          });
        }).catch(function(error) {
          log('Failed to initialize PCA (OfficeRuntime unavailable)', error);
          reject(error);
        });
      }
    });
  }
}


function set_body(str)
{
  Office.context.mailbox.item.body.setAsync
  (
	get_cal_offset() + str,

	{
		coercionType: Office.CoercionType.Html
	},

	function (asyncResult)
	{
	  console.log("set_body - " + JSON.stringify(asyncResult));
	}
  );
}

function set_signature(str)
{
  Office.context.mailbox.item.body.setSignatureAsync
  (
	str,

	{
		coercionType: Office.CoercionType.Html
	},

	function (asyncResult)
	{
	  console.log("set_signature - " + JSON.stringify(asyncResult));
	}
  );
}

function insert_signature(str)
{
  if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Appointment)
  {
	set_body(str);
  }
  else
  {
	set_signature(str);
  }
}

function test_template_A()
{
	let str = get_template_A_str(_user_info);
	console.log("test_template_A - " + str);

	insert_signature(str);
}

function test_template_B()
{
	let str = get_template_B_str(_user_info);
	console.log("test_template_B - " + str);

	insert_signature(str);
}

function test_template_C()
{
	let str = get_template_C_str(_user_info);
	console.log("test_template_C - " + str);

	insert_signature(str);
}

async function fetchSignatureFromSyncSignature() {
  try {
      console.log("Fetching signature from SyncSignature...");
      
      // Retrieve user info from localStorage
      let user_info_str = localStorage.getItem('user_info');
      if (!user_info_str) {
          console.warn("No user_info found in localStorage.");
          return null;
      }

      let _user_info;
      try {
          _user_info = JSON.parse(user_info_str);
          console.log("Parsed user_info:", _user_info);
      } catch (parseError) {
          console.error("Error parsing user_info from localStorage:", parseError);
          return null;
      }

      if (!_user_info || !_user_info.email) {
          console.warn("User info is missing or does not contain an email.");
          return null;
      }

      // Store user info in Office roaming settings
      Office.context.roamingSettings.set('user_info', user_info_str);
      console.log("User info set in roaming settings.");

      save_user_settings_to_roaming_settings();
      console.log("User settings saved to roaming settings.");

      disable_client_signatures_if_necessary();
      console.log("Checked and disabled client signatures if necessary.");

      const apiUrl = `https://server.dev.syncsignature.com/main-server/api/syncsignature?email=${encodeURIComponent(_user_info.email)}`;
      console.log("Making API request to:", apiUrl);

      // Get authentication headers with token
      let authHeaders;
      try {
          authHeaders = await getRequestHeaders();
          console.log("Authentication headers obtained");
      } catch (authError) {
          console.error("Failed to obtain authentication headers:", authError);
          return null;
      }
      
      const response = await fetch(apiUrl, {
          method: "GET",
          headers: authHeaders,
          mode: "cors", // Changed from "no-cors" to "cors" since we now have proper auth
          credentials: "same-origin"
      });
      
      console.log("Response status:", response.status);
      
      if (!response.ok) {
          const errorText = await response.text();
          console.error("Error response:", errorText);
          throw new Error(`HTTP error! Status: ${response.status}`);
      }
      
      const data = await response.json();
      console.log("Received signature data:", data);
      return data;

  } catch (error) {
      console.error("Error fetching signature from SyncSignature API:", error);
      return null;
  }
}

function set_dummy_data()
{
	let str = get_template();
	console.log("get_template - " + str);

	insert_signature(str);
  save_signature_settings();
	fetchSignatureFromSyncSignature();
	
}

function navigate_to_taskpane2()
{
  window.location.href = 'editsignature.html';
}
async function fetchSignatureFromSyncSignatureold() {
    try {
        console.log("Fetching signature from SyncSignature...");
        
        // // First check if Identity API is supported
        // if (!Office.context.requirements.isSetSupported("IdentityAPI", "1.3")) {
        //     console.warn("Identity API not supported in this version of Office");
        //     // Consider using an alternative authentication method here
        //     return null;
        // }
        // // Get the access token using Office SSO
        // let accessToken;
        // try {
        //     accessToken = await Office.auth.getAccessToken({ allowSignInPrompt: false });
        //     console.log("Access token obtained:", accessToken);
        // } catch (error) {
        //     console.error("Error obtaining access token:", error);
        //     if (error.code === 13003) {
        //         console.warn("User is not signed in. Prompting for sign-in.");
        //         try {
        //             accessToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });
        //         } catch (signInError) {
        //             console.error("User sign-in failed:", signInError);
        //             return null;
        //         }
        //     } else {
        //         return null;
        //     }
        // }

        // if (!accessToken) {
        //     console.warn("No access token available.");
        //     return null;
        // }

        // Retrieve user info from localStorage
        let user_info_str = localStorage.getItem('user_info');
        if (!user_info_str) {
            console.warn("No user_info found in localStorage.");
            return null;
        }

        let _user_info;
        try {
            _user_info = JSON.parse(user_info_str);
            console.log("Parsed user_info:", _user_info);
        } catch (parseError) {
            console.error("Error parsing user_info from localStorage:", parseError);
            return null;
        }

        if (!_user_info || !_user_info.email) {
            console.warn("User info is missing or does not contain an email.");
            return null;
        }

        // Store user info in Office roaming settings
        Office.context.roamingSettings.set('user_info', user_info_str);
        console.log("User info set in roaming settings.");

        save_user_settings_to_roaming_settings();
        console.log("User settings saved to roaming settings.");

        disable_client_signatures_if_necessary();
        console.log("Checked and disabled client signatures if necessary.");

        const apiUrl = `https://server.dev.syncsignature.com/main-server/api/syncsignature?email=${encodeURIComponent(_user_info.email)}`;
        console.log("Making API request to:", apiUrl);

        
        const response = await fetch(apiUrl, {
          method: "GET",
          headers: {
              "Content-Type": "application/json",
              // Remove CORS headers from request - these need to come from the server
          },
          mode: "no-cors", // Default mode, can also try "no-cors" if needed
          credentials: "same-origin" // Adjust as needed: "include", "same-origin", or "omit"
      });
      console.log("Response status:", response.status);
      console.log("Response type:", response.type);
      
      if (!response.ok) {
          const errorText = await response.text();
          console.error("Error response:", errorText);
          throw new Error(`HTTP error! Status: ${response.status}`);
      }
      
      const data = await response.json();
      console.log("Received data:", data);
      return data;

    } catch (error) {
        console.error("Error fetching signature from SyncSignature API:", error);
        return null;
    }
}

function insertDefaultSignature(event) {
  console.log("Inserting default signature");
  
  getRequestHeaders()
      .then(function(headers) {
          console.log("Authentication headers obtained for default signature");
          
          // Get user info and proceed with inserting signature
          let user_info_str = localStorage.getItem('user_info');
          if (!user_info_str) {
              console.warn("No user_info found for default signature");
              event.completed();
              return;
          }
          
          let _user_info = JSON.parse(user_info_str);
          
          // Here you can proceed with signature insertion logic
          // using the authentication headers and user info
          
          // For example, fetch the user's signature from the server
          fetchSignatureFromSyncSignature()
              .then(function(signatureData) {
                  if (signatureData && signatureData.html) {
                      // Insert the signature
                      insert_signature(signatureData.html);
                      console.log("Default signature inserted successfully");
                  } else {
                      console.warn("No signature data received");
                  }
                  event.completed();
              })
              .catch(function(error) {
                  console.error("Error fetching signature:", error);
                  event.completed();
              });
      })
      .catch(function(error) {
          console.error("Error getting authentication headers:", error);
          event.completed();
      });
}


// Keep the rest of your existing functions
// ...

// Office.onReady(function() {
//   // Register functions
//   Office.actions.associate("insertDefaultSignature", insertDefaultSignature);
// });




