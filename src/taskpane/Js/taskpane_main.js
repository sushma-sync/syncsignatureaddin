// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function save_user_settings_to_roaming_settings()
{
  Office.context.roamingSettings.saveAsync(function (asyncResult)
  {
	console.log("save_user_info_str_to_roaming_settings - " + JSON.stringify(asyncResult));
  });
}

function disable_client_signatures_if_necessary()
{
  if ($("#overrideOutlookSignature").prop("checked") === true)
  {
	Office.context.mailbox.item.disableClientSignatureAsync(function (asyncResult)
	{
	  console.log("disable_client_signature_if_necessary - " + JSON.stringify(asyncResult));
	});
  }
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
  
	// Save signature type preferences
	Office.context.roamingSettings.set('newMail', $("#newMailSignature").prop('checked') ? 'syncsignature' : 'none');
	Office.context.roamingSettings.set('reply', $("#replySignature").prop('checked') ? 'syncsignature' : 'none');
	Office.context.roamingSettings.set('forward', $("#forwardSignature").prop('checked') ? 'syncsignature' : 'none');
	Office.context.roamingSettings.set('override_olk_signature', $("#overrideOutlookSignature").prop('checked'));

	save_user_settings_to_roaming_settings();

	disable_client_signatures_if_necessary();

	$("#message").show("slow");
  }
  else
  {
	// TBD display an error somewhere?
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

function load_signature_settings()
{
  // Load saved settings and update checkboxes
  let newMailSetting = Office.context.roamingSettings.get("newMail");
  let replySetting = Office.context.roamingSettings.get("reply");
  let forwardSetting = Office.context.roamingSettings.get("forward");
  let overrideSetting = Office.context.roamingSettings.get("override_olk_signature");

  $("#newMailSignature").prop('checked', newMailSetting === 'syncsignature');
  $("#replySignature").prop('checked', replySetting === 'syncsignature');
  $("#forwardSignature").prop('checked', forwardSetting === 'syncsignature');
  $("#overrideOutlookSignature").prop('checked', overrideSetting === true);
}

async function set_syncsignature()
{
  // Get user email from Outlook API
  let userEmail = Office.context.mailbox ? Office.context.mailbox.userProfile.emailAddress : "Unknown User";
  let user_info = {
      name: userEmail.split("@")[0], 
      email: userEmail
  };
  localStorage.setItem('user_info', JSON.stringify(user_info));
  console.log("User Info:", user_info);
  
  // Load existing settings first
  load_signature_settings();
  
  //let str = get_template_image();
  let signature = await fetchSignatureFromSyncSignature();
  console.log("signature >> ", signature)
  if(signature)
  {
    document.getElementById("dummy_signature").innerHTML = signature;
    // Only save settings, don't insert signature immediately
    save_signature_settings();
    
    // Show success message
    alert("Signature settings saved successfully! The signature will be applied based on your selected email types.");
  }
}

function navigate_to_taskpane2()
{
  window.location.href = 'editsignature.html';
}


async function fetchSignatureFromSyncSignature() {
    try {

        console.log("Fetching signature from SyncSignature...");
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

        const apiUrl = `https://server.syncsignature.com/main-server/api/syncsignature?email=${encodeURIComponent(_user_info.email)}`;
        console.log("Making API request to:", apiUrl);

        const response = await fetch(apiUrl, {
          method: "GET",
          headers: {
              "Content-Type": "application/json",
              "Accept": "application/json"
          },
          mode: "cors", 
        });

        const dummySignatureDiv = document.getElementById("dummy_signature");
        const submitButton = document.getElementById("submit_button");
        
        if (!response.ok) {
            console.log(response.status)
            if (response.status === 404) {
              const errorText = await response.text();
              dummySignatureDiv.innerHTML = "Signature Not Configure at SyncSignature";
              submitButton.disabled = true; // Disable the submit button
              return "Signature Not Configure at SyncSignature";
           }
            
        }
        else{
            console.log("Response status:", response.status);
        }
        console.log(response)
        const data = await response.json();
        console.log("Received data:", data);
        console.log("Received data:", data.html);
        dummySignatureDiv.innerHTML = data.html || "No signature available";
        submitButton.disabled = !data.html;
        return data.html;

    } catch (error) {
        console.error("Error fetching signature from SyncSignature API:", error);
        return null;
    }
}

Office.onReady(function() {
  // Register functions
  Office.actions.associate("insertDefaultSignature", insertDefaultSignature);
  
  // Load signature settings when page loads
  $(document).ready(function() {
    load_signature_settings();
  });
});
function insertDefaultSignature(event) {
  // Get user identity token silently (if already logged in)
  console.log("sushma")
  Office.context.mailbox.getUserIdentityTokenAsync(function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          // Get the token
          var exchangeToken = result.value;
          console.log("Exchange token:", exchangeToken);
          
      } else {
          console.error("Error getting identity token:", result.error);
          event.completed();
          showNotification("Error authenticating. Please try again.");
      }
  });
}




