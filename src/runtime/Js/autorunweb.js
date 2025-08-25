// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file contains code only used by autorunweb.html when loaded in Outlook on the web.

Office.initialize = function (reason) {};

/**
 * For Outlook on the web, insert signature into appointment or message.
 * Outlook on the web does not support using setSignatureAsync on appointments,
 * so this method will update the body directly.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
async function insert_auto_signature(compose_type, user_info, eventObj) {
  console.log("insert auto signature (web)")

  // Check if we should override Outlook signatures (same logic as desktop)
  let shouldOverride = false;
  let stored_preferences = localStorage.getItem('signature_preferences');
  
  if (stored_preferences) {
    try {
      let preferences = JSON.parse(stored_preferences);
      shouldOverride = preferences.override_olk_signature || false;
      console.log("Override Outlook signature setting (web):", shouldOverride);
    } catch (e) {
      console.error("Error parsing preferences for override setting:", e);
      shouldOverride = Office.context.roamingSettings.get("override_olk_signature") || false;
    }
  } else {
    shouldOverride = Office.context.roamingSettings.get("override_olk_signature") || false;
  }

  let template_name = get_template_name(compose_type);
  console.log("Template name (web):", template_name)
  let signatureDetails = await get_signature_info(template_name, user_info);
  console.log("Signature Info (web) >>", signatureDetails)

  // Only proceed if template is not 'none'
  if (template_name !== 'none' && signatureDetails) {
    if (Office.context.mailbox.item.itemType == "appointment") {
      set_bodynew(signatureDetails, eventObj);
    } else {
      addTemplateSignatureNew(signatureDetails, eventObj);
    }
  } else {
    console.log("No signature to insert for compose type:", compose_type);
    eventObj.completed();
  }
}

/**
 * For Outlook on the web, set signature for current appointment
* @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj Office event object
 */
function set_body(signatureDetails, eventObj) {

  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) { 
        Office.context.mailbox.item.body.setAsync(
        "<br/><br/>" + signatureDetails.signature,
        {
          coercionType: "html",
          asyncContext: eventObj,
        },
        function (asyncResult) {

          asyncResult.asyncContext.completed();
        }
      );
    });
  } else {
    Office.context.mailbox.item.body.setAsync(
      "<br/><br/>" + signatureDetails.signature,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {

        asyncResult.asyncContext.completed();
      }
    );
  }
  
}

function set_bodynew(signatureDetails, eventObj) {
 
   console.log("set_bodynew - " + JSON.stringify(signatureDetails));
    Office.context.mailbox.item.body.setAsync(
      "<br/><br/>" + signatureDetails,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {

        asyncResult.asyncContext.completed();
      }
    );
  
}

/**
 * Gets template name mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns 'syncsignature' if enabled for that type, 'none' otherwise
 */
function get_template_name(compose_type) {
  console.log("Compose type (web):", compose_type)
  let templateName = 'none';
  
  // First, try to get preferences from localStorage
  let stored_preferences = localStorage.getItem('signature_preferences');
  let preferences = null;
  
  if (stored_preferences) {
    try {
      preferences = JSON.parse(stored_preferences);
      console.log("Using localStorage preferences (web):", preferences);
      
      // Use localStorage preferences
      if (compose_type === "reply") {
        templateName = preferences.reply ? 'syncsignature' : 'none';
      } else if (compose_type === "forward") {
        templateName = preferences.forward ? 'syncsignature' : 'none';
      } else {
        // For newMail and other types
        templateName = preferences.newMail ? 'syncsignature' : 'none';
      }
      
    } catch (e) {
      console.error("Error parsing stored preferences (web):", e);
      preferences = null;
    }
  }
  
  // Fallback to roaming settings if no localStorage preferences
  if (!preferences) {
    console.log("Using roaming settings fallback (web)");
    if (compose_type === "reply") {
      templateName = Office.context.roamingSettings.get("reply") || 'none';
    } else if (compose_type === "forward") {
      templateName = Office.context.roamingSettings.get("forward") || 'none';
    } else {
      // For newMail, default to syncsignature if not explicitly set
      let newMailSetting = Office.context.roamingSettings.get("newMail");
      templateName = (newMailSetting === 'none') ? 'none' : 'syncsignature';
    }
  }
  
  console.log("Template name for", compose_type, ":", templateName);
  return templateName;
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {*} template_name Which template format to use ('syncsignature' or 'none')
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format, or null if disabled
 */
async function get_signature_info(template_name, user_info) {
  // Return null if signature is disabled for this compose type
  if (template_name === 'none') {
    console.log("Signature disabled for this compose type (web)");
    return null;
  }
  
  // Fetch signature from SyncSignature API
  let signature = await fetchSignatureFromSyncSignature(user_info)
  console.log("Fetched signature (web):", signature)
  return signature;
}

async function fetchSignatureFromSyncSignature(user_info) {
  try {
    
      console.log("Fetching signature from SyncSignature (web)...");
      let user_info_str = user_info;
      if (!user_info_str) {
          console.warn("No user_info found (web)");
          return null;
      }
      const apiUrl = `https://server.syncsignature.com/main-server/api/syncsignature?email=${encodeURIComponent(user_info_str.email)}`;
      console.log("Making API request to (web):", apiUrl);

      const response = await fetch(apiUrl, {
        method: "GET",
        headers: {
            "Content-Type": "application/json",
            "Accept": "application/json"
        },
        mode: "cors", 
      });

      if (!response.ok) {
          console.log("Web response:", response)
          const errorText = await response.text();
          console.error("Error response (web):", errorText);
          throw new Error(`HTTP error! Status: ${response.status}`);
      }
      else{
          console.log("Response status (web):", response.status);
      }
      const data = await response.json();
      console.log("Received data (web):", data);
      console.log("Received data (web):", data.html);
      return data.html;

  } catch (error) {
      console.error("Error fetching signature from SyncSignature API (web):", error);
      return null;
  }
}

/**
 * For web version, add signature using setAsync instead of setSignatureAsync for appointments
 */
function addTemplateSignatureNew(signatureDetails, eventObj) {
  console.log("addTemplateSignatureNew function (web) >>")
  
  if (Office.context.mailbox.item.itemType == "appointment") {
    // For appointments on web, use setAsync
    Office.context.mailbox.item.body.setAsync(
      "<br/><br/>" + signatureDetails,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  } else {
    // For emails, use setSignatureAsync if available, otherwise setAsync
    if (Office.context.mailbox.item.body.setSignatureAsync) {
      Office.context.mailbox.item.body.setSignatureAsync(
        signatureDetails,
        {
          coercionType: "html",
          asyncContext: eventObj,
        },
        function (asyncResult) {
          asyncResult.asyncContext.completed();
        }
      );
    } else {
      // Fallback to setAsync for older versions
      Office.context.mailbox.item.body.setAsync(
        "<br/><br/>" + signatureDetails,
        {
          coercionType: "html",
          asyncContext: eventObj,
        },
        function (asyncResult) {
          asyncResult.asyncContext.completed();
        }
      );
    }
  }
}
