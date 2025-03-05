// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on Outlook on web, on Windows, and on Mac (new UI preview).

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */
function checkSignature(eventObj) {
  let user_info_str = Office.context.roamingSettings.get("user_info");
  console.log("Autorun process",user_info_str)
  if (!user_info_str) {
    display_insight_infobar();
  } else {
    let user_info = JSON.parse(user_info_str);

    if (Office.context.mailbox.item.getComposeTypeAsync) {
      //Find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
      Office.context.mailbox.item.getComposeTypeAsync(
        {
          asyncContext: {
            user_info: user_info,
            eventObj: eventObj,
          },
        },
        function (asyncResult) {
          if (asyncResult.status === "succeeded") {
            insert_auto_signature(
              asyncResult.value.composeType,
              asyncResult.asyncContext.user_info,
              asyncResult.asyncContext.eventObj
            );
          }
        }
      );
    } else {
      // Appointment item. Just use newMail pattern
      let user_info = JSON.parse(user_info_str);
      insert_auto_signature("newMail", user_info, eventObj);
    }
  }
}

/**
 * For Outlook on Windows and on Mac only. Insert signature into appointment or message.
 * Outlook on Windows and on Mac can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  console.log(template_name)
  let signature_info = get_signature_info(template_name, user_info);
  console.log("Signature Info >>")
  console.log(signature_info)
  addTemplateSignatureNew(signature_info, eventObj);
}

/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */
function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) {
        //After image is attached, insert the signature
        Office.context.mailbox.item.body.setSignatureAsync(
          signatureDetails.signature,
          {
            coercionType: "html",
            asyncContext: eventObj,
          },
          function (asyncResult) {
            asyncResult.asyncContext.completed();
          }
        );
      }
    );
  } else {
    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureDetails.signature,
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

function addTemplateSignatureNew(signatureDetails, eventObj, signatureImageBase64) {
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
}

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: "Please set your signature with the Office Add-ins sample.",
    icon: "Icon.16x16",
    actions: [
      {
        actionType: "showTaskPane",
        actionText: "Set signatures",
        commandId: get_command_id(),
        contextData: "{''}",
      },
    ],
  });
}

/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
  console.log(compose_type)
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */
function get_signature_info(template_name, user_info) {
  // if (template_name === "templateB") return get_template_B_info(user_info);
  // if (template_name === "templateC") return get_template_C_info(user_info);
  // return get_template_A_info(user_info);
  console.log(get_template_image())
  return get_template_image();
}

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

/**
 * Gets HTML string for template A
 * Embeds the signature logo image into the HTML string
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template A,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 */
function get_template_A_info(user_info) {
  const logoFileName = "sample-logo.png";
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Embed the logo using <img src='cid:...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='cid:" +
    logoFileName +
    "' alt='MS Logo' width='24' height='24' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += is_valid_data(user_info.job) ? user_info.job + "<br/>" : "";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  // return object with signature HTML, logo image base64 string, and filename to reference it with.
  return {
    signature: str,
    logoBase64:
      "iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAEeSURBVFhHzdhBEoIwDIVh4EoeQJd6YrceQM+kvo5hQNokLymO/4aF0/ajlBl1fL4bEp0uj3K9XQ/lGi0MEcB3UdD0uVK1EEj7TIuGeBaKYCgIswCLcUMid8mMcUEiCMk71oRYE+Etsd4UD0aFeBBSFtOEMAgpg6lCIggpitlAMggpgllBeiAkFjNDeiIkBlMgeyAkL6Z6WJdlEJJnjvF4vje/BvRALNN23tyRXzVpd22dHSZtLhjMHemB8cxRINZZyGCssbL2vCN7YLwItHo0PTEMAm3OSA8Mi0DVw5rBRBCoCkERTBSBmhDEYDII5PqlZy1iZSGQuiOSZ6JW3rEuCIpgmDFuCGImZuEUBHkWiOweDUHaQhEE+pM/aobhBZaOpYLJeeeoAAAAAElFTkSuQmCC",
    logoFileName: logoFileName,
  };
}

/**
 * Gets HTML string for template B
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template B,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */
function get_template_B_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Reference the logo using a URI to the web server <img src='https://...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Gets HTML string for template C
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template C,
    "logoBase64": null since there is no image,
    "logoFileName": null since there is no image
 */
function get_template_C_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

function get_template_image() {
  return `
    <table width="600px" cellpadding="0" cellspacing="0" border="0" style="font-family: Arial, Helvetica, sans-serif; user-select: none;">
    <tbody>
        <tr>
            <td aria-label="main-content" style="border-collapse: collapse; font-size: inherit; padding-bottom: 2px;">
                <table cellpadding="0" cellspacing="0" border="0">
                    <tbody>
                        <tr>
                            <td>
                                <table cellpadding="0px" cellspacing="0" border="0">
                                    <tbody>
                                        <tr>
                                            <td style="vertical-align: top; padding-right: 14px;">
                                                <table align="center" cellpadding="0" cellspacing="0" style="font-size: inherit;"></table>
                                                <table align="center" cellpadding="0" cellspacing="0" style="font-size: inherit;">
                                                    <tbody>
                                                        <tr>
                                                            <td style="padding-bottom: 8px;"><img src="https://static.sendsig.com/signatures/db7e1098-9a04-4228-846a-5e86048d1ca6/company-logo-1741175789012.png" alt="company-logo" width="80" title="Company logo" style="display: block;" /></td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                            <td style="vertical-align: top;">
                                                <table cellpadding="0px" cellspacing="0" border="0" style="border-collapse: collapse;">
                                                    <tbody>
                                                        <tr>
                                                            <td>
                                                                <table cellpadding="0" cellspacing="0" border="0" style="font-size: inherit; padding-bottom: 6px;">
                                                                    <tbody>
                                                                        <tr>
                                                                            <td id="name-text-id" style="font-weight: 700; padding-bottom: 2px; font-size: 13px; line-height: 1.09; color: rgb(127, 86, 217); font-family: inherit;">Olivia Bolton</td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td id="position-text-id" style="font-size: 12px; line-height: 1.09; color: rgb(0, 0, 0); font-family: inherit; padding-bottom: 2px;">Marketing Manager</td>
                                                                        </tr>
                                                                    </tbody>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="separator-text-id" colspan="2" style="border-top: 1px solid rgb(127, 86, 217); line-height: 0;"></td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <table cellpadding="0" cellspacing="0" border="0" style="font-size: inherit; padding-top: 8px;">
                                                                    <tbody>
                                                                        <tr>
                                                                            <td valign="middle" style="padding: 0px 0px 6px; text-align: left; align-items: center;"><span style="vertical-align: middle; display: inline-block;"><img src="https://static.sendsig.com/icons/db7e1098-9a04-4228-846a-5e86048d1ca6/phone-icon-84fd5c3e-e85a-4413-9134-2a565ee6ae59.png?timestamp=1741175793582" alt="phone" height="16" width="16" /></span> <span id="phone-text-id" class="item-center justify-center" align="top" style="font-size: 12px; color: rgb(0, 0, 0); text-align: left; text-decoration: none; vertical-align: middle;">212-323</span></td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="padding: 0px 0px 6px;"><span style="vertical-align: middle; display: inline-block;"><img src="https://static.sendsig.com/icons/db7e1098-9a04-4228-846a-5e86048d1ca6/email-icon-a088a992-7db1-48db-a1a4-0f0e743b8d2d.png?timestamp=1741175793638" alt="email" height="16" width="16" /></span> <a id="email-text-id" valign="middle" href="https://server.utags.co/mhFwEJQh" target="_blank" rel="noreferrer" style="font-size: 12px; color: rgb(0, 0, 0); text-align: left; text-decoration: none; vertical-align: middle;">dummyemail@dummy.com</a></td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="align-items: center; padding: 0px 0px 6px;"><span style="vertical-align: middle;"><img src="https://static.sendsig.com/icons/db7e1098-9a04-4228-846a-5e86048d1ca6/website-icon-e4dd73ed-6c26-4d04-8a12-418923ce4b8c.png?timestamp=1741175793685" alt="website" height="16" width="16" /></span> <a id="website-text-id" href="https://server.utags.co/tYDUZlfI" style="font-size: 12px; color: rgb(0, 0, 0); text-align: left; text-decoration: none; vertical-align: middle;">yourwebsite.com</a></td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="padding-bottom: 6px;">
                                                                                <table cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-size: inherit;">
                                                                                    <tbody>
                                                                                        <tr>
                                                                                            <td style="padding-right: 8px;"><a href="https://server.utags.co/DMDMyzSd" style="display: inline-block;"><img src="https://static.sendsig.com/icons/db7e1098-9a04-4228-846a-5e86048d1ca6/facebook-icon-77fd0edb-6f79-4fb9-8274-aa6448d0ea9e.png?timestamp=1741175793709" alt="facebook" height="22" width="20" /></a></td>
                                                                                            <td style="padding-right: 8px;"><a href="https://server.utags.co/yTazrdmg" style="display: inline-block;"><img src="https://static.sendsig.com/icons/db7e1098-9a04-4228-846a-5e86048d1ca6/instagram-icon-e6c94c74-33e3-4128-948e-e35c8b8b4db4.png?timestamp=1741175793713" alt="instagram" height="22" width="20" /></a></td>
                                                                                            <td><a href="https://server.utags.co/MbBRZaEL" style="display: inline-block;"><img src="https://static.sendsig.com/icons/db7e1098-9a04-4228-846a-5e86048d1ca6/linkedin-icon-30abf2e5-f527-4a3a-b9dd-d76394745e4a.png?timestamp=1741175793748" alt="linkedin" height="22" width="20" /></a></td>
                                                                                        </tr>
                                                                                    </tbody>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                    </tbody>
                                                                </table>
                                                            </td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </td>
                        </tr>
                    </tbody>
                </table>
            </td>
        </tr>
    </tbody>
</table>
  `;
}

/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */
function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);
