// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function get_template_A_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str +=   "<tr>";
  str +=     "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>";
  str +=     "<td style='padding-left: 5px;'>";
  str +=	   "<strong>" + user_info.name + "</strong>";
  str +=     is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str +=     "<br/>";
  str +=	   is_valid_data(user_info.job) ? user_info.job + "<br/>" : "";
  str +=	   user_info.email + "<br/>";
  str +=	   is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str +=     "</td>";
  str +=   "</tr>";
  str += "</table>";

  return str;
}

function get_template_B_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str +=   "<tr>";
  str +=     "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>";
  str +=     "<td style='padding-left: 5px;'>";
  str +=	   "<strong>" + user_info.name + "</strong>";
  str +=     is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str +=     "<br/>";
  str +=	   user_info.email + "<br/>";
  str +=	   is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str +=     "</td>";
  str +=   "</tr>";
  str += "</table>";

  return str;
}

function get_template_C_str(user_info)
{
  let str = "";
  if (is_valid_data(user_info.greeting))
  {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;
  
  return str;
}


function get_template() {
  return `
    <table width="600px" cellpadding="0" cellspacing="0" border="0" style="font-family: 'Lucida Console', Monaco, monospace; user-select: none;">
      <tbody>
        <tr>
          <td>
            <table cellpadding="0" cellspacing="0" border="0" role="presentation" style="font-size: inherit; padding-bottom: 2px;">
              <tbody>
                <tr>
                  <td>
                    <table cellpadding="0px" style="font-size: inherit;">
                      <tbody>
                        <tr>
                          <td>
                            <table cellpadding="0px" cellspacing="0" border="0" role="presentation" style="font-size: inherit; padding-bottom: 6px;">
                              <tbody>
                                <tr>
                                  <td id="name-text-id" style="font-weight: 700; padding-bottom: 2px; font-size: 14px; line-height: 1.25; color: rgb(0, 188, 111);">
                                    Sushma Reshamwala
                                  </td>
                                </tr>
                                <tr>
                                  <td id="position-text-id" style="font-size: 11px; line-height: 1.25; color: rgb(0, 0, 0); padding-right: 8px; padding-bottom: 2px;">
                                    Software Engineer
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                          </td>
                        </tr>
                        <tr>
                          <td class="separator-text-id" colspan="2" style="border-top: 1px solid rgb(0, 188, 111); padding-bottom: 8px;"></td>
                        </tr>
                        <tr>
                          <td>
                            <table cellpadding="0px" cellspacing="0" border="0" role="presentation" style="font-size: inherit; padding-bottom: 2px;">
                              <tbody>
                                <tr>
                                  <td valign="middle" style="font-family: inherit; padding: 0px 0px 6px; text-align: left; align-items: center;">
                                    <span id="phone-text-id" class="item-center justify-center" align="top" style="font-size: 11px; line-height: 1.25; color: rgb(0, 0, 0); text-align: left; text-decoration: none; vertical-align: middle;">
                                      (345) 087 - 1239
                                    </span>
                                  </td>
                                </tr>
                                <tr>
                                  <td style="font-family: inherit; padding: 0px 0px 6px;">
                                    <a id="email-text-id" valign="middle" href="mailto:j@innovatechlabs.com" target="_blank" rel="noreferrer" style="font-size: 11px; line-height: 1.25; color: rgb(0, 0, 0); text-align: left; text-decoration: none; vertical-align: middle;">
                                      j@innovatechlabs.com
                                    </a>
                                  </td>
                                </tr>
                                <tr>
                                  <td style="font-family: inherit; align-items: center; padding: 0px 0px 6px;">
                                    <a id="website-text-id" href="https://innovatechlabs.com" style="font-size: 11px; line-height: 1.25; color: rgb(0, 0, 0); text-align: left; text-decoration: none; vertical-align: middle;">
                                      innovatechlabs.com
                                    </a>
                                  </td>
                                </tr>
                                <tr>
                                  <td valign="middle" style="text-align: left; align-items: center; padding: 0px 0px 6px; font-family: inherit;">
                                    <span id="address-text-id" class="item-center justify-center" style="font-size: 11px; line-height: 1.25; color: rgb(0, 0, 0); text-align: left; vertical-align: middle;">
                                      984 Penn Rd. NY 102
                                    </span>
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
    <table cellpadding="0" cellspacing="0" border="0" style="font-family: 'Lucida Console', Monaco, monospace; font-size: 11px; user-select: none;">
      <tbody>
        <tr>
          <td aria-label="cta" style="font-family: inherit; padding-top: 8px; padding-bottom: 16px;">
            <a aria-label="link" href="https://server.utags.co/qJulrvGP" target="_blank" rel="noreferrer" style="color: transparent; text-decoration: none; display: inline-block;">
              <span style="font-family: inherit; width: 120px; border-radius: 4px; background-color: transparent; color: rgb(0, 188, 111); padding: 0px; text-decoration: underline; font-weight: 500; justify-content: left; display: flex; align-items: center; font-style: normal; line-height: 16px; cursor: pointer;">
                Schedule a call
              </span>
            </a>
          </td>
        </tr>
      </tbody>
    </table>
  `;
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
