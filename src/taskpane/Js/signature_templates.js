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
