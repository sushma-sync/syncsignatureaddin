/* eslint semi: 2 */
/* global Office, OfficeRuntime */
import { createNestablePublicClientApplication } from '@azure/msal-browser';

var signatureHtmlSettingName = 'signature_html';
var appEnv = process.env.ENV;
var apiHost = process.env.initializePCA;
var logLines = [];
var version = '3.8.0';
var requestId = uuidv4();
var debug = false;
var fromEmail = '';
var composeType = 'newMail';
var isInternal = false;
var currentEvent = '';
var alreadyInsertedFromCache = false;

let pca = undefined;
let isPCAInitialized = false;

function initializePCA() {
  return new Office.Promise(function (resolve, reject) {
    if (isPCAInitialized) {
      resolve();
    }
    // Initialize the public client application.
    createNestablePublicClientApplication({
      auth: {
        clientId:  process.env.MICROSOFT_CLIENT_ID,
        authority: 'https://login.microsoftonline.com/common',
      },
    }).then(function(localPca) {
      pca = localPca;
      isPCAInitialized = true;
      resolve();
    }).catch(function(error) {
      log('Error creating pca', error);
      reject();
    });
  });
}

/////////////////////////////////////////////////////////
// Utilities
/////////////////////////////////////////////////////////
function uuidv4() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
    var r = (Math.random() * 16) | 0,
      v = c == 'x' ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

if (!('toJSON' in Error.prototype)) {
  // eslint-disable-next-line no-extend-native
  Object.defineProperty(Error.prototype, 'toJSON', {
    value: function () {
      var alt = {};

      Object.getOwnPropertyNames(this).forEach(function (key) {
        alt[key] = this[key];
      }, this);

      return alt;
    },
    configurable: true,
    writable:     true,
  });
}

function getClient() {
  if (Office.context.diagnostics.platform === Office.PlatformType.Mac) {
    return 'outlook_mac';
  } else if (Office.context.diagnostics.platform === Office.PlatformType.OfficeOnline) {
    if (Office.context.mailbox.diagnostics.hostName === 'newOutlookWindows') {
      return 'outlook_windows';
    } else {
      return 'outlook_web';
    }
  } else if (Office.context.diagnostics.platform === Office.PlatformType.PC) {
    return 'outlook_windows';
  } else if (Office.context.diagnostics.platform === Office.PlatformType.Android) {
    return 'outlook_android';
  } else if (Office.context.diagnostics.platform === Office.PlatformType.iOS) {
    return 'outlook_ios';
  }
  return null;
}

function getMetaData() {
  return {
    version: Office.context.diagnostics.version,
  };
}

function log(message, error) {
  if (appEnv === 'development' && debug) {
    var item = Office.context.mailbox.item;
    item.body.setSelectedDataAsync(
      '<p>Message:' + message + ' Error:' + error + '</p>',
      { coercionType: Office.CoercionType.Html, asyncContext: { var3: 1, var4: 2 } },
      function (asyncResult) {}
    );
    console.log(message, error);
  }
  logLines.push({
    timestamp: Date.now().toString(),
    line:      message + (error ? ' ' + JSON.stringify(error) : ''),
    level:     error ? 'ERROR' : 'INFO',
  });
}

function sendLogsAndComplete(event) {
  post('/log?email=' + Office.context.mailbox.userProfile.emailAddress.toLowerCase(), {
    lines:    JSON.stringify(logLines),
    now:      Date.now().toString(),
    metadata: {
      env:            appEnv,
      email:          Office.context.mailbox.userProfile.emailAddress.toLowerCase(),
      from_email:     fromEmail,
      is_internal:    isInternal,
      compose_type:   composeType,
      office_client:  Office.context.mailbox.diagnostics.hostName === 'newOutlookWindows' ? 'new_outlook_windows' : (Office.context.diagnostics.platform === Office.PlatformType.PC ? 'old_outlook_windows' : getClient()),
      office_version: Office.context.diagnostics.version,
      version:        version,
      request_id:     requestId,
      event:          currentEvent,
    },
  })
    .then(function () {
      event.completed();
    })
    .catch(function () {
      event.completed();
    });
}

function isInternalEmail() {
  return new Office.Promise(function (resolve) {
    Office.context.mailbox.item.to.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        var allRecipients = asyncResult.value;
        if (fromEmail === '' || allRecipients.length === 0) {
          resolve(false);
        } else {
          var internalEmails = allRecipients.filter((recipient) => recipient.emailAddress.toLowerCase().includes(fromEmail.split('@')[1]));
          if (internalEmails.length === allRecipients.length) {
            resolve(true);
          } else {
            resolve(false);
          }
        }
      } else {
        log('Error getting recipients', asyncResult.error);
        resolve(false);
      }
    });
  });
}

/////////////////////////////////////////////////////////
// Main methods
/////////////////////////////////////////////////////////
function getSignatureHtmlSettingName() {
  return signatureHtmlSettingName + '_' + composeType + '_' + fromEmail + '_' + isInternal ? 'internal' : 'external';
}

function setSignatureHtmlInCache(html) {
  return new Office.Promise(function (resolve, reject) {
    Office.context.roamingSettings.remove(signatureHtmlSettingName);
    Office.context.roamingSettings.remove(signatureHtmlSettingName + '_' + fromEmail);
    Office.context.roamingSettings.remove(signatureHtmlSettingName + '_' + composeType + '_' + fromEmail);
    Office.context.roamingSettings.remove(signatureHtmlSettingName + '_' + Office.context.mailbox.userProfile.emailAddress.toLowerCase());

    Office.context.roamingSettings.set(getSignatureHtmlSettingName(), html);
    Office.context.roamingSettings.saveAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        log('Signature not set in cache', result);
        reject();
      } else {
        log('Signature set in cache');
        resolve();
      }
    });
  });
}

function getSignatureHtmlFromCache() {
  return Office.context.roamingSettings.get(getSignatureHtmlSettingName());
}

function headers(token) {
  return {
    Accept:           'application/json',
    'Content-Type':   'application/json',
    'Access-Token':  token,
    'X-Api-Version':  4,
    'X-Client-Type':  'OutlookAddinEvent',
  };
}

function getRequestHeaders() {
  const tokenRequest = {
    scopes: ['User.Read', 'openid', 'profile', 'email'],
  };
  if (Office.context.diagnostics.platform === Office.PlatformType.Android || Office.context.diagnostics.platform === Office.PlatformType.iOS) {
    return new Office.Promise(function (resolve, reject) {
      initializePCA().then(function() {
        pca.acquireTokenSilent(tokenRequest).then(function(userAccount) {
          log('Token acquired silently');
          resolve(headers(userAccount.idToken));
        }).catch(function(error) {
          log('Error acquiring token silently', error);
          reject(error);
        });
      });
    });
  } else {
    return new Office.Promise(function (resolve, reject) {
      OfficeRuntime.auth
        .getAccessToken()
        .then(function (token) {
          log('Token acquired from OfficeRuntime');
          resolve(headers(token));
        }).catch(function (error) {
          log('Error acquiring token from OfficeRuntime', error);
          initializePCA().then(function() {
            pca.acquireTokenSilent(tokenRequest).then(function(userAccount) {
              log('Token acquired silently');
              resolve(headers(userAccount.idToken));
            }).catch(function(error) {
              log('Error acquiring token silently', error);
              reject(error);
            });
          });
        });
    });
  }
}
    
  


function get(path) {
  return new Office.Promise(function (resolve, reject) {
    getRequestHeaders()
      .then(function (headers) {
        resolve(
          fetch(apiHost + path, {
            headers: headers,
          })
        );
      })
      .catch(function (error) {
        reject(error);
      });
  });
}

function post(path, body) {
  return new Office.Promise(function (resolve, reject) {
    getRequestHeaders()
      .then(function (headers) {
        resolve(
          fetch(apiHost + path, {
            method:  'POST',
            headers: headers,
            body:    JSON.stringify(body),
          })
        );
      })
      .catch(function (error) {
        reject(error);
      });
  });
}

function createSignatureInstallation() {
  return post(
    '/email_signature_installation?email=' +
      Office.context.mailbox.userProfile.emailAddress.toLowerCase() +
      '&from_email=' +
      fromEmail + '&' +
      'is_internal=' +
      isInternal,
    {
      installation: {
        email_client: getClient(),
        metadata:     getMetaData(),
      },
    }
  );
}

function insertSignatureFromCacheAsync(event) {
  if (alreadyInsertedFromCache) {
    log('Signature already inserted from cache but failed');
    sendLogsAndComplete(event);
  } else {
    alreadyInsertedFromCache = true;
    var signatureHtml = getSignatureHtmlFromCache();
    if (!signatureHtml && signatureHtml !== '') {
      log('Signature not set - Missing from cache');
      sendLogsAndComplete(event);
    } else {
      Office.context.mailbox.item.body.setSignatureAsync(
        signatureHtml,
        {
          coercionType: Office.CoercionType.Html,
          asyncContext: event,
        },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            log('Signature set from cache');
          } else {
            log('Signature not set from cache', asyncResult.error);
          }
          sendLogsAndComplete(asyncResult.asyncContext);
        }
      );
    }
  }
}

function getAndInstallSignature(event) {
  get(
    '/email_signature_raw_html?email=' +
      Office.context.mailbox.userProfile.emailAddress.toLowerCase() +
      '&from_email=' +
      fromEmail +
      '&compose_type=' +
      composeType + '&' +
      'is_internal=' +
      isInternal
  )
    .then(function (signatureResponse) {
      if (!signatureResponse.ok) {
        log('Get signature error: ' + signatureResponse.status);
        if (signatureResponse.status != 401 && signatureResponse.status != 403) {
          log('Inserting from cache for error status: ' + signatureResponse.status);
          insertSignatureFromCacheAsync(event);
        } else {
          sendLogsAndComplete(event);
        }
      } else {
        signatureResponse.json().then(function (signatureResponseJson) {
          if (signatureResponseJson.images) {
            signatureResponseJson.images.forEach((image) => {
              Office.context.mailbox.item.addFileAttachmentAsync(
                image.url,
                image.name,
                {
                  isInline: true,
                }
              );
            });
          } else {
            setSignatureHtmlInCache(signatureResponseJson.raw_html);
          }
          Office.context.mailbox.item.body.setSignatureAsync(
            signatureResponseJson.raw_html,
            {
              coercionType: Office.CoercionType.Html,
              asyncContext: event,
            },
            function (asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                log('Signature inserted successfully');
                createSignatureInstallation()
                  .then(function () {
                    log('Signature installation created successfully');
                    sendLogsAndComplete(asyncResult.asyncContext);
                  })
                  .catch(function (error) {
                    log('Signature installation not created', error);
                    sendLogsAndComplete(asyncResult.asyncContext);
                  });
              } else {
                log('Signature not inserted', asyncResult.error);
                sendLogsAndComplete(asyncResult.asyncContext);
              }
            }
          );
        });
      }
    })
    .catch(function (error) {
      log('Other error within get signatures', error);
      insertSignatureFromCacheAsync(event);
    });
}

function insertSignature(event) {
  log('Started ' + currentEvent);
  try {

    Office.context.mailbox.item.getComposeTypeAsync(function (getComposeTypeAsyncResult) {
      if (getComposeTypeAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        composeType = getComposeTypeAsyncResult.value.composeType;
      } else {
        log('Error getting compose type', getComposeTypeAsyncResult);
      }
      Office.context.mailbox.item.from.getAsync({ asyncContext: event }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          fromEmail = result.value.emailAddress.toLowerCase();
        } else {
          log('Error getting from email', result);
          fromEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
        }
        isInternalEmail().then((isInternalResult) => {
          isInternal = isInternalResult;
          getAndInstallSignature(event);
        });
      });
    });
  } catch (error) {
    log('Nothing happened', error);
    insertSignatureFromCacheAsync(event);
  }
}

// To remove when everybody updated
function onMessageComposeHandler(event) {
  currentEvent = 'onMessageCompose';
  insertSignature(event);
}

function onNewMessageComposeHandler(event) {
  currentEvent = 'onNewMessageCompose';
  insertSignature(event);
}

function onMessageFromChangedHandler(event) {
  currentEvent = 'onMessageFromChanged';
  insertSignature(event);
}

function onMessageRecipientsChangedHandler(event) {
  currentEvent = 'onMessageRecipientsChanged';
  insertSignature(event);
}

Office.onReady().then(function () {
  Office.actions.associate('onMessageComposeHandler', onMessageComposeHandler); // To remove when everybody updated
  Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
  Office.actions.associate('onMessageFromChangedHandler', onMessageFromChangedHandler);
  Office.actions.associate('onMessageRecipientsChangedHandler', onMessageRecipientsChangedHandler);
});

Office.actions.associate('onMessageComposeHandler', onMessageComposeHandler); // To remove when everybody updated
Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
Office.actions.associate('onMessageFromChangedHandler', onMessageFromChangedHandler);
Office.actions.associate('onMessageRecipientsChangedHandler', onMessageRecipientsChangedHandler);
