/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/



// ==========================
//   Event Based Activation
// ==========================
// - Event Based Activation is NOT supported for addins downloaded from the office store
// - Event Based Activation will disable your active client signature. a consequence of this is that the store addin and the COM-Addin will be in conflict
//   because the COM-Addin uses the active client signature
// - Event Based Activation does not use a page, that means the document object, and its properties are undefined
//   - This means localStorage, location, etc., are all not available
//   - This means everything, including Office.initialize, can not be in a document ready method
// - To concatenate numbers with strings, you must cast the number to a string with .toString()
// - This script contains libraries, which are modified and/or stripped and simplified.
//
// ==========================
//   Documentation and info
// ==========================
// - Documentation             https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch
// - Behaviour and Limitations https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch#event-based-activation-behavior-and-limitations
//
// ====================
//   Notes and Todo's
// ====================
// TODO: clear storage option, so you can manually delete(/update) the signature, before the expiry date. is in the /Addin/Outlook/Index.cshtml
// NOTE: Office Storage seems to be available only in Outlook Desktop? it should be available in Event-Based Activation, find out why it doesnt exist.

// TODO: it seems OfficeRuntime.auth.getAccessToken is now supported in event based activation. we may be able to sign in before even opening the addin.
// https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/use-sso-in-event-based-activation

// ========================
//   Initialize Libraries
// ========================
var eformity = {
    common: {},
    http: {},
    ui: {},
    localstorage: {},
    localization: {},
    fallbacks: {},
    host: { _messages: {} },
    hosts: { officeAddin: { outlook: {} } },
    hostinfo: { support: {} },
    office: { outlook: {} }
};


// ==================
//   Office Helpers
// ==================
function _getGlobal() {
    return typeof self !== 'undefined' ? self : typeof window !== 'undefined' ? window : typeof global !== 'undefined' ? global : null;
}

function _bindAction(action, handler) {
    if (Office.actions && Office.actions.associate) {
        Office.actions.associate(action, handler);
    } else {
        var g = _getGlobal();
        if (g) {
            g[action] = handler;
        }
    }
}

function _invoke(callback) {
    var args = Array.prototype.slice.call(arguments, 1);
    callback && callback.apply(this, args);
}


// =====================
//   Initialize Office
// =====================
Office.initialize = function () {
    // BUG: initialize is too late for Desktop. just do it inside the bindAction
    //eformity.hostinfo.initialize();
    
};

// bind an event on message compose, which will insert the users default signature
_bindAction('onMessageComposeHandler', function (eventObj) {
    eformity.hostinfo.initialize();
      // Add the created signature to the message.
      const signature = "testestestesteste";
      item.body.setSignatureAsync(signature, { coercionType: Office.CoercionType.Html }, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log(result.error.message);
          event.completed();
          return;
        }
});

function onNewMessageComposeHandler(event) {
    eformity.office.outlook.insertDefaultSignature(function (result) {
        result.completed();
    }, eventObj);
}


// ==============================
//   eformity.office.outlook.js
// ==============================
(function () {

    // ===========================
    //   Outlook Library Methods
    // ===========================
    this.setSubject = function (subject, callback, asyncContext) {
        try {
            eformity.hosts.officeAddin.outlook.setSubject(subject, function (result, asyncContext) {
                if (!result) {
                    throw eformity.localization.get('office.setSubjectFailed');
                }

                if (callback) {
                    callback(asyncContext);
                }
            }, asyncContext);
        } catch (e) {
            _handleError(e.message, e);
        }
    };

    this.insertDefaultSignature = function (callback, asyncContext) {
        try {
            // first we need to ensure the client signature is disabled
            eformity.hosts.officeAddin.outlook.disableClientSignature(function (asyncContext) {
                // then we must get the compose type
                eformity.hosts.officeAddin.outlook.getComposeType(function (composeType) {
                    // 'forward' signatures dont exist now, so we should return the 'new' signature
                    if (composeType == ComposeType.forward) {
                        composeType = ComposeType.new;
                    }

                    // then we get the signature, either from (local)storage or from the server
                    _getSignature(composeType, function (signature) {
                        _insertSignature(signature, callback, asyncContext);
                    });
                }, asyncContext);
            }, asyncContext);
        } catch (e) {
            _handleError(e.message, e);
        }
    };

    this.showNotification = function (message, callback, options) {
        _showNotification(message, callback, options);
    };

    // =====================
    //   Signature Helpers
    // =====================
    function _getSignature(composeType, callback) {
        var storageKey = ComposeType.getStorageKey(composeType);

        // check if the storage is flagged to be cleared, which is done after e.g. reloading the list of signatures in the addin
        _validateSignaturesStorage(function (clearedStorage) {
            // try to retrieve the signature from storage
            _getSignatureFromStorage(storageKey, function (storedSignature) {
                // TEST: storage can be turned of for testing by setting useStoredSignature to false
                //var useStoredSignature = false;
                var useStoredSignature = _isValidSignature(storedSignature) && !clearedStorage;

                if (useStoredSignature) {
                    callback(storedSignature);
                } else {
                    _getSignatureByUrl(composeType, function (response) {
                        _handleSignatureResponse(response, storageKey, callback, composeType);
                    }, function () {
                        // when the GET failed, insert the expired stored signature if valid, or show the  'no signature open taskpane' message
                        if (useStoredSignature) {
                            callback(storedSignature);
                        } else {
                            _showNoSignatureNotification();
                        }
                    });
                }
            });
        });
    }

    function _handleSignatureResponse(response, storageKey, callback, composeType) {
        if (_isValidSignature(response)) {
            // store the new signature if its found and valid
            _setStoredSignature(storageKey, response, function () {
                callback(response);
            });
        } else {
            // or, when we didnt retrieve a 'reply' signature, try getting the 'new' signature instead
            if (composeType == ComposeType.reply) {
                _getSignature(ComposeType.new, callback);
            } else {
                // otherwise invoke callback with null when there is no signature retrieved
                callback(null);
            }
        }
    }

    function _validateSignaturesStorage(callback) {
        // INFO: EF2035, EF2038: When switching the language in the addin, or selecting a different employee record, the stored signatures deprecate.
        // To reload the stored signatures, we set a flag in the Add-In, when changing the document language or employee record for example,
        // Then we check if that flag is set, we will clear the signatures from the storage.
        eformity.hosts.officeAddin.outlook.getRoamingSetting(StorageKey.clearStorageFlag, function (clearStorage) {
            if (clearStorage) {
                var keys = [
                    StorageKey.newMailStorageKey,
                    StorageKey.replyMailStorageKey
                ];

                eformity.host.removeStorageItems(keys, function () {
                    eformity.hosts.officeAddin.outlook.removeRoamingSetting(StorageKey.clearStorageFlag, function () {
                        callback(clearStorage);
                    });
                });
            } else {
                callback(false);
            }
        });
    }

    function _getSignatureByUrl(composeType, callback, fallback) {
        _loginOrGetToken(function (actionCredentialToken) {
            actionCredentialToken = actionCredentialToken || '';

            var useUrl = !eformity.hostinfo.support.addFileAttachmentFromBase64Async;
            var attachmentType = useUrl ? AttachmentType.url : AttachmentType.cid;

            // retrieve the owners' email of the mailbox
            var mailbox = Office.context.mailbox.userProfile.emailAddress;

            // the new token format is {username}:{token}
            var items = actionCredentialToken.split(':');
            var username = (items[1] && encodeURIComponent(items[0])) || '';
            var token = encodeURIComponent(items[1] || items[0]) || '';

            _get('/addin/outlook/default?attachmentType={0}&composeType={1}&actionCredentialToken={2}&mailbox={3}&username={4}'.format(attachmentType, composeType, token, mailbox, username), callback, {
                error: function (xhr, status, error) {
                    var message = eformity.localization.get('office.insertSignatureHttpError');
                    _showError(message, null, true);
                }
            });
        }, fallback);
    }

    function _trySilentLogin(callback, errorHandler) {
        // We should skip the login if GET:/status returns a 200, because we are signed in already in another tab or from earlier,
        // The reason we dont want to sign in again, is because you might be signed in with an eformity account, and not an MSAD account,
        // and it might fail to sign in with MSAD when you dont use the Microsoft AD provider

        //_showSuccess('_trySilentLogin', null, true);

        // TODO: for now, we dont support getAccessToken on PC, as it does not seem to work.
        // https://github.com/OfficeDev/office-js/issues/3706
        var diagnostics = Office.context.diagnostics;
        if (diagnostics.platform !== 'OfficeOnline') {
            errorHandler();
            return;
        }

        _get('/Status', function (status) {

            //_showSuccess('_trySilentLogin.Status: ' + (status && status.isAuthenticated), null, true);

            if (status && status.isAuthenticated) {
                // We are signed in, so no need for the actionCredentialToken, so we return '' to let the normal flow continue
                callback('');
            } else {
                // try to login silently, so we can GET the signature from our server
                _getAccessToken(function (token) {
                    _get('/MicrosoftOAuth/SigninOnBehalfOf', function (result) {
                        // if silent login was successful, there is no need for the actionCredentialToken, so we return '' to let the normal flow continue
                        if (result) {
                            callback('');
                        } else {
                            errorHandler();
                        }
                    }, { authorization: 'Bearer ' + token, error: errorHandler });
                }, errorHandler);
            }
        });
    }

    function _getAccessToken(callback, errorHandler) {
        OfficeRuntime.auth.getAccessToken({ forMSGraphAccess: true }).then(callback).catch(errorHandler);

        //OfficeRuntime.auth.getAccessToken({ forMSGraphAccess: true }).then(callback, errorHandler);

        //OfficeRuntime.auth.getAccessToken().then(callback).catch(errorHandler);

        //_showSuccess('getAccessToken', null, true);

        //OfficeRuntime.auth.getAccessToken({ forMSGraphAccess: true }).then(function (token) {
        //    _showSuccess('getAccessToken.then: ' + token, null, true);

        //    callback(token);
        //}, function (e1) {
        //    _showSuccess('getAccessToken.catch:1', null, true);

        //    errorHandler(e1);
        //}).catch(function (e2) {
        //    _showSuccess('getAccessToken.catch:2', null, true);

        //    errorHandler(e2);
        //});
    }

    function _loginOrGetToken(callback, errorHandler) {
        _trySilentLogin(callback, function () {
            eformity.hosts.officeAddin.outlook.getRoamingSetting(StorageKey.actionCredentialToken, function (actionCredentialToken) {
                if (actionCredentialToken) {
                    callback(actionCredentialToken);
                } else {
                    errorHandler();
                }
            });
        })
    }

    function _insertSignature(signature, callback, asyncContext) {
        // inserts the signature, if it has content
        if (_isValidSignature(signature)) {
            eformity.hosts.officeAddin.outlook.insertSignature(signature, function (result, asyncContext) {
                if (result) {
                    var message = eformity.localization.get('office.insertSignatureSuccess');
                    _showSuccess(message, null, true);
                }

                if (callback) {
                    callback(asyncContext);
                }
            }, asyncContext);
        } else {
            _showInvalidSignatureNotification();
        }
    }

    function _getSignatureFromStorage(key, callback) {
        eformity.host.getStorageItem(key, function (signature) {
            if (_isValidSignature(signature)) {
                callback(signature);
            } else {
                // remove the signature from storage if its not valid
                if (signature) {
                    eformity.host.removeStorageItem(key);
                }

                callback(null);
            }
        });
    }

    function _setStoredSignature(key, signature, callback) {
        if (_isValidSignature(signature)) {
            eformity.host.setStorageItem(key, signature, callback, eformity.fallbacks.daysToMiliseconds(1), true);
        } else {
            callback();
        }
    }

    function _isValidSignature(signature) {
        if (signature && signature.hasOwnProperty('content')) {
            return true;
        }

        return false;
    }

    // ========================
    //   Notification Helpers
    // ========================
    function _showSuccess(message, callback, showTaskPaneLink) {
        _showNotification(message, callback, {
            showTaskPane: showTaskPaneLink
        });
    }

    function _showError(message, callback, showTaskPaneLink) {
        _showNotification(message, callback, {
            showTaskPane: showTaskPaneLink,
            type: 'errorMessage'
        });
    }

    function _showNotification(message, callback, options) {
        try {
            eformity.hosts.officeAddin.outlook.showNotification(message, callback, options);
        } catch (e) {
            _handleError(e.message, e);
        }
    }

    function _showNoSignatureNotification() {
        var message = eformity.localization.get('office.noDefaultSignature');
        _showSuccess(message, null, true);
    }

    function _showInvalidSignatureNotification() {
        var message = eformity.localization.get('office.invalidDefaultSignature');
        _showSuccess(message, null, true);
    }

    // ===================
    //   Request Helpers
    // ===================
    function _get(url, callback, options) {
        eformity.http.get(url, callback, _getOptions(options));
    }

    function _post(url, callback, data, options) {
        eformity.http.post(url, callback, data, _getOptions(options));
    }

    function _getOptions(options) {
        return eformity.fallbacks.extend({
            responseType: 'json'
        }, options);
    }

    // ===================
    //   Library Helpers
    // ===================
    function _handleError(message, data) {
        if (!message) {
            message = message || eformity.localization.get('office.unknownError');
        }

        if (data) {
            console.error('[eformity.office.outlook] ' + message, data);
        } else {
            console.error('[eformity.office.outlook] ' + message);
        }
    }

}).call(eformity.office.outlook);


// ============================
//   eformity.localstorage.js
// ============================
(function () {

    this.get = function (key) {
        var expiryMessage = '';

        var schedule = _deserialize(localStorage.getItem(key + '_Expires'));
        if (schedule) {
            if (_isExpired(schedule)) {
                console.log('[outlook:eformity.localstorage] the item {0} is expired'.format(key));

                this.remove(key);

                return null;
            } else {
                expiryMessage = ' (expires on: {0})'.format(new Date(schedule.expires).datetimeNow());
            }
        }

        var result = _deserialize(localStorage.getItem(key));

        console.log('[outlook:eformity.localstorage] retrieved item: {0}{1}'.format(key, expiryMessage), result);

        return result;
    };

    this.set = function (key, value, expires) {
        var expiryMessage = '';

        if (expires) {
            var now = Date.now();
            var schedule = {
                created: now,
                expires: now + Math.abs(expires)
            };

            localStorage.setItem(key + '_Expires', _serialize(schedule));

            expiryMessage = ' (expires on: {0})'.format(new Date(schedule.expires).datetimeNow());
        }

        localStorage.setItem(key, _serialize(value));

        console.log('[outlook:eformity.localstorage] stored item: {0}{1}'.format(key, expiryMessage)); //, value);
    };

    this.remove = function (key) {
        localStorage.removeItem(key + '_Expires');
        localStorage.removeItem(key);

        console.log('[outlook:eformity.localstorage] removed item: {0}'.format(key));
    };

    this.getSubItem = function (key, subKey) {
        var item = eformity.localstorage.get(key);
        if (!item) {
            return null;
        }

        return item[subKey];
    };

    this.setSubItem = function (key, subKey, value) {
        var item = eformity.localstorage.get(key) || {};

        item[subKey] = value;

        eformity.localstorage.set(key, item);
    };

    this.removeSubItem = function (key, subKey) {
        var item = eformity.localstorage.get(key);
        if (!item) {
            return null;
        }

        delete item[subKey];

        eformity.localstorage.set(key, item);
    };

    this.clear = function () {
        localStorage.clear();

        console.log('[outlook:eformity.localstorage] cleared localstorage');
    };

    function _serialize(value) {
        try {
            return JSON.stringify(value);
        } catch (e) {
            return value;
        }
    }

    function _deserialize(value) {
        try {
            return JSON.parse(value);
        } catch (e) {
            return value;
        }
    }

    function _isExpired(schedule) {
        return schedule && schedule.expires < Date.now();
    }

}).call(eformity.localstorage);


// ========================
//   eformity.hostinfo.js
// ========================
(function () {

    this.initialize = function () {
        var mailbox110 = Office.context.requirements.isSetSupported('Mailbox', '1.10');
        var mailbox18 = Office.context.requirements.isSetSupported('Mailbox', '1.8');

        eformity.hostinfo.support.setSignature = mailbox110;
        eformity.hostinfo.support.composeType = mailbox110;

        //eformity.hostinfo.support.addFileAttachmentFromBase64Async = mailbox18;
        eformity.hostinfo.support.addFileAttachmentFromBase64Async = !!Office.context && !!Office.context.mailbox && !!Office.context.mailbox.item && !!Office.context.mailbox.item.addFileAttachmentFromBase64Async;

        eformity.hostinfo.support.storage = false;
        try {
            eformity.hostinfo.support.storage = !!OfficeRuntime && !!OfficeRuntime.storage;
        } catch (e) { }
    };

}).call(eformity.hostinfo);


// =========================================
//   eformity.hosts.officeAddin.outlook.js
// =========================================
(function () {

    var app = {
        showError: eformity.ui.errorHandler
    };

    this.name = 'Outlook';

    // ====================
    //   Insert Signature
    // ====================
    this.insertSignature = function (signature, callback, asyncContext) {
        // allow a string 'content' or an object { content: '', images: {} } to be passed
        if (eformity.common.isString(signature)) {
            signature = { content: signature };
        }

        // FIX: for email models having different property names for the same purpose.
        // TODO: rather use .content instead of .html, but for now lets use html. Do when email signatures are refactored
        if (eformity.common.isString(signature.html)) {
            signature.content = signature.html;
        }

        var limit = 30000;
        var func = Office.context.mailbox.item.body.setSignatureAsync;
        var content = signature.content || '';

        // if the signature length exceeds the limit (or setSignature is not supported), we should be able to fallback to setSelectetData which supports longer length
        if (!eformity.hostinfo.support.setSignature || content.length > limit) {
            limit = 1000000;
            func = Office.context.mailbox.item.body.setSelectedDataAsync;
            content = '<br />' + signature.content + '<br />';
        }

        if (content.length > limit) {
            app.showError(eformity.host._messages.get('signatureTooLarge'));

            if (callback) {
                callback(false);
            }

            return;
        }

        var setSignature = function (content) {
            func(content, { coercionType: Office.CoercionType.Html, asyncContext: asyncContext }, function (asyncResult) {
                var result = true;

                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    eformity.ui.errorHandler(asyncResult.error);
                    result = false;
                }

                if (callback) {
                    callback(result, asyncResult.asyncContext);
                }
            })
        };

        if (signature.images && signature.images.length) {
            this._addImagesAsAttachments(signature.images, function () {
                setSignature(content);
            });
        } else {
            setSignature(content);
        }
    }

    this.getInsertSizeLimit = function () {
        if (eformity.hostinfo.support.setSignature) {
            // .setSignatureAsync() allows 30_000 characters
            return 30000;
        } else {
            // .setSelectedDataAsync() allows 1_000_000 characters
            return 1000000;
        }
    };

    // ============================
    //   Disable Client Signature
    // ============================
    this.disableClientSignature = function (callback, asyncContext) {
        if (eformity.hostinfo.support.setSignature) {
            this.isClientSignatureEnabled(function (isEnabled, asyncContext) {
                if (isEnabled) {
                    Office.context.mailbox.item.disableClientSignatureAsync({ asyncContext: asyncContext }, function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                            eformity.ui.errorHandler(asyncResult.error);
                        } else {
                            if (callback) {
                                callback(asyncResult.asyncContext);
                            }
                        }
                    });
                } else {
                    // the client signature is already disabled
                    callback(asyncContext);
                }
            }, asyncContext);
        } else {
            _invoke(callback, asyncContext)
        }
    };

    this.isClientSignatureEnabled = function (callback, asyncContext) {
        if (eformity.hostinfo.support.setSignature) {
            Office.context.mailbox.item.isClientSignatureEnabledAsync({ asyncContext: asyncContext }, function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    eformity.ui.errorHandler(asyncResult.error);
                } else {
                    if (callback) {
                        callback(asyncResult.value, asyncResult.asyncContext);
                    }
                }
            });
        } else {
            _invoke(callback, true, asyncContext)
        }
    };

    // ==================
    //   Insert Content
    // ==================
    this.insertContent = function (content, callback, asyncContext) {
        Office.context.mailbox.item.body.setSelectedDataAsync(content, { coercionType: Office.CoercionType.Html, asyncContext: asyncContext }, function (asyncResult) {
            var result = true;

            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                eformity.ui.errorHandler(asyncResult.error);
                result = false;
            }

            if (callback) {
                callback(result, asyncResult.asyncContext);
            }
        });
    }

    // ===============
    //   Set Subject
    // ===============
    this.setSubject = function (subject, callback, asyncContext) {
        Office.context.mailbox.item.subject.setAsync(subject, { asyncContext: asyncContext }, function (asyncResult) {
            var result = true;

            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                eformity.ui.errorHandler(asyncResult.error);
                result = false;
            }

            if (callback) {
                callback(result, asyncResult.asyncContext);
            }
        });
    }

    // ================
    //   Insert Image
    // ================
    this.insertImage = function (data, callback, width, height, contentType) {
        var img = '<img src="data:{0};base64,{1}" />'.format(contentType || 'image/jpeg', data);

        // The data string in office js is limited to 1.000.000 characters
        if (img.length > 1000000) {
            app.showError(eformity.host._messages.get('imageTooLarge'));

            if (callback) {
                callback(false);
            }

            return;
        }

        this.insertContent(img, callback);
    }

    this._addImagesAsAttachments = function (images, callback) {
        var me = this;
        var img = images.pop();

        if (!img) {
            // no images left, we are done successfully
            if (callback) {
                callback(true);
            }
            return;
        }

        var type;
        var data;

        if (eformity.hostinfo.support.addFileAttachmentFromBase64Async && img.data) {
            type = AttachmentType.cid;
            data = img.data;
        } else if (img.url) {
            type = AttachmentType.url;
            data = eformity.common.toAbsoluteUrl('{0}?sessionId={1}'.format(img.url, $('html').attr('data-sessionId') || ''));
        }

        this.addAttachment(type, data, img.id, function (result) {
            if (!result || result.error) {
                if (callback) {
                    callback(false);
                }
            } else {
                if (!images.length) {
                    // no images left, we are done successfully
                    if (callback) {
                        callback(true);
                    }
                } else {
                    me._addImagesAsAttachments(images, callback);
                }
            }
        }, { isInline: true, preventDuplicate: true });
    }

    // ==================
    //   Add Attachment
    // ==================
    this.addAttachmentFromUrl = function (url, name, callback, options) {
        this.addAttachment(AttachmentType.url, url, name, callback, options);
    };

    this.addAttachmentFromBase64 = function (data, name, callback, options) {
        this.addAttachment(AttachmentType.cid, data, name, callback, options);
    };

    this.addAttachment = function (attachmentType, data, name, callback, options) {
        _addAttachment(attachmentType, data, name, callback, options);
    };

    function _addAttachment(attachmentType, data, name, callback, options) {
        options = options || {};

        var exists = eformity.common.isBoolean(options.preventDuplicate) && options.preventDuplicate && attachmentExists(name);
        if (!exists) {
            var isBase64 = attachmentType == AttachmentType.cid;
            var isUrl = attachmentType == AttachmentType.url;
            var isInline = eformity.common.isBoolean(options.isInline) ? options.isInline : false;

            var addAttachmentMethod;
            var addAttachmentOptions;

            // urls can not use addFileAttachmentFromBase64Async
            if (eformity.hostinfo.support.addFileAttachmentFromBase64Async && isBase64) {
                addAttachmentMethod = Office.context.mailbox.item.addFileAttachmentFromBase64Async;
                addAttachmentOptions = {
                    isInline: isInline
                };
            } else {
                addAttachmentMethod = Office.context.mailbox.item.addFileAttachmentAsync;
                addAttachmentOptions = {
                    isInline: isInline
                };
            }

            addAttachmentMethod(data, name, addAttachmentOptions, function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    if (isUrl && data.indexOf('://localhost') > -1) {
                        var msg = eformity.localization.get('office.errorUrlAttachmentLocalhost');
                        eformity.ui.errorHandler(msg);
                    } else {
                        eformity.ui.errorHandler(asyncResult.error);
                    }
                } else {
                    if (callback) {
                        callback(true);
                    }
                }
            });
        } else {
            if (callback) {
                callback(false);
            }
        }
    }

    // =====================
    //   Show Notification
    // =====================
    // TODO: implement correctly
    // https://docs.microsoft.com/en-us/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-preview
    var _notificationKeys = [];
    var _notificationCount = 0;
    var _notificationMessageTypes = {
        progressIndicator: 'progressIndicator',
        informationalMessage: 'informationalMessage',
        errorMessage: 'errorMessage',
        insightMessage: 'insightMessage' // Allows actions. Preview API set
    };

    this.showNotification = function (message, callback, options) {
        if (!options) {
            options = {};
        }

        if (!message) {
            message = '';
        }

        // office notification message can only be 150 characters long
        var msg = eformity.common.truncate(message, 150);

        var details = {
            type: options.type || _notificationMessageTypes.informationalMessage,
            message: msg
        };

        if (options.showTaskPane === true) {
            details.actions = [{
                actionText: eformity.localization.get('office.showTaskPane'),
                actionType: 'showTaskPane',
                commandId: 'eformity.TaskpaneButton'
            }];
        }

        // if actions are specified, we must use insightMessage
        if (details.actions && details.actions.length) {
            details.type = _notificationMessageTypes.insightMessage;
        }

        // only informationalMessage and insightMessage allow an icon
        if (details.type == _notificationMessageTypes.informationalMessage || details.type == _notificationMessageTypes.insightMessage) {
            details.icon = 'eformity.tpicon_32x32';
        } else if (typeof details.icon !== 'undefined') {
            delete details.icon;
        }

        // force persistent only for informational messages
        if (details.type == _notificationMessageTypes.informationalMessage && typeof details.persistent === 'undefined') {
            details.persistent = false;
        } else if (details.type != _notificationMessageTypes.informationalMessage && typeof details.persistent !== 'undefined') {
            delete details.persistent;
        }

        // store and return a unique key which can be used to remove the notification, for now, loop through 5 notifications, by replacing the oldest one
        if (_notificationCount > 4) {
            _notificationCount = 0;
        }

        var key = 'notification_' + _notificationCount.toString();

        _notificationKeys.push(key);
        _notificationCount = _notificationCount + 1;

        // only call the replaceAsync with callback parameter if defined otherwise it causes an error
        if (callback) {
            Office.context.mailbox.item.notificationMessages.replaceAsync(key, details, callback);
        } else {
            Office.context.mailbox.item.notificationMessages.replaceAsync(key, details);
        }

        return key;
    }

    // TODO: fix closeNotification, as it doesnt work
    this.closeNotification = function (key, callback) {
        if (!key || !_notificationKeys.length) {
            if (callback) {
                callback();
            }
            return;
        }

        // cant use slice in this ecma version, so use own implementation
        if (!eformity.fallbacks.includes(_notificationKeys, key)) {
            if (callback) {
                callback();
            }
            return;
        }

        // cant use slice in this ecma version, so use own implementation
        _notificationKeys = eformity.fallbacks.removeByValue(_notificationKeys, key);

        // only call the replaceAsync with callback parameter if defined otherwise it causes an error
        if (callback) {
            Office.context.mailbox.item.notificationMessages.removeAsync(key, callback);
        } else {
            Office.context.mailbox.item.notificationMessages.removeAsync(key);
        }
    }

    // ================
    //   Compose Type
    // ================
    this.getComposeType = function (callback, asyncContext) {
        if (!callback) {
            return;
        }

        // find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
        if (eformity.hostinfo.support.composeType) {
            Office.context.mailbox.item.getComposeTypeAsync({ asyncContext: asyncContext, }, function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    eformity.ui.errorHandler(asyncResult.error);
                    return;
                }

                var value = asyncResult.value.composeType;
                var composeType = ComposeType.fromOfficeComposeType(value);

                callback(composeType, asyncResult.asyncContext);
            });
        } else {
            // fallback to use 'new' when getComposeTypeAsync does not exist
            callback(ComposeType.new, asyncContext);
        }
    }

    function attachmentExists(name) {
        // TODO: The 'attachments' property does not work/returns null(?). Find out why.
        var attachments = Office.context.mailbox.item.attachments;
        if (attachments && attachments.length) {
            for (i = 0; i < attachments.length; i++) {
                if (name == attachments[i].name) {
                    return true;
                }
            }
        }

        return false;
    }

    // ============================
    //   Outlook Roaming Settings 
    // ============================
    this.getRoamingSetting = function (key, callback) {
        var value = Office.context.roamingSettings.get(key);

        var result = eformity.fallbacks.decodeJSON(value);

        _invoke(callback, result);

        console.log('[outlook:roamingSettings] retrieved setting: {0}'.format(key), { key: key, value: result });
    };

    this.setRoamingSetting = function (key, value, callback) {
        var result = eformity.fallbacks.encodeJSON(value);

        Office.context.roamingSettings.set(key, result);

        this.saveRoamingSettings(callback, function () {
            console.log('[outlook:roamingSettings] updated setting: {0}'.format(key), { key: key, value: value });
        });
    };

    this.removeRoamingSetting = function (key, callback) {
        Office.context.roamingSettings.remove(key);

        this.saveRoamingSettings(callback, function () {
            console.log('[outlook:roamingSettings] removed setting: {0}'.format(key));
        });
    };

    this.removeRoamingSettings = function (keys, callback) {
        for (var i = 0; i < keys.length; i++) {
            Office.context.roamingSettings.remove(keys[i]);
        }

        this.saveRoamingSettings(callback, function () {
            console.log('[outlook:roamingSettings] removed settings: {0}'.format(keys.join(', \n')));
        });
    };

    this.saveRoamingSettings = function (callback, log) {
        Office.context.roamingSettings.saveAsync(function (asyncResult) {
            var isSuccess = asyncResult.status === Office.AsyncResultStatus.Succeeded;
            _invoke(callback, isSuccess);

            _invoke(log);
        });
    };

}).call(eformity.hosts.officeAddin.outlook);


// ====================
//   eformity.host.js
// ====================
(function () {

    this._messages.get = function (message) {
        switch (message) {
            case 'imageTooLarge':
                return 'Image too large';
            default:
                return 'An error occured: ' + message;
        }
    };

    // ==================
    //   Office Storage 
    // ==================
    this.getStorageItem = function (key, callback) {
        if (eformity.hostinfo.support.storage) {
            _getItem.call(this, key, callback);
        } else {
            var val = eformity.localstorage.get(key);
            _invoke(callback, val);
        }
    };

    this.setStorageItem = function (key, value, callback, expires, softExpire) {
        if (eformity.hostinfo.support.storage) {
            _setItem.call(this, key, value, callback, expires, softExpire);
        } else {
            eformity.localstorage.set(key, value, expires);
            _invoke(callback);
        }
    };

    this.removeStorageItem = function (key, callback) {
        if (eformity.hostinfo.support.storage) {
            _removeItem.call(this, key, callback);
        } else {
            eformity.localstorage.remove(key);
            _invoke(callback);
        }
    };

    this.removeStorageItems = function (keys, callback) {
        if (eformity.hostinfo.support.storage) {
            _removeItems.call(this, keys, callback);
        } else {
            for (var i = 0; i < keys.length; i++) {
                eformity.localstorage.remove(keys[i]);
            }
            _invoke(callback);
        }
    };

    function _getItem(key, callback) {
        var me = this;

        var expiryMessage = '';
        var isExpired = false;

        _internalGetItem(key + '_Expires', function (schedule) {
            if (schedule) {
                isExpired = _isExpired(schedule);

                if (isExpired) {
                    console.log('[outlook:eformity.host.storage] the item {0} is expired'.format(key));

                    if (!schedule.softExpire) {
                        me.removeStorageItem(key, callback);

                        return null;
                    }
                } else {
                    expiryMessage = ' ({0} on: {1})'.format(isExpired ? 'expired' : 'expires', new Date(schedule.expires).datetimeNow());
                }
            }

            _internalGetItem(key, function (result) {
                if (isExpired) {
                    result.isExpired = true;
                }

                _invoke(callback, result);

                console.log('[outlook:eformity.host.storage] retrieved item: {0}{1}'.format(key, expiryMessage), { value: result });
            });
        });
    }

    function _setItem(key, value, callback, expires, softExpire) {
        var expiryMessage = '';

        if (expires) {
            var now = Date.now();

            var schedule = {
                created: now,
                expires: now + Math.abs(expires),
                softExpire: softExpire
            };

            _internalSetItem(key + '_Expires', schedule, callback);

            expiryMessage = ' (expires on: {0})'.format(new Date(schedule.expires).datetimeNow());
        }

        _internalSetItem(key, value, callback);

        console.log('[outlook:eformity.host.storage] stored item: {0}{1}'.format(key, expiryMessage));
    }

    function _removeItem(key, callback) {
        var keys = [key, key + '_Expires'];

        _internalRemoveItems(keys, callback);

        console.log('[outlook:eformity.host.storage] removed item: {0}'.format(key));
    }

    function _removeItems(keys, callback) {
        var items = keys;
        var length = keys.length;

        // for each key, we also need to remove the {key}_Expires item, so duplicate all keys and append _Expires
        for (var i = 0; i < length; i++) {
            items.push(keys[i] + '_Expires');
        }

        _internalRemoveItems(keys, callback);

        console.log('[outlook:eformity.host.storage] removed items: {0}'.format(keys.join(', \n')), { keys: keys });
    }

    function _internalGetItem(key, callback) {
        OfficeRuntime.storage.getItem(key).then(function (value) {
            var result = _deserialize(value);
            _invoke(callback, _deserialize(value));
        }, callback);
    }

    function _internalSetItem(key, value, callback) {
        var result = _serialize(value);
        OfficeRuntime.storage.setItem(key, result).then(callback, callback);
    }

    function _internalRemoveItem(key, callback) {
        OfficeRuntime.storage.removeItem(key).then(callback, callback);
    }

    function _internalRemoveItems(keys, callback) {
        OfficeRuntime.storage.removeItems(keys).then(callback, callback);
    }

    function _serialize(value) {
        try {
            return JSON.stringify(value);
        } catch (e) {
            return value;
        }
    }

    function _deserialize(value) {
        try {
            return JSON.parse(value);
        } catch (e) {
            return value;
        }
    }

    function _isExpired(schedule) {
        return schedule && schedule.expires < Date.now();
    }

}).call(eformity.host);


// ======================
//   eformity.common.js
// ======================
(function () {

    // =====================
    //   Data type helpers
    // =====================
    this.isNull = function (value) {
        return value === null;
    }

    this.isUndefined = function (value) {
        return typeof value === 'undefined';
    }

    this.isNullOrUndefined = function (value) {
        return this.isNull(value) || this.isUndefined(value);
    }

    this.isNullOrEmpty = function (value) {
        return this.isNullOrUndefined(value) || !value.length;
    }

    this.isNullOrWhiteSpace = function (value) {
        return this.isNullOrEmpty((value || '').trim());
    }

    this.isString = function (value) {
        return typeof value === 'string';
    }

    this.isBoolean = function (value) {
        return typeof value === 'boolean';
    }

    this.isNumber = function (value) {
        return typeof value === 'number' && isFinite(value);
    }

    this.isNumeric = function (value) {
        return !this.isNullOrUndefined(value) && !isNaN(value) && isFinite(value);
    }

    this.isArray = function (value) {
        return value && typeof value === 'object' && value.constructor === Array;
    }

    this.isEnumerable = function (value) {
        return this.isArray(value) || this.isJQuery(value);
    }

    this.isObject = function (value) {
        return value && typeof value === 'object' && value.constructor === Object;
    }

    this.isFunction = function (value) {
        return typeof value === 'function';
    }

    this.isJson = function (value) {
        try {
            return eformity.common.isObject(JSON.parse(value));
        } catch (e) {
            return false;
        }
    }

    // ================
    //  String helpers
    // ================
    this.trim = function (source, char) {
        return this.trimStart(this.trimEnd(source, char), char);
    }

    this.trimEnd = function (source, char) {
        if (eformity.common.isNullOrEmpty(char)) {
            char = ' ';
        }
        while (source.charAt(source.length - 1) == char) {
            source = source.substring(0, source.length - 1);
        }
        return source;
    }

    this.trimStart = function (source, char) {
        if (eformity.common.isNullOrEmpty(char)) {
            char = ' ';
        }
        while (source.charAt(0) == char) {
            source = source.substring(1);
        }
        return source;
    }

    this.truncate = function (source, limit) {
        if (!eformity.common.isString(source)) {
            return null;
        }

        // string length is allowed
        if (source.length <= limit) {
            return source;
        }

        // limits the string length and replaces the last charcter with an ellipsis (…) character, if the string is too long
        return source.substr(0, limit - 1) + '…';
    };

    // ===============
    //   URL helpers
    // ===============
    this.removeReservedCharacters = function (source, replacement) {
        // These characters are not allowed for filenames, so encode and replace dots: \ / : * ? " < > | .
        return source.replace(/[/\\:*?"<>|.]/g, replacement || '_');
    }

    this.isAbsoluteUrl = function (url) {
        if (!eformity.common.isString(url)) {
            return false;
        }

        return url.indexOf('://') > -1;
    }

    this.toAbsoluteUrl = function (url) {
        if (!eformity.common.isString(url)) {
            return null;
        }

        if (eformity.common.isAbsoluteUrl(url)) {
            return url;
        }

        return '{0}/{1}'.format(window.location.origin, eformity.common.trimStart(url, '/'));
    }

}).call(eformity.common);

String.prototype.format = function () {
    var args = arguments;

    // to support an array as arguments
    if (args[0] instanceof Array) {
        args = args[0];
    }

    // to support easy resource localization of parameters starting with $$
    //var original = args;
    //try {
    //    if (eformity.localization) {
    //        for (var i = 0; i < args.length; i++) {
    //            args[i] = eformity.localization.tryParse(args[i]);
    //        }
    //    }
    //} catch (e) {
    //    // restore original if parsing failed just to be sure
    //    args = original;
    //}

    return this.replace(/{(\d+)}/g, function (match, number) {
        return !eformity.common.isUndefined(args[number]) ? args[number] : match;
    });
};

/** get the date in day/month/year format */
Date.prototype.datetimeNow = function () {
    return this.today() + ' ' + this.timeNow();
};

/** get the date in day/month/year format */
Date.prototype.today = function () {
    return ((this.getDate() < 10) ? '0' : '') + this.getDate() + '/' + (((this.getMonth() + 1) < 10) ? '0' : '') + (this.getMonth() + 1) + '/' + this.getFullYear();
};

/** get the time in hour:minute:second format */
Date.prototype.timeNow = function () {
    return ((this.getHours() < 10) ? '0' : '') + this.getHours() + ':' + ((this.getMinutes() < 10) ? '0' : '') + this.getMinutes() + ':' + ((this.getSeconds() < 10) ? '0' : '') + this.getSeconds();
};


// ==================
//   eformity.ui.js
// ==================
(function () {

    this.errorHandler = function (error) {
        var unknownError = eformity.localization.get('office.unknownError');

        if (eformity.common.isObject(error)) {
            var msg = error.message || unknownError;
            console.error('[outlook:eformity.ui] ' + msg, error);
        } else if (eformity.common.isString(error)) {
            console.error('[outlook:eformity.ui] ' + error);
        } else {
            console.error('[outlook:eformity.ui] ' + unknownError);
        }
    };

}).call(eformity.ui);


// ============================
//   eformity.localization.js
// ============================
(function () {

    var _resources = {
        office: {
            insertSignatureSuccess: {
                en: 'The default signature is automatically added. Open the task pane for more features.',
                nl: 'De standaardhandtekening is automatisch ingevoegd. Open het zijpaneel voor meer functionaliteiten.'
            },
            noDefaultSignature: {
                en: 'The default signature is not set. Open the task pane and set your signature.',
                nl: 'De standaardhandtekening is niet ingesteld. Open het zijpaneel om deze in te stellen.'
            },
            invalidDefaultSignature: {
                //en: 'The default signature is not valid and could not be inserted. Open the task pane and set your signature.',
                //nl: 'De standaard handtekening is niet geldig en kon niet worden ingevoegd. Open het zijpaneel om deze in te stellen.'
                en: 'The default signature is not set. Open the task pane and set your signature.',
                nl: 'De standaardhandtekening is niet ingesteld. Open het zijpaneel om deze in te stellen.'
            },
            noSignatureStorageAndGet: {
                en: 'There is no signature found in local storage, and no signature could not be retrieved from the server.',
                nl: 'Er is geen handtekening gevonden in de lokale opslag en er kan geen handtekening worden opgehaald van de server.'
            },
            insertSignatureHttpError: {
                en: 'Failed to insert the default signature. Open the task pane and validate your signature.',
                nl: 'Kan de standaardhandtekening niet invoegen. Open het zijpaneel en controleer uw handtekening.'
            },
            errorUrlAttachmentLocalhost: {
                en: 'The attachment cannot be added to the item. This is because the server can\'t reach your localhost to download the image. This does work in production environments.',
                nl: 'De bijlage kan niet worden toegevoegd aan het item. Dit komt omdat de server niet bij je localhost kan om de afbeelding te downloaden. Dit werkt wel op productie omgevingen.'
            },
            showTaskPane: {
                en: 'Open the task pane',
                nl: 'Open het zijpaneel'
            },
            showNotificationFailed: {
                en: 'Notification could not be opened',
                nl: 'Notificatie kan niet worden geopend'
            },
            setSubjectFailed: {
                en: 'Subject could not be set',
                nl: 'Onderwep kan niet worden ingesteld'
            },
            unknownError: {
                en: 'An unknown error occurred.',
                nl: 'Er is een onbekende fout opgetreden.'
            }
        }
    };

    this.defaultCulture = 'en';

    this.get = function (name) {
        var cul;
        try {
            cul = (Office.context.contentLanguage || Office.context.displayLanguage).split('-')[0];
        } catch (e) {
            cul = this.defaultCulture;
        }

        var culture = cul && cul.split('-')[0];

        var resource = getResource(name, culture) || '';

        if (arguments.length == 1) {
            return resource;
        }

        var args = Array.prototype.slice.call(arguments, 1);
        return resource.format(args);
    }

    function getResource(name, culture) {
        if (!name || !culture) {
            return null;
        }

        var names = name.split('.');
        var resources = _resources;
        var exists = true;

        for (var i = 0; i < names.length; i++) {
            if (exists && resources) {
                resources = resources[names[i]];
            } else {
                exists = false;
            }
        }

        if (!exists || !resources) {
            return null;
        }

        if (resources.hasOwnProperty(culture)) {
            return resources[culture];
        }

        return resources[eformity.localization.defaultCulture];
    }

}).call(eformity.localization);


// =========================
//   eformity.fallbacks.js
// =========================
// Miscellaneous functions that are needed, but we dont need the entire library from
// This is to prevent overhead. Event-Based Activation script should be small.
// Probably smaller than it is already...
(function () {

    this.daysToMiliseconds = function (days) {
        return days * 24 * 60 * 60 * 1000;
    };

    this.includes = function (source, value) {
        for (var i = 0; i < source.length; i++) {
            if (source[i] === value) {
                return true;
            }
        }

        return false;
    };

    this.indexOf = function (source, value) {
        for (var i = 0; i < source.length; i++) {
            if (source[i] === value) {
                return i;
            }
        }

        return -1;
    };

    this.removeByValue = function (source, value) {
        var result = [];

        for (var i = 0; i < source.length; i++) {
            if (source[i] !== value) {
                result.push(source[i]);
            }
        }

        return result;
    };

    this.removeByIndex = function (source, index) {
        var result = [];

        for (var i = 0; i < source.length; i++) {
            if (i != index) {
                result.push(source[i]);
            }
        }

        return result;
    };

    this.extend = function (target, source) {
        var _extend = function (target, source) {
            for (var prop in source) {
                if (Object.prototype.hasOwnProperty.call(source, prop)) {
                    target[prop] = source[prop];
                }
            }

            return target;
        };

        return _extend(_extend({}, target), source);
    };

    // =================
    //   JSON Encoding
    // =================
    this.encodeJSON = function (value) {
        try {
            return JSON.stringify(value);
        } catch (e) {
            return value;
        }
    };

    this.decodeJSON = function (value) {
        try {
            return JSON.parse(value);
        } catch (e) {
            return value;
        }
    };

}).call(eformity.fallbacks);


// ====================
//   eformity.http.js
// ====================
// custom HTTP Request implementation
// https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch#requesting-external-data
(function () {

    this.get = function (url, callback, options) {
        this.request('GET', url, null, callback, options);
    };

    this.post = function (url, data, callback, options) {
        this.request('POST', url, data, callback, options);
    };

    this.request = function (method, url, data, callback, options) {
        options = options || {};

        //console.log('[outlook:eformity.http] {0} {1}'.format(method, url));

        if (eformity.common.isBoolean(options.withCredentials)) {
            xhr.withCredentials = options.withCredentials;
        }

        options.url = url;
        options.method = method;

        var xhr = new XMLHttpRequest();

        xhr.onreadystatechange = function () {
            _onReadyStateChanged.call(this, xhr, callback, options);
        };

        xhr.open(method, url, true);

        if (options.withCredentials === true) {
            xhr.withCredentials = true;
        }

        _appendHeaders(xhr, options);

        var body = _encodeBody(data);
        if (body) {
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded; charset=UTF-8');

            xhr.send(body);
        } else {
            xhr.send(null);
        }
    };

    function _onReadyStateChanged(xhr, callback, options) {
        options = options || {};

        if (xhr.readyState == 1) {
            _invoke(options.beforeSend);
        }

        if (xhr.readyState == 4) {
            _invoke(options.complete, xhr, xhr.status);

            if (xhr.status == 200 || xhr.status == 0) {
                if (eformity.common.isFunction(options.success) || eformity.common.isFunction(callback)) {
                    var response = xhr.responseText;
                    try {
                        if (response) {
                            if (options.responseType && options.responseType.indexOf('json') > -1 || options.contentType && options.contentType.indexOf('json')) {
                                response = JSON.parse(response);
                            }
                        }
                    } catch (e) {
                        console.error('[outlook:eformity.http:{0}] Error: {1}'.format(options.method, options.url), { responseText: xhr.responseText });
                    }

                    console.log('[outlook:eformity.http:{0}] {1} {2} {3}'.format(options.method, xhr.status, xhr.statusText, options.url), { response: response });

                    _invoke(options.success, response);
                    _invoke(callback, response);
                }
            } else {
                _invoke(options.error, xhr, xhr.status, xhr.responseText);
            }
        }
    }

    function _appendHeaders(xhr, options) {
        var headers = options.headers || {};

        if (eformity.common.isString(options.authorization)) {
            headers['Authorization'] = options.authorization;
        }

        for (var header in headers) {
            if (headers.hasOwnProperty(header)) {
                xhr.setRequestHeader(header, headers[header]);
            }
        }
    }

    function _encodeBody(data) {
        if (!data) {
            return null;
        }

        var result = [];

        for (var key in data) {
            var val = encodeURIComponent(key) + '=' + encodeURIComponent(data[key]);
            result.push(val);
        }

        return result.join('&');
    }

}).call(eformity.http);


// ==================================
//   eformity.hostinfo.constants.js
// ==================================
const StorageKey = {
    newMailStorageKey: 'eformity.signatures.new',
    replyMailStorageKey: 'eformity.signatures.reply',
    clearStorageFlag: 'eformity.signatures.clearStorage',
    actionCredentialToken: 'actionCredentialToken'
};

// eformity.api.v2.EmailSignatureAttachmentType
const AttachmentType = {
    cid: 0,
    url: 1,
    embeddedBase64: 2,
    embeddedUrl: 3
};

// eformity.api.v2.EmailSignatureComposeType
const ComposeType = {
    new: 0,
    reply: 1,
    forward: 2,
    fromOfficeComposeType: function (composeType) {
        if (composeType == 'newMail') {
            return ComposeType.new;
        } else if (composeType == 'reply') {
            return ComposeType.reply;
        } else if (composeType == 'forward') {
            return ComposeType.forward;
        }

        return ComposeType.new;
    },
    getStorageKey: function (composeType) {
        // we only support reply and new compose types. forward should use the new signature
        if (composeType == ComposeType.reply) {
            return StorageKey.replyMailStorageKey;
        }

        return StorageKey.newMailStorageKey;
    }
};

// ===========
//   General 
// ===========
// freeze the constants to prevent them from being modified
if (Object.freeze) {
    Object.freeze(StorageKey);
    Object.freeze(AttachmentType);
    Object.freeze(ComposeType);
}
