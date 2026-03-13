/**
 * Loan Support Assistant — Microsoft authentication + Copilot Studio chat
 * - Login page: Sign in with Microsoft (redirect to index after login)
 * - Chat page: Protected; requires MSAL session; shows chatbot and passes token to agent
 */
(function () {
  'use strict';

  var LOGIN_PAGE_URL = 'login.html';
  var CHAT_PAGE_URL = 'index.html';

  // ----- Login page -----
  var loginBtn = document.getElementById('btn-signin-ms');
  if (loginBtn) {
    var loginError = document.getElementById('login-error');
    loginBtn.addEventListener('click', function () {
      if (loginError) loginError.textContent = '';
      loginBtn.disabled = true;
      var popupPath = (window.location.pathname || '').replace(/login\.html$/i, 'auth-popup.html') || 'auth-popup.html';
      var popupUrl = window.location.origin + popupPath;
      var w = window.open(popupUrl, 'msalSignIn', 'width=500,height=600,scrollbars=yes');
      if (!w) {
        if (loginError) loginError.textContent = 'Popup blocked. Allow popups for this site and try again, or use the redirect flow.';
        loginBtn.disabled = false;
        return;
      }
      var checkClosed = setInterval(function () {
        if (w.closed) {
          clearInterval(checkClosed);
          loginBtn.disabled = false;
          window.location.href = (window.location.pathname || '').replace(/login\.html$/i, 'index.html') || 'index.html';
        }
      }, 500);
    });
  }

  // ----- Chat page -----
  if (document.body.classList.contains('chat-page')) {
    var chatLoading = document.getElementById('chat-loading');
    var chatContent = document.getElementById('chat-content');
    var userNameEl = document.getElementById('user-name');
    var userBadge = document.getElementById('user-badge');
    var logoutBtn = document.getElementById('btn-logout');

    function redirectToLogin() {
      window.location.href = LOGIN_PAGE_URL;
    }

    function showChatUI(account, msalInstance) {
      if (userNameEl) userNameEl.textContent = account.name || account.username || 'Signed in';
      if (chatLoading) chatLoading.hidden = true;
      if (chatContent) chatContent.hidden = false;
      if (logoutBtn) {
        logoutBtn.addEventListener('click', function () {
          if (typeof window.LoanChatAuth !== 'undefined' && window.LoanChatAuth.signOut) {
            window.LoanChatAuth.signOut(msalInstance);
          } else {
            redirectToLogin();
          }
        });
      }
      initProgrammaticChat(msalInstance);
    }

    if (typeof window.LoanChatAuth === 'undefined' || !window.LoanChatAuth.handleRedirect) {
      if (chatLoading) chatLoading.textContent = 'Auth not loaded. Redirecting…';
      setTimeout(redirectToLogin, 1500);
      return;
    }

    function hasOAuthHash() {
      var h = window.location.hash || '';
      return h.indexOf('state=') !== -1 || h.indexOf('code=') !== -1 || h.indexOf('access_token=') !== -1 || h.indexOf('error=') !== -1;
    }

    function tryShowChat(attempt) {
      attempt = attempt || 0;
      var maxAttempts = 4;
      if (chatLoading) chatLoading.textContent = attempt > 0 ? 'Completing sign-in…' : 'Checking sign-in…';
      return window.LoanChatAuth.handleRedirect().then(function (result) {
        if (result && result.account) {
          try {
            showChatUI(result.account, result.instance);
          } catch (e) {
            if (typeof console !== 'undefined' && console.error) console.error('Loan Support: showChatUI failed', e);
            if (chatLoading) chatLoading.textContent = 'Could not load chat. Please try again.';
            var errEl = document.getElementById('chat-error');
            if (errEl) { errEl.textContent = (e && e.message) || 'Error loading chat'; errEl.hidden = false; }
          }
          return;
        }
        if (hasOAuthHash() && attempt < maxAttempts - 1) {
          return new Promise(function (resolve) {
            setTimeout(function () { tryShowChat(attempt + 1).then(resolve).catch(resolve); }, 600);
          });
        }
        redirectToLogin();
      }).catch(function (err) {
        if (typeof console !== 'undefined' && console.warn) console.warn('Loan Support: handleRedirect failed', err);
        if (hasOAuthHash() && attempt < maxAttempts - 1) {
          return new Promise(function (resolve) {
            setTimeout(function () { tryShowChat(attempt + 1).then(resolve).catch(function () { redirectToLogin(); }); }, 600);
          });
        }
        redirectToLogin();
      });
    }

    tryShowChat(0);
  }

  /**
   * Programmatic chat: Conversations API with Bearer token from MSAL.
   * Requires LOAN_CHAT_CONFIG (config.js) and an initialized MSAL instance (from handleRedirect).
   */
  function initProgrammaticChat(msalInstance) {
    var messagesEl = document.getElementById('chat-messages');
    var inputEl = document.getElementById('chat-input');
    var sendBtn = document.getElementById('btn-send');
    var errorEl = document.getElementById('chat-error');
    if (!messagesEl || !inputEl || !sendBtn) return;

    var config = typeof LOAN_CHAT_CONFIG !== 'undefined' ? LOAN_CHAT_CONFIG : null;
    if (!config || !config.CONVERSATIONS_API_BASE) {
      showChatError('Missing config: set CONVERSATIONS_API_BASE in config.js');
      return;
    }
    if (!msalInstance) {
      showChatError('Not signed in. Please sign in again.');
      return;
    }

    var baseUrl = config.CONVERSATIONS_API_BASE.replace(/\/conversations.*$/, '');
    var apiVersion = config.API_VERSION || '2022-03-01-preview';
    var query = '?api-version=' + encodeURIComponent(apiVersion);
    var conversationId = null;
    var watermark = null;
    var shownActivityIds = {};
    var pollTimer = null;
    var msalConfig = config.MSAL || {};
    var userId = msalInstance.getAllAccounts()[0]?.localAccountId || ('user-' + Math.random().toString(36).slice(2, 10));

    function showChatError(msg) {
      if (errorEl) {
        errorEl.textContent = msg;
        errorEl.hidden = false;
      }
    }
    function clearChatError() {
      if (errorEl) { errorEl.textContent = ''; errorEl.hidden = true; }
    }

    function appendMessage(text, isUser) {
      var div = document.createElement('div');
      div.className = 'chat-msg ' + (isUser ? 'chat-msg-user' : 'chat-msg-bot');
      div.setAttribute('role', 'listitem');
      var inner = document.createElement('div');
      inner.className = 'chat-msg-bubble';
      inner.textContent = text;
      div.appendChild(inner);
      messagesEl.appendChild(div);
      messagesEl.scrollTop = messagesEl.scrollHeight;
    }

    function getAccessToken() {
      var redirectUri = msalConfig.redirectUri || (window.location.origin + window.location.pathname);
      var request = {
        scopes: [msalConfig.scope || 'https://api.powerplatform.com/.default'],
        redirectUri: redirectUri
      };
      return msalInstance.acquireTokenSilent(request).catch(function () {
        return msalInstance.acquireTokenPopup(request);
      });
    }

    function apiRequest(method, path, body) {
      return getAccessToken().then(function (tokenResponse) {
        var url = baseUrl + path + query;
        var opts = {
          method: method,
          headers: {
            'Authorization': 'Bearer ' + tokenResponse.accessToken,
            'Content-Type': 'application/json'
          }
        };
        if (body) opts.body = JSON.stringify(body);
        return fetch(url, opts);
      });
    }

    function createConversation() {
      return apiRequest('POST', '/conversations', {}).then(function (res) {
        if (!res.ok) {
          return res.text().then(function (body) {
            var msg = 'Create conversation failed: ' + res.status + ' ' + res.statusText;
            if (body) {
              try {
                var j = JSON.parse(body);
                if (j.message) msg += ' — ' + j.message;
                if (j.error) msg += ' — ' + j.error;
                if (j.innererror && j.innererror.message) msg += ' (' + j.innererror.message + ')';
              } catch (e) {
                msg += ' — ' + body.slice(0, 300);
              }
            }
            throw new Error(msg);
          });
        }
        return res.json();
      }).then(function (data) {
        var id = data.conversationId || data.id || data.ConversationId;
        if (id) return id;
        throw new Error('No conversationId in response');
      });
    }

    function sendActivity(text) {
      if (!conversationId) return Promise.reject(new Error('No conversation'));
      var activity = {
        type: 'message',
        from: { id: userId, name: 'User' },
        text: text
      };
      return apiRequest('POST', '/conversations/' + encodeURIComponent(conversationId) + '/activities', activity).then(function (res) {
        if (!res.ok) throw new Error('Send failed: ' + res.status);
        return res.json();
      });
    }

    function getActivities() {
      if (!conversationId) return Promise.resolve([]);
      var path = '/conversations/' + encodeURIComponent(conversationId) + '/activities';
      var q = query;
      if (watermark) q += (q.indexOf('?') >= 0 ? '&' : '?') + 'watermark=' + encodeURIComponent(watermark);
      return apiRequest('GET', path + q, null).then(function (res) {
        if (!res.ok) return [];
        return res.json();
      }).then(function (data) {
        var activities = data.activities || data || [];
        if (Array.isArray(activities)) {
          activities.forEach(function (a) {
            if (a.id && shownActivityIds[a.id]) return;
            if (a.from && (a.from.role === 'bot' || !a.from.role) && a.text) {
              appendMessage(a.text, false);
              if (a.id) shownActivityIds[a.id] = true;
            }
            if (a.id) watermark = a.id;
          });
        }
        return activities;
      }).catch(function () { return []; });
    }

    function pollForMessages() {
      if (pollTimer) clearTimeout(pollTimer);
      getActivities().then(function () {
        pollTimer = setTimeout(pollForMessages, 2000);
      });
    }

    createConversation().then(function (id) {
      conversationId = id;
      appendMessage('Connected. How can I help with your loan today?', false);
      pollForMessages();
    }).catch(function (err) {
      showChatError('Could not start conversation: ' + (err.message || err));
    });

    sendBtn.addEventListener('click', function () {
      var text = (inputEl.value || '').trim();
      if (!text || !conversationId) return;
      inputEl.value = '';
      appendMessage(text, true);
      sendActivity(text).then(function () {
        pollForMessages();
      }).catch(function (err) {
        appendMessage('Error sending: ' + (err.message || err), false);
      });
    });
    inputEl.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        sendBtn.click();
      }
    });
  }
})();
