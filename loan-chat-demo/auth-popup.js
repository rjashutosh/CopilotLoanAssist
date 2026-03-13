/**
 * Popup window: starts sign-in or handles callback, then redirects opener to index.
 * Add this page's URL as a redirect URI in Azure (e.g. .../loan-chat-demo/auth-popup.html).
 */
(function () {
  'use strict';

  function getConfig() {
    return (typeof LOAN_CHAT_CONFIG !== 'undefined' && LOAN_CHAT_CONFIG && LOAN_CHAT_CONFIG.MSAL) ? LOAN_CHAT_CONFIG.MSAL : null;
  }

  function getBasePath() {
    var path = window.location.pathname || '';
    return path.replace(/auth-popup\.html$/i, '') || '/';
  }

  function getIndexUrl() {
    var base = getBasePath();
    return window.location.origin + (base.slice(-1) === '/' ? base : base + '/') + 'index.html';
  }

  function redirectOpenerToIndex() {
    if (window.opener && !window.opener.closed) {
      window.opener.location.href = getIndexUrl();
    }
    window.close();
  }

  function showError(msg) {
    var el = document.getElementById('msg');
    if (el) el.textContent = msg || 'Sign-in failed. You can close this window.';
  }

  var hash = window.location.hash || '';
  var hasCallback = hash.indexOf('state=') !== -1 || hash.indexOf('code=') !== -1 || hash.indexOf('access_token=') !== -1 || hash.indexOf('error=') !== -1;

  if (hasCallback) {
    var config = getConfig();
    if (!config || !config.clientId) {
      showError('Config missing. Close and try again.');
      return;
    }
    var redirectUri = window.location.origin + window.location.pathname;
    var opts = {
      auth: {
        clientId: config.clientId,
        authority: config.authority || 'https://login.microsoftonline.com/common',
        redirectUri: redirectUri
      },
      cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false }
    };
    var instance = new window.msal.PublicClientApplication(opts);
    instance.initialize().then(function () {
      return instance.handleRedirectPromise();
    }).then(function (result) {
      if (result && result.account) {
        document.getElementById('msg').textContent = 'Success! Opening chat…';
        redirectOpenerToIndex();
      } else {
        var accounts = instance.getAllAccounts();
        if (accounts.length > 0) {
          redirectOpenerToIndex();
        } else {
          showError('Could not complete sign-in. Close and try again.');
        }
      }
    }).catch(function (err) {
      showError('Sign-in error: ' + (err && err.message ? err.message : 'Unknown'));
    });
  } else {
    var cfg = getConfig();
    if (!cfg || !cfg.clientId || cfg.clientId === 'YOUR_CLIENT_ID') {
      showError('Configure config.js (clientId).');
      return;
    }
    var popupRedirectUri = window.location.origin + window.location.pathname;
    var authOpts = {
      auth: {
        clientId: cfg.clientId,
        authority: cfg.authority || 'https://login.microsoftonline.com/common',
        redirectUri: popupRedirectUri
      },
      cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false }
    };
    var app = new window.msal.PublicClientApplication(authOpts);
    app.initialize().then(function () {
      return app.loginRedirect({
        scopes: [cfg.scope || 'https://api.powerplatform.com/.default'],
        redirectUri: popupRedirectUri
      });
    }).catch(function (err) {
      showError(err && err.message ? err.message : 'Sign-in failed.');
    });
  }
})();
