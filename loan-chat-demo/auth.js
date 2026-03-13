/**
 * Microsoft authentication (MSAL) for Loan Support Assistant.
 * Handles redirect flow: login.html starts sign-in, index.html is redirect target and runs chat.
 */
(function () {
  'use strict';

  var MSAL_CONFIG_KEY = 'loanChatMsalConfig';

  function getMsalConfig() {
    if (typeof LOAN_CHAT_CONFIG !== 'undefined' && LOAN_CHAT_CONFIG && LOAN_CHAT_CONFIG.MSAL) {
      return LOAN_CHAT_CONFIG.MSAL;
    }
    return null;
  }

  /**
   * Create and initialize MSAL PublicClientApplication.
   * redirectUri must be the full URL of the page where the user lands after login (index.html).
   */
  function createMsalInstance(redirectUri) {
    var config = getMsalConfig();
    if (!config || !config.clientId || config.clientId === 'YOUR_CLIENT_ID') {
      return null;
    }
    var authority = config.authority || 'https://login.microsoftonline.com/common';
    var opts = {
      auth: {
        clientId: config.clientId,
        authority: authority,
        redirectUri: redirectUri || config.redirectUri || config.redirectUrl || (window.location.origin + window.location.pathname)
      },
      cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
    };
    var instance = new window.msal.PublicClientApplication(opts);
    return instance;
  }

  /**
   * Run on index.html: handle redirect from Microsoft login, return account or null.
   */
  window.LoanChatAuth = {
    getConfig: getMsalConfig,
    createMsalInstance: createMsalInstance,

    /**
     * Must be called once on the redirect target page (index.html) after MSAL is loaded.
     * @param {string} [redirectUri] - Optional; defaults to current page or config.
     * @returns {Promise<{ account: object, instance: object }|null>}
     */
    handleRedirect: function (redirectUri) {
      var config = getMsalConfig();
      if (!config) return Promise.resolve(null);
      var uri = redirectUri || (window.location.origin + window.location.pathname) || config.redirectUri;
      var instance = createMsalInstance(uri);
      if (!instance) return Promise.resolve(null);
      function tryHandleRedirect() {
        return instance.initialize().then(function () {
          return instance.handleRedirectPromise();
        }).then(function (result) {
          if (result && result.account) {
            return { account: result.account, instance: instance };
          }
          var accounts = instance.getAllAccounts();
          if (accounts.length > 0) {
            return { account: accounts[0], instance: instance };
          }
          return null;
        });
      }
      return tryHandleRedirect().catch(function () {
        var hash = window.location.hash || '';
        if (hash.indexOf('state=') !== -1 || hash.indexOf('code=') !== -1 || hash.indexOf('access_token=') !== -1) {
          return new Promise(function (resolve) {
            setTimeout(function () {
              tryHandleRedirect().then(resolve).catch(function () { resolve(null); });
            }, 800);
          });
        }
        return null;
      });
    },

    /**
     * Redirect to Microsoft sign-in. Call from login.html; redirectUri should be index.html.
     */
    signInRedirect: function () {
      var config = getMsalConfig();
      if (!config || !config.clientId || config.clientId === 'YOUR_CLIENT_ID') {
        return Promise.reject(new Error('Configure MSAL in config.js (clientId and redirectUri).'));
      }
      var indexPath = (window.location.pathname || '').replace(/login\.html$/i, 'index.html') || '/index.html';
      var redirectUri = config.redirectUri || config.redirectUrl || (window.location.origin + indexPath);
      var instance = createMsalInstance(redirectUri);
      if (!instance) return Promise.reject(new Error('MSAL not configured.'));
      var scope = config.scope || 'https://api.powerplatform.com/.default';
      return instance.initialize().then(function () {
        return instance.loginRedirect({
          scopes: [scope],
          redirectUri: redirectUri
        });
      });
    },

    /**
     * Sign out and redirect to login page.
     */
    signOut: function (msalInstance) {
      if (msalInstance && typeof msalInstance.logoutRedirect === 'function') {
        var config = getMsalConfig();
        var loginUri = (window.location.origin + window.location.pathname).replace(/index\.html$/i, 'login.html');
        msalInstance.logoutRedirect({ postLogoutRedirectUri: loginUri });
      } else {
        window.location.href = 'login.html';
      }
    }
  };
})();
