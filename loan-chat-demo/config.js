/**
 * Configuration for Loan Support Assistant
 * - Copilot Studio Conversations API (authenticated agent)
 * - Microsoft Entra ID (Azure AD) for sign-in and agent token
 *
 * CONVERSATIONS_API_BASE: From Copilot Studio → Publish → Channels → connection string.
 * Use the base URL (no /conversations, no ?api-version=...).
 */
var LOAN_CHAT_CONFIG = {
  CONVERSATIONS_API_BASE:'https://default5d41fd7cb2914130ac2b9170e1c4c0.3e.environment.api.powerplatform.com/copilotstudio/dataverse-backed/authenticated/bots/auto_agent_UqpTR/conversations?api-version=2022-03-01-preview',
    API_VERSION: '2022-03-01-preview',

  /**
   * Microsoft Entra ID (Azure AD) app registration for MSAL.js.
   * redirectUri must be the full URL of index.html (where the user lands after Microsoft sign-in).
   */
  MSAL: {
    clientId: '80b8b3f5-45dd-4648-9d00-1d4a570221c1',
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: 'http://localhost:8080/loan-chat-demo/index.html',
    scope: 'https://api.powerplatform.com/.default'
  }
};
