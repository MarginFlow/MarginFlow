window.FIREFLY_CONFIG = {
  appEnvironment: "production",
  clientId: "YOUR_MICROSOFT_APP_CLIENT_ID",
  tenantId: "common",
  redirectUri: `${window.location.origin}/`,
  graphScopes: ["openid", "profile", "email", "User.Read", "Mail.Read", "Mail.Send"]
};
