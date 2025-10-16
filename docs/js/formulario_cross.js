// js/msal-app.js
export const msalConfig = {
  auth: {
    clientId: "a64f76bf-cb87-4348-8db0-d3959ca6cd54",
    authority: "https://login.microsoftonline.com/f88cba71-d226-4d73-be87-c972ecafc1f5",
    redirectUri: "https://planeamientocontrol-debug.github.io/calidad_formularios/01_crosselling.html",
    // Para pruebas locales, agrega esta URI en Azure y cámbiala temporalmente:
    // redirectUri: "http://localhost:5500/docs/index.html"
  },
  cache: { cacheLocation: "localStorage" }
};

export const graphScopes = [
  "openid", "profile", "email",
  "offline_access",
  "Files.ReadWrite",
  "Files.ReadWrite.All",  // OneDrive personal y compartidos
  "Sites.ReadWrite.All"   // (recomendado) SharePoint sites/links compartidos
];

export const msalInstance = new msal.PublicClientApplication(msalConfig);

msalInstance.handleRedirectPromise()
  .then((resp) => { if (resp?.account) msalInstance.setActiveAccount(resp.account); })
  .catch(console.error);

export async function login() {
  try {
    await msalInstance.loginRedirect({ scopes: graphScopes });
  } catch (err) {
    console.error("Error al iniciar sesión:", err);
  }
}

export async function getToken() {
  let account = msalInstance.getActiveAccount();
  if (!account) {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      await login();
      return null;
    }
    account = accounts[0];
    msalInstance.setActiveAccount(account);
  }

  try {
    const response = await msalInstance.acquireTokenSilent({ account, scopes: graphScopes });
    return response.accessToken;
  } catch (err) {
    console.warn("Token expirado o requiere interacción:", err);
    await msalInstance.acquireTokenRedirect({ scopes: graphScopes });
    return null;
  }
}

export async function debugToken() {
  try {
    const t = await getToken();
    if (!t) {
      console.warn("[DEBUG] Aún no hay token (probable redirect).");
      return null;
    }
    console.log("[DEBUG] Token OK. Scopes:", (JSON.parse(atob(t.split(".")[1]))?.scp || "(sin scp)"));
    console.log("[DEBUG] Token (primeros 60):", t.slice(0, 60) + "...");
    return t;
  } catch (e) {
    console.error("[DEBUG] Error obteniendo token:", e);
    return null;
  }
}

// útil para inspección en dev
window.__msal = { msalInstance, msalConfig };
