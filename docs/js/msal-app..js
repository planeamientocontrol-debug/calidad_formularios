export const msalConfig = {
  auth: {
    clientId: "168fa127-022d-4535-9acd-cf17a64cee20", // tu Id. de aplicación
    authority: "https://login.microsoftonline.com/f88cba71-d226-4d73-be87-c972ecafc1f5", // 👈 URL completa del inquilino
    redirectUri: "https://planeamientocontrol-debug.github.io/calidad_formularios/01_crosselling.html" // 👈 tu URL de producción (GitHub Pages)
    // redirectUri: "http://localhost:5500/docs/index.html" // 👈 para pruebas locales
  },
  cache: {
    cacheLocation: "localStorage", // mantiene la sesión activa
  }
};

// Scopes o permisos necesarios
export const graphScopes = [
  "openid", "profile", "email",
  "offline_access",
  "Files.ReadWrite",  // permite leer y escribir archivos propios
  "Files.ReadWrite.All", // opcional, si necesitas escribir en compartidos
];

// Crear la instancia principal de MSAL
export const msalInstance = new msal.PublicClientApplication(msalConfig);

// Manejo del redireccionamiento
msalInstance.handleRedirectPromise().then((resp) => {
  if (resp?.account) msalInstance.setActiveAccount(resp.account);
}).catch(console.error);

// Función para iniciar sesión
export async function login() {
  try {
    await msalInstance.loginRedirect({ scopes: graphScopes });
  } catch (err) {
    console.error("Error al iniciar sesión:", err);
  }
}

// Función para obtener el token
export async function getToken() {
  let account = msalInstance.getActiveAccount();
  if (!account) {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      await login(); // redirige si no hay sesión
      return null;
    }
    account = accounts[0];
    msalInstance.setActiveAccount(account);
  }

  try {
    const response = await msalInstance.acquireTokenSilent({
      account,
      scopes: graphScopes
    });
    return response.accessToken;
  } catch (err) {
    console.warn("Token expirado o requiere interacción:", err);
    await msalInstance.acquireTokenRedirect({ scopes: graphScopes });
    return null;
  }
}

// Vincula el botón de login
document.getElementById("btnLogin")?.addEventListener("click", (e) => {
  e.preventDefault();
  login();
});

// Debug
console.log("MSAL inicializado con redirectUri:", msalConfig.auth.redirectUri);

// al final de js/msal-app.js
window.__msal = { msalInstance, msalConfig }; // para inspeccionar desde consola

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
