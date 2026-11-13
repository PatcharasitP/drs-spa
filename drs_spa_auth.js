// drs_spa_auth.js ทาง J ของ DR (13/11/2026 ฟ้า) หน้า DR เปิดจากที่อยู่นอก SharePoint แล้วอ่านเขียน list ด้วยสิทธิ์ของผู้ใช้เอง
// ล็อกอินด้วย MSAL (แอปที่ลงทะเบียนในเทนแนนต์ ขอสิทธิ์ SharePoint AllSites.Write แบบผู้ใช้ยินยอมเอง พิสูจน์ 11/11/2026)
// ได้ token แล้วค่อยโหลดสคริปต์ของหน้า (แท็ก <script type="text/drs-wait" data-src>) ตามลำดับ หน้าจึงไม่เริ่มก่อนล็อกอินเสร็จ
// token แนบเฉพาะคำขอที่ไปโดเมน SharePoint ของไซต์เท่านั้น
(function (root) {
  "use strict";

  function makeAuthFetch(getToken, baseFetch, siteUrl) {
    const origin = new URL(siteUrl).origin;
    return async function (input, init) {
      const url = typeof input === "string" ? input : input && input.url ? input.url : String(input);
      let same = false;
      try { same = new URL(url, siteUrl).origin === origin; } catch (e) { same = false; }
      if (!same) return baseFetch(input, init);
      const send = async (force) => {
        const tok = await getToken(force);
        const o = Object.assign({}, init || {});
        o.headers = Object.assign({}, (init && init.headers) || {}, { Authorization: "Bearer " + tok });
        o.credentials = "omit";
        return baseFetch(url, o);
      };
      const r = await send(false);
      return r && r.status === 401 ? send(true) : r;
    };
  }

  const api = { makeAuthFetch };
  if (typeof module === "object" && module.exports) { module.exports = api; return; }
  root.DRS = root.DRS || {}; root.DRS.spaAuth = api;

  const d = root.document, el = d && d.documentElement;
  const client = el && el.getAttribute("data-drs-auth-client");
  if (!client) return;
  const site = el.getAttribute("data-drs-site");
  const tenant = el.getAttribute("data-drs-auth-tenant") || "organizations";
  const scopes = [new URL(site).origin + "/AllSites.Write"];
  const say = (t) => { const b = d.getElementById("drs-auth-msg"); if (b) b.textContent = t; };
  const pca = new root.msal.PublicClientApplication({
    auth: { clientId: client, authority: "https://login.microsoftonline.com/" + tenant, redirectUri: root.location.origin + root.location.pathname.replace(/index\.html$/i, "") },
    cache: { cacheLocation: "sessionStorage" }, // token หายเมื่อปิดแท็บ (ความเสี่ยงข้อ 1 ของพี่ปริม 11/11) ล็อกอินใหม่เงียบด้วย SSO ของ Microsoft
  });
  let account = null;
  async function getToken(force) {
    try {
      return (await pca.acquireTokenSilent({ scopes, account, forceRefresh: !!force })).accessToken;
    } catch (e) {
      say("ต้องล็อกอินใหม่ กำลังพาไปหน้าล็อกอิน");
      await pca.acquireTokenRedirect({ scopes, account });
      return new Promise(() => {});
    }
  }
  function loadPageScripts() {
    const wait = Array.from(d.querySelectorAll('script[type="text/drs-wait"]'));
    for (const s of wait) {
      const n = d.createElement("script");
      n.src = s.getAttribute("data-src"); n.async = false;
      s.parentNode.insertBefore(n, s.nextSibling);
    }
  }
  (async () => {
    try {
      await pca.initialize();
      const red = await pca.handleRedirectPromise();
      account = (red && red.account) || pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
      if (!account) { say("กำลังพาไปหน้าล็อกอิน Microsoft"); await pca.loginRedirect({ scopes }); return; }
      pca.setActiveAccount(account);
      await getToken(false);
      root.DRS_AUTH_FETCH = makeAuthFetch(getToken, root.fetch.bind(root), site);
      root.DRS_AUTH_ACCOUNT = { name: account.name || "", user: account.username || "" };
      const b = d.getElementById("drs-auth-msg"); if (b) b.remove();
      loadPageScripts();
    } catch (e) {
      say("ล็อกอินไม่สำเร็จ " + (e && (e.errorCode || e.message) || e));
      try { console.error("[DRS-SPA]", e); } catch (x) {}
    }
  })();
})(typeof globalThis !== "undefined" ? globalThis : this);
