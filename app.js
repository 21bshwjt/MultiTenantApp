require("dotenv").config(); // must be first line

const express = require("express");
const session = require("express-session");
const { ConfidentialClientApplication } = require("@azure/msal-node");

const app = express();

app.use(
  session({
    secret: "testsecret",
    resave: false,
    saveUninitialized: false,
  })
);

// ─── MSAL Config ─────────────────────────────────────────────────────────────

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: "https://login.microsoftonline.com/common",
    clientSecret: process.env.CLIENT_SECRET,
  },
};

const cca = new ConfidentialClientApplication(msalConfig);
const REDIRECT_URI = "http://localhost:3000/auth/redirect";

// ─── HTML Helpers ─────────────────────────────────────────────────────────────

const FONTS = `<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">`;

const BASE_CSS = `
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'DM Sans', sans-serif;
    background: #0D1117;
    min-height: 100vh;
    color: #E6EDF3;
  }
  .dot-bg {
    position: fixed; inset: 0; z-index: 0;
    background-image: radial-gradient(circle, rgba(255,255,255,.055) 1px, transparent 1px);
    background-size: 28px 28px;
    pointer-events: none;
  }
  .glow {
    position: fixed; top: -120px; left: 50%; transform: translateX(-50%);
    width: 600px; height: 360px; z-index: 0;
    background: radial-gradient(ellipse, rgba(56,139,253,.14) 0%, transparent 70%);
    pointer-events: none;
  }
  .content { position: relative; z-index: 1; }
</style>`;

function loginPage() {
  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Sign in</title>
  ${FONTS}
  ${BASE_CSS}
  <style>
    .outer {
      display: flex; align-items: center; justify-content: center;
      min-height: 100vh; padding: 40px 20px;
    }
    .card {
      background: #161B22;
      border: 0.5px solid rgba(255,255,255,.1);
      border-radius: 16px;
      padding: 40px;
      width: 100%; max-width: 400px;
    }
    .brand { display: flex; align-items: center; gap: 10px; margin-bottom: 32px; }
    .brand-icon {
      width: 36px; height: 36px; border-radius: 8px;
      background: linear-gradient(135deg, #388BFD, #58A6FF);
      display: flex; align-items: center; justify-content: center;
    }
    .brand-icon svg { width: 18px; height: 18px; fill: none; stroke: white; stroke-width: 2; stroke-linecap: round; stroke-linejoin: round; }
    .brand-name { font-size: 15px; font-weight: 600; color: #E6EDF3; letter-spacing: -.2px; }
    h1 { font-size: 22px; font-weight: 600; color: #E6EDF3; margin-bottom: 6px; letter-spacing: -.3px; }
    .subtitle { font-size: 14px; color: #7D8590; margin-bottom: 28px; line-height: 1.55; }
    .ms-btn {
      display: flex; align-items: center; justify-content: center; gap: 10px;
      width: 100%; padding: 11px 20px;
      background: #1C2128; border: 0.5px solid rgba(255,255,255,.15);
      border-radius: 10px; cursor: pointer;
      font-size: 14px; font-weight: 500; color: #E6EDF3;
      font-family: 'DM Sans', sans-serif;
      text-decoration: none;
      transition: background .15s, border-color .15s;
    }
    .ms-btn:hover { background: #21262D; border-color: rgba(56,139,253,.5); }
    .ms-logo { width: 18px; height: 18px; flex-shrink: 0; }
    .divider { display: flex; align-items: center; gap: 12px; margin: 20px 0; }
    .divider span { font-size: 12px; color: #484F58; white-space: nowrap; }
    .divider-line { flex: 1; height: 0.5px; background: rgba(255,255,255,.08); }
    .meta {
      background: #0D1117; border: 0.5px solid rgba(255,255,255,.06);
      border-radius: 8px; padding: 12px 14px;
      font-size: 12px; color: #484F58; line-height: 1.8;
    }
    .meta strong { color: #7D8590; font-weight: 500; }
    .footer { font-size: 11px; color: #3D444D; text-align: center; margin-top: 20px; line-height: 1.6; }
  </style>
</head>
<body>
  <div class="dot-bg"></div>
  <div class="glow"></div>
  <div class="content outer">
    <div class="card">
      <div class="brand">
        <div class="brand-icon">
          <svg viewBox="0 0 24 24"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>
        </div>
        <span class="brand-name">MyApp</span>
      </div>
      <h1>Welcome back</h1>
      <p class="subtitle">Sign in with your Microsoft account to continue to your dashboard.</p>
      <a href="/login" class="ms-btn">
        <svg class="ms-logo" viewBox="0 0 23 23" xmlns="http://www.w3.org/2000/svg">
          <path fill="#F25022" d="M1 1h10v10H1z"/>
          <path fill="#00A4EF" d="M12 1h10v10H12z"/>
          <path fill="#7FBA00" d="M1 12h10v10H1z"/>
          <path fill="#FFB900" d="M12 12h10v10H12z"/>
        </svg>
        Continue with Microsoft
      </a>
      <div class="divider">
        <div class="divider-line"></div>
        <span>multi-tenant</span>
        <div class="divider-line"></div>
      </div>
      <div class="meta">
        <strong>Authority:</strong> login.microsoftonline.com/common<br>
        <strong>Scopes:</strong> openid &middot; profile &middot; User.Read
      </div>
      <p class="footer">Your credentials are handled securely by Microsoft.<br>We never see your password.</p>
    </div>
  </div>
</body>
</html>`;
}

function successPage(account, idTokenClaims) {
  const initials = account.name
    .split(" ")
    .slice(0, 2)
    .map((n) => n[0].toUpperCase())
    .join("");
  const authTime  = new Date(idTokenClaims.iat * 1000).toLocaleString();
  const expiresAt = new Date(idTokenClaims.exp * 1000).toLocaleString();
  const homeTenant = account.homeAccountId.split(".")[1] ?? "N/A";

  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Authenticated — ${account.name}</title>
  ${FONTS}
  ${BASE_CSS}
  <style>
    .page { max-width: 520px; margin: 0 auto; padding: 32px 20px 48px; }
    header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 24px; }
    .brand { display: flex; align-items: center; gap: 8px; }
    .brand-icon {
      width: 28px; height: 28px; border-radius: 6px;
      background: linear-gradient(135deg, #388BFD, #58A6FF);
      display: flex; align-items: center; justify-content: center;
    }
    .brand-icon svg { width: 13px; height: 13px; fill: none; stroke: white; stroke-width: 2.2; stroke-linecap: round; stroke-linejoin: round; }
    .brand-name { font-size: 13px; font-weight: 500; color: #7D8590; }
    .badge {
      background: rgba(63,185,80,.1); border: 0.5px solid rgba(63,185,80,.25);
      border-radius: 20px; padding: 4px 11px;
      font-size: 11px; color: #3FB950; font-weight: 500;
      display: flex; align-items: center; gap: 5px;
    }
    .badge::before { content: ''; width: 6px; height: 6px; border-radius: 50%; background: #3FB950; display: inline-block; }
    .user-card {
      background: #161B22; border: 0.5px solid rgba(255,255,255,.08);
      border-radius: 12px; padding: 20px;
      display: flex; align-items: center; gap: 14px; margin-bottom: 14px;
    }
    .avatar {
      width: 48px; height: 48px; border-radius: 50%;
      background: rgba(56,139,253,.12); border: 1px solid rgba(56,139,253,.25);
      display: flex; align-items: center; justify-content: center;
      font-size: 15px; font-weight: 600; color: #58A6FF; flex-shrink: 0;
    }
    .user-name { font-size: 15px; font-weight: 600; color: #E6EDF3; letter-spacing: -.2px; }
    .user-email { font-size: 13px; color: #7D8590; margin-top: 2px; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 14px; }
    .cell { background: #161B22; border: 0.5px solid rgba(255,255,255,.08); border-radius: 10px; padding: 14px; }
    .cell-label { font-size: 10px; font-weight: 500; color: #484F58; text-transform: uppercase; letter-spacing: .5px; margin-bottom: 5px; }
    .cell-value { font-size: 12px; color: #7D8590; font-family: 'DM Mono', monospace; word-break: break-all; line-height: 1.5; }
    .cell-value.blue { color: #58A6FF; }
    .token-box { background: #0D1117; border: 0.5px solid rgba(255,255,255,.06); border-radius: 10px; padding: 4px 14px; margin-bottom: 14px; }
    .token-row { display: flex; align-items: center; justify-content: space-between; padding: 10px 0; border-bottom: 0.5px solid rgba(255,255,255,.05); }
    .token-row:last-child { border-bottom: none; }
    .token-key { font-size: 12px; color: #484F58; }
    .token-val { font-size: 12px; color: #7D8590; font-family: 'DM Mono', monospace; text-align: right; max-width: 60%; word-break: break-all; }
    .sign-out {
      display: block; width: 100%; padding: 10px;
      background: transparent; border: 0.5px solid rgba(255,255,255,.1);
      border-radius: 8px; color: #7D8590; font-size: 13px; cursor: pointer;
      font-family: 'DM Sans', sans-serif; text-decoration: none;
      text-align: center; transition: background .15s, color .15s;
    }
    .sign-out:hover { background: #161B22; color: #E6EDF3; }
  </style>
</head>
<body>
  <div class="dot-bg"></div>
  <div class="glow"></div>
  <div class="content page">
    <header>
      <div class="brand">
        <div class="brand-icon">
          <svg viewBox="0 0 24 24"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>
        </div>
        <span class="brand-name">MyApp</span>
      </div>
      <div class="badge">Authenticated</div>
    </header>

    <div class="user-card">
      <div class="avatar">${initials}</div>
      <div>
        <div class="user-name">${account.name}</div>
        <div class="user-email">${account.username}</div>
      </div>
    </div>

    <div class="grid">
      <div class="cell">
        <div class="cell-label">Tenant ID</div>
        <div class="cell-value">${account.tenantId}</div>
      </div>
      <div class="cell">
        <div class="cell-label">Object ID</div>
        <div class="cell-value">${account.localAccountId}</div>
      </div>
      <div class="cell">
        <div class="cell-label">Environment</div>
        <div class="cell-value blue">${account.environment}</div>
      </div>
      <div class="cell">
        <div class="cell-label">Home Tenant</div>
        <div class="cell-value blue">${homeTenant}</div>
      </div>
    </div>

    <div class="token-box">
      <div class="token-row">
        <span class="token-key">Issuer</span>
        <span class="token-val">${idTokenClaims.iss}</span>
      </div>
      <div class="token-row">
        <span class="token-key">Auth time</span>
        <span class="token-val">${authTime}</span>
      </div>
      <div class="token-row">
        <span class="token-key">Expires at</span>
        <span class="token-val">${expiresAt}</span>
      </div>
      <div class="token-row">
        <span class="token-key">Nonce</span>
        <span class="token-val">${idTokenClaims.nonce ?? "N/A"}</span>
      </div>
    </div>

    <a href="/" class="sign-out">Sign out</a>
  </div>
</body>
</html>`;
}

// ─── Routes ───────────────────────────────────────────────────────────────────

app.get("/", (req, res) => {
  res.send(loginPage());
});

app.get("/login", async (req, res) => {
  try {
    const authCodeUrlParameters = {
      scopes: ["openid", "profile", "User.Read"],
      redirectUri: REDIRECT_URI,
    };
    const authUrl = await cca.getAuthCodeUrl(authCodeUrlParameters);
    res.redirect(authUrl);
  } catch (err) {
    console.error(err);
    res.status(500).send(`<pre>${JSON.stringify(err, null, 2)}</pre>`);
  }
});

app.get("/auth/redirect", async (req, res) => {
  try {
    const tokenRequest = {
      code: req.query.code,
      scopes: ["openid", "profile", "User.Read"],
      redirectUri: REDIRECT_URI,
    };
    const response = await cca.acquireTokenByCode(tokenRequest);
    res.send(successPage(response.account, response.idTokenClaims));
  } catch (err) {
    console.error(err);
    res.status(500).send(`<pre>${JSON.stringify(err, null, 2)}</pre>`);
  }
});

// ─── Start ────────────────────────────────────────────────────────────────────

app.listen(3000, () => {
  console.log("Running on http://localhost:3000");
});