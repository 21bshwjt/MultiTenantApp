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
  :root {
    --bg: #0D1117;
    --text: #E6EDF3;
    --card-bg: #161B22;
    --muted: #7D8590;
    --subdued: #484F58;
    --accent: #58A6FF;
    --accent-weak: rgba(56,139,253,.14);
    --border-weak: rgba(255,255,255,.08);
    --border-med: rgba(255,255,255,.1);
    --positive: #3FB950;
    --muted-2: #3D444D;
  }

  /* Theme overrides */
  .theme-light {
    --bg: #F6F8FA; --text: #0D1117; --card-bg: #FFFFFF; --muted: #5A6470; --subdued: #6B7280; --accent: #2563EB; --accent-weak: rgba(37,99,235,.08); --border-weak: rgba(13,17,23,.06);
  }
  .theme-solar {
    --bg: #FFFBF0; --text: #2B2B2B; --card-bg: #FFF6E8; --muted: #7A5A2A; --subdued: #8B6A3E; --accent: #D97706; --accent-weak: rgba(217,119,6,.08); --border-weak: rgba(43,43,43,.06);
  }
  .theme-forest {
    --bg: #0F1F14; --text: #E6F7E9; --card-bg: #0B1A10; --muted: #9CC4A8; --subdued: #7BA887; --accent: #38A169; --accent-weak: rgba(56,161,105,.08); --border-weak: rgba(230,247,233,.04);
  }

  body {
    font-family: 'DM Sans', sans-serif;
    background: var(--bg);
    min-height: 100vh;
    color: var(--text);
  }
  .dot-bg {
    position: fixed; inset: 0; z-index: 0;
    background-image: radial-gradient(circle, rgba(255,255,255,.035) 1px, transparent 1px);
    background-size: 28px 28px;
    pointer-events: none;
    opacity: 0.9;
  }
  .glow {
    position: fixed; top: -120px; left: 50%; transform: translateX(-50%);
    width: 600px; height: 360px; z-index: 0;
    background: radial-gradient(ellipse, var(--accent-weak) 0%, transparent 70%);
    pointer-events: none;
  }
  .content { position: relative; z-index: 1; }

  /* Common helpers use variables in page-specific styles */
</style>`;

const THEME_SCRIPT = `
<script>
  (function(){
    const themeColors = {
      default: getComputedStyle(document.documentElement).getPropertyValue('--accent').trim() || '#58A6FF',
      'theme-light': '#2563EB',
      'theme-solar': '#D97706',
      'theme-forest': '#38A169'
    };

    function applyTheme(name){
      document.documentElement.classList.remove('theme-light','theme-solar','theme-forest');
      if(name && name !== 'default') document.documentElement.classList.add(name);
    }

    function updateButtonColor(name){
      const btn = document.getElementById('theme-button');
      if(!btn) return;
      const c = themeColors[name] || themeColors.default;
      btn.style.background = c;
      btn.setAttribute('data-theme', name);
    }

    const stored = localStorage.getItem('theme') || 'default';
    applyTheme(stored === 'default' ? null : stored);

    window.setAppTheme = function(name){
      localStorage.setItem('theme', name);
      applyTheme(name === 'default' ? null : name);
      updateButtonColor(name);
      const sel = document.getElementById('theme-select'); if(sel) sel.value = name;
    };

    document.addEventListener('DOMContentLoaded', function(){
      const sel = document.getElementById('theme-select');
      if(sel){ sel.value = localStorage.getItem('theme') || 'default';
        sel.addEventListener('change', function(){ window.setAppTheme(this.value); });
      }

      const btn = document.getElementById('theme-button');
      const pop = document.getElementById('theme-popover');
      if(btn && pop){
        // populate option colors
        Array.from(pop.querySelectorAll('.theme-option')).forEach(function(o){
          const t = o.getAttribute('data-theme');
          o.style.background = themeColors[t] || themeColors.default;
          o.addEventListener('click', function(e){ e.stopPropagation(); window.setAppTheme(t); pop.style.display = 'none'; });
        });

        updateButtonColor(localStorage.getItem('theme') || 'default');

        btn.addEventListener('click', function(e){ e.stopPropagation(); pop.style.display = pop.style.display === 'flex' ? 'none' : 'flex'; });

        // close on outside click
        document.addEventListener('click', function(){ if(pop) pop.style.display = 'none'; });
        pop.addEventListener('click', function(e){ e.stopPropagation(); });
      }
    });
  })();
</script>`;

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
      background: var(--card-bg);
      border: 0.5px solid var(--border-med);
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
    .brand-name { font-size: 15px; font-weight: 600; color: var(--text); letter-spacing: -.2px; }
    h1 { font-size: 22px; font-weight: 600; color: var(--text); margin-bottom: 6px; letter-spacing: -.3px; }
    .subtitle { font-size: 14px; color: var(--muted); margin-bottom: 28px; line-height: 1.55; }
    .ms-btn {
      display: flex; align-items: center; justify-content: center; gap: 10px;
      width: 100%; padding: 11px 20px;
      background: var(--card-bg);
      border: 0.5px solid var(--border-med);
      border-radius: 10px; cursor: pointer;
      font-size: 14px; font-weight: 500; color: #E6EDF3;
      font-family: 'DM Sans', sans-serif;
      text-decoration: none;
      transition: background .15s, border-color .15s;
    }
    .ms-btn:hover { background: rgba(255,255,255,.02); border-color: rgba(56,139,253,.5); }
    .ms-logo { width: 18px; height: 18px; flex-shrink: 0; }
    .divider { display: flex; align-items: center; gap: 12px; margin: 20px 0; }
    .divider span { font-size: 12px; color: var(--subdued); white-space: nowrap; }
    .divider-line { flex: 1; height: 0.5px; background: var(--border-weak); }
    .meta {
      background: var(--bg); border: 0.5px solid var(--border-weak);
      border-radius: 8px; padding: 12px 14px;
      font-size: 12px; color: var(--subdued); line-height: 1.8;
    }
    .meta strong { color: var(--muted); font-weight: 500; }
    .footer { font-size: 11px; color: var(--muted-2); text-align: center; margin-top: 20px; line-height: 1.6; }
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
  ${THEME_SCRIPT}
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
      background: linear-gradient(135deg, #388BFD, var(--accent));
      display: flex; align-items: center; justify-content: center;
    }
    .brand-icon svg { width: 13px; height: 13px; fill: none; stroke: white; stroke-width: 2.2; stroke-linecap: round; stroke-linejoin: round; }
    .brand-name { font-size: 13px; font-weight: 500; color: var(--subdued); }
    .badge {
      background: rgba(63,185,80,.08); border: 0.5px solid rgba(63,185,80,.18);
      border-radius: 20px; padding: 4px 11px;
      font-size: 11px; color: var(--positive); font-weight: 500;
      display: flex; align-items: center; gap: 5px;
    }
    .badge::before { content: ''; width: 6px; height: 6px; border-radius: 50%; background: #3FB950; display: inline-block; }
    .user-card {
      background: var(--card-bg); border: 0.5px solid var(--border-weak);
      border-radius: 12px; padding: 20px;
      display: flex; align-items: center; gap: 14px; margin-bottom: 14px;
    }
    .avatar {
      width: 48px; height: 48px; border-radius: 50%;
      background: rgba(56,139,253,.08); border: 1px solid rgba(56,139,253,.18);
      display: flex; align-items: center; justify-content: center;
      font-size: 15px; font-weight: 600; color: #58A6FF; flex-shrink: 0;
    }
    .user-name { font-size: 15px; font-weight: 600; color: var(--text); letter-spacing: -.2px; }
    .user-email { font-size: 13px; color: var(--muted); margin-top: 2px; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 14px; }
    .cell { background: #161B22; border: 0.5px solid rgba(255,255,255,.08); border-radius: 10px; padding: 14px; }
    .cell { background: var(--card-bg); border: 0.5px solid var(--border-weak); border-radius: 10px; padding: 14px; }
    .cell-label { font-size: 10px; font-weight: 500; color: var(--subdued); text-transform: uppercase; letter-spacing: .5px; margin-bottom: 5px; }
    .cell-value { font-size: 12px; color: var(--muted); font-family: 'DM Mono', monospace; word-break: break-all; line-height: 1.5; }
    .cell-value.blue { color: var(--accent); }
    .token-box { background: #0D1117; border: 0.5px solid rgba(255,255,255,.06); border-radius: 10px; padding: 4px 14px; margin-bottom: 14px; }
    .token-box { background: var(--bg); border: 0.5px solid var(--border-weak); border-radius: 10px; padding: 4px 14px; margin-bottom: 14px; }
    .token-row { display: flex; align-items: center; justify-content: space-between; padding: 10px 0; border-bottom: 0.5px solid rgba(255,255,255,.03); }
    .token-row:last-child { border-bottom: none; }
    .token-key { font-size: 12px; color: var(--subdued); }
    .token-val { font-size: 12px; color: var(--muted); font-family: 'DM Mono', monospace; text-align: right; max-width: 60%; word-break: break-all; }
    .sign-out {
      display: block; width: 100%; padding: 10px;
      background: transparent; border: 0.5px solid var(--border-med);
      border-radius: 8px; color: var(--muted); font-size: 13px; cursor: pointer;
      font-family: 'DM Sans', sans-serif; text-decoration: none;
      text-align: center; transition: background .15s, color .15s;
    }
    .sign-out:hover { background: var(--card-bg); color: var(--text); }
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
      <div style="display:flex;align-items:center;gap:12px;position:relative">
        <div class="badge">Authenticated</div>
        <div id="theme-widget" style="position:relative">
          <button id="theme-button" aria-label="Theme" style="width:36px;height:36px;border-radius:50%;border:0.5px solid var(--border-weak);background:var(--accent);display:flex;align-items:center;justify-content:center;cursor:pointer;padding:0"></button>
          <div id="theme-popover" style="position:absolute;right:0;top:44px;background:var(--card-bg);border:0.5px solid var(--border-weak);padding:8px;border-radius:10px;display:none;gap:8px;box-shadow:0 6px 18px rgba(0,0,0,.12);flex-direction:row;">
            <button class="theme-option" data-theme="default" title="Default" style="width:28px;height:28px;border-radius:50%;border:1px solid rgba(0,0,0,.06);cursor:pointer"></button>
            <button class="theme-option" data-theme="theme-light" title="Light" style="width:28px;height:28px;border-radius:50%;border:1px solid rgba(0,0,0,.06);cursor:pointer"></button>
            <button class="theme-option" data-theme="theme-solar" title="Solar" style="width:28px;height:28px;border-radius:50%;border:1px solid rgba(0,0,0,.06);cursor:pointer"></button>
            <button class="theme-option" data-theme="theme-forest" title="Forest" style="width:28px;height:28px;border-radius:50%;border:1px solid rgba(0,0,0,.06);cursor:pointer"></button>
          </div>
        </div>
      </div>
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

  </div>
    <a href="/" class="sign-out">Sign out</a>
  </div>
  ${THEME_SCRIPT}
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