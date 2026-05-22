<p align="center">
  <img src="https://img.shields.io/badge/Platform-Microsoft%20Azure-0078D4?style=flat-square&logo=microsoftazure&logoColor=white"/>
  <img src="https://img.shields.io/badge/Language-JavaScript-F7DF1E?style=flat-square&logo=javascript&logoColor=black"/>
  <img src="https://img.shields.io/badge/Runtime-Node.js-339933?style=flat-square&logo=nodedotjs&logoColor=white"/>
  <img src="https://img.shields.io/badge/Framework-Express-000000?style=flat-square&logo=express&logoColor=white"/>
  <img src="https://img.shields.io/badge/Auth-MSAL%20Node%20v5-0078D4?style=flat-square&logo=microsoft&logoColor=white"/>
  <img src="https://img.shields.io/badge/Identity-Microsoft%20Entra%20ID-00A4EF?style=flat-square&logo=microsoftazure&logoColor=white"/>
  <img src="https://img.shields.io/badge/License-ISC-green?style=flat-square"/>
</p>

<h1 align="center">🔐 Entra ID - Multi Tenant App</h1>
<p align="center">
  A Node.js + Express application demonstrating <strong>Microsoft Entra ID Multi-Tenant Authentication</strong> using MSAL Node.
</p>

---

## 📋 Overview

**MultiTenantApp** is a reference implementation that showcases how to integrate **Microsoft Entra ID (formerly Azure AD)** multi-tenant OAuth 2.0 / OpenID Connect authentication into a Node.js web application. It uses the official Microsoft Authentication Library (`@azure/msal-node`) to handle token acquisition and session management across multiple tenants.

---

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────┐
│                    Browser / Client                  │
└────────────────────────┬────────────────────────────┘
                         │ HTTP
                         ▼
┌─────────────────────────────────────────────────────┐
│              Express Web Server (app.js)             │
│                                                     │
│  ┌──────────────────┐    ┌───────────────────────┐  │
│  │   Routes / Pages  │    │   Session Management  │  │
│  │  (login, callback │    │  (express-session)    │  │
│  │   , home, logout) │    └───────────────────────┘  │
│  └────────┬─────────┘                               │
│           │                                         │
│  ┌────────▼──────────────────────────────────────┐  │
│  │           MSAL Node  (@azure/msal-node)        │  │
│  │   ConfidentialClientApplication               │  │
│  └────────────────────┬──────────────────────────┘  │
└───────────────────────┼─────────────────────────────┘
                        │ HTTPS / OAuth 2.0
                        ▼
┌─────────────────────────────────────────────────────┐
│           Microsoft Entra ID  (Multi-Tenant)         │
│          Authorization & Token Endpoint              │
└─────────────────────────────────────────────────────┘
```

---

## ✨ Features

- ✅ **Multi-Tenant Support** — Accept sign-ins from any Microsoft Entra ID tenant
- ✅ **OAuth 2.0 Authorization Code Flow** — Secure, standards-compliant auth
- ✅ **MSAL Node v5** — Latest Microsoft Authentication Library for Node.js
- ✅ **Session Management** — Persistent login sessions with `express-session`
- ✅ **Environment-based Config** — Credentials managed via `.env` / `dotenv`
- ✅ **Express 5** — Built on the latest Express framework

---

## 🛠️ Tech Stack

| Component | Technology |
|-----------|-----------|
| Runtime | Node.js |
| Web Framework | Express 5 |
| Authentication | @azure/msal-node v5 |
| Identity Provider | Microsoft Entra ID |
| Session | express-session |
| Config | dotenv |

---

## 🚀 Getting Started

### Prerequisites

- **Node.js** v18 or later
- A **Microsoft Entra ID** (Azure AD) App Registration with:
  - Multi-tenant support enabled (`Accounts in any organizational directory`)
  - A redirect URI set (e.g. `http://localhost:3000/redirect`)
  - A client secret generated

### 1. Clone the Repository

```bash
git clone https://github.com/21bshwjt/MultiTenantApp.git
cd MultiTenantApp
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Configure Environment Variables

Create a `.env` file (or update the existing one) with your Entra app registration details:

```env
CLIENT_ID=<your-application-client-id>
CLIENT_SECRET=<your-client-secret>
TENANT_ID=common
REDIRECT_URI=http://localhost:3000/redirect
POST_LOGOUT_REDIRECT_URI=http://localhost:3000
SESSION_SECRET=<a-strong-random-secret>
```

> ⚠️ **Never commit your `.env` file to source control.** Add it to `.gitignore`.

### 4. Run the Application

```bash
node app.js
```

Visit `http://localhost:3000` in your browser.

---

## 🔑 Authentication Flow

```
User visits /login
      │
      ▼
MSAL generates Auth URL  ──►  Microsoft Entra ID login page
                                        │
                              User authenticates
                                        │
                                        ▼
                          Entra redirects to /redirect
                          with authorization code
                                        │
                                        ▼
                    MSAL exchanges code for tokens
                                        │
                                        ▼
                     Token cached in session storage
                                        │
                                        ▼
                          User redirected to home page
```

---

## 📁 Project Structure

```
MultiTenantApp/
├── app.js               # Main Express application & MSAL configuration
├── .env                 # Environment variables (not committed)
├── package.json         # Project metadata & dependencies
├── package-lock.json    # Dependency lock file
└── node_modules/        # Installed packages
```

---

## 📦 Dependencies

```json
{
  "@azure/msal-node": "^5.2.2",
  "dotenv": "^17.4.2",
  "express": "^5.2.1",
  "express-session": "^1.19.0"
}
```

---

## 🔒 Security Notes

- Store `CLIENT_SECRET` and `SESSION_SECRET` securely — use **Azure Key Vault** or a secrets manager in production
- Use **HTTPS** in production environments
- Set `secure: true` and `sameSite: 'strict'` on session cookies in production
- Validate the `tid` (tenant ID) claim in tokens if you want to restrict access to specific tenants

---

## 📚 References

- [Microsoft Entra ID Documentation](https://learn.microsoft.com/en-us/entra/identity/)
- [MSAL Node on GitHub](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-node)
- [Multi-tenant app registration guide](https://learn.microsoft.com/en-us/entra/identity-platform/howto-convert-app-to-be-multi-tenant)
- [OAuth 2.0 Authorization Code Flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)

---

## 📄 License

This project is licensed under the **ISC License**.

