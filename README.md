# 🔐 M365 Conditional Access Baseline

A PowerShell script to deploy a baseline set of Conditional Access policies for Microsoft 365 / Entra ID. Built and maintained by **Christopher Johnston** (christopher.johnston@it1.com).

---

## 📋 Overview

This repository contains PowerShell scripts to build and manage a Conditional Access baseline for Microsoft 365 tenants. Policies are deployed as **Disabled** so you can review Sign-in Logs before enabling them in production.

### Scripts

| Script | Description |
|---|---|
| `Build-ConditionalAccessBaseline.ps1` | Deploys the full CA baseline (security group, named location, policies) |
| `Remove-AllConditionalAccessPolicies.ps1` | Removes all CA policies and named locations from the tenant |

---

## 🛡️ What Gets Deployed

### Resources Created
- **Security Group** — `CA-Exclusions` (empty break-glass group, excluded from all policies)
- **Named Location** — `Allowed-Country-USA` (United States only, unknown countries blocked)

### Policies (All created as Disabled)

| # | Policy Name | License | Description |
|---|---|---|---|
| 1 | BASELINE - Blocking Legacy Authentication | P1 | Blocks Exchange ActiveSync and all other legacy auth protocols |
| 2 | BASELINE - GEO Block (Only USA Allowed) | P1 | Blocks sign-ins from all countries except the United States |
| 3 | BASELINE - MFA All Users | P1 | Requires MFA for all users across all applications |
| 4 | BASELINE - Only Allow Windows, Mac, iOS, Android | P1 | Blocks sign-ins from unsupported device platforms |
| 5 | BASELINE - Phishing-Resistant MFA for Admin Portals | P1 | Requires phishing-resistant MFA for all admin roles on Microsoft admin portals |
| 6 | BASELINE - Risky Users Risk - High (Block) | **P2** | Blocks sign-in for users with high identity risk score |
| 7 | BASELINE - Risky Users Risk Med-Low (Enforce MFA) | **P2** | Requires MFA + password change for medium/low risk users |

> **Note:** Policies 6 and 7 require **Entra ID P2** (or Microsoft 365 E5). The script automatically detects your license and skips these policies on P1 tenants.

---

## ✅ Prerequisites

### Licensing
- **Entra ID P1** — minimum for policies 1–5
- **Entra ID P2** or **Microsoft 365 E5** — required for policies 6–7 (Risky Users)

### PowerShell Module
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

### Required Graph API Permissions (Delegated)
The script will prompt for consent on first run:

| Permission | Purpose |
|---|---|
| `Policy.Read.All` | Read existing CA policies |
| `Policy.ReadWrite.ConditionalAccess` | Create and modify CA policies |
| `Policy.ReadWrite.AuthenticationMethod` | Read and modify Security Defaults |
| `Group.ReadWrite.All` | Create the CA-Exclusions security group |
| `Directory.Read.All` | Read tenant directory information |
| `Application.Read.All` | Required for policies targeting specific applications |

### Important — Security Defaults
Security Defaults and Conditional Access **cannot run together**. The script checks for this automatically and will prompt you to disable Security Defaults before proceeding.

---

## 🚀 Usage

### Deploy the Baseline

```powershell
# 1. Clone the repository
git clone https://github.com/IT1-CJ/Conditional-Access-Baselines.git
cd Conditional-Access-Baselines

# 2. Run the deployment script
pwsh ./Build-ConditionalAccessBaseline.ps1
```

### Remove All Policies (Clean Slate)

```powershell
pwsh ./Remove-AllConditionalAccessPolicies.ps1
```

> ⚠️ **Warning:** This deletes ALL Conditional Access policies and named locations in your tenant. Use with caution.

---

## 📋 Recommended Enable Order

After deploying, monitor **Sign-in Logs** in Entra ID for **7–14 days** before enabling any policy. Enable in this order to minimize risk of lockout:

```
1. Blocking Legacy Authentication      ← Safest, enable first
2. MFA All Users
3. Only Allow Windows, Mac, iOS, Android
4. Phishing-Resistant MFA for Admin Portals
5. Risky Users - High (Block)          ← P2 only
6. Risky Users Med-Low (Enforce MFA)   ← P2 only
7. GEO Block (Only USA Allowed)        ← Enable last, highest lockout risk
```

### How to Enable a Policy
1. Go to [Entra ID → Conditional Access → Policies](https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade)
2. Click the policy name
3. Change **State** from `Off` → `Report-only` (monitor for 7–14 days)
4. Change **State** from `Report-only` → `On`

---

## 🔒 Break-Glass Accounts

The `CA-Exclusions` security group is excluded from **every policy**. Before enabling any policy:

1. Create at least **2 break-glass admin accounts**
2. Add them to the `CA-Exclusions` group in Entra ID
3. Store credentials securely offline
4. Test them periodically to ensure access

> ⚠️ Never add your day-to-day admin account to CA-Exclusions permanently.

---

## 📁 Repository Structure

```
conditional-access-baseline/
│
├── README.md
├── Build-ConditionalAccessBaseline.ps1
└── Remove-AllConditionalAccessPolicies.ps1
```

---

## 🔄 Changelog

| Version | Date | Changes |
|---|---|---|
| 2.3 | 2026-05-06 | Added P1/P2 license detection, Risky Users policies auto-created on P2 |
| 2.2 | 2026-05-06 | Removed PATCH step, added named location replication wait, fixed app IDs |
| 2.1 | 2026-05-06 | Added Application.Read.All scope, removed P2 policies for P1 tenants |
| 2.0 | 2026-05-06 | Removed description from POST body (Graph API rejects it) |
| 1.9 | 2026-05-06 | All variables use script scope to fix scoping issues |
| 1.8 | 2026-05-06 | Fixed Graph API schema — added required excludeUsers:[] field |

---

## ⚠️ Disclaimer

These scripts are provided as-is. Always test in a non-production tenant first. Review and understand each policy before enabling in production. The author is not responsible for any lockouts or access issues.

---

## 👤 Author

**Christopher Johnston**
christopher.johnston@it1.com
