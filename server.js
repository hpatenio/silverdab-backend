const express = require("express");
const ldap = require("ldapjs");
const jwt = require("jsonwebtoken");
const cors = require("cors");
require("dotenv").config();

const app = express();
app.use(express.json());
app.use(cors());

// ─── Config ────────────────────────────────────────────────────────────────────
const AD_URL = process.env.AD_URL;
const AD_BASE_DN = process.env.AD_BASE_DN;
const AD_DOMAIN = process.env.AD_DOMAIN;
const AD_SERVICE_USER = process.env.AD_SERVICE_USER;
const AD_SERVICE_PASS = process.env.AD_SERVICE_PASS;
const JWT_SECRET = process.env.JWT_SECRET;
const PORT = process.env.PORT || 3000;

console.log("=== Backend Config ===");
console.log("AD_URL:", AD_URL);
console.log("AD_BASE_DN:", AD_BASE_DN);
console.log("AD_DOMAIN:", AD_DOMAIN);
console.log("AD_SERVICE_USER:", AD_SERVICE_USER);
console.log("AD_SERVICE_PASS:", AD_SERVICE_PASS ? "✅ set" : "❌ missing");
console.log("======================");

// ─── Role Mapping from OCGBIM AD Groups ───────────────────────────────────────
function getRoleFromGroups(memberOf) {
  if (!memberOf) return "employee";
  const groups = Array.isArray(memberOf) ? memberOf : [memberOf];
  const g = groups.map((x) => x.toLowerCase());

  if (g.some((x) => x.includes("ocgbim_it_users") || x.includes("it admin") || x.includes("it_admin") || x.includes("it installer")))
    return "it";
  if (g.some((x) => x.includes("ocgbim_adminstaff_users") || x.includes("ocgbim_ceo_users") || x.includes("ocgbim_local_administrator")))
    return "admin";
  return "employee";
}

// ─── Department Mapping from OCGBIM AD Groups ─────────────────────────────────
function getDepartmentFromGroups(memberOf) {
  if (!memberOf) return "General";
  const groups = Array.isArray(memberOf) ? memberOf : [memberOf];
  const g = groups.map((x) => x.toLowerCase());

  if (g.some((x) => x.includes("ocgbim_it_users"))) return "IT";
  if (g.some((x) => x.includes("ocgbim_adminstaff_users"))) return "Admin Staff";
  if (g.some((x) => x.includes("ocgbim_ceo_users"))) return "CEO";
  if (g.some((x) => x.includes("ocgbim_accounting_users"))) return "Accounting";
  if (g.some((x) => x.includes("ocgbim_nscr_users"))) return "NSCR";
  if (g.some((x) => x.includes("ocgbim_production_users"))) return "Production";
  return "General";
}

// ─── Create LDAP Client ────────────────────────────────────────────────────────
function createLDAPClient() {
  return ldap.createClient({
    url: AD_URL,
    timeout: 5000,
    connectTimeout: 10000,
    tlsOptions: { rejectUnauthorized: false },
  });
}

// ─── Get Service Client (tries multiple formats) ───────────────────────────────
async function getServiceClient() {
  const rawUser = AD_SERVICE_USER || "";
  const username = rawUser.includes("@") ? rawUser.split("@")[0] : rawUser;

  // Try all common AD bind formats
  const formats = [
    `${username}@${AD_DOMAIN}`,
    `${AD_DOMAIN.split(".")[0].toUpperCase()}\\${username}`,
    username,
    rawUser,
  ];

  console.log("🔑 Trying service account formats...");

  for (const dn of formats) {
    const client = createLDAPClient();
    try {
      await new Promise((resolve, reject) => {
        client.bind(dn, AD_SERVICE_PASS, (err) => {
          if (err) {
            client.destroy();
            reject(err);
          } else {
            resolve();
          }
        });
      });
      console.log("✅ Service account connected:", dn);
      return client;
    } catch (err) {
      console.log(`⚠ Format failed (${dn}): code ${err.code}`);
    }
  }

  throw new Error("Service account failed. Check AD_SERVICE_USER and AD_SERVICE_PASS in .env");
}

// ─── Parse LDAP Entry ─────────────────────────────────────────────────────────
function parseEntry(entry) {
  if (entry.pojo) {
    return Object.fromEntries(
      entry.pojo.attributes.map((a) => [
        a.type,
        a.values.length === 1 ? a.values[0] : a.values,
      ])
    );
  }
  return entry.object;
}

// ─── Search Single User ───────────────────────────────────────────────────────
function searchUser(client, username) {
  return new Promise((resolve, reject) => {
    const opts = {
      scope: "sub",
      filter: `(sAMAccountName=${username})`,
      attributes: ["sAMAccountName", "displayName", "mail", "department", "title", "memberOf", "givenName", "sn", "telephoneNumber"],
    };

    console.log("🔍 Searching user:", username);
    client.search(AD_BASE_DN, opts, (err, res) => {
      if (err) return reject(err);
      const entries = [];
      res.on("searchEntry", (e) => {
        const u = parseEntry(e);
        console.log("✅ Found:", u.sAMAccountName, "|", u.displayName);
        entries.push(u);
      });
      res.on("error", reject);
      res.on("end", () => {
        if (entries.length === 0) reject(new Error("User not found in Active Directory."));
        else resolve(entries[0]);
      });
    });
  });
}

// ─── Search Employees ─────────────────────────────────────────────────────────
function searchEmployees(client, filter, limit = 200) {
  return new Promise((resolve, reject) => {
    const opts = {
      scope: "sub",
      filter,
      sizeLimit: limit,
      attributes: ["sAMAccountName", "displayName", "mail", "department", "title", "telephoneNumber", "memberOf", "givenName", "sn"],
    };

    const entries = [];
    client.search(AD_BASE_DN, opts, (err, res) => {
      if (err) return reject(err);
      res.on("searchEntry", (e) => {
        const u = parseEntry(e);
        entries.push({
          username: u.sAMAccountName,
          displayName: u.displayName,
          email: u.mail,
          department: u.department || getDepartmentFromGroups(u.memberOf),
          title: u.title,
          phone: u.telephoneNumber,
          role: getRoleFromGroups(u.memberOf),
        });
      });
      res.on("error", reject);
      res.on("end", () => resolve(entries));
    });
  });
}

// ─── POST /auth/login ─────────────────────────────────────────────────────────
app.post("/auth/login", async (req, res) => {
  const { username, password } = req.body;

  console.log("\n=== Login Attempt ===");
  console.log("Username:", username);

  if (!username || !password) {
    return res.status(400).json({ success: false, message: "Username and password are required." });
  }

  const userDN = `${username}@${AD_DOMAIN}`;
  let client;

  try {
    // Step 1: Verify user credentials
    console.log("Step 1: Verifying credentials...");
    client = createLDAPClient();

    await new Promise((resolve, reject) => {
      client.bind(userDN, password, (err) => {
        if (err) {
          console.error("❌ User bind error:", err.message, "| Code:", err.code);
          reject(new Error(err.code === 49 ? "Invalid username or password." : `Connection error: ${err.message}`));
        } else {
          console.log("✅ Credentials verified!");
          resolve();
        }
      });
    });

    // Step 2: Fetch user details via service account
    console.log("Step 2: Fetching user details...");
    const serviceClient = await getServiceClient();
    const userInfo = await searchUser(serviceClient, username);
    serviceClient.destroy();

    // Step 3: Map role and department
    const role = getRoleFromGroups(userInfo.memberOf);
    const department = userInfo.department || getDepartmentFromGroups(userInfo.memberOf);
    console.log("✅ Role:", role, "| Department:", department);

    // Step 4: Generate JWT
    const token = jwt.sign(
      { username: userInfo.sAMAccountName, displayName: userInfo.displayName, email: userInfo.mail, department, title: userInfo.title, phone: userInfo.telephoneNumber, role },
      JWT_SECRET,
      { expiresIn: "8h" }
    );

    console.log("✅ Login successful:", username);

    return res.json({
      success: true,
      token,
      user: {
        username: userInfo.sAMAccountName,
        displayName: userInfo.displayName,
        email: userInfo.mail,
        department,
        title: userInfo.title,
        phone: userInfo.telephoneNumber,
        role,
      },
    });
  } catch (err) {
    console.error("❌ Login failed:", err.message);
    return res.status(401).json({ success: false, message: err.message || "Authentication failed." });
  } finally {
    if (client) client.destroy();
  }
});

// ─── GET /auth/verify ─────────────────────────────────────────────────────────
app.get("/auth/verify", (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try {
    const decoded = jwt.verify(authHeader.split(" ")[1], JWT_SECRET);
    return res.json({ success: true, user: decoded });
  } catch {
    return res.status(401).json({ success: false, message: "Invalid or expired token." });
  }
});

// ─── GET /employees ───────────────────────────────────────────────────────────
app.get("/employees", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  try {
    const serviceClient = await getServiceClient();
    const filter = "(&(objectClass=user)(objectCategory=person)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))";
    const employees = await searchEmployees(serviceClient, filter);
    serviceClient.destroy();
    return res.json({ success: true, count: employees.length, employees });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── GET /employees/search?q= ─────────────────────────────────────────────────
app.get("/employees/search", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const { q } = req.query;
  if (!q) return res.status(400).json({ success: false, message: "?q= is required." });

  try {
    const serviceClient = await getServiceClient();
    const filter = `(&(objectClass=user)(objectCategory=person)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(|(displayName=*${q}*)(sAMAccountName=*${q}*)(department=*${q}*)(mail=*${q}*)))`;
    const employees = await searchEmployees(serviceClient, filter);
    serviceClient.destroy();
    return res.json({ success: true, count: employees.length, employees });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── GET /employees/:username ─────────────────────────────────────────────────
app.get("/employees/:username", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  try {
    const serviceClient = await getServiceClient();
    const employees = await searchEmployees(serviceClient, `(&(objectClass=user)(sAMAccountName=${req.params.username}))`, 1);
    serviceClient.destroy();
    if (employees.length === 0)
      return res.status(404).json({ success: false, message: "Employee not found." });
    return res.json({ success: true, employee: employees[0] });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── GET /test-service ────────────────────────────────────────────────────────
app.get("/test-service", async (req, res) => {
  const username = AD_SERVICE_USER?.includes("@") ? AD_SERVICE_USER.split("@")[0] : AD_SERVICE_USER;
  const tryBind = (dn) => new Promise((resolve) => {
    const c = createLDAPClient();
    c.bind(dn, AD_SERVICE_PASS, (err) => {
      c.destroy();
      resolve(err ? { dn, success: false, code: err.code } : { dn, success: true });
    });
  });
  const results = await Promise.all([
    tryBind(`${username}@${AD_DOMAIN}`),
    tryBind(`${AD_DOMAIN.split(".")[0].toUpperCase()}\\${username}`),
    tryBind(username),
  ]);
  return res.json({ service_user: AD_SERVICE_USER, pass_length: AD_SERVICE_PASS?.length, results });
});

// ─── GET /health ──────────────────────────────────────────────────────────────
app.get("/health", (req, res) => {
  res.json({ status: "Backend is running", domain: AD_DOMAIN, adUrl: AD_URL });
});

// ─── Start ────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n✅ Silverdab backend running on port ${PORT}`);
  console.log(`📡 AD Server: ${AD_URL}`);
  console.log(`🌐 Domain: ${AD_DOMAIN}`);
});
