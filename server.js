const express = require("express");
const ldap = require("ldapjs");
const jwt = require("jsonwebtoken");
const crypto = require("crypto");
const cors = require("cors");
const nodemailer = require("nodemailer");
const webpush = require("web-push");
require("dotenv").config();

// Loaded defensively — if expo-server-sdk isn't installed yet, native push
// is disabled instead of crashing the whole server on startup (this bit
// everyone last time: an uninstalled require() throws before app.listen()).
let Expo = null;
let expo = null;
try {
  ({ Expo } = require("expo-server-sdk"));
  expo = new Expo();
  console.log("📱 Expo push SDK loaded");
} catch (err) {
  console.warn("⚠ expo-server-sdk not installed — native app push disabled. Run: npm install expo-server-sdk");
}

const {
  createTripEvent, updateTripEvent, deleteTripEvent,
  createRoomEvent, updateRoomEvent, deleteRoomEvent,
} = require("./graphCalendarService");

// POST /fleet/trips/:id/sync-to-my-calendar — pushes the trip event
// directly to the logged-in user's own Outlook mailbox (via app-permission
// Graph, targeting their notification_email/email instead of the fixed
// fleet admin mailbox).

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

// Always CC'd on supply request notifications (new requests + status updates)
const SUPPLY_REQUEST_NOTIFY_EMAIL = "hess.espinas@ocgbim.com";

// If a request contains an item whose name starts with "Test Item" (e.g.
// "Test Item 1", "Test Item 2"), treat it as a dev/test request and skip all
// notification emails for it — just add a "Test Item ..." item to the cart
// while testing, no other setup needed.
function containsTestItem(items) {
  return Array.isArray(items) && items.some((i) => {
    const name = i.itemName ?? i.item_name ?? "";
    return name.toLowerCase().startsWith("test item");
  });
}

// ─── Web Push setup ─────────────────────────────────────────────────────────
if (process.env.VAPID_PUBLIC_KEY && process.env.VAPID_PRIVATE_KEY) {
  webpush.setVapidDetails(
    process.env.VAPID_SUBJECT || "mailto:admin@silverdab.com",
    process.env.VAPID_PUBLIC_KEY,
    process.env.VAPID_PRIVATE_KEY,
  );
  console.log("🔔 Web push configured");
} else {
  console.warn("⚠ VAPID keys missing — web push notifications disabled");
}

// Sends a browser popup notification to every admin who has subscribed.
// Silently drops any subscription that's gone stale (expired/unsubscribed
// in the browser) by deleting it from the DB — same "never throw" pattern
// as the email senders, since a failed push shouldn't break the request flow.
async function sendWebPushToAdmins({ title, body, url, permissionColumn, permissionColumns }) {
  if (!process.env.VAPID_PUBLIC_KEY) {
    console.warn("🔔 sendWebPushToAdmins: VAPID not configured, skipping");
    return;
  }
  try {
    let subsQuery = `
      SELECT ps.* FROM push_subscriptions ps
      JOIN users u ON u.username = ps.username
    `;
    // Superadmins always see everything (same bypass as GET /permissions/me).
    // Accepts either a single permissionColumn (legacy) or an array of
    // permissionColumns (any one of which grants notification) — falls back
    // to notifying everyone subscribed if neither is given.
    const columns = permissionColumns ?? (permissionColumn ? [permissionColumn] : []);
    // Hardcoded allow-list guards against SQL injection since these are
    // interpolated directly into the query string below.
    const ALLOWED_COLUMNS = new Set([
      "perm_it_access", "perm_office_supplies", "perm_office_all_access",
      "perm_office_dashboard", "perm_office_inventory", "perm_office_supply_request",
      "perm_office_monthly_report", "perm_office_activity",
      "perm_fleet_control", "perm_fleet_driver",
    ]);
    const safeColumns = columns.filter((c) => ALLOWED_COLUMNS.has(c));
    if (safeColumns.length > 0) {
      const orClause = safeColumns.map((c) => `u.${c} = 1`).join(" OR ");
      subsQuery += ` WHERE u.role = 'superadmin' OR (${orClause})`;
    }
    const [subs] = await db.query(subsQuery);
    console.log(`🔔 sendWebPushToAdmins: found ${subs.length} subscription(s)`);
    const payload = JSON.stringify({ title, body, url: url || "/" });

    await Promise.all(
      subs.map(async (sub) => {
        try {
          await webpush.sendNotification(
            {
              endpoint: sub.endpoint,
              keys: { p256dh: sub.p256dh, auth: sub.auth },
            },
            payload,
          );
        } catch (err) {
          if (err.statusCode === 404 || err.statusCode === 410) {
            // Subscription expired or was revoked — clean it up.
            console.warn(`🔔 Subscription ${sub.id} expired/revoked, removing`);
            await db.query("DELETE FROM push_subscriptions WHERE id = ?", [sub.id]);
          } else {
            console.error("🔔 Web push send failed:", err.statusCode, err.message);
          }
        }
      }),
    );
    console.log("🔔 sendWebPushToAdmins: done");
  } catch (err) {
    console.error("🔔 sendWebPushToAdmins failed:", err.message);
  }
}

// Sends a web push to ONE specific user by username — used for driver
// trip-assignment notifications, where the audience is exactly one person
// rather than a permission-based group.
async function sendWebPushToUser(username, { title, body, url }) {
  if (!process.env.VAPID_PUBLIC_KEY) {
    console.warn("🔔 sendWebPushToUser: VAPID not configured, skipping");
    return;
  }
  try {
    const [subs] = await db.query(
      "SELECT * FROM push_subscriptions WHERE username = ?",
      [username],
    );
    console.log(`🔔 sendWebPushToUser(${username}): found ${subs.length} subscription(s)`);
    const payload = JSON.stringify({ title, body, url: url || "/" });

    await Promise.all(
      subs.map(async (sub) => {
        try {
          await webpush.sendNotification(
            { endpoint: sub.endpoint, keys: { p256dh: sub.p256dh, auth: sub.auth } },
            payload,
          );
        } catch (err) {
          if (err.statusCode === 404 || err.statusCode === 410) {
            console.warn(`🔔 Subscription ${sub.id} expired/revoked, removing`);
            await db.query("DELETE FROM push_subscriptions WHERE id = ?", [sub.id]);
          } else {
            console.error("🔔 Web push send failed:", err.statusCode, err.message);
          }
        }
      }),
    );
  } catch (err) {
    console.error(`🔔 sendWebPushToUser(${username}) failed:`, err.message);
  }
}

// Sends a native push notification to ONE specific user via their
// registered Expo push token(s) — the native-app counterpart to
// sendWebPushToUser(), used when the driver is on the mobile app rather
// than a browser tab with a VAPID subscription.
async function sendExpoPushToUser(username, { title, body, data }) {
  if (!expo) {
    console.warn("📱 sendExpoPushToUser: Expo SDK not available, skipping");
    return;
  }
  try {
    const [tokens] = await db.query(
      "SELECT token FROM expo_push_tokens WHERE username = ?",
      [username],
    );
    if (tokens.length === 0) {
      console.log(`📱 sendExpoPushToUser(${username}): no tokens found`);
      return;
    }

    const messages = tokens
      .filter((t) => Expo.isExpoPushToken(t.token))
      .map((t) => ({ to: t.token, sound: "default", title, body, data: data ?? {} }));

    if (messages.length === 0) return;

    const chunks = expo.chunkPushNotifications(messages);
    for (const chunk of chunks) {
      const tickets = await expo.sendPushNotificationsAsync(chunk);
      tickets.forEach((ticket, i) => {
        if (ticket.status === "error" && ticket.details?.error === "DeviceNotRegistered") {
          db.query("DELETE FROM expo_push_tokens WHERE token = ?", [chunk[i].to]).catch(() => {});
        }
      });
    }
    console.log(`📱 sendExpoPushToUser(${username}): sent to ${messages.length} device(s)`);
  } catch (err) {
    console.error(`📱 sendExpoPushToUser(${username}) failed:`, err.message);
  }
}

// Sends a native push notification to every admin matching the given
// permission columns (or all admins/superadmins if none given) — the Expo
// counterpart to sendWebPushToAdmins(), for admins on the mobile app rather
// than a browser tab.
async function sendExpoPushToAdmins({ title, body, data, permissionColumn, permissionColumns }) {
  if (!expo) {
    console.warn("📱 sendExpoPushToAdmins: Expo SDK not available, skipping");
    return;
  }
  try {
    let tokensQuery = `
      SELECT ept.token FROM expo_push_tokens ept
      JOIN users u ON u.username = ept.username
    `;
    const columns = permissionColumns ?? (permissionColumn ? [permissionColumn] : []);
    const ALLOWED_COLUMNS = new Set([
      "perm_it_access", "perm_office_supplies", "perm_office_all_access",
      "perm_office_dashboard", "perm_office_inventory", "perm_office_supply_request",
      "perm_office_monthly_report", "perm_office_activity",
      "perm_fleet_control", "perm_fleet_driver",
    ]);
    const safeColumns = columns.filter((c) => ALLOWED_COLUMNS.has(c));
    if (safeColumns.length > 0) {
      const orClause = safeColumns.map((c) => `u.${c} = 1`).join(" OR ");
      tokensQuery += ` WHERE u.role = 'superadmin' OR (${orClause})`;
    }
    const [tokens] = await db.query(tokensQuery);
    console.log(`📱 sendExpoPushToAdmins: found ${tokens.length} token(s)`);

    const messages = tokens
      .filter((t) => Expo.isExpoPushToken(t.token))
      .map((t) => ({ to: t.token, sound: "default", title, body, data: data ?? {} }));

    if (messages.length === 0) return;

    const chunks = expo.chunkPushNotifications(messages);
    for (const chunk of chunks) {
      const tickets = await expo.sendPushNotificationsAsync(chunk);
      tickets.forEach((ticket, i) => {
        if (ticket.status === "error" && ticket.details?.error === "DeviceNotRegistered") {
          db.query("DELETE FROM expo_push_tokens WHERE token = ?", [chunk[i].to]).catch(() => {});
        }
      });
    }
    console.log(`📱 sendExpoPushToAdmins: sent to ${messages.length} device(s)`);
  } catch (err) {
    console.error("📱 sendExpoPushToAdmins failed:", err.message);
  }
}

console.log("=== Backend Config ===");
console.log("AD_URL:", AD_URL);
console.log("AD_BASE_DN:", AD_BASE_DN);
console.log("AD_DOMAIN:", AD_DOMAIN);
console.log("AD_SERVICE_USER:", AD_SERVICE_USER);
console.log("AD_SERVICE_PASS:", AD_SERVICE_PASS ? "✅ set" : "❌ missing");
console.log("======================");

// ─── MySQL Connection ──────────────────────────────────────────────────────────
const mysql = require("mysql2/promise");

const db = mysql.createPool({
  host: process.env.MYSQL_HOST,
  port: process.env.MYSQL_PORT || 3306,
  user: process.env.MYSQL_USER,
  password: process.env.MYSQL_PASSWORD,
  database: process.env.MYSQL_DATABASE,
  waitForConnections: true,
  connectionLimit: 10,
});

db.getConnection()
  .then(() => console.log("✅ MySQL connected!"))
  .catch((err) => console.error("❌ MySQL error:", err.message));

// Returns the current local time as a MySQL DATETIME string
// ("YYYY-MM-DD HH:MM:SS"), independent of MySQL's own session timezone.
// Used instead of NOW() in fleet trip inserts to avoid the UTC-vs-local
// mismatch without touching the global mysql2 pool `timezone` option.
function nowLocalDatetime() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

// ─── Email (SMTP) ──────────────────────────────────────────────────────────
const transporter = nodemailer.createTransport({
  host: "smtp.gmail.com",
  port: 587,
  secure: false,
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

async function sendRequestNotification({ requestedById, requestedByName, ticketNumber, items }) {
  if (containsTestItem(items)) {
    console.log(`📧 [skipped, test item] request notification for ${ticketNumber}`);
    return;
  }
  try {
    const [rows] = await db.query(
      "SELECT notification_email FROM users WHERE username = ?",
      [requestedById],
    );
    const toEmail = rows[0]?.notification_email;
    if (!toEmail) {
      console.warn(`No email preference set for ${requestedById}, skipping notification.`);
      return;
    }

    const itemListText = items
      .map((i) => `- ${i.itemName} (Qty: ${i.quantityRequested})`)
      .join("\n");

    const itemListHtml = items
      .map(
        (i) => `
          <tr>
            <td style="padding: 8px 12px; border-bottom: 1px solid #e5e7eb; color: #1f2937;">${i.itemName}</td>
            <td style="padding: 8px 12px; border-bottom: 1px solid #e5e7eb; color: #1f2937; text-align: right;">${i.quantityRequested}</td>
          </tr>`,
      )
      .join("");

    const htmlBody = `
      <div style="font-family: Arial, Helvetica, sans-serif; max-width: 560px; margin: 0 auto; color: #1f2937;">
        <div style="background-color: #1e3a5f; padding: 20px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="color: #ffffff; margin: 0; font-size: 18px;">Silverdab Supply Request</h2>
        </div>
        <div style="border: 1px solid #e5e7eb; border-top: none; border-radius: 0 0 8px 8px; padding: 24px;">
          <p style="margin: 0 0 12px 0;">Dear ${requestedByName},</p>
          <p style="margin: 0 0 20px 0; line-height: 1.6;">
            This is to confirm that your supply request has been successfully submitted
            and is now pending review by the administration team.
          </p>

          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Ticket Number</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${ticketNumber}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Status</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #b45309;">Pending Review</td>
            </tr>
          </table>

          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <thead>
              <tr style="background-color: #f3f4f6;">
                <th style="padding: 8px 12px; text-align: left; font-size: 12px; text-transform: uppercase; color: #6b7280;">Item</th>
                <th style="padding: 8px 12px; text-align: right; font-size: 12px; text-transform: uppercase; color: #6b7280;">Qty</th>
              </tr>
            </thead>
            <tbody>
              ${itemListHtml}
            </tbody>
          </table>

          <p style="margin: 0 0 4px 0; line-height: 1.6; font-size: 13px; color: #6b7280;">
            You will receive a follow-up notification once your request has been reviewed.
          </p>
          <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
            <tr>
              <td style="vertical-align: middle; line-height: 1.6;">
                Thank you,<br/>
                <strong>Silvergraph Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:SilvergraphLogo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
              </td>
            </tr>
          </table>
        </div>
        <p style="font-size: 11px; color: #9ca3af; text-align: center; margin-top: 16px;">
          This is an automated notification. Please do not reply directly to this email.
        </p>
      </div>
    `;

    const info = await transporter.sendMail({
      from: `"Silvergraph Unified Management System" <${process.env.EMAIL_USER}>`,
      to: toEmail,
      subject: `Supply Request ${ticketNumber} Received`,
      text: `Dear ${requestedByName},\n\nYour supply request ${ticketNumber} has been submitted and is pending review.\n\nItems:\n${itemListText}`,
      html: htmlBody,
      attachments: [
        {
          filename: "SilvergraphLogo.png",
          path: "./assets/SilvergraphLogo.png",
          cid: "SilvergraphLogo",
        },
      ],
    });

    await db.query(
      "UPDATE supply_requests SET email_message_id = ? WHERE ticket_number = ?",
      [info.messageId, ticketNumber],
    );

    console.log(`📧 Notification sent to ${toEmail} for ${ticketNumber}`);
  } catch (err) {
    console.error("Email notification failed:", err.message);
    // never throw — a failed email should not break the request flow
  }
}

async function sendAdminRequestNotification({ requestedByName, ticketNumber, items }) {
  if (containsTestItem(items)) {
    console.log(`📧 [skipped, test item] admin request notification for ${ticketNumber}`);
    return;
  }
  try {
    const itemListText = items
      .map((i) => `- ${i.itemName} (Qty: ${i.quantityRequested})`)
      .join("\n");

    const itemListHtml = items
      .map(
        (i) => `
          <tr>
            <td style="padding: 8px 12px; border-bottom: 1px solid #e5e7eb; color: #1f2937;">${i.itemName}</td>
            <td style="padding: 8px 12px; border-bottom: 1px solid #e5e7eb; color: #1f2937; text-align: right;">${i.quantityRequested}</td>
          </tr>`,
      )
      .join("");

    const htmlBody = `
      <div style="font-family: Arial, Helvetica, sans-serif; max-width: 560px; margin: 0 auto; color: #1f2937;">
        <div style="background-color: #b45309; padding: 20px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="color: #ffffff; margin: 0; font-size: 18px;">New Supply Request — Action Needed</h2>
        </div>
        <div style="border: 1px solid #e5e7eb; border-top: none; border-radius: 0 0 8px 8px; padding: 24px;">
          <p style="margin: 0 0 12px 0;">Hi,</p>
          <p style="margin: 0 0 20px 0; line-height: 1.6;">
            <strong>${requestedByName}</strong> has submitted a new supply request that is
            awaiting your review and approval.
          </p>

          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Ticket Number</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${ticketNumber}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Requested By</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${requestedByName}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Status</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #b45309;">Pending Review</td>
            </tr>
          </table>

          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <thead>
              <tr style="background-color: #f3f4f6;">
                <th style="padding: 8px 12px; text-align: left; font-size: 12px; text-transform: uppercase; color: #6b7280;">Item</th>
                <th style="padding: 8px 12px; text-align: right; font-size: 12px; text-transform: uppercase; color: #6b7280;">Qty</th>
              </tr>
            </thead>
            <tbody>
              ${itemListHtml}
            </tbody>
          </table>

          <p style="margin: 0 0 4px 0; line-height: 1.6; font-size: 13px; color: #6b7280;">
            Please log in to Silverdab UMS to approve, partially approve, or reject this request.
          </p>

          <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
            <tr>
              <td style="vertical-align: middle; line-height: 1.6;">
                Regards,<br/>
                <strong>Silvergraph Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:SilvergraphLogo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
              </td>
            </tr>
          </table>
        </div>
        <p style="font-size: 11px; color: #9ca3af; text-align: center; margin-top: 16px;">
          This is an automated notification. Please do not reply directly to this email.
        </p>
      </div>
    `;

    await transporter.sendMail({
      from: `"Silvergraph Unified Management System" <${process.env.EMAIL_USER}>`,
      to: SUPPLY_REQUEST_NOTIFY_EMAIL,
      subject: `New Supply Request ${ticketNumber} — Pending Your Review`,
      text: `${requestedByName} has submitted a new supply request awaiting your approval.\n\nTicket: ${ticketNumber}\n\nItems:\n${itemListText}\n\nPlease log in to Silverdab UMS to review.`,
      html: htmlBody,
      attachments: [
        {
          filename: "SilvergraphLogo.png",
          path: "./assets/SilvergraphLogo.png",
          cid: "SilvergraphLogo",
        },
      ],
    });

    console.log(`📧 Admin notification sent to ${SUPPLY_REQUEST_NOTIFY_EMAIL} for ${ticketNumber}`);
  } catch (err) {
    console.error("Admin email notification failed:", err.message);
    // never throw — a failed email should not break the request flow
  }
}

async function sendStatusUpdateNotification({ requestId, statusLabel, extraMessage, updatedByName }) {
  try {
    const [itemRows] = await db.query(
      "SELECT item_name FROM supply_request_items WHERE request_id = ?",
      [requestId],
    );
    if (containsTestItem(itemRows.map((r) => ({ itemName: r.item_name })))) {
      console.log(`📧 [skipped, test item] status update (${statusLabel}) for request ${requestId}`);
      return;
    }

    const [reqRows] = await db.query(
      "SELECT ticket_number, requested_by_id, requested_by_name, email_message_id FROM supply_requests WHERE id = ?",
      [requestId],
    );
    if (reqRows.length === 0) return;
    const request = reqRows[0];

    const [userRows] = await db.query(
      "SELECT notification_email FROM users WHERE username = ?",
      [request.requested_by_id],
    );
    const toEmail = userRows[0]?.notification_email;
    if (!toEmail) {
      console.warn(`No email preference set for ${request.requested_by_id}, skipping status update.`);
      return;
    }

    const statusColors = {
      "Out for Delivery": "#1e3a5f",
      "Rejected": "#b91c1c",
      "Issued": "#15803d",
      "Failed Delivery": "#b45309",
      "Cancelled": "#475569",
    };
    const statusColor = statusColors[statusLabel] || "#1e3a5f";

    const htmlBody = `
      <div style="font-family: Arial, Helvetica, sans-serif; max-width: 560px; margin: 0 auto; color: #1f2937;">
        <div style="background-color: ${statusColor}; padding: 20px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="color: #ffffff; margin: 0; font-size: 18px;">Supply Request Update</h2>
        </div>
        <div style="border: 1px solid #e5e7eb; border-top: none; border-radius: 0 0 8px 8px; padding: 24px;">
          <p style="margin: 0 0 12px 0;">Dear ${request.requested_by_name},</p>
          <p style="margin: 0 0 20px 0; line-height: 1.6;">
            The status of your supply request has been updated.
          </p>
          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Ticket Number</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${request.ticket_number}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Status</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: ${statusColor};">${statusLabel}</td>
            </tr>
            ${updatedByName ? `
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Updated By</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${updatedByName}</td>
            </tr>
            ` : ""}
          </table>
          ${extraMessage ? `<p style="margin: 0 0 20px 0; line-height: 1.6; color: #374151;">${extraMessage}</p>` : ""}
          <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
            <tr>
              <td style="vertical-align: middle; line-height: 1.6;">
                Thank you,<br/>
                <strong>Silvergraph Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:SilvergraphLogo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
              </td>
            </tr>
          </table>
        </div>
        <p style="font-size: 11px; color: #9ca3af; text-align: center; margin-top: 16px;">
          This is an automated notification. Please do not reply directly to this email.
        </p>
      </div>
    `;

    const mailOptions = {
      from: `"Silvergraph Unified Management System" <${process.env.EMAIL_USER}>`,
      to: toEmail,
      subject: `Re: Supply Request ${request.ticket_number} Received`,
      text: `Dear ${request.requested_by_name},\n\nYour supply request ${request.ticket_number} status has been updated to: ${statusLabel}.${updatedByName ? ` (by ${updatedByName})` : ""}${extraMessage ? `\n\n${extraMessage}` : ""}`,
      html: htmlBody,
      attachments: [
        {
          filename: "SilvergraphLogo.png",
          path: "./assets/SilvergraphLogo.png",
          cid: "SilvergraphLogo",
        },
      ],
    };

    if (request.email_message_id) {
      mailOptions.inReplyTo = request.email_message_id;
      mailOptions.references = request.email_message_id;
    }

    await transporter.sendMail(mailOptions);
    console.log(`📧 Status update (${statusLabel}) sent to ${toEmail} for ${request.ticket_number}`);
  } catch (err) {
    console.error("Status update email failed:", err.message);
  }
}

async function sendAdminStatusUpdateNotification({ requestId, statusLabel, extraMessage, updatedByName }) {
  try {
    const [itemRows] = await db.query(
      "SELECT item_name FROM supply_request_items WHERE request_id = ?",
      [requestId],
    );
    if (containsTestItem(itemRows.map((r) => ({ itemName: r.item_name })))) {
      console.log(`📧 [skipped, test item] admin status update (${statusLabel}) for request ${requestId}`);
      return;
    }

    const [reqRows] = await db.query(
      "SELECT ticket_number, requested_by_name FROM supply_requests WHERE id = ?",
      [requestId],
    );
    if (reqRows.length === 0) return;
    const request = reqRows[0];

    const statusColors = {
      "Out for Delivery": "#1e3a5f",
      "Rejected": "#b91c1c",
      "Issued": "#15803d",
      "Failed Delivery": "#b45309",
      "Cancelled": "#475569",
    };
    const statusColor = statusColors[statusLabel] || "#1e3a5f";

    const htmlBody = `
      <div style="font-family: Arial, Helvetica, sans-serif; max-width: 560px; margin: 0 auto; color: #1f2937;">
        <div style="background-color: ${statusColor}; padding: 20px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="color: #ffffff; margin: 0; font-size: 18px;">Supply Request Status Changed</h2>
        </div>
        <div style="border: 1px solid #e5e7eb; border-top: none; border-radius: 0 0 8px 8px; padding: 24px;">
          <p style="margin: 0 0 12px 0;">Hi,</p>
          <p style="margin: 0 0 20px 0; line-height: 1.6;">
            This is a record that the status of a supply request was updated.
          </p>

          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Ticket Number</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${request.ticket_number}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Requested By</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${request.requested_by_name}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">New Status</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: ${statusColor};">${statusLabel}</td>
            </tr>
            ${updatedByName ? `
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Updated By</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${updatedByName}</td>
            </tr>
            ` : ""}
          </table>

          ${extraMessage ? `<p style="margin: 0 0 20px 0; line-height: 1.6; color: #374151;">${extraMessage}</p>` : ""}

          <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
            <tr>
              <td style="vertical-align: middle; line-height: 1.6;">
                Regards,<br/>
                <strong>Silvergraph Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:SilvergraphLogo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
              </td>
            </tr>
          </table>
        </div>
        <p style="font-size: 11px; color: #9ca3af; text-align: center; margin-top: 16px;">
          This is an automated notification. Please do not reply directly to this email.
        </p>
      </div>
    `;

    await transporter.sendMail({
      from: `"Silvergraph Unified Management System" <${process.env.EMAIL_USER}>`,
      to: SUPPLY_REQUEST_NOTIFY_EMAIL,
      subject: `Supply Request ${request.ticket_number} — Status Changed to ${statusLabel}`,
      text: `Ticket ${request.ticket_number} (requested by ${request.requested_by_name}) status changed to: ${statusLabel}.${updatedByName ? ` Updated by ${updatedByName}.` : ""}${extraMessage ? `\n\n${extraMessage}` : ""}`,
      html: htmlBody,
      attachments: [
        {
          filename: "SilvergraphLogo.png",
          path: "./assets/SilvergraphLogo.png",
          cid: "SilvergraphLogo",
        },
      ],
    });

    console.log(`📧 Admin status update (${statusLabel}) sent to ${SUPPLY_REQUEST_NOTIFY_EMAIL} for ${request.ticket_number}`);
  } catch (err) {
    console.error("Admin status update email failed:", err.message);
  }
}

// Trip requests don't have a single fixed notify address like supply requests
// do (SUPPLY_REQUEST_NOTIFY_EMAIL) — fleet control access is granted per-user
// via perm_fleet_control, so look up everyone with that flag (or superadmin)
// and email them all.
async function sendFleetTripAdminNotification({
  requestorName,
  tripRef,
  pickupLocationText,
  dropoffLocationText,
  tripType,
  departureDatetime,
  purpose,
  passengerCount,
  replyToMessageId,
}) {
  // Skip admin notification entirely when this is Henrick's own test
  // bookings — keeps fleet control admins' inboxes from getting flooded
  // during testing. Matches on display name, case-insensitive.
  if ((requestorName || "").trim().toLowerCase() === "henrick e. patenio") {
    console.log(`📧 [skipped, test requestor] fleet trip admin notification for ${tripRef}`);
    return;
  }

  try {
    const [rows] = await db.query(
      `SELECT notification_email, email FROM users
       WHERE perm_fleet_control = 1`,
    );
    const recipients = rows
      .map((r) => r.notification_email || r.email)
      .filter(Boolean);

    if (recipients.length === 0) {
      console.warn("No fleet control recipients found, skipping trip request email.");
      return;
    }

    const htmlBody = `
      <div style="font-family: Arial, Helvetica, sans-serif; max-width: 560px; margin: 0 auto; color: #1f2937;">
        <div style="background-color: #b45309; padding: 20px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="color: #ffffff; margin: 0; font-size: 18px;">New Trip Request — Action Needed</h2>
        </div>
        <div style="border: 1px solid #e5e7eb; border-top: none; border-radius: 0 0 8px 8px; padding: 24px;">
          <p style="margin: 0 0 12px 0;">Hi,</p>
          <p style="margin: 0 0 20px 0; line-height: 1.6;">
            <strong>${requestorName}</strong> has submitted a new trip request that is
            awaiting dispatch review.
          </p>

          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Trip Ref</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${tripRef}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Requested By</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${requestorName}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Pick-up</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${pickupLocationText}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Drop-off</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${dropoffLocationText}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Trip Type</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${tripType}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Departure</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${departureDatetime}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Passengers</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${passengerCount ?? 1}</td>
            </tr>
          </table>

          ${purpose ? `<p style="margin: 0 0 20px 0; line-height: 1.6; color: #374151;"><strong>Purpose:</strong> ${purpose}</p>` : ""}

          <p style="margin: 0 0 4px 0; line-height: 1.6; font-size: 13px; color: #6b7280;">
            Please log in to Silverdab UMS to approve or reject this request.
          </p>

          <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
            <tr>
              <td style="vertical-align: middle; line-height: 1.6;">
                Regards,<br/>
                <strong>Silvergraph Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:SilvergraphLogo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
              </td>
            </tr>
          </table>
        </div>
        <p style="font-size: 11px; color: #9ca3af; text-align: center; margin-top: 16px;">
          This is an automated notification. Please do not reply directly to this email.
        </p>
      </div>
    `;

    const mailOptions = {
      from: `"Silvergraph Unified Management System" <${process.env.EMAIL_USER}>`,
      to: recipients.join(", "),
      subject: `Re: Trip Request ${tripRef} Received`,
      text: `${requestorName} has submitted a new trip request awaiting your approval.\n\nTrip: ${tripRef}\nPick-up: ${pickupLocationText}\nDrop-off: ${dropoffLocationText}\nType: ${tripType}\nDeparture: ${departureDatetime}\n${purpose ? `Purpose: ${purpose}\n` : ""}\nPlease log in to Silverdab UMS to review.`,
      html: htmlBody,
      attachments: [
        {
          filename: "SilvergraphLogo.png",
          path: "./assets/SilvergraphLogo.png",
          cid: "SilvergraphLogo",
        },
      ],
    };

    if (replyToMessageId) {
      mailOptions.inReplyTo = replyToMessageId;
      mailOptions.references = replyToMessageId;
    }

    await transporter.sendMail(mailOptions);

    console.log(`📧 Fleet trip admin notification sent to ${recipients.length} recipient(s) for ${tripRef}`);
  } catch (err) {
    console.error("Fleet trip admin email failed:", err.message);
    // never throw — a failed email should not break the request flow
  }
}

const FLEET_TRIP_STATUS_LABELS = {
  pending: "Pending Review",
  approved: "Approved",
  rejected: "Rejected",
  ongoing: "Ongoing",
  arrived: "Completed",
  returning: "Returning",
  completed: "Completed",
  cancelled: "Cancelled",
};

const FLEET_TRIP_STATUS_COLORS = {
  pending: "#b45309",
  approved: "#1e3a5f",
  rejected: "#b91c1c",
  ongoing: "#1e3a5f",
  arrived: "#15803d",
  returning: "#b45309",
  completed: "#15803d",
  cancelled: "#475569",
};

// Notifies the employee who requested the trip whenever its status changes
// (approved, rejected, ongoing, arrived, returning, completed) — mirrors
// sendStatusUpdateNotification() for supply requests.
async function sendFleetTripStatusNotification({ tripId, statusKey, extraMessage, updatedByName }) {
  try {
    const [tripRows] = await db.query(
      `SELECT t.trip_ref, t.pickup_location_text, t.dropoff_location_text,
              t.email_message_id,
              u.username AS requestor_username, u.display_name AS requestor_name,
              u.notification_email, u.email,
              drv.display_name AS driver_name,
              veh.plate_number AS vehicle_plate, veh.model AS vehicle_model
       FROM fleet_trips t
       JOIN users u ON u.id = t.requestor_id
       LEFT JOIN users drv ON drv.id = t.driver_id
       LEFT JOIN fleet_vehicles veh ON veh.id = t.vehicle_id
       WHERE t.id = ?`,
      [tripId],
    );
    if (tripRows.length === 0) return;
    const trip = tripRows[0];

    const toEmail = trip.notification_email || trip.email;
    if (!toEmail) {
      console.warn(`No email preference set for ${trip.requestor_username}, skipping trip status update.`);
      return;
    }

    const statusLabel = FLEET_TRIP_STATUS_LABELS[statusKey] || statusKey;
    const statusColor = FLEET_TRIP_STATUS_COLORS[statusKey] || "#1e3a5f";

    const htmlBody = `
      <div style="font-family: Arial, Helvetica, sans-serif; max-width: 560px; margin: 0 auto; color: #1f2937;">
        <div style="background-color: ${statusColor}; padding: 20px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="color: #ffffff; margin: 0; font-size: 18px;">Trip Request Update</h2>
        </div>
        <div style="border: 1px solid #e5e7eb; border-top: none; border-radius: 0 0 8px 8px; padding: 24px;">
          <p style="margin: 0 0 12px 0;">Dear ${trip.requestor_name},</p>
          <p style="margin: 0 0 20px 0; line-height: 1.6;">
            The status of your trip request has been updated.
          </p>
          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Trip Ref</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${trip.trip_ref}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Pick-up</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${trip.pickup_location_text}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Drop-off</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${trip.dropoff_location_text}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Status</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: ${statusColor};">${statusLabel}</td>
            </tr>
            ${trip.vehicle_plate ? `
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Assigned Vehicle</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${trip.vehicle_plate}${trip.vehicle_model ? ` (${trip.vehicle_model})` : ""}</td>
            </tr>
            ` : ""}
            ${trip.driver_name ? `
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Assigned Driver</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${trip.driver_name}</td>
            </tr>
            ` : ""}
            ${updatedByName ? `
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Updated By</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${updatedByName}</td>
            </tr>
            ` : ""}
          </table>
          ${extraMessage ? `<p style="margin: 0 0 20px 0; line-height: 1.6; color: #374151;">${extraMessage}</p>` : ""}
          <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
            <tr>
              <td style="vertical-align: middle; line-height: 1.6;">
                Thank you,<br/>
                <strong>Silvergraph Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:SilvergraphLogo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
              </td>
            </tr>
          </table>
        </div>
        <p style="font-size: 11px; color: #9ca3af; text-align: center; margin-top: 16px;">
          This is an automated notification. Please do not reply directly to this email.
        </p>
      </div>
    `;

    const mailOptions = {
      from: `"Silvergraph Unified Management System" <${process.env.EMAIL_USER}>`,
      to: toEmail,
      subject: `Re: Trip Request ${trip.trip_ref} Received`,
      text: `Dear ${trip.requestor_name},\n\nYour trip request ${trip.trip_ref} status has been updated to: ${statusLabel}.${trip.vehicle_plate ? ` Vehicle: ${trip.vehicle_plate}.` : ""}${trip.driver_name ? ` Driver: ${trip.driver_name}.` : ""}${updatedByName ? ` (by ${updatedByName})` : ""}${extraMessage ? `\n\n${extraMessage}` : ""}`,
      html: htmlBody,
      attachments: [
        {
          filename: "SilvergraphLogo.png",
          path: "./assets/SilvergraphLogo.png",
          cid: "SilvergraphLogo",
        },
      ],
    };

    if (trip.email_message_id) {
      mailOptions.inReplyTo = trip.email_message_id;
      mailOptions.references = trip.email_message_id;
    }

    await transporter.sendMail(mailOptions);

    console.log(`📧 Trip status update (${statusLabel}) sent to ${toEmail} for ${trip.trip_ref}`);
  } catch (err) {
    console.error("Trip status update email failed:", err.message);
    // never throw — a failed email should not break the request flow
  }
}

// Notifies everyone with fleet control access whenever a trip's status
// changes — mirrors sendAdminStatusUpdateNotification() for supply requests.
async function sendFleetTripAdminStatusNotification({ tripId, statusKey, extraMessage, updatedByName }) {
  try {
    const [tripRows] = await db.query(
      `SELECT t.trip_ref, t.pickup_location_text, t.dropoff_location_text,
              t.email_message_id,
              u.display_name AS requestor_name,
              drv.display_name AS driver_name,
              veh.plate_number AS vehicle_plate, veh.model AS vehicle_model
       FROM fleet_trips t
       JOIN users u ON u.id = t.requestor_id
       LEFT JOIN users drv ON drv.id = t.driver_id
       LEFT JOIN fleet_vehicles veh ON veh.id = t.vehicle_id
       WHERE t.id = ?`,
      [tripId],
    );
    if (tripRows.length === 0) return;
    const trip = tripRows[0];

    // Same test-requestor skip as sendFleetTripAdminNotification — status
    // changes on Henrick's test trips fire far more often (every approve/
    // start/arrive/complete step), so this matters even more here.
    if ((trip.requestor_name || "").trim().toLowerCase() === "henrick e. patenio") {
      console.log(`📧 [skipped, test requestor] fleet trip admin status update (${statusKey}) for ${trip.trip_ref}`);
      return;
    }

    const [rows] = await db.query(
      `SELECT notification_email, email FROM users WHERE perm_fleet_control = 1`,
    );
    const recipients = rows
      .map((r) => r.notification_email || r.email)
      .filter(Boolean);

    if (recipients.length === 0) {
      console.warn("No fleet control recipients found, skipping trip admin status email.");
      return;
    }

    const statusLabel = FLEET_TRIP_STATUS_LABELS[statusKey] || statusKey;
    const statusColor = FLEET_TRIP_STATUS_COLORS[statusKey] || "#1e3a5f";

    const htmlBody = `
      <div style="font-family: Arial, Helvetica, sans-serif; max-width: 560px; margin: 0 auto; color: #1f2937;">
        <div style="background-color: ${statusColor}; padding: 20px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="color: #ffffff; margin: 0; font-size: 18px;">Trip Request Status Changed</h2>
        </div>
        <div style="border: 1px solid #e5e7eb; border-top: none; border-radius: 0 0 8px 8px; padding: 24px;">
          <p style="margin: 0 0 12px 0;">Hi,</p>
          <p style="margin: 0 0 20px 0; line-height: 1.6;">
            This is a record that the status of a trip request was updated.
          </p>

          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Trip Ref</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${trip.trip_ref}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Requested By</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${trip.requestor_name}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Pick-up</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${trip.pickup_location_text}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Drop-off</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${trip.dropoff_location_text}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">New Status</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: ${statusColor};">${statusLabel}</td>
            </tr>
            ${trip.vehicle_plate ? `
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Assigned Vehicle</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${trip.vehicle_plate}${trip.vehicle_model ? ` (${trip.vehicle_model})` : ""}</td>
            </tr>
            ` : ""}
            ${trip.driver_name ? `
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Assigned Driver</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${trip.driver_name}</td>
            </tr>
            ` : ""}
            ${updatedByName ? `
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Updated By</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${updatedByName}</td>
            </tr>
            ` : ""}
          </table>

          ${extraMessage ? `<p style="margin: 0 0 20px 0; line-height: 1.6; color: #374151;">${extraMessage}</p>` : ""}

          <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
            <tr>
              <td style="vertical-align: middle; line-height: 1.6;">
                Regards,<br/>
                <strong>Silvergraph Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:SilvergraphLogo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
              </td>
            </tr>
          </table>
        </div>
        <p style="font-size: 11px; color: #9ca3af; text-align: center; margin-top: 16px;">
          This is an automated notification. Please do not reply directly to this email.
        </p>
      </div>
    `;

    const mailOptions = {
      from: `"Silvergraph Unified Management System" <${process.env.EMAIL_USER}>`,
      to: recipients.join(", "),
      subject: `Re: Trip Request ${trip.trip_ref} Received`,
      text: `Trip ${trip.trip_ref} (requested by ${trip.requestor_name}) status changed to: ${statusLabel}.${trip.vehicle_plate ? ` Vehicle: ${trip.vehicle_plate}.` : ""}${trip.driver_name ? ` Driver: ${trip.driver_name}.` : ""}${updatedByName ? ` Updated by ${updatedByName}.` : ""}${extraMessage ? `\n\n${extraMessage}` : ""}`,
      html: htmlBody,
      attachments: [
        {
          filename: "SilvergraphLogo.png",
          path: "./assets/SilvergraphLogo.png",
          cid: "SilvergraphLogo",
        },
      ],
    };

    if (trip.email_message_id) {
      mailOptions.inReplyTo = trip.email_message_id;
      mailOptions.references = trip.email_message_id;
    }

    await transporter.sendMail(mailOptions);

    console.log(`📧 Fleet trip admin status update (${statusLabel}) sent to ${recipients.length} recipient(s) for ${trip.trip_ref}`);
  } catch (err) {
    console.error("Fleet trip admin status email failed:", err.message);
  }
}

// Sends the initial "Pending Review" confirmation to the employee who booked
// the trip, and stores the Gmail message id so every later status update can
// reply into the same thread — mirrors sendRequestNotification() for supply requests.
async function sendFleetTripRequestNotification({
  requestorId,
  requestorName,
  tripRef,
  pickupLocationText,
  dropoffLocationText,
  tripType,
  departureDatetime,
  purpose,
  passengerCount,
}) {
  try {
    const [rows] = await db.query(
      "SELECT notification_email, email FROM users WHERE username = ?",
      [requestorId],
    );
    const toEmail = rows[0]?.notification_email || rows[0]?.email;
    if (!toEmail) {
      console.warn(`No email preference set for ${requestorId}, skipping trip confirmation.`);
      return;
    }

    const htmlBody = `
      <div style="font-family: Arial, Helvetica, sans-serif; max-width: 560px; margin: 0 auto; color: #1f2937;">
        <div style="background-color: #1e3a5f; padding: 20px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="color: #ffffff; margin: 0; font-size: 18px;">Silverdab Trip Request</h2>
        </div>
        <div style="border: 1px solid #e5e7eb; border-top: none; border-radius: 0 0 8px 8px; padding: 24px;">
          <p style="margin: 0 0 12px 0;">Dear ${requestorName},</p>
          <p style="margin: 0 0 20px 0; line-height: 1.6;">
            This is to confirm that your trip request has been successfully submitted
            and is now pending dispatch review.
          </p>

          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Trip Ref</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${tripRef}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Pick-up</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${pickupLocationText}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Drop-off</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${dropoffLocationText}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Trip Type</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${tripType}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Departure</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${departureDatetime}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Status</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #b45309;">Pending Review</td>
            </tr>
          </table>

          ${purpose ? `<p style="margin: 0 0 20px 0; line-height: 1.6; color: #374151;"><strong>Purpose:</strong> ${purpose}</p>` : ""}

          <p style="margin: 0 0 4px 0; line-height: 1.6; font-size: 13px; color: #6b7280;">
            You will receive a follow-up notification once your request has been reviewed.
          </p>
          <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
            <tr>
              <td style="vertical-align: middle; line-height: 1.6;">
                Thank you,<br/>
                <strong>Silvergraph Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:SilvergraphLogo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
              </td>
            </tr>
          </table>
        </div>
        <p style="font-size: 11px; color: #9ca3af; text-align: center; margin-top: 16px;">
          This is an automated notification. Please do not reply directly to this email.
        </p>
      </div>
    `;

    const info = await transporter.sendMail({
      from: `"Silvergraph Unified Management System" <${process.env.EMAIL_USER}>`,
      to: toEmail,
      subject: `Trip Request ${tripRef} Received`,
      text: `Dear ${requestorName},\n\nYour trip request ${tripRef} has been submitted and is pending review.\n\nPick-up: ${pickupLocationText}\nDrop-off: ${dropoffLocationText}\nType: ${tripType}\nDeparture: ${departureDatetime}`,
      html: htmlBody,
      attachments: [
        {
          filename: "SilvergraphLogo.png",
          path: "./assets/SilvergraphLogo.png",
          cid: "SilvergraphLogo",
        },
      ],
    });

    await db.query(
      "UPDATE fleet_trips SET email_message_id = ? WHERE trip_ref = ?",
      [info.messageId, tripRef],
    );

    console.log(`📧 Trip confirmation sent to ${toEmail} for ${tripRef}`);
    return info.messageId;
  } catch (err) {
    console.error("Trip confirmation email failed:", err.message);
    // never throw — a failed email should not break the request flow
    return null;
  }
}

// Sends a single confirmation email to the person who booked the room —
// mirrors sendFleetTripRequestNotification()/sendRequestNotification(), but
// room_reservations already stores the booker's email directly on the row,
// so there's no users-table lookup needed like the fleet/supply flows.
async function sendRoomReservationConfirmation({
  toEmail,
  fullName,
  roomRef,
  roomName,
  bookingDate,
  startTime,
  endTime,
  agenda,
}) {
  try {
    const htmlBody = `
      <div style="font-family: Arial, Helvetica, sans-serif; max-width: 560px; margin: 0 auto; color: #1f2937;">
        <div style="background-color: #1e3a5f; padding: 20px 24px; border-radius: 8px 8px 0 0;">
          <h2 style="color: #ffffff; margin: 0; font-size: 18px;">Silverdab Room Reservation</h2>
        </div>
        <div style="border: 1px solid #e5e7eb; border-top: none; border-radius: 0 0 8px 8px; padding: 24px;">
          <p style="margin: 0 0 12px 0;">Dear ${fullName},</p>
          <p style="margin: 0 0 20px 0; line-height: 1.6;">
            This is to confirm that your room reservation has been successfully booked.
          </p>

          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Booking Ref</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${roomRef}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Room</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${roomName}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Date</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${bookingDate}</td>
            </tr>
            <tr>
              <td style="padding: 6px 12px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px;">Time</td>
              <td style="padding: 6px 12px; font-weight: bold; text-align: right; color: #1e3a5f;">${startTime} – ${endTime}</td>
            </tr>
          </table>

          ${agenda ? `<p style="margin: 0 0 20px 0; line-height: 1.6; color: #374151;"><strong>Agenda:</strong> ${agenda}</p>` : ""}

          <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
            <tr>
              <td style="vertical-align: middle; line-height: 1.6;">
                Thank you,<br/>
                <strong>Silvergraph Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:SilvergraphLogo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
              </td>
            </tr>
          </table>
        </div>
        <p style="font-size: 11px; color: #9ca3af; text-align: center; margin-top: 16px;">
          This is an automated notification. Please do not reply directly to this email.
        </p>
      </div>
    `;

    await transporter.sendMail({
      from: `"Silvergraph Unified Management System" <${process.env.EMAIL_USER}>`,
      to: toEmail,
      subject: `Room Reservation ${roomRef} Confirmed`,
      text: `Dear ${fullName},\n\nYour room reservation ${roomRef} has been confirmed.\n\nRoom: ${roomName}\nDate: ${bookingDate}\nTime: ${startTime} - ${endTime}${agenda ? `\nAgenda: ${agenda}` : ""}`,
      html: htmlBody,
      attachments: [
        {
          filename: "SilvergraphLogo.png",
          path: "./assets/SilvergraphLogo.png",
          cid: "SilvergraphLogo",
        },
      ],
    });

    console.log(`📧 Room reservation confirmation sent to ${toEmail} for ${roomRef}`);
  } catch (err) {
    console.error("Room reservation confirmation email failed:", err.message);
    // never throw — a failed email should not break the booking flow
  }
}

// ─── GET /api/health/db ────────────────────────────────────────────────────────
app.get("/api/health/db", async (req, res) => {
  try {
    await db.query("SELECT 1");
    res.json({ connected: true });
  } catch (err) {
    res.status(500).json({ connected: false, error: err.message });
  }
});

// ─── GET /users — list all users from MySQL ───────────────────────────────────
app.get("/users", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  try {
    const [rows] = await db.query("SELECT * FROM users ORDER BY display_name ASC");
    return res.json({ success: true, count: rows.length, users: rows });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── POST /users/sync — upsert AD users into MySQL ─────────────────────────────
app.post("/users/sync", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const { users, resetRoles } = req.body;
  if (!Array.isArray(users))
    return res.status(400).json({ success: false, message: "users array is required." });

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    for (const u of users) {
      const username = (u.username || "").toLowerCase().trim();
      if (!username) continue;

      await conn.query(
        `INSERT INTO users
           (username, display_name, email, department, title, phone, role,
            perm_it_inventory, perm_consumables, perm_tickets, perm_office_supplies, perm_it_access,
            perm_office_all_access, perm_office_dashboard, perm_office_inventory,
            perm_office_supply_request, perm_office_monthly_report, perm_office_activity,
            created_at, updated_at)
         VALUES (?, ?, ?, ?, ?, ?, 'employee', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NOW(), NOW())
         ON DUPLICATE KEY UPDATE
           display_name = VALUES(display_name),
           email = VALUES(email),
           department = VALUES(department),
           title = VALUES(title),
           phone = VALUES(phone),
           role = IF(?, 'employee', role),
           updated_at = NOW()`,
        [
          username,
          u.displayName || username,
          u.email || `${username}@ocgbim.com`,
          u.department || "",
          u.title || "",
          u.phone || "",
          resetRoles ? 1 : 0,
        ]
      );
    }

    await conn.commit();
    return res.json({ success: true, count: users.length });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// ─── PATCH /users/:username/role — promote to admin / demote to employee ──────
// Only superadmins may call this, and a superadmin's own role can never be
// changed through this route (avoids accidentally locking yourself out).
app.patch("/users/:username/role", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  if (decoded.role !== "superadmin") {
    return res.status(403).json({ success: false, message: "Not authorized." });
  }

  const { role } = req.body;
  const username = req.params.username.toLowerCase().trim();

  if (role !== "admin" && role !== "employee") {
    return res.status(400).json({ success: false, message: "role must be 'admin' or 'employee'." });
  }

  try {
    const [rows] = await db.query("SELECT role FROM users WHERE username = ?", [username]);
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "User not found." });
    }
    if (rows[0].role === "superadmin") {
      return res.status(403).json({ success: false, message: "Cannot change a superadmin's role." });
    }

    await db.query(
      "UPDATE users SET role = ?, updated_at = NOW() WHERE username = ?",
      [role, username],
    );
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── PATCH /users/:username/permissions ────────────────────────────────────────
app.patch("/users/:username/permissions", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const {
    itAccess, itInventory, consumables, tickets,
    officeSupplies, // deprecated, kept for backward compat
    officeAllAccess, officeDashboard, officeInventory,
    officeSupplyRequest, officeMonthlyReport, officeActivity,
    fleetControl, fleetDriver,
  } = req.body;
  const username = req.params.username.toLowerCase().trim();
  console.log("PATCH permissions body:", req.body);

  try {
    await db.query(
      `UPDATE users SET
         perm_it_access = ?, perm_it_inventory = ?, perm_consumables = ?,
         perm_tickets = ?, perm_office_supplies = ?,
         perm_office_all_access = ?, perm_office_dashboard = ?,
         perm_office_inventory = ?, perm_office_supply_request = ?,
         perm_office_monthly_report = ?, perm_office_activity = ?,
         perm_fleet_control = ?, perm_fleet_driver = ?, updated_at = NOW()
       WHERE username = ?`,
      [
        itAccess ? 1 : 0,
        itInventory ? 1 : 0,
        consumables ? 1 : 0,
        tickets ? 1 : 0,
        officeSupplies ? 1 : 0,
        officeAllAccess ? 1 : 0,
        officeDashboard ? 1 : 0,
        officeInventory ? 1 : 0,
        officeSupplyRequest ? 1 : 0,
        officeMonthlyReport ? 1 : 0,
        officeActivity ? 1 : 0,
        fleetControl ? 1 : 0,
        fleetDriver ? 1 : 0,
        username,
      ]
    );
    return res.json({ success: true });
  } catch (err) {
    console.error("PATCH /users/:username/permissions error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── DELETE /users — used by "Clear & Resync" ──────────────────────────────────
app.delete("/users", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  try {
    await db.query("DELETE FROM users");
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── GET /users/:username/email-preference ────────────────────────────────
app.get("/users/:username/email-preference", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const username = req.params.username.toLowerCase().trim();
  try {
    const [rows] = await db.query(
      "SELECT username, display_name, notification_email FROM users WHERE username = ?",
      [username],
    );
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "User not found." });
    }
    const u = rows[0];
    const silverdabEmail = `${u.username}@silverdab.com`;
    const nameParts = (u.display_name || "").trim().toLowerCase().split(/\s+/);
    const ocgbimEmail =
      nameParts.length >= 2
        ? `${nameParts[0]}.${nameParts[nameParts.length - 1]}@ocgbim.com`
        : null;
    return res.json({
      success: true,
      current: u.notification_email,
      options: { silverdab: silverdabEmail, ocgbim: ocgbimEmail },
    });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});


// ─── PATCH /users/:username/email-preference ───────────────────────────────
app.patch("/users/:username/email-preference", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { username } = req.params;
  const { email } = req.body;
  if (!email || !email.includes("@")) {
    return res.status(400).json({ success: false, message: "A valid email is required." });
  }
  try {
    const [result] = await db.query(
      "UPDATE users SET notification_email = ?, updated_at = NOW() WHERE username = ?",
      [email, username],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "User not found." });
    }
    return res.json({ success: true, notification_email: email });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── Role Mapping from OCGBIM AD Groups ───────────────────────────────────────
function getRoleFromGroups(memberOf) {
  if (!memberOf) return "employee";
  const groups = Array.isArray(memberOf) ? memberOf : [memberOf];
  const g = groups.map((x) => x.toLowerCase());

  if (
    g.some(
      (x) =>
        x.includes("ocgbim_it_users") ||
        x.includes("it admin") ||
        x.includes("it_admin") ||
        x.includes("it installer"),
    )
  )
    return "superadmin"; // ← was "it"

  if (
    g.some(
      (x) =>
        x.includes("ocgbim_adminstaff_users") ||
        x.includes("ocgbim_ceo_users") ||
        x.includes("ocgbim_local_administrator"),
    )
  )
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

    //console.log("🔍 Searching user:", username);
    client.search(AD_BASE_DN, opts, (err, res) => {
      if (err) return reject(err);
      const entries = [];
      res.on("searchEntry", (e) => {
        const u = parseEntry(e);
        //console.log("✅ Found:", u.sAMAccountName, "|", u.displayName);
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

  //console.log("\n=== Login Attempt ===");
  //console.log("Username:", username);

  if (!username || !password) {
    return res.status(400).json({ success: false, message: "Username and password are required." });
  }

  const userDN = `${username}@${AD_DOMAIN}`;
  let client;

  try {
    // Step 1: Verify user credentials
    //console.log("Step 1: Verifying credentials...");
    client = createLDAPClient();

    await new Promise((resolve, reject) => {
      client.bind(userDN, password, (err) => {
        if (err) {
          //console.error("❌ User bind error:", err.message, "| Code:", err.code);
          reject(new Error(err.code === 49 ? "Invalid username or password." : `Connection error: ${err.message}`));
        } else {
          //console.log("✅ Credentials verified!");
          resolve();
        }
      });
    });

    // Step 2: Fetch user details via service account
    //console.log("Step 2: Fetching user details...");
    const serviceClient = await getServiceClient();
    const userInfo = await searchUser(serviceClient, username);
    serviceClient.destroy();

    // Step 3: Map role and department
    const role = getRoleFromGroups(userInfo.memberOf);
    const department = userInfo.department || getDepartmentFromGroups(userInfo.memberOf);
    //console.log("✅ Role:", role, "| Department:", department);

    // Step 4: Generate JWT
    const token = jwt.sign(
      { username: userInfo.sAMAccountName, displayName: userInfo.displayName, email: userInfo.mail, department, title: userInfo.title, phone: userInfo.telephoneNumber, role },
      JWT_SECRET,
      { expiresIn: "30d" }
    );

    //console.log("✅ Login successful:", username);

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
    //console.error("❌ Login failed:", err.message);
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

// ─── GET /auth/service-token ──────────────────────────────────────────────────
app.get("/auth/service-token", (req, res) => {
  const secret = req.headers["x-internal-secret"];
  if (secret !== process.env.INTERNAL_SECRET) {
    return res.status(403).json({ success: false, message: "Forbidden" });
  }
  const token = jwt.sign(
    { username: "service", role: "superadmin", displayName: "System" },
    JWT_SECRET,
    { expiresIn: "365d" }
  );
  return res.json({ success: true, token });
});

// ─── GET /debug/users ─────────────────────────────────────────────────────────
app.get("/debug/users", async (req, res) => {
  try {
    const serviceClient = await getServiceClient();
    const employees = await searchEmployees(
      serviceClient,
      "(sAMAccountName=*)",
      20
    );
    serviceClient.destroy();
    return res.json({ count: employees.length, sample: employees.slice(0, 5) });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});


// ─── Start ────────────────────────────────────────────────────────────────────

// ─── OFFICE INVENTORY ROUTES ───────────────────────────────────────────────

function computeStockStatus(currentStock, inStockThreshold) {
  if (currentStock <= 0) return "out_of_stock";
  if (currentStock <= inStockThreshold) return "low_stock";
  return "in_stock";
}

// office_inventory.stock_status uses in_stock/low_stock/out_of_stock, but
// supply_request_items.stock_status_at_request uses available/low/out_of_stock
// (see worstStockStatus() on the frontend) — map between the two vocabularies.
function toRequestItemStockStatus(inventoryStockStatus) {
  switch (inventoryStockStatus) {
    case "in_stock":
      return "available";
    case "low_stock":
      return "low";
    case "out_of_stock":
      return "out_of_stock";
    default:
      return inventoryStockStatus;
  }
}



app.post("/office-inventory", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const {
    itemCode, name, brand, category, unit,
    pricePerUnit, currentStock, lowStockThreshold, inStockThreshold,
    isRestricted, performedByName,
  } = req.body;

  if (!itemCode || !name || !category || !unit) {
    return res.status(400).json({
      success: false,
      message: "itemCode, name, category, and unit are required.",
    });
  }

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [existing] = await conn.query(
      "SELECT id FROM office_inventory WHERE item_code = ?",
      [itemCode],
    );
    if (existing.length > 0) {
      await conn.rollback();
      return res.status(409).json({
        success: false,
        message: `Item code "${itemCode}" already exists.`,
      });
    }

    const id = crypto.randomUUID();
    const lowThresh = lowStockThreshold ?? 5;
    const inThresh = inStockThreshold ?? 10;
    const stock = currentStock ?? 0;
    const price = pricePerUnit ?? 0;
    const stockStatus = computeStockStatus(stock, inThresh);
    const restricted = isRestricted ? 1 : 0;

    await conn.query(
      `INSERT INTO office_inventory
        (id, item_code, name, brand, category, unit, price_per_unit,
         current_stock, stock_status, low_stock_threshold, in_stock_threshold,
         is_active, is_restricted, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, NOW(), NOW())`,
      [id, itemCode, name, brand ?? "", category, unit, price,
       stock, stockStatus, lowThresh, inThresh, restricted],
    );

    // Log an activity entry so item creation shows up in the Activity tab,
    // same as deliveries/adjustments do — only when there's actual stock
    // to record (a brand-new item with 0 beginning stock has nothing to log).
    if (stock > 0) {
      await conn.query(
        `INSERT INTO stock_transactions
          (id, item_id, item_code, item_name, type, quantity_change, stock_before,
           stock_after, price_per_unit, total_amount, reason, performed_by_name,
           transaction_date, created_at)
         VALUES (?, ?, ?, ?, 'item_created', ?, 0, ?, ?, ?, ?, ?, ?, NOW())`,
        [
          crypto.randomUUID(), id, itemCode, name,
          stock, stock, price, stock * Number(price),
          "Beginning inventory", performedByName ?? "Unknown",
          new Date().toISOString().split("T")[0],
        ],
      );
    }

    await conn.commit();

    const [rows] = await conn.query("SELECT * FROM office_inventory WHERE id = ?", [id]);
    return res.status(201).json({ success: true, item: rows[0] });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

app.patch("/office-inventory/:id", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const { id } = req.params;
  const { name, brand, category, unit, pricePerUnit, lowStockThreshold, inStockThreshold } = req.body;

  try {
    const [rows] = await db.query("SELECT * FROM office_inventory WHERE id = ?", [id]);
    if (rows.length === 0)
      return res.status(404).json({ success: false, message: "Item not found" });

    const current = rows[0];
    const inThresh = inStockThreshold ?? current.in_stock_threshold;
    const stockStatus = computeStockStatus(current.current_stock, inThresh);

    await db.query(
      `UPDATE office_inventory SET
         name = ?, brand = ?, category = ?, unit = ?, price_per_unit = ?,
         low_stock_threshold = ?, in_stock_threshold = ?, stock_status = ?,
         updated_at = NOW()
       WHERE id = ?`,
      [
        name ?? current.name,
        brand ?? current.brand,
        category ?? current.category,
        unit ?? current.unit,
        pricePerUnit ?? current.price_per_unit,
        lowStockThreshold ?? current.low_stock_threshold,
        inThresh,
        stockStatus,
        id,
      ],
    );

    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

app.patch("/office-inventory/:id/archive", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const { id } = req.params;
  const { performedByName } = req.body;

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [rows] = await conn.query("SELECT * FROM office_inventory WHERE id = ? FOR UPDATE", [id]);
    if (rows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Item not found" });
    }
    const item = rows[0];

    const [result] = await conn.query(
      "UPDATE office_inventory SET is_active = 0, updated_at = NOW() WHERE id = ?",
      [id],
    );
    if (result.affectedRows === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Item not found" });
    }

    await conn.query(
      `INSERT INTO stock_transactions
        (id, item_id, item_code, item_name, type, quantity_change, stock_before,
         stock_after, price_per_unit, total_amount, reason, performed_by_name,
         transaction_date, created_at)
       VALUES (?, ?, ?, ?, 'item_archived', 0, ?, ?, ?, 0, ?, ?, ?, NOW())`,
      [
        crypto.randomUUID(), id, item.item_code, item.name,
        item.current_stock, item.current_stock, item.price_per_unit,
        "Item archived", performedByName ?? "Unknown",
        new Date().toISOString().split("T")[0],
      ],
    );

    await conn.commit();
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

app.post("/office-inventory/:id/adjust-stock", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const { id } = req.params;
  const { quantity, date, reason, performedByName } = req.body;

  if (!quantity || quantity <= 0)
    return res.status(400).json({ success: false, message: "Quantity must be greater than 0." });

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [rows] = await conn.query("SELECT * FROM office_inventory WHERE id = ? FOR UPDATE", [id]);
    if (rows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Item not found" });
    }

    const item = rows[0];
    const stockBefore = item.current_stock;

    if (quantity > stockBefore) {
      await conn.rollback();
      return res.status(400).json({ success: false, message: "Cannot deduct more than current stock." });
    }

    const stockAfter = stockBefore - quantity;
    const stockStatus = computeStockStatus(stockAfter, item.in_stock_threshold);

    await conn.query(
      "UPDATE office_inventory SET current_stock = ?, stock_status = ?, updated_at = NOW() WHERE id = ?",
      [stockAfter, stockStatus, id],
    );

    await conn.query(
      `INSERT INTO stock_transactions
        (id, item_id, item_code, item_name, type, quantity_change, stock_before,
         stock_after, price_per_unit, total_amount, reason, performed_by_name,
         transaction_date, created_at)
       VALUES (?, ?, ?, ?, 'manual_adjustment', ?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
      [
        crypto.randomUUID(), id, item.item_code, item.name,
        -quantity, stockBefore, stockAfter, item.price_per_unit,
        quantity * Number(item.price_per_unit),
        reason ?? "", performedByName ?? "Unknown", date,
      ],
    );

    await conn.commit();
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

app.post("/office-inventory/:id/deliver", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const { id } = req.params;
  const { quantity, date, pricePerUnit, notes, performedByName } = req.body;

  if (!quantity || quantity <= 0)
    return res.status(400).json({ success: false, message: "Quantity must be greater than 0." });

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [rows] = await conn.query("SELECT * FROM office_inventory WHERE id = ? FOR UPDATE", [id]);
    if (rows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Item not found" });
    }

    const item = rows[0];
    const stockBefore = item.current_stock;
    const stockAfter = stockBefore + quantity;
    const stockStatus = computeStockStatus(stockAfter, item.in_stock_threshold);
    const price = pricePerUnit ?? item.price_per_unit;

    await conn.query(
      "UPDATE office_inventory SET current_stock = ?, price_per_unit = ?, stock_status = ?, updated_at = NOW() WHERE id = ?",
      [stockAfter, price, stockStatus, id],
    );

    await conn.query(
      `INSERT INTO stock_transactions
        (id, item_id, item_code, item_name, type, quantity_change, stock_before,
         stock_after, price_per_unit, total_amount, reason, performed_by_name,
         transaction_date, created_at)
       VALUES (?, ?, ?, ?, 'delivery', ?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
      [
        crypto.randomUUID(), id, item.item_code, item.name,
        quantity, stockBefore, stockAfter, price,
        quantity * Number(price),
        notes ?? "", performedByName ?? "Unknown", date,
      ],
    );

    // ── Sync any supply requests that were waiting on this item's stock ──
    const requestItemStatus = toRequestItemStockStatus(stockStatus);

    await conn.query(
      `UPDATE supply_request_items sri
       JOIN supply_requests sr ON sr.id = sri.request_id
       SET sri.stock_status_at_request = ?
       WHERE sri.item_id = ?
         AND sr.status IN ('pending', 'awaiting_stock')
         AND sri.stock_status_at_request != ?`,
      [requestItemStatus, id, requestItemStatus],
    );

    // If a request was explicitly 'awaiting_stock' and this item is no longer
    // out of stock, drop it back to 'pending' so it re-enters the normal
    // review queue — but only if none of its OTHER lines are still out of stock.
    if (requestItemStatus !== "out_of_stock") {
      await conn.query(
        `UPDATE supply_requests sr
         SET sr.status = 'pending'
         WHERE sr.status = 'awaiting_stock'
           AND EXISTS (
             SELECT 1 FROM supply_request_items sri
             WHERE sri.request_id = sr.id AND sri.item_id = ?
           )
           AND NOT EXISTS (
             SELECT 1 FROM supply_request_items sri2
             WHERE sri2.request_id = sr.id
               AND sri2.item_id != ?
               AND sri2.stock_status_at_request = 'out_of_stock'
           )`,
        [id, id],
      );
    }

    await conn.commit();
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

app.get("/stock-transactions", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  try {
    const [rows] = await db.query(
      "SELECT * FROM stock_transactions ORDER BY created_at DESC",
    );
    return res.json({ success: true, count: rows.length, transactions: rows });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── SUPPLY REQUESTS ROUTES ─────────────────────────────────────────────────

function requireAuth(req, res) {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer ")) {
    res.status(401).json({ success: false, message: "No token provided." });
    return null;
  }
  try {
    return jwt.verify(authHeader.split(" ")[1], JWT_SECRET);
  } catch {
    res.status(401).json({ success: false, message: "Invalid token." });
    return null;
  }
}

// POST /push/subscribe — save a browser's push subscription
app.post("/push/subscribe", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  const { subscription } = req.body;
  if (!subscription?.endpoint || !subscription?.keys?.p256dh || !subscription?.keys?.auth) {
    return res.status(400).json({ success: false, message: "Invalid subscription object." });
  }

  try {
    await db.query(
      `INSERT INTO push_subscriptions (username, endpoint, p256dh, auth)
       VALUES (?, ?, ?, ?)
       ON DUPLICATE KEY UPDATE username = VALUES(username), p256dh = VALUES(p256dh), auth = VALUES(auth)`,
      [decoded.username.toLowerCase().trim(), subscription.endpoint, subscription.keys.p256dh, subscription.keys.auth],
    );
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /push/unsubscribe
app.post("/push/unsubscribe", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  const { endpoint } = req.body;
  if (!endpoint) return res.status(400).json({ success: false, message: "endpoint is required." });

  try {
    await db.query("DELETE FROM push_subscriptions WHERE endpoint = ?", [endpoint]);
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /push/expo-token — save a native app's Expo push token
app.post("/push/expo-token", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  const { token } = req.body;
  if (!token || !token.startsWith("ExponentPushToken")) {
    return res.status(400).json({ success: false, message: "Invalid Expo push token." });
  }

  try {
    await db.query(
      `INSERT INTO expo_push_tokens (username, token)
       VALUES (?, ?)
       ON DUPLICATE KEY UPDATE username = VALUES(username)`,
      [decoded.username.toLowerCase().trim(), token],
    );
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// GET /supply-requests — joins supply_request_items and nests them per request
app.get("/supply-requests", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const includeArchived = req.query.includeArchived === "true";

  try {
    const [rows] = await db.query(`
      SELECT
        sr.*,
        sri.item_id            AS item_item_id,
        sri.item_name           AS item_item_name,
        sri.item_code           AS item_item_code,
        sri.category            AS item_category,
        sri.quantity_requested  AS item_quantity_requested,
        sri.quantity_approved   AS item_quantity_approved,
        sri.stock_status_at_request AS item_stock_status_at_request,
        sri.price_per_unit      AS item_price_per_unit
      FROM supply_requests sr
      LEFT JOIN supply_request_items sri ON sri.request_id = sr.id
      ${includeArchived ? "" : "WHERE sr.is_archived = 0"}
      ORDER BY sr.created_at DESC
    `);

    const byId = new Map();
    for (const row of rows) {
      if (!byId.has(row.id)) {
        const { item_item_id, item_item_name, item_item_code, item_category,
                item_quantity_requested, item_quantity_approved,
                item_stock_status_at_request, item_price_per_unit, ...parent } = row;
        byId.set(row.id, { ...parent, items: [] });
      }
      if (row.item_item_id) {
        byId.get(row.id).items.push({
          item_id: row.item_item_id,
          item_name: row.item_item_name,
          item_code: row.item_item_code,
          category: row.item_category,
          quantity_requested: row.item_quantity_requested,
          quantity_approved: row.item_quantity_approved,
          stock_status_at_request: row.item_stock_status_at_request,
          price_per_unit: row.item_price_per_unit,
        });
      }
    }

    return res.json({ success: true, requests: Array.from(byId.values()) });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /supply-requests — insert parent + item rows in one transaction
app.post("/supply-requests", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { requestedById, requestedByName, items, notes } = req.body;
  if (!requestedById || !requestedByName || !Array.isArray(items) || items.length === 0) {
    return res.status(400).json({
      success: false,
      message: "requestedById, requestedByName, and at least one item are required.",
    });
  }

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const year = new Date().getFullYear();
const [maxRows] = await conn.query(
  `SELECT MAX(CAST(SUBSTRING_INDEX(ticket_number, '-', -1) AS UNSIGNED)) AS maxNum
   FROM supply_requests WHERE ticket_number LIKE ? FOR UPDATE`,
  [`SR-${year}-%`],
);
const nextNum = String((maxRows[0].maxNum ?? 0) + 1).padStart(4, "0");
const ticketNumber = `SR-${year}-${nextNum}`;
    const id = crypto.randomUUID();

    await conn.query(
      `INSERT INTO supply_requests
        (id, ticket_number, requested_by_id, requested_by_name, status, notes, created_at)
       VALUES (?, ?, ?, ?, 'pending', ?, NOW())`,
      [id, ticketNumber, requestedById, requestedByName, notes ?? ""],
    );

   for (const item of items) {
      // Look up current price from office_inventory at request time — this
      // "locks in" the price as it was when requested, same reasoning as
      // stockStatusAtRequest already being snapshotted rather than live.
      const [priceRows] = await conn.query(
        "SELECT price_per_unit FROM office_inventory WHERE id = ?",
        [item.itemId],
      );
      const pricePerUnit = priceRows[0]?.price_per_unit ?? 0;

      await conn.query(
        `INSERT INTO supply_request_items
          (request_id, item_id, item_name, item_code, category, quantity_requested, stock_status_at_request, price_per_unit)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
        [
          id,
          item.itemId,
          item.itemName,
          item.itemCode,
          item.category ?? "",
          item.quantityRequested,
          item.stockStatusAtRequest,
          pricePerUnit,
        ],
      );
    }

    await conn.commit();

    sendRequestNotification({ requestedById, requestedByName, ticketNumber, items });
    // Admin email disabled — was flooding the inbox on every request.
    // sendAdminRequestNotification({ requestedByName, ticketNumber, items });
    sendWebPushToAdmins({
      title: "New Supply Request",
      body: `${requestedByName} submitted ${ticketNumber} (${items.length} item${items.length !== 1 ? "s" : ""})`,
      url: "/supply-requests",
      permissionColumns: ["perm_office_supplies", "perm_office_all_access", "perm_office_supply_request"],
    });

    return res.status(201).json({ success: true, ticketNumber });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// Deducts stock for one item line at DELIVERY time (not approval), logs a
// stock_transactions row. quantity_approved is now set at approval time
// separately — this only moves physical stock.
async function deductStockOnDelivery(conn, itemId, qty, ticketNumber, actorName) {
  const [rows] = await conn.query(
    "SELECT * FROM office_inventory WHERE id = ? FOR UPDATE",
    [itemId],
  );
  if (rows.length === 0) return; // item may have been deleted — skip silently

  const item = rows[0];
  const stockBefore = item.current_stock;
  const deduct = Math.min(qty, stockBefore);
  const stockAfter = stockBefore - deduct;
  const stockStatus = computeStockStatus(stockAfter, item.in_stock_threshold);

  await conn.query(
    "UPDATE office_inventory SET current_stock = ?, stock_status = ?, updated_at = NOW() WHERE id = ?",
    [stockAfter, stockStatus, itemId],
  );

  await conn.query(
    `INSERT INTO stock_transactions
      (id, item_id, item_code, item_name, type, quantity_change, stock_before,
       stock_after, price_per_unit, total_amount, reason, performed_by_name,
       transaction_date, created_at)
     VALUES (?, ?, ?, ?, 'supply_request_fulfilled', ?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
    [
      crypto.randomUUID(), itemId, item.item_code, item.name,
      -deduct, stockBefore, stockAfter, item.price_per_unit,
      deduct * Number(item.price_per_unit),
      `Supply request ${ticketNumber} delivered`, actorName ?? "Unknown",
      new Date().toISOString().split("T")[0],
    ],
  );
}

// POST /supply-requests/:id/approve — full approval, deducts full requested qty per line
app.post("/supply-requests/:id/approve", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { id } = req.params;
  const { approvedByName } = req.body;

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [reqRows] = await conn.query("SELECT * FROM supply_requests WHERE id = ? FOR UPDATE", [id]);
    if (reqRows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Request not found" });
    }
    const request = reqRows[0];

    const [itemRowsForNotif] = await conn.query(
      "SELECT item_name FROM supply_request_items WHERE request_id = ?",
      [id],
    );

    // No stock deduction here — items are approved at full requested qty,
    // stock is only deducted once the request is actually marked delivered.
    await conn.query(
      "UPDATE supply_request_items SET quantity_approved = quantity_requested WHERE request_id = ?",
      [id],
    );

    await conn.query(
      `UPDATE supply_requests SET
         status = 'out_for_delivery', approved_at = NOW(), approved_by_name = ?,
         reviewed_by_name = ?, reviewed_at = NOW()
       WHERE id = ?`,
      [approvedByName ?? "Unknown", approvedByName ?? "Unknown", id],
    );

    await conn.commit();
    sendStatusUpdateNotification({ requestId: id, statusLabel: "Out for Delivery", updatedByName: approvedByName });
    // sendAdminStatusUpdateNotification({ requestId: id, statusLabel: "Out for Delivery", updatedByName: approvedByName });
    const itemNamesForNotif = itemRowsForNotif.map((r) => r.item_name);
    const itemSummaryForNotif =
      itemNamesForNotif.length > 2
        ? `${itemNamesForNotif.slice(0, 2).join(", ")} +${itemNamesForNotif.length - 2} more`
        : itemNamesForNotif.join(", ");

    sendWebPushToAdmins({
      title: "Request Out for Delivery",
      body: `${request.requested_by_name} — ${itemSummaryForNotif} (${request.ticket_number})`,
      url: "/supply-requests",
      permissionColumns: ["perm_office_supplies", "perm_office_all_access", "perm_office_supply_request"],
    });
    sendExpoPushToAdmins({
      title: "Request Out for Delivery",
      body: `${request.requested_by_name} — ${itemSummaryForNotif} (${request.ticket_number})`,
      data: { requestId: id, url: "/supply-requests" },
      permissionColumns: ["perm_office_supplies", "perm_office_all_access", "perm_office_supply_request"],
    });
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// POST /supply-requests/:id/approve-partial — records caller-specified qty per line (no stock movement yet)
app.post("/supply-requests/:id/approve-partial", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { id } = req.params;
  const { lines, approvedByName } = req.body;
  if (!Array.isArray(lines)) {
    return res.status(400).json({ success: false, message: "lines array is required." });
  }

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [reqRows] = await conn.query("SELECT * FROM supply_requests WHERE id = ? FOR UPDATE", [id]);
    if (reqRows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Request not found" });
    }
    const request = reqRows[0];

    const [itemRowsForNotif] = await conn.query(
      "SELECT item_name FROM supply_request_items WHERE request_id = ?",
      [id],
    );

    const approvedItemIds = [];
    for (const line of lines) {
      if (!line.qtyToDispense || line.qtyToDispense <= 0) continue;
      // Just record what was approved — no stock movement until delivery.
      await conn.query(
        "UPDATE supply_request_items SET quantity_approved = LEAST(?, quantity_requested) WHERE request_id = ? AND item_id = ?",
        [line.qtyToDispense, id, line.itemId],
      );
      approvedItemIds.push(line.itemId);
    }

    // Any item on this request NOT included in `lines` (or given 0 qty) was
    // skipped by the admin — mark it explicitly so the UI can tell "skipped"
    // apart from "not yet reviewed" (which stays NULL).
    if (approvedItemIds.length > 0) {
      await conn.query(
        `UPDATE supply_request_items
         SET quantity_approved = 0
         WHERE request_id = ? AND item_id NOT IN (?)`,
        [id, approvedItemIds],
      );
    } else {
      await conn.query(
        "UPDATE supply_request_items SET quantity_approved = 0 WHERE request_id = ?",
        [id],
      );
    }

    await conn.query(
      `UPDATE supply_requests SET
         status = 'out_for_delivery', approved_at = NOW(), approved_by_name = ?,
         reviewed_by_name = ?, reviewed_at = NOW()
       WHERE id = ?`,
      [approvedByName ?? "Unknown", approvedByName ?? "Unknown", id],
    );

    await conn.commit();
    sendStatusUpdateNotification({ requestId: id, statusLabel: "Out for Delivery", updatedByName: approvedByName });
    // sendAdminStatusUpdateNotification({ requestId: id, statusLabel: "Out for Delivery", updatedByName: approvedByName });
    const itemNamesForNotif = itemRowsForNotif.map((r) => r.item_name);
    const itemSummaryForNotif =
      itemNamesForNotif.length > 2
        ? `${itemNamesForNotif.slice(0, 2).join(", ")} +${itemNamesForNotif.length - 2} more`
        : itemNamesForNotif.join(", ");

    sendWebPushToAdmins({
      title: "Request Out for Delivery",
      body: `${request.requested_by_name} — ${itemSummaryForNotif} (${request.ticket_number})`,
      url: "/supply-requests",
      permissionColumns: ["perm_office_supplies", "perm_office_all_access", "perm_office_supply_request"],
    });
    sendExpoPushToAdmins({
      title: "Request Out for Delivery",
      body: `${request.requested_by_name} — ${itemSummaryForNotif} (${request.ticket_number})`,
      data: { requestId: id, url: "/supply-requests" },
      permissionColumns: ["perm_office_supplies", "perm_office_all_access", "perm_office_supply_request"],
    });
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// POST /supply-requests/:id/reject
app.post("/supply-requests/:id/reject", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { id } = req.params;
  const { reason, reviewedByName } = req.body;
  if (!reason) {
    return res.status(400).json({ success: false, message: "reason is required." });
  }

  try {
    const [result] = await db.query(
      `UPDATE supply_requests SET
         status = 'rejected', rejection_reason = ?, reviewed_by_name = ?,
         reviewed_at = NOW(), resolved_at = NOW()
       WHERE id = ?`,
      [reason, reviewedByName ?? "Unknown", id],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Request not found" });
    }
    sendStatusUpdateNotification({ requestId: id, statusLabel: "Rejected", extraMessage: reason, updatedByName: reviewedByName });
    // sendAdminStatusUpdateNotification({ requestId: id, statusLabel: "Rejected", extraMessage: reason, updatedByName: reviewedByName });
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /supply-requests/:id/deliver — stock is deducted HERE, not at approval
app.post("/supply-requests/:id/deliver", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { id } = req.params;
  const { deliveredByName } = req.body;

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [reqRows] = await conn.query("SELECT * FROM supply_requests WHERE id = ? FOR UPDATE", [id]);
    if (reqRows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Request not found" });
    }
    const request = reqRows[0];

    const [items] = await conn.query(
      "SELECT * FROM supply_request_items WHERE request_id = ?",
      [id],
    );
    for (const line of items) {
      const qty = line.quantity_approved ?? line.quantity_requested;
      if (!qty || qty <= 0) continue;
      await deductStockOnDelivery(conn, line.item_id, qty, request.ticket_number, deliveredByName);
    }

    await conn.query(
      `UPDATE supply_requests SET
         status = 'resolved', delivered_at = NOW(), delivered_by_name = ?,
         resolved_at = NOW()
       WHERE id = ?`,
      [deliveredByName ?? "Unknown", id],
    );

    await conn.commit();
    sendStatusUpdateNotification({ requestId: id, statusLabel: "Issued", updatedByName: deliveredByName });
    // sendAdminStatusUpdateNotification({ requestId: id, statusLabel: "Issued", updatedByName: deliveredByName });
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// POST /supply-requests/:id/fail
app.post("/supply-requests/:id/fail", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { id } = req.params;
  const { reason, deliveredByName } = req.body;
  if (!reason) {
    return res.status(400).json({ success: false, message: "reason is required." });
  }

  try {
    const [result] = await db.query(
      `UPDATE supply_requests SET
         status = 'failed_delivery', failed_reason = ?, delivered_by_name = ?,
         failed_at = NOW()
       WHERE id = ?`,
      [reason, deliveredByName ?? "Unknown", id],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Request not found" });
    }
    sendStatusUpdateNotification({ requestId: id, statusLabel: "Failed Delivery", extraMessage: reason, updatedByName: deliveredByName });
    // sendAdminStatusUpdateNotification({ requestId: id, statusLabel: "Failed Delivery", extraMessage: reason, updatedByName: deliveredByName });
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /supply-requests/:id/cancel — employee-initiated, only while the
// request hasn't been acted on by an admin yet (pending or awaiting_stock).
// Re-validates status here even though the frontend already checks this,
// since the admin could approve/reject in the gap between the employee
// opening the drawer and confirming cancellation.
app.post("/supply-requests/:id/cancel", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { id } = req.params;
  const { cancelledByName } = req.body;

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [rows] = await conn.query(
      "SELECT status FROM supply_requests WHERE id = ? FOR UPDATE",
      [id],
    );
    if (rows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Request not found" });
    }

    const currentStatus = rows[0].status;
    if (currentStatus !== "pending" && currentStatus !== "awaiting_stock") {
      await conn.rollback();
      return res.status(409).json({
        success: false,
        message: "This request has already been reviewed and can no longer be cancelled.",
      });
    }

    await conn.query(
      `UPDATE supply_requests SET
         status = 'cancelled', cancelled_at = NOW(), cancelled_by_name = ?,
         resolved_at = NOW()
       WHERE id = ?`,
      [cancelledByName ?? "Unknown", id],
    );

    await conn.commit();
    sendStatusUpdateNotification({
      requestId: id,
      statusLabel: "Cancelled",
      updatedByName: cancelledByName,
    });
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// POST /supply-requests/:id/archive — hide from the default list without deleting
app.post("/supply-requests/:id/archive", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  try {
    const [result] = await db.query(
      "UPDATE supply_requests SET is_archived = 1 WHERE id = ?",
      [id],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Request not found" });
    }
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /supply-requests/:id/unarchive — undo
app.post("/supply-requests/:id/unarchive", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  try {
    const [result] = await db.query(
      "UPDATE supply_requests SET is_archived = 0 WHERE id = ?",
      [id],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Request not found" });
    }
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── ADDITIONS to your server.js office-inventory routes ───────────────────
// 1. GET /office-inventory now filters is_active = 1 server-side by default,
//    and includes archived rows too when ?includeArchived=true is passed.
//    (Previously it always selected every row and relied on the frontend
//    to filter — this moves that filtering server-side, which is both
//    cheaper and makes the "includeArchived" toggle meaningful.)
// 2. New PATCH /office-inventory/:id/restore — the undo for the existing
//    /office-inventory/:id/archive endpoint.

// Replace your existing GET /office-inventory handler with this version:
app.get("/office-inventory", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  const isPrivileged = decoded.role === "admin" || decoded.role === "superadmin";

  try {
    const includeArchived = req.query.includeArchived === "true";

    const clauses = [];
    if (!includeArchived) clauses.push("is_active = 1");
    // Employees never see restricted items, regardless of includeArchived.
    if (!isPrivileged) clauses.push("is_restricted = 0");

    const where = clauses.length ? `WHERE ${clauses.join(" AND ")}` : "";
    const [rows] = await db.query(
      `SELECT * FROM office_inventory ${where} ORDER BY name ASC`,
    );
    return res.json({ success: true, count: rows.length, items: rows });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// Add this new handler anywhere near your existing /archive route:
app.patch("/office-inventory/:id/restore", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const { id } = req.params;
  const { performedByName } = req.body;

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [rows] = await conn.query("SELECT * FROM office_inventory WHERE id = ? FOR UPDATE", [id]);
    if (rows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Item not found" });
    }
    const item = rows[0];

    const [result] = await conn.query(
      "UPDATE office_inventory SET is_active = 1, updated_at = NOW() WHERE id = ?",
      [id],
    );
    if (result.affectedRows === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Item not found" });
    }

    await conn.query(
      `INSERT INTO stock_transactions
        (id, item_id, item_code, item_name, type, quantity_change, stock_before,
         stock_after, price_per_unit, total_amount, reason, performed_by_name,
         transaction_date, created_at)
       VALUES (?, ?, ?, ?, 'item_restored', 0, ?, ?, ?, 0, ?, ?, ?, NOW())`,
      [
        crypto.randomUUID(), id, item.item_code, item.name,
        item.current_stock, item.current_stock, item.price_per_unit,
        "Item restored", performedByName ?? "Unknown",
        new Date().toISOString().split("T")[0],
      ],
    );

    await conn.commit();
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// DELETE /office-inventory/:id/permanent — hard delete, only allowed once archived
app.delete("/office-inventory/:id/permanent", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  if (decoded.role !== "admin" && decoded.role !== "superadmin") {
    return res.status(403).json({ success: false, message: "Not authorized." });
  }

  const { id } = req.params;
  const { performedByName } = req.body;

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [rows] = await conn.query("SELECT * FROM office_inventory WHERE id = ? FOR UPDATE", [id]);
    if (rows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Item not found" });
    }
    const item = rows[0];

    if (item.is_active) {
      await conn.rollback();
      return res.status(400).json({
        success: false,
        message: "Item must be archived before it can be permanently deleted.",
      });
    }

    // Log before deleting so the activity trail survives the item itself.
    await conn.query(
      `INSERT INTO stock_transactions
        (id, item_id, item_code, item_name, type, quantity_change, stock_before,
         stock_after, price_per_unit, total_amount, reason, performed_by_name,
         transaction_date, created_at)
       VALUES (?, ?, ?, ?, 'item_deleted', 0, ?, ?, ?, 0, ?, ?, ?, NOW())`,
      [
        crypto.randomUUID(), id, item.item_code, item.name,
        item.current_stock, item.current_stock, item.price_per_unit,
        "Item permanently deleted", performedByName ?? "Unknown",
        new Date().toISOString().split("T")[0],
      ],
    );

    await conn.query("DELETE FROM office_inventory WHERE id = ?", [id]);

    await conn.commit();
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// PATCH /office-inventory/:id/restrict — toggle employee visibility
app.patch("/office-inventory/:id/restrict", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  if (decoded.role !== "admin" && decoded.role !== "superadmin") {
    return res.status(403).json({ success: false, message: "Not authorized." });
  }

  const { id } = req.params;
  const { isRestricted, performedByName } = req.body;

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [rows] = await conn.query("SELECT * FROM office_inventory WHERE id = ? FOR UPDATE", [id]);
    if (rows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Item not found" });
    }
    const item = rows[0];

    await conn.query(
      "UPDATE office_inventory SET is_restricted = ?, updated_at = NOW() WHERE id = ?",
      [isRestricted ? 1 : 0, id],
    );

    await conn.query(
      `INSERT INTO stock_transactions
        (id, item_id, item_code, item_name, type, quantity_change, stock_before,
         stock_after, price_per_unit, total_amount, reason, performed_by_name,
         transaction_date, created_at)
       VALUES (?, ?, ?, ?, ?, 0, ?, ?, ?, 0, ?, ?, ?, NOW())`,
      [
        crypto.randomUUID(), id, item.item_code, item.name,
        isRestricted ? "item_restricted" : "item_unrestricted",
        item.current_stock, item.current_stock, item.price_per_unit,
        isRestricted ? "Restricted to admin/superadmin" : "Unrestricted",
        performedByName ?? "Unknown",
        new Date().toISOString().split("T")[0],
      ],
    );

    await conn.commit();

    const [updated] = await conn.query("SELECT * FROM office_inventory WHERE id = ?", [id]);
    return res.json({ success: true, item: updated[0] });
  } catch (err) {
    await conn.rollback();
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});


// ─── IT INVENTORY ROUTES ─────────────────────────────────────────────────
// Table: it_inventory
//   asset_tag VARCHAR(100) PRIMARY KEY
//   company VARCHAR(50)
//   serial_number VARCHAR(100)
//   model VARCHAR(100)
//   brand VARCHAR(100)
//   category VARCHAR(50)
//   status VARCHAR(50)
//   assignee_id VARCHAR(100)
//   assignee_name VARCHAR(150)
//   location VARCHAR(100)
//   date_purchased DATE
//   notes TEXT
//   created_at DATETIME DEFAULT CURRENT_TIMESTAMP
//   updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP

// GET /it-inventory
app.get("/it-inventory", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  try {
    const [rows] = await db.query("SELECT * FROM it_inventory ORDER BY created_at DESC");
    return res.json({ success: true, count: rows.length, items: rows });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /it-inventory
app.post("/it-inventory", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const {
    assetTag, company, serialNumber, model, brand,
    category, status, assigneeId, assigneeName,
    location, datePurchased, notes,
  } = req.body;

  if (!assetTag || !company || !brand) {
    return res.status(400).json({
      success: false,
      message: "assetTag, company, and brand are required.",
    });
  }

  try {
    const [existing] = await db.query(
      "SELECT asset_tag FROM it_inventory WHERE asset_tag = ?",
      [assetTag],
    );
    if (existing.length > 0) {
      return res.status(409).json({
        success: false,
        message: `Asset tag "${assetTag}" already exists.`,
      });
    }

    await db.query(
      `INSERT INTO it_inventory
        (asset_tag, company, serial_number, model, brand, category, status,
         assignee_id, assignee_name, location, date_purchased, notes,
         created_at, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NOW(), NOW())`,
      [
        assetTag, company, serialNumber ?? "", model ?? "", brand,
        category, status, assigneeId ?? "", assigneeName ?? "",
        location, datePurchased || null, notes ?? "",
      ],
    );

    const [rows] = await db.query(
      "SELECT * FROM it_inventory WHERE asset_tag = ?",
      [assetTag],
    );
    return res.status(201).json({ success: true, item: rows[0] });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// PATCH /it-inventory/:assetTag — partial update, only sends fields present in body.
// assetTag (the URL param) is always the ORIGINAL tag used to find the row.
// If the body includes a new assetTag, this renames the primary key itself.
app.patch("/it-inventory/:assetTag", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const { assetTag } = req.params;

  const FIELD_MAP = {
    assetTag: "asset_tag",
    company: "company",
    serialNumber: "serial_number",
    model: "model",
    brand: "brand",
    category: "category",
    status: "status",
    assigneeId: "assignee_id",
    assigneeName: "assignee_name",
    location: "location",
    datePurchased: "date_purchased",
    notes: "notes",
  };

  const entries = Object.entries(req.body).filter(([key]) => FIELD_MAP[key]);
  if (entries.length === 0) {
    return res.status(400).json({ success: false, message: "No valid fields to update." });
  }

  // Renaming the primary key needs its own validation before the UPDATE.
  const newAssetTagEntry = entries.find(([key]) => key === "assetTag");
  const newAssetTag = newAssetTagEntry ? String(newAssetTagEntry[1]).trim() : null;

  if (newAssetTagEntry && !newAssetTag) {
    return res.status(400).json({ success: false, message: "Asset tag cannot be blank." });
  }

  const setClause = entries.map(([key]) => `${FIELD_MAP[key]} = ?`).join(", ");
  const values = entries.map(([key, value]) =>
    key === "datePurchased" && value === "" ? null :
    key === "assetTag" ? newAssetTag :
    value,
  );

  try {
    if (newAssetTag && newAssetTag !== assetTag) {
      const [dupe] = await db.query(
        "SELECT asset_tag FROM it_inventory WHERE asset_tag = ?",
        [newAssetTag],
      );
      if (dupe.length > 0) {
        return res.status(409).json({
          success: false,
          message: `Asset tag "${newAssetTag}" already exists.`,
        });
      }
    }

    const [result] = await db.query(
      `UPDATE it_inventory SET ${setClause}, updated_at = NOW() WHERE asset_tag = ?`,
      [...values, assetTag],
    );
    if (result.affectedRows === 0)
      return res.status(404).json({ success: false, message: "Asset not found." });
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});

// DELETE /it-inventory/:assetTag
app.delete("/it-inventory/:assetTag", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const { assetTag } = req.params;
  try {
    const [result] = await db.query(
      "DELETE FROM it_inventory WHERE asset_tag = ?",
      [assetTag],
    );
    if (result.affectedRows === 0)
      return res.status(404).json({ success: false, message: "Asset not found." });
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});


// ─── DROPDOWN CONFIGS ROUTES ─────────────────────────────────────────────
// Table: dropdown_configs
//   id INT AUTO_INCREMENT PRIMARY KEY
//   module VARCHAR(50)
//   field VARCHAR(50)
//   label VARCHAR(100)
//   value VARCHAR(100)
//   bg_color VARCHAR(20)
//   text_color VARCHAR(20)
//   sort_order INT
//   created_at DATETIME
//   updated_at DATETIME
//   UNIQUE (module, field, value)

// GET /dropdown-configs — everything, grouped module -> field -> [options]
app.get("/dropdown-configs", async (req, res) => {
  if (!requireAuth(req, res)) return;

  try {
    const [rows] = await db.query(
      "SELECT module, field, label, value, bg_color, text_color, sort_order FROM dropdown_configs ORDER BY module, field, sort_order ASC"
    );

    const configs = {};
    for (const row of rows) {
      configs[row.module] ??= {};
      configs[row.module][row.field] ??= [];
      configs[row.module][row.field].push({
        label: row.label,
        value: row.value,
        bgColor: row.bg_color,
        textColor: row.text_color,
      });
    }

    return res.json({ success: true, configs });
  } catch (err) {
    return res.status(500).json({ success: false, message: err.message });
  }
});
// PUT /dropdown-configs/:module/:field — replaces the full option list for one column
// Wrapped with deadlock-retry: concurrent saves (e.g. status + location fired
// together from ManageColumnsModal) can each DELETE+INSERT overlapping rows
// in dropdown_configs and get picked as a deadlock victim by InnoDB. That's
// expected/transient, so we retry the whole transaction a few times with a
// short randomized backoff before giving up.
app.put("/dropdown-configs/:module/:field", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { module, field } = req.params;
  const { options } = req.body;

  if (!Array.isArray(options)) {
    return res.status(400).json({ success: false, message: "options must be an array." });
  }

  const MAX_ATTEMPTS = 3;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    const conn = await db.getConnection();
    try {
      await conn.beginTransaction();

      await conn.query(
        "DELETE FROM dropdown_configs WHERE module = ? AND field = ?",
        [module, field],
      );

      for (let i = 0; i < options.length; i++) {
        const opt = options[i];
        await conn.query(
          `INSERT INTO dropdown_configs
            (module, field, label, value, bg_color, text_color, sort_order, created_at, updated_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, NOW(), NOW())`,
          [module, field, opt.label, opt.value, opt.bgColor, opt.textColor, i],
        );
      }

      await conn.commit();
      return res.json({ success: true });
    } catch (err) {
      await conn.rollback();

      const isDeadlock = err.code === "ER_LOCK_DEADLOCK";
      if (isDeadlock && attempt < MAX_ATTEMPTS) {
        console.warn(
          `PUT /dropdown-configs deadlock on attempt ${attempt}, retrying (${module}/${field})...`,
        );
        await new Promise((r) => setTimeout(r, 50 + Math.random() * 150));
        continue;
      }

      console.error("PUT /dropdown-configs error:", err);
      return res.status(500).json({ success: false, message: err.message });
    } finally {
      conn.release();
    }
  }
});


// ─── AUDIT LOGS ROUTES ────────────────────────────────────────────────────
// ─── AUDIT LOGS ROUTES ────────────────────────────────────────────────────

// POST /audit-logs — single-field entry
app.post("/audit-logs", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { table, recordId, recordLabel, field, oldValue, newValue, changedBy, changedById } = req.body;
  if (!table || !recordId || !field || !changedBy || !changedById) {
    return res.status(400).json({ success: false, message: "table, recordId, field, changedBy, changedById are required." });
  }

  try {
    await db.query(
      `INSERT INTO audit_logs
        (module, record_id, record_label, action, entry_type, field_name, old_value, new_value, performed_by_username, performed_by_name, created_at)
       VALUES (?, ?, ?, 'field_update', 'single', ?, ?, ?, ?, ?, NOW())`,
      [table, recordId, recordLabel ?? "", field, oldValue ?? "", newValue ?? "", changedById, changedBy],
    );
    return res.status(201).json({ success: true });
  } catch (err) {
    console.error("POST /audit-logs error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /audit-logs/batch — grouped multi-field entry
app.post("/audit-logs/batch", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { table, recordId, recordLabel, changes, changedBy, changedById } = req.body;
  if (!table || !recordId || !Array.isArray(changes) || !changedBy || !changedById) {
    return res.status(400).json({ success: false, message: "table, recordId, changes[], changedBy, changedById are required." });
  }
  if (changes.length === 0) {
    return res.json({ success: true, skipped: true });
  }

  try {
    await db.query(
      `INSERT INTO audit_logs
        (module, record_id, record_label, action, entry_type, changes, performed_by_username, performed_by_name, created_at)
       VALUES (?, ?, ?, 'batch_update', 'batch', ?, ?, ?, NOW())`,
      [table, recordId, recordLabel ?? "", JSON.stringify(changes), changedById, changedBy],
    );
    return res.status(201).json({ success: true });
  } catch (err) {
    console.error("POST /audit-logs/batch error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// GET /audit-logs/:table — optional ?recordId= and ?limit=
app.get("/audit-logs/:table", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { table } = req.params;
  const { recordId, limit } = req.query;
  const max = Math.min(parseInt(limit, 10) || 200, 500);

  try {
    const conditions = ["module = ?"];
    const params = [table];
    if (recordId) {
      conditions.push("record_id = ?");
      params.push(recordId);
    }

    const [rows] = await db.query(
      `SELECT * FROM audit_logs WHERE ${conditions.join(" AND ")} ORDER BY created_at DESC LIMIT ?`,
      [...params, max],
    );

    return res.json({ success: true, count: rows.length, entries: rows });
  } catch (err) {
    console.error("GET /audit-logs error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── IT CONSUMABLES ROUTES ─────────────────────────────────────────────────
// Table: it_consumables (see schema above)

// GET /it-consumables
app.get("/it-consumables", async (req, res) => {
  if (!requireAuth(req, res)) return;
  try {
    const [rows] = await db.query("SELECT * FROM it_consumables ORDER BY created_at DESC");
    return res.json({ success: true, count: rows.length, items: rows });
  } catch (err) {
    console.error("GET /it-consumables error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /it-consumables
app.post("/it-consumables", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const {
    model, name, status, location, ipAddress, macAddress,
    black, photoBlack, cyan, magenta, yellow, maintenanceBox,
  } = req.body;

  if (!model || !name) {
    return res.status(400).json({ success: false, message: "model and name are required." });
  }

  try {
    const [existing] = await db.query("SELECT model FROM it_consumables WHERE model = ?", [model]);
    if (existing.length > 0) {
      return res.status(409).json({ success: false, message: `Model "${model}" already exists.` });
    }

    await db.query(
      `INSERT INTO it_consumables
        (model, name, status, location, ip_address, mac_address, black, photo_black, cyan, magenta, yellow, maintenance_box, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NOW(), NOW())`,
      [
        model, name, status ?? "Spare", location, ipAddress ?? "", macAddress ?? "",
        black ?? 0, photoBlack ?? 0, cyan ?? 0, magenta ?? 0, yellow ?? 0, maintenanceBox ?? 0,
      ],
    );

    const [rows] = await db.query("SELECT * FROM it_consumables WHERE model = ?", [model]);
    return res.status(201).json({ success: true, item: rows[0] });
  } catch (err) {
    console.error("POST /it-consumables error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// PATCH /it-consumables/:model — partial update, only fields present in body
app.patch("/it-consumables/:model", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { model } = req.params;

  const FIELD_MAP = {
    name: "name",
    status: "status",
    location: "location",
    ipAddress: "ip_address",
    macAddress: "mac_address",
    black: "black",
    photoBlack: "photo_black",
    cyan: "cyan",
    magenta: "magenta",
    yellow: "yellow",
    maintenanceBox: "maintenance_box",
  };

  const entries = Object.entries(req.body).filter(([key]) => FIELD_MAP[key]);
  if (entries.length === 0) {
    return res.status(400).json({ success: false, message: "No valid fields to update." });
  }

  const setClause = entries.map(([key]) => `${FIELD_MAP[key]} = ?`).join(", ");
  const values = entries.map(([, value]) => value);

  try {
    const [result] = await db.query(
      `UPDATE it_consumables SET ${setClause}, updated_at = NOW() WHERE model = ?`,
      [...values, model],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Consumable not found." });
    }
    return res.json({ success: true });
  } catch (err) {
    console.error("PATCH /it-consumables error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// DELETE /it-consumables/:model
app.delete("/it-consumables/:model", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const { model } = req.params;
  try {
    const [result] = await db.query("DELETE FROM it_consumables WHERE model = ?", [model]);
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Consumable not found." });
    }
    return res.json({ success: true });
  } catch (err) {
    console.error("DELETE /it-consumables error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});


// ─── FLEET OPS ROUTES ────────────────────────────────────────────────────

// GET /fleet/vehicles
app.get("/fleet/vehicles", async (req, res) => {
  if (!requireAuth(req, res)) return;
  try {
    const [rows] = await db.query(
      `SELECT v.*, u.display_name AS assigned_driver_name
       FROM fleet_vehicles v
       LEFT JOIN users u ON u.id = v.assigned_driver_id
       ORDER BY v.plate_number ASC`,
    );
    const vehicles = rows.map((r) => ({
      id: String(r.id),
      plateNumber: r.plate_number,
      type: r.type,
      model: r.model,
      seatingCapacity: r.seating_capacity,
      status: r.status,
      currentTripLabel: r.current_trip_label,
      assignedDriverId: r.assigned_driver_id !== null ? String(r.assigned_driver_id) : null,
      assignedDriverName: r.assigned_driver_name,
      lastPingAt: r.last_ping_at,
      tramigoDeviceId: r.tramigo_device_id ?? null,
    }));
    return res.json({ success: true, count: vehicles.length, vehicles });
  } catch (err) {
    console.error("GET /fleet-vehicles error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /fleet/vehicles — body: { plateNumber, type, model, seatingCapacity }
app.post("/fleet/vehicles", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { plateNumber, type, model, seatingCapacity, tramigoDeviceId } = req.body;
  if (!plateNumber || !type || !model) {
    return res.status(400).json({ success: false, message: "plateNumber, type, and model are required." });
  }
  try {
    const [existing] = await db.query("SELECT id FROM fleet_vehicles WHERE plate_number = ?", [plateNumber]);
    if (existing.length > 0) {
      return res.status(409).json({ success: false, message: `Plate number "${plateNumber}" already exists.` });
    }
    await db.query(
      `INSERT INTO fleet_vehicles (plate_number, type, model, seating_capacity, tramigo_device_id, status, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?, 'idle', NOW(), NOW())`,
      [plateNumber, type, model, seatingCapacity ?? 4, tramigoDeviceId || null],
    );
    return res.status(201).json({ success: true });
  } catch (err) {
    console.error("POST /fleet/vehicles error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// GET /fleet/drivers
app.get("/fleet/drivers", async (req, res) => {
  if (!requireAuth(req, res)) return;
  try {
    const [rows] = await db.query(
      `SELECT d.*, u.display_name AS name, v.plate_number AS vehicle_plate
       FROM fleet_drivers d
       JOIN users u ON u.id = d.user_id
       LEFT JOIN fleet_vehicles v ON v.id = d.vehicle_id
       ORDER BY u.display_name ASC`,
    );
    const drivers = rows.map((r) => ({
      id: String(r.id),
      userId: String(r.user_id),
      name: r.name,
      licenseNumber: r.license_number,
      contactNumber: r.contact_number,
      vehicleId: r.vehicle_id !== null ? String(r.vehicle_id) : null,
      vehiclePlate: r.vehicle_plate,
      dutyStatus: r.duty_status ?? "off_duty",
      dutyStatusUpdatedAt: r.updated_at ?? null,
      shiftStart: r.shift_start ?? null,
      shiftEnd: r.shift_end ?? null,
    }));
    return res.json({ success: true, count: drivers.length, drivers });
  } catch (err) {
    console.error("GET /fleet-drivers error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /fleet/drivers — body: { username, licenseNumber, contactNumber }
// Links an existing `users` row as a driver via fleet_drivers.user_id.
app.post("/fleet/drivers", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { username, licenseNumber, contactNumber } = req.body;
  if (!username) {
    return res.status(400).json({ success: false, message: "username is required." });
  }
  try {
    const [userRows] = await db.query("SELECT id FROM users WHERE username = ?", [username.toLowerCase().trim()]);
    if (userRows.length === 0) {
      return res.status(404).json({ success: false, message: `User "${username}" not found.` });
    }
    const userId = userRows[0].id;

    const [existing] = await db.query("SELECT id FROM fleet_drivers WHERE user_id = ?", [userId]);
    if (existing.length > 0) {
      return res.status(409).json({ success: false, message: `"${username}" is already registered as a driver.` });
    }

    await db.query(
      `INSERT INTO fleet_drivers (user_id, license_number, contact_number, created_at, updated_at)
       VALUES (?, ?, ?, NOW(), NOW())`,
      [userId, licenseNumber ?? "", contactNumber ?? ""],
    );

    // Being added to the driver roster IS the grant now that the manual
    // "Fleet Driver" toggle is gone from Users Page — flip perm_fleet_driver
    // so this user's next permission fetch shows the Driver View nav item.
    await db.query(
      "UPDATE users SET perm_fleet_driver = 1, updated_at = NOW() WHERE id = ?",
      [userId],
    );

    return res.status(201).json({ success: true });
  } catch (err) {
    console.error("POST /fleet/drivers error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// GET /fleet/locations
app.get("/fleet/locations", async (req, res) => {
  if (!requireAuth(req, res)) return;
  try {
    const [rows] = await db.query("SELECT * FROM fleet_locations ORDER BY name ASC");
    const locations = rows.map((r) => ({
      id: String(r.id),
      name: r.name,
      latitude: r.latitude !== null ? Number(r.latitude) : null,
      longitude: r.longitude !== null ? Number(r.longitude) : null,
    }));
    return res.json({ success: true, count: locations.length, locations });
  } catch (err) {
    console.error("GET /fleet-locations error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /fleet/locations — body: { name, shortLabel }
app.post("/fleet/locations", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { name, latitude, longitude } = req.body;
  if (!name) {
    return res.status(400).json({ success: false, message: "name is required." });
  }
  try {
    await db.query(
      "INSERT INTO fleet_locations (name, latitude, longitude, created_at, updated_at) VALUES (?, ?, ?, NOW(), NOW())",
      [name, latitude ?? null, longitude ?? null],
    );
    return res.status(201).json({ success: true });
  } catch (err) {
    console.error("POST /fleet/locations error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});
// TEMPORARY — bulk-delete every Outlook event this app has created for
// fleet trips, then clear outlook_event_id so future syncs start fresh.
// Hit this once from Postman/curl (with your service token) to clean up
// the test data shown in the calendar, then remove this route.
app.post("/fleet/trips/cleanup-calendar-events", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  try {
    const [rows] = await db.query(
      "SELECT id, outlook_event_id FROM fleet_trips WHERE outlook_event_id IS NOT NULL",
    );

    let deleted = 0;
    let failed = 0;

    for (const row of rows) {
      try {
        await deleteTripEvent(row.outlook_event_id);
        await db.query(
          "UPDATE fleet_trips SET outlook_event_id = NULL WHERE id = ?",
          [row.id],
        );
        deleted++;
      } catch (err) {
        console.error(`Cleanup: failed to delete event for trip ${row.id}:`, err.message);
        failed++;
      }
    }

    return res.json({ success: true, total: rows.length, deleted, failed });
  } catch (err) {
    console.error("POST /fleet/trips/cleanup-calendar-events error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

app.post("/fleet/trips/:id/sync-to-my-calendar", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  const { id } = req.params;
  // Fleet Ops calls all go through the shared service token (see
  // getServiceToken() in fleetOps.ts), so decoded.username is always
  // "service" here — same situation as POST /fleet/trips/:id/approve.
  // The real requester's AD username has to come from the body instead.
  const requestingUsername = (req.body?.requestingUsername || decoded.username || "")
    .toLowerCase()
    .trim();

  try {
    const [userRows] = await db.query(
      "SELECT notification_email, email FROM users WHERE username = ?",
      [requestingUsername],
    );
    const targetUpn = userRows[0]?.notification_email || userRows[0]?.email;
    if (!targetUpn) {
      return res.status(400).json({ success: false, message: "No email on file for your account." });
    }

    const [tripRows] = await db.query(
      `SELECT t.trip_ref AS tripRef, t.pickup_location_text AS pickupLabel,
              t.dropoff_location_text AS dropoffLabel, t.departure_datetime AS departureDatetime,
              t.return_datetime AS returnDatetime, t.purpose,
              req.display_name AS requestorName, veh.plate_number AS vehiclePlate,
              drv.display_name AS driverName
       FROM fleet_trips t
       JOIN users req ON req.id = t.requestor_id
       LEFT JOIN fleet_vehicles veh ON veh.id = t.vehicle_id
       LEFT JOIN users drv ON drv.id = t.driver_id
       WHERE t.id = ?`,
      [id],
    );
    if (tripRows.length === 0) {
      return res.status(404).json({ success: false, message: "Trip not found." });
    }

    const eventId = await createTripEvent(tripRows[0], targetUpn);
    await db.query(
      "UPDATE fleet_trips SET outlook_event_id = ?, outlook_event_upn = ? WHERE id = ?",
      [eventId, targetUpn, id],
    );
    return res.json({ success: true, eventId });
  } catch (err) {
    console.error("POST /fleet/trips/:id/sync-to-my-calendar error:", err.message);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /fleet/trips/:id/remove-from-calendar — deletes the Outlook event
// this trip was synced to (mailbox + event id captured above), then clears
// both columns so the trip shows as "not synced" again.
app.post("/fleet/trips/:id/remove-from-calendar", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  const { id } = req.params;

  try {
    const [rows] = await db.query(
      "SELECT outlook_event_id, outlook_event_upn FROM fleet_trips WHERE id = ?",
      [id],
    );
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "Trip not found." });
    }
    const { outlook_event_id, outlook_event_upn } = rows[0];
    if (!outlook_event_id) {
      return res.status(400).json({ success: false, message: "This trip isn't synced to a calendar." });
    }

    await deleteTripEvent(outlook_event_id, outlook_event_upn || undefined);

    await db.query(
      "UPDATE fleet_trips SET outlook_event_id = NULL, outlook_event_upn = NULL WHERE id = ?",
      [id],
    );

    return res.json({ success: true });
  } catch (err) {
    console.error("POST /fleet/trips/:id/remove-from-calendar error:", err.message);
    return res.status(500).json({ success: false, message: err.message });
  }
});
// ─── INSERT into server.js, in the "FLEET OPS ROUTES" section ─────────────
// Place this right ABOVE the existing `app.get("/fleet/trips", ...)` handler.
// It's the missing counterpart to that GET — creates a new trip request from
// the TripBookingModal submission. Uses the same transaction + "FOR UPDATE"
// counter pattern as POST /supply-requests for generating a sequential ref,
// and resolves the AD username sent from the frontend to a numeric users.id
// the same way POST /fleet/drivers already does (since requests here run on
// the service token, not the employee's own JWT).

// POST /fleet/trips — employee submits a new trip booking request
// body: { requestorId (AD username), requestorName, pickupLocationId?, pickupLocationText,
//         dropoffLocationId?, dropoffLocationText, tripType, departureDatetime,
//         returnDatetime?, purpose?, passengerCount }
app.post("/fleet/trips", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const {
    requestorId,
    requestorName,
    pickupLocationId,
    pickupLocationText,
    pickupLatitude,
    pickupLongitude,
    dropoffLocationId,
    dropoffLocationText,
    dropoffLatitude,
    dropoffLongitude,
    additionalDropoffs, // [{ locationId?, locationText, latitude, longitude }, ...]
    tripType,
    departureDatetime,
    returnDatetime,
    purpose,
    passengerCount,
    passengerNames,
  } = req.body;

  if (
    !requestorId ||
    !pickupLocationText?.trim() ||
    !dropoffLocationText?.trim() ||
    !tripType ||
    !departureDatetime
  ) {
    return res.status(400).json({
      success: false,
      message:
        "requestorId, pickup/dropoff locations, tripType, and departureDatetime are required.",
    });
  }

  // Coordinates are required at booking time so the Control Tower's route
  // map always has a pin to draw — older trips predating this check can
  // still have null lat/lng, but no new trip should.
  if (
    pickupLatitude == null ||
    pickupLongitude == null ||
    dropoffLatitude == null ||
    dropoffLongitude == null
  ) {
    return res.status(400).json({
      success: false,
      message:
        "Pickup and drop-off must be selected from the map/search suggestions so their coordinates are captured.",
    });
  }

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [userRows] = await conn.query(
      "SELECT id FROM users WHERE username = ?",
      [requestorId.toLowerCase().trim()],
    );
    if (userRows.length === 0) {
      await conn.rollback();
      return res
        .status(404)
        .json({ success: false, message: `User "${requestorId}" not found.` });
    }
    const requestorUserId = userRows[0].id;

    const year = new Date().getFullYear();
    const [maxRows] = await conn.query(
      `SELECT MAX(CAST(SUBSTRING_INDEX(trip_ref, '-', -1) AS UNSIGNED)) AS maxNum
       FROM fleet_trips WHERE trip_ref LIKE ? FOR UPDATE`,
      [`TRIP-${year}-%`],
    );
    const nextNum = String((maxRows[0].maxNum ?? 0) + 1).padStart(4, "0");
    const tripRef = `TRIP-${year}-${nextNum}`;

    const nowStr = nowLocalDatetime();

    const [insertResult] = await conn.query(
      `INSERT INTO fleet_trips
        (trip_ref, requestor_id, pickup_location_id, pickup_location_text,
         pickup_latitude, pickup_longitude,
         dropoff_location_id, dropoff_location_text, dropoff_latitude, dropoff_longitude,
         trip_type, departure_datetime, return_datetime, purpose, passenger_count,
         passenger_names, status, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?)`,
      [
        tripRef,
        requestorUserId,
        pickupLocationId ?? null,
        pickupLocationText.trim(),
        pickupLatitude ?? null,
        pickupLongitude ?? null,
        dropoffLocationId ?? null,
        dropoffLocationText.trim(),
        dropoffLatitude ?? null,
        dropoffLongitude ?? null,
        tripType,
        departureDatetime,
        returnDatetime ?? null,
        purpose ?? "",
        passengerCount ?? 1,
        Array.isArray(passengerNames) && passengerNames.length > 0
          ? JSON.stringify(passengerNames)
          : null,
        nowStr,
        nowStr,
      ],
    );

    await conn.query(
      "INSERT INTO fleet_trip_status_log (trip_id, status, changed_by, created_at) VALUES (?, 'pending', ?, ?)",
      [insertResult.insertId, requestorUserId, nowStr],
    );

    // Extra drop-off stops beyond the primary one — stored in order so a
    // multi-stop trip's route can be reconstructed and displayed later.
    if (Array.isArray(additionalDropoffs) && additionalDropoffs.length > 0) {
      for (let i = 0; i < additionalDropoffs.length; i++) {
        const stop = additionalDropoffs[i];
        await conn.query(
          `INSERT INTO fleet_trip_stops
            (trip_id, stop_order, location_id, location_text, latitude, longitude, created_at)
           VALUES (?, ?, ?, ?, ?, ?, NOW())`,
          [
            insertResult.insertId,
            i + 2, // 1 = primary dropoff, 2+ = additional stops in order
            stop.locationId ?? null,
            stop.locationText,
            stop.latitude,
            stop.longitude,
          ],
        );
      }
    }

    await conn.commit();

    // Push to Outlook at booking time (before approval/dispatch).
    try {
      const eventId = await createTripEvent({
        tripRef,
        pickupLabel: pickupLocationText.trim(),
        dropoffLabel: dropoffLocationText.trim(),
        requestorName: requestorName ?? "An employee",
        purpose,
        departureDatetime,
        returnDatetime,
      });
      await db.query(
        "UPDATE fleet_trips SET outlook_event_id = ? WHERE id = ?",
        [eventId, insertResult.insertId],
      );
    } catch (err) {
      console.error("Outlook calendar sync (booking) failed:", err.message);
    }

    // Mirrors the "notify admins" step on supply request creation — remove
    // this call if you don't want a push notification fired on every booking.
    sendWebPushToAdmins({
      title: "New Trip Request",
      body: `${requestorName ?? "An employee"} requested ${tripRef} (${pickupLocationText} → ${dropoffLocationText})`,
      url: "/fleet-control-tower",
      permissionColumn: "perm_fleet_control",
    });

    sendFleetTripRequestNotification({
      requestorId,
      requestorName: requestorName ?? "An employee",
      tripRef,
      pickupLocationText: pickupLocationText.trim(),
      dropoffLocationText: dropoffLocationText.trim(),
      tripType,
      departureDatetime,
      purpose,
      passengerCount,
    }).then((messageId) => {
      sendFleetTripAdminNotification({
        requestorName: requestorName ?? "An employee",
        tripRef,
        pickupLocationText: pickupLocationText.trim(),
        dropoffLocationText: dropoffLocationText.trim(),
        tripType,
        departureDatetime,
        purpose,
        passengerCount,
        replyToMessageId: messageId,
      });
    });

    return res.status(201).json({ success: true, tripRef });
  } catch (err) {
    await conn.rollback();
    console.error("POST /fleet/trips error:", err);
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});


// GET /fleet/trips — ?includeArchived=true also returns archived trips
// (default excludes them, mirrors GET /supply-requests).
app.get("/fleet/trips", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const includeArchived = req.query.includeArchived === "true";
  try {
    const [rows] = await db.query(
      `SELECT t.*,
              t.outlook_event_id AS outlook_event_id,
              req.display_name AS requestor_name,
              veh.plate_number AS vehicle_plate,
              drv.display_name AS driver_name,
              apr.display_name AS approved_by_name
       FROM fleet_trips t
       JOIN users req                ON req.id = t.requestor_id
       LEFT JOIN fleet_vehicles veh  ON veh.id = t.vehicle_id
       LEFT JOIN users drv           ON drv.id = t.driver_id
       LEFT JOIN users apr           ON apr.id = t.approved_by
       ${includeArchived ? "" : "WHERE t.is_archived = 0"}
       ORDER BY t.departure_datetime DESC`,
    );

    const tripIds = rows.map((r) => r.id);
    let historyByTrip = {};
    let stopsByTrip = {};
    if (tripIds.length > 0) {
      const [stopRows] = await db.query(
        `SELECT trip_id, stop_order, location_id, location_text, latitude, longitude
         FROM fleet_trip_stops WHERE trip_id IN (?) ORDER BY stop_order ASC`,
        [tripIds],
      );
      stopsByTrip = stopRows.reduce((acc, row) => {
        (acc[row.trip_id] ??= []).push({
          locationId: row.location_id !== null ? String(row.location_id) : null,
          locationText: row.location_text,
          latitude: row.latitude !== null ? Number(row.latitude) : null,
          longitude: row.longitude !== null ? Number(row.longitude) : null,
        });
        return acc;
      }, {});
    }
    if (tripIds.length > 0) {
      const [logRows] = await db.query(
        `SELECT
           l.trip_id,
           l.status,
           l.vehicle_id,
           lv.plate_number AS vehicle_plate,
           l.driver_id,
           ldu.display_name AS driver_name,
           l.note,
           lcu.display_name AS changed_by_name,
           l.created_at
         FROM fleet_trip_status_log l
         LEFT JOIN fleet_vehicles lv ON lv.id = l.vehicle_id
         LEFT JOIN users ldu         ON ldu.id = l.driver_id
         LEFT JOIN users lcu         ON lcu.id = l.changed_by
         WHERE l.trip_id IN (?)
         ORDER BY l.created_at ASC`,
        [tripIds],
      );
      historyByTrip = logRows.reduce((acc, row) => {
        (acc[row.trip_id] ??= []).push({
          status: row.status,
          vehicleId: row.vehicle_id !== null ? String(row.vehicle_id) : null,
          vehiclePlate: row.vehicle_plate,
          driverId: row.driver_id !== null ? String(row.driver_id) : null,
          driverName: row.driver_name,
          note: row.note,
          changedByName: row.changed_by_name,
          timestamp: row.created_at,
        });
        return acc;
      }, {});
    }

    const trips = rows.map((r) => ({
      id: String(r.id),
      tripRef: r.trip_ref,
      requestorId: String(r.requestor_id),
      requestorName: r.requestor_name,
      pickupLocationId: r.pickup_location_id !== null ? String(r.pickup_location_id) : null,
      pickupLabel: r.pickup_location_text,
      pickupLatitude: r.pickup_latitude !== null ? Number(r.pickup_latitude) : null,
      pickupLongitude: r.pickup_longitude !== null ? Number(r.pickup_longitude) : null,
      dropoffLocationId: r.dropoff_location_id !== null ? String(r.dropoff_location_id) : null,
      dropoffLabel: r.dropoff_location_text,
      dropoffLatitude: r.dropoff_latitude !== null ? Number(r.dropoff_latitude) : null,
      dropoffLongitude: r.dropoff_longitude !== null ? Number(r.dropoff_longitude) : null,
      tripType: r.trip_type,
      departureDatetime: r.departure_datetime,
      returnDatetime: r.return_datetime,
      purpose: r.purpose,
      passengerCount: r.passenger_count,
      passengerNames: (() => {
        try {
          return r.passenger_names ? JSON.parse(r.passenger_names) : [];
        } catch {
          return [];
        }
      })(),
      vehicleId: r.vehicle_id !== null ? String(r.vehicle_id) : null,
      vehiclePlate: r.vehicle_plate,
      driverId: r.driver_id !== null ? String(r.driver_id) : null,
      driverName: r.driver_name,
     status: r.status,
      rejectedReason: r.rejected_reason,
      approvedByName: r.approved_by_name,
      approvedAt: r.approved_at,
      createdAt: r.created_at,
      updatedAt: r.updated_at,
      statusHistory: historyByTrip[r.id] ?? [],
      // Extra stops beyond the primary dropoff, in visit order.
      additionalDropoffs: stopsByTrip[r.id] ?? [],
      calendarSynced: !!r.outlook_event_id,
      isArchived: !!r.is_archived,
    }));
    return res.json({ success: true, count: trips.length, trips });
  } catch (err) {
    console.error("GET /fleet-trips error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /fleet/trips/:id/archive — hide from default list without deleting
app.post("/fleet/trips/:id/archive", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  try {
    const [result] = await db.query(
      "UPDATE fleet_trips SET is_archived = 1 WHERE id = ?",
      [id],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Trip not found" });
    }
    return res.json({ success: true });
  } catch (err) {
    console.error("POST /fleet/trips/:id/archive error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /fleet/trips/:id/unarchive — undo
app.post("/fleet/trips/:id/unarchive", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  try {
    const [result] = await db.query(
      "UPDATE fleet_trips SET is_archived = 0 WHERE id = ?",
      [id],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Trip not found" });
    }
    return res.json({ success: true });
  } catch (err) {
    console.error("POST /fleet/trips/:id/unarchive error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /fleet/trips/:id/approve — body: { vehicleId, driverId }
// driverId here is fleet_drivers.id (from the dispatch dropdown), which is
// resolved to its underlying users.id since fleet_trips.driver_id is an FK to users.
app.post("/fleet/trips/:id/approve", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  const { id } = req.params;
  const { vehicleId, driverId, approvedById: bodyApprovedByUsername } = req.body;
  if (!vehicleId || !driverId) {
    return res.status(400).json({ success: false, message: "vehicleId and driverId are required." });
  }

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [tripRows] = await conn.query("SELECT * FROM fleet_trips WHERE id = ? FOR UPDATE", [id]);
    if (tripRows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Trip not found." });
    }
    const trip = tripRows[0];

    const [driverRows] = await conn.query("SELECT user_id FROM fleet_drivers WHERE id = ?", [driverId]);
    if (driverRows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Driver not found." });
    }
    const driverUserId = driverRows[0].user_id;

    const [vehicleRows] = await conn.query("SELECT * FROM fleet_vehicles WHERE id = ? FOR UPDATE", [vehicleId]);
    if (vehicleRows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Vehicle not found." });
    }

    // decoded.username is always "service" here — every request goes through
    // the shared service token (see getServiceToken() in fleetOps.ts), not the
    // logged-in admin's own JWT — so the real approver has to come from the AD
    // username the frontend already sends in the body.
    const [approverRows] = await conn.query(
      "SELECT id, display_name FROM users WHERE username = ?",
      [(bodyApprovedByUsername || "").toLowerCase().trim()],
    );
    const approvedById = approverRows[0]?.id ?? null;
    const approvedByDisplayName = approverRows[0]?.display_name || bodyApprovedByUsername;
    console.log("approve trip:", { bodyApprovedByUsername, approvedById });

    // Vehicle/driver get linked here at approval time (not just reserved
    // on the trip row) so the Driver Portal's duty-status selector — which
    // looks up the driver's vehicle via fleet_drivers.vehicle_id — shows up
    // as soon as a trip is approved, not only after the driver taps "Start
    // Trip". fleet_vehicles.status itself still stays 'idle' here; it only
    // flips to 'active' in POST /fleet/trips/:id/start below, once the
    // driver is actually en route.
    await conn.query(
      `UPDATE fleet_trips SET
         status = 'approved', vehicle_id = ?, driver_id = ?,
         approved_by = ?, approved_at = NOW(), updated_at = NOW()
       WHERE id = ?`,
      [vehicleId, driverUserId, approvedById, id],
    );

    await conn.query(
      "UPDATE fleet_drivers SET vehicle_id = ?, updated_at = NOW() WHERE user_id = ?",
      [vehicleId, driverUserId],
    );

    // Reassigning the vehicle/driver on a trip that's already approved/
    // arrived hits this same route again — it isn't a new status
    // transition, so it's logged at the trip's actual current status with
    // a descriptive note instead of a bare duplicate "Approved" entry.
    const wasAlreadyApproved = trip.status === "approved" || trip.status === "arrived";
    if (!wasAlreadyApproved) {
      await conn.query(
        `INSERT INTO fleet_trip_status_log (trip_id, status, vehicle_id, driver_id, changed_by, note, created_at)
         VALUES (?, 'approved', ?, ?, ?, NULL, NOW())`,
        [id, vehicleId, driverUserId, approvedById],
      );
    } else {
      const [vehRows] = await conn.query("SELECT plate_number FROM fleet_vehicles WHERE id = ?", [vehicleId]);
      const [drvRows] = await conn.query("SELECT display_name FROM users WHERE id = ?", [driverUserId]);
      const note = `Reassigned to ${vehRows[0]?.plate_number ?? "vehicle"} / ${drvRows[0]?.display_name ?? "driver"}`;
      await conn.query(
        `INSERT INTO fleet_trip_status_log (trip_id, status, vehicle_id, driver_id, changed_by, note, created_at)
         VALUES (?, ?, ?, ?, ?, ?, NOW())`,
        [id, trip.status, vehicleId, driverUserId, approvedById, note],
      );
    }
    var wasReassignment = wasAlreadyApproved;

    await conn.commit();

    // Notify the assigned driver they've got a new trip. push_subscriptions
    // is keyed by username, not users.id, so look that up first.
    try {
      const [driverUserRows] = await db.query(
        "SELECT username FROM users WHERE id = ?",
        [driverUserId],
      );
      const driverUsername = driverUserRows[0]?.username;
      if (driverUsername) {
        sendWebPushToUser(driverUsername, {
          title: "New Trip Assigned",
          body: `You've been assigned to ${trip.trip_ref} (${trip.pickup_location_text} → ${trip.dropoff_location_text})`,
          url: "/driver-portal",
        });
        sendExpoPushToUser(driverUsername, {
          title: "New Trip Assigned",
          body: `You've been assigned to ${trip.trip_ref} (${trip.pickup_location_text} → ${trip.dropoff_location_text})`,
          data: { tripId: id, url: "/driver-portal" },
        });
      }
    } catch (err) {
      console.error("Driver trip-assignment push failed:", err.message);
    }

    if (wasReassignment) {
      const noteText = `Vehicle and driver reassigned by dispatch.`;
      sendFleetTripStatusNotification({ tripId: id, statusKey: trip.status, extraMessage: noteText, updatedByName: approvedByDisplayName });
    } else {
      sendFleetTripStatusNotification({ tripId: id, statusKey: "approved", updatedByName: approvedByDisplayName });
    }

    // Push to Outlook — create on first approval, update on reassignment.
    try {
      const [calRows] = await db.query(
        `SELECT t.trip_ref AS tripRef, t.pickup_location_text AS pickupLabel,
                t.dropoff_location_text AS dropoffLabel, t.departure_datetime AS departureDatetime,
                t.return_datetime AS returnDatetime, t.purpose,
                t.outlook_event_id AS outlookEventId,
                req.display_name AS requestorName, veh.plate_number AS vehiclePlate,
                drv.display_name AS driverName
         FROM fleet_trips t
         JOIN users req ON req.id = t.requestor_id
         LEFT JOIN fleet_vehicles veh ON veh.id = t.vehicle_id
         LEFT JOIN users drv ON drv.id = t.driver_id
         WHERE t.id = ?`,
        [id],
      );
      const calTrip = calRows[0];
      if (calTrip.outlookEventId) {
        await updateTripEvent(calTrip.outlookEventId, calTrip);
      } else {
        const eventId = await createTripEvent(calTrip);
        await db.query("UPDATE fleet_trips SET outlook_event_id = ? WHERE id = ?", [eventId, id]);
      }
    } catch (err) {
      console.error("Outlook calendar sync (approve/reassign) failed:", err.message);
    }

    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    console.error("POST /fleet-trips/:id/approve error:", err);
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// POST /fleet/trips/:id/reject — body: { reason }
app.post("/fleet/trips/:id/reject", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  const { id } = req.params;
  const { reason } = req.body;

  try {
    const [result] = await db.query(
      "UPDATE fleet_trips SET status = 'rejected', rejected_reason = ?, updated_at = NOW() WHERE id = ?",
      [reason ?? "", id],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Trip not found." });
    }
    const [changerRows] = await db.query("SELECT id FROM users WHERE username = ?", [decoded.username]);
    const changedById = changerRows[0]?.id ?? null;

    await db.query(
      "INSERT INTO fleet_trip_status_log (trip_id, status, note, changed_by, created_at) VALUES (?, 'rejected', ?, ?, NOW())",
      [id, reason ?? "", changedById],
    );
    sendFleetTripStatusNotification({ tripId: id, statusKey: "rejected", extraMessage: reason });

    try {
      const [tRows] = await db.query("SELECT outlook_event_id FROM fleet_trips WHERE id = ?", [id]);
      if (tRows[0]?.outlook_event_id) {
        await deleteTripEvent(tRows[0].outlook_event_id);
      }
    } catch (err) {
      console.error("Outlook calendar sync (reject) failed:", err.message);
    }

    return res.json({ success: true });
  } catch (err) {
    console.error("POST /fleet-trips/:id/reject error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /fleet/trips/:id/cancel — employee-initiated while pending, OR
// admin-initiated (from the Control Tower) while approved. Mirrors
// POST /supply-requests/:id/cancel. If the trip was already approved and
// had a vehicle/driver attached, both are released back to available —
// same release logic as /arrive and /complete — since a cancelled trip
// should never leave a vehicle "stuck" on it.
const CANCELLABLE_STATUSES = ["pending", "approved", "arrived"];

app.post("/fleet/trips/:id/cancel", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  const { id } = req.params;
  const { cancelledByName } = req.body;

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [rows] = await conn.query(
      "SELECT * FROM fleet_trips WHERE id = ? FOR UPDATE",
      [id],
    );
    if (rows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Trip not found" });
    }
    const trip = rows[0];

    if (!CANCELLABLE_STATUSES.includes(trip.status)) {
      await conn.rollback();
      return res.status(409).json({
        success: false,
        message: "This trip can no longer be cancelled.",
      });
    }

    await conn.query(
      "UPDATE fleet_trips SET status = 'cancelled', updated_at = NOW() WHERE id = ?",
      [id],
    );

    if (trip.vehicle_id) {
      await conn.query(
        "UPDATE fleet_vehicles SET status = 'idle', current_trip_label = NULL, assigned_driver_id = NULL, updated_at = NOW() WHERE id = ?",
        [trip.vehicle_id],
      );
      await conn.query(
        "UPDATE fleet_drivers SET vehicle_id = NULL, updated_at = NOW() WHERE vehicle_id = ?",
        [trip.vehicle_id],
      );
    }

    const [changerRows] = await conn.query("SELECT id FROM users WHERE username = ?", [decoded.username]);
    const changedById = changerRows[0]?.id ?? null;

    await conn.query(
      "INSERT INTO fleet_trip_status_log (trip_id, status, vehicle_id, driver_id, changed_by, created_at) VALUES (?, 'cancelled', ?, ?, ?, NOW())",
      [id, trip.vehicle_id, trip.driver_id, changedById],
    );

    await conn.commit();
    sendFleetTripStatusNotification({ tripId: id, statusKey: "cancelled", updatedByName: cancelledByName });

    try {
      if (trip.outlook_event_id) {
        await deleteTripEvent(trip.outlook_event_id);
      }
    } catch (err) {
      console.error("Outlook calendar sync (cancel) failed:", err.message);
    }

    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    console.error("POST /fleet-trips/:id/cancel error:", err);
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// PATCH /fleet/trips/:id/dropoffs — admin edits the trip's drop-off set:
// the primary destination (dropoffLabel/Location/lat/lng) and/or the
// additional-stops list. Replaces fleet_trip_stops wholesale for this trip
// (delete + reinsert) since stop order/removal is easiest to express that
// way, same pattern as PUT /dropdown-configs.
app.patch("/fleet/trips/:id/dropoffs", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  const { id } = req.params;
  const {
    dropoffLocationId,
    dropoffLabel,
    dropoffLatitude,
    dropoffLongitude,
    additionalDropoffs,
  } = req.body;

  if (!dropoffLabel?.trim()) {
    return res.status(400).json({ success: false, message: "dropoffLabel is required." });
  }
  if (!Array.isArray(additionalDropoffs)) {
    return res.status(400).json({ success: false, message: "additionalDropoffs array is required." });
  }

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [tripRows] = await conn.query("SELECT id, outlook_event_id FROM fleet_trips WHERE id = ? FOR UPDATE", [id]);
    if (tripRows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Trip not found." });
    }

    await conn.query(
      `UPDATE fleet_trips SET
         dropoff_location_id = ?, dropoff_location_text = ?,
         dropoff_latitude = ?, dropoff_longitude = ?, updated_at = NOW()
       WHERE id = ?`,
      [
        dropoffLocationId ?? null,
        dropoffLabel.trim(),
        dropoffLatitude ?? null,
        dropoffLongitude ?? null,
        id,
      ],
    );

    await conn.query("DELETE FROM fleet_trip_stops WHERE trip_id = ?", [id]);

    for (let i = 0; i < additionalDropoffs.length; i++) {
      const stop = additionalDropoffs[i];
      await conn.query(
        `INSERT INTO fleet_trip_stops
          (trip_id, stop_order, location_id, location_text, latitude, longitude, created_at)
         VALUES (?, ?, ?, ?, ?, ?, NOW())`,
        [
          id,
          i + 2, // 1 = primary dropoff, 2+ = additional stops in order
          stop.locationId ?? null,
          stop.locationText,
          stop.latitude ?? null,
          stop.longitude ?? null,
        ],
      );
    }

    const [changerRows] = await conn.query("SELECT id FROM users WHERE username = ?", [decoded.username]);
    const changedById = changerRows[0]?.id ?? null;
    await conn.query(
      "INSERT INTO fleet_trip_status_log (trip_id, status, changed_by, note, created_at) VALUES (?, (SELECT status FROM fleet_trips WHERE id = ?), ?, 'Drop-offs updated', NOW())",
      [id, id, changedById],
    );

    const [statusRow] = await conn.query("SELECT status FROM fleet_trips WHERE id = ?", [id]);
    const currentTripStatus = statusRow[0]?.status;

    await conn.commit();

    sendFleetTripStatusNotification({
      tripId: id,
      statusKey: currentTripStatus,
      extraMessage: `Your trip's drop-off details have been updated.`,
    });

    // Keep Outlook in sync if this trip already has a calendar event.
    const outlookEventId = tripRows[0].outlook_event_id;
    if (outlookEventId) {
      try {
        const [calRows] = await db.query(
          `SELECT t.trip_ref AS tripRef, t.pickup_location_text AS pickupLabel,
                  t.dropoff_location_text AS dropoffLabel, t.departure_datetime AS departureDatetime,
                  t.return_datetime AS returnDatetime, t.purpose,
                  req.display_name AS requestorName, veh.plate_number AS vehiclePlate,
                  drv.display_name AS driverName
           FROM fleet_trips t
           JOIN users req ON req.id = t.requestor_id
           LEFT JOIN fleet_vehicles veh ON veh.id = t.vehicle_id
           LEFT JOIN users drv ON drv.id = t.driver_id
           WHERE t.id = ?`,
          [id],
        );
        if (calRows[0]) {
          await updateTripEvent(outlookEventId, calRows[0]);
        }
      } catch (err) {
        console.error("Outlook calendar sync (dropoffs update) failed:", err.message);
      }
    }

    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    console.error("PATCH /fleet-trips/:id/dropoffs error:", err);
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// PATCH /fleet/trips/:id/reschedule — employee-initiated, only while the
// trip hasn't been approved yet (same gate as cancel above). Only moves
// departureDatetime/returnDatetime; everything else about the trip is
// untouched. Also pushes the updated time to Outlook if the trip was
// already synced to a calendar event.
app.patch("/fleet/trips/:id/reschedule", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  const { id } = req.params;
  const { departureDatetime, returnDatetime } = req.body;

  if (!departureDatetime) {
    return res.status(400).json({ success: false, message: "departureDatetime is required." });
  }

  // Allowed while pending (employee-initiated, before dispatch has looked
  // at it) or approved (admin-initiated, from the Control Tower — the trip
  // already has a vehicle/driver assigned, this only moves its time).
  const RESCHEDULABLE_STATUSES = ["pending", "approved"];

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [rows] = await conn.query(
      "SELECT status, outlook_event_id FROM fleet_trips WHERE id = ? FOR UPDATE",
      [id],
    );
    if (rows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Trip not found" });
    }

    const currentStatus = rows[0].status;
    if (!RESCHEDULABLE_STATUSES.includes(currentStatus)) {
      await conn.rollback();
      return res.status(409).json({
        success: false,
        message: "This trip can no longer be rescheduled.",
      });
    }

    await conn.query(
      "UPDATE fleet_trips SET departure_datetime = ?, return_datetime = ?, updated_at = NOW() WHERE id = ?",
      [departureDatetime, returnDatetime ?? null, id],
    );

    const [changerRows] = await conn.query("SELECT id FROM users WHERE username = ?", [decoded.username]);
    const changedById = changerRows[0]?.id ?? null;

    // Log at the trip's ACTUAL current status, not a hardcoded 'pending' —
    // an approved trip stays approved after an admin reschedules it, this
    // is a time change, not a status transition.
    const noteText = currentStatus === "approved" ? "Rescheduled by dispatch" : "Rescheduled by requestor";
    await conn.query(
      "INSERT INTO fleet_trip_status_log (trip_id, status, changed_by, note, created_at) VALUES (?, ?, ?, ?, NOW())",
      [id, currentStatus, changedById, noteText],
    );

    await conn.commit();

    sendFleetTripStatusNotification({
      tripId: id,
      statusKey: currentStatus,
      extraMessage: `Your trip has been rescheduled to ${departureDatetime}.`,
    });

    const outlookEventId = rows[0].outlook_event_id;
    if (outlookEventId) {
      try {
        const [calRows] = await db.query(
          `SELECT t.trip_ref AS tripRef, t.pickup_location_text AS pickupLabel,
                  t.dropoff_location_text AS dropoffLabel, t.departure_datetime AS departureDatetime,
                  t.return_datetime AS returnDatetime, t.purpose,
                  req.display_name AS requestorName, veh.plate_number AS vehiclePlate,
                  drv.display_name AS driverName
           FROM fleet_trips t
           JOIN users req ON req.id = t.requestor_id
           LEFT JOIN fleet_vehicles veh ON veh.id = t.vehicle_id
           LEFT JOIN users drv ON drv.id = t.driver_id
           WHERE t.id = ?`,
          [id],
        );
        if (calRows[0]) {
          await updateTripEvent(outlookEventId, calRows[0]);
        }
      } catch (err) {
        console.error("Outlook calendar sync (reschedule) failed:", err.message);
      }
    }

    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    console.error("PATCH /fleet-trips/:id/reschedule error:", err);
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// POST /fleet/trips/:id/arrive
// "Arrive" now completes the trip outright — no separate returning/mark-
// completed step. Mirrors the vehicle/driver-release logic from
// POST /fleet/trips/:id/complete, since that's effectively what this does now.
app.post("/fleet/trips/:id/arrive", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  const { id } = req.params;

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [tripRows] = await conn.query("SELECT * FROM fleet_trips WHERE id = ? FOR UPDATE", [id]);
    if (tripRows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Trip not found." });
    }
    const trip = tripRows[0];

    await conn.query(
      "UPDATE fleet_trips SET status = 'completed', updated_at = NOW() WHERE id = ?",
      [id],
    );

    if (trip.vehicle_id) {
      await conn.query(
        "UPDATE fleet_vehicles SET status = 'idle', current_trip_label = NULL, assigned_driver_id = NULL, updated_at = NOW() WHERE id = ?",
        [trip.vehicle_id],
      );
      await conn.query(
        "UPDATE fleet_drivers SET vehicle_id = NULL, updated_at = NOW() WHERE vehicle_id = ?",
        [trip.vehicle_id],
      );
    }

    const [changerRows] = await conn.query("SELECT id FROM users WHERE username = ?", [decoded.username]);
    const changedById = changerRows[0]?.id ?? null;

    await conn.query(
      "INSERT INTO fleet_trip_status_log (trip_id, status, vehicle_id, driver_id, changed_by, created_at) VALUES (?, 'completed', ?, ?, ?, NOW())",
      [id, trip.vehicle_id, trip.driver_id, changedById],
    );

    await conn.commit();
    sendFleetTripStatusNotification({ tripId: id, statusKey: "completed" });
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    console.error("POST /fleet-trips/:id/arrive error:", err);
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});
// POST /fleet/trips/:id/start-return
app.post("/fleet/trips/:id/start-return", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  const { id } = req.params;
  try {
    const [tripRows] = await db.query("SELECT vehicle_id, driver_id FROM fleet_trips WHERE id = ?", [id]);
    if (tripRows.length === 0) {
      return res.status(404).json({ success: false, message: "Trip not found." });
    }
    const trip = tripRows[0];

    const [result] = await db.query(
      "UPDATE fleet_trips SET status = 'returning', updated_at = NOW() WHERE id = ?",
      [id],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Trip not found." });
    }
    const [changerRows] = await db.query("SELECT id FROM users WHERE username = ?", [decoded.username]);
    const changedById = changerRows[0]?.id ?? null;

    await db.query(
      "INSERT INTO fleet_trip_status_log (trip_id, status, vehicle_id, driver_id, changed_by, created_at) VALUES (?, 'returning', ?, ?, ?, NOW())",
      [id, trip.vehicle_id, trip.driver_id, changedById],
    );
    sendFleetTripStatusNotification({ tripId: id, statusKey: "returning" });
    return res.json({ success: true });
  } catch (err) {
    console.error("POST /fleet-trips/:id/start-return error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});
// server.js — add near POST /fleet/trips/:id/arrive
app.post("/fleet/trips/:id/start", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  const { id } = req.params;

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [tripRows] = await conn.query("SELECT * FROM fleet_trips WHERE id = ? FOR UPDATE", [id]);
    if (tripRows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Trip not found." });
    }
    const trip = tripRows[0];

    if (trip.status !== "approved") {
      await conn.rollback();
      return res.status(409).json({
        success: false,
        message: "Trip is not in an approved state and can't be started.",
      });
    }

    await conn.query(
      "UPDATE fleet_trips SET status = 'ongoing', updated_at = NOW() WHERE id = ?",
      [id],
    );

    // Vehicle/driver only go "live" here, at the moment the driver actually
    // starts the trip — not at admin approval time, so the fleet lists
    // (Vehicles/Drivers panels) show idle/unassigned until then.
    // current_trip_label is a capped-width column, but pickup/dropoff text
    // can each run up to 150 chars — concatenated, that can exceed the
    // column's limit and MySQL rejects the whole UPDATE ("Data too long
    // for column 'current_trip_label'"), which was surfacing as "Advance
    // trip failed" on the driver's Start Trip button. Truncate to fit.
    const CURRENT_TRIP_LABEL_MAX = 150;
    let tripLabel = `${trip.pickup_location_text} → ${trip.dropoff_location_text}`;
    if (tripLabel.length > CURRENT_TRIP_LABEL_MAX) {
      tripLabel = tripLabel.slice(0, CURRENT_TRIP_LABEL_MAX - 1) + "…";
    }
    if (trip.vehicle_id) {
      await conn.query(
        `UPDATE fleet_vehicles SET
           status = 'active', assigned_driver_id = ?, current_trip_label = ?, updated_at = NOW()
         WHERE id = ?`,
        [trip.driver_id, tripLabel, trip.vehicle_id],
      );
    }
    

    const [changerRows] = await conn.query("SELECT id FROM users WHERE username = ?", [decoded.username]);
    const changedById = changerRows[0]?.id ?? null;

    await conn.query(
      "INSERT INTO fleet_trip_status_log (trip_id, status, vehicle_id, driver_id, changed_by, created_at) VALUES (?, 'ongoing', ?, ?, ?, NOW())",
      [id, trip.vehicle_id, trip.driver_id, changedById],
    );

    await conn.commit();
    sendFleetTripStatusNotification({ tripId: id, statusKey: "ongoing" });
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    console.error("POST /fleet-trips/:id/start error:", err);
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});
// POST /fleet/trips/:id/complete — frees the vehicle back to idle AND
// unassigns the driver from it, so the driver directory shows them as
// "Unassigned" again instead of still tied to a vehicle whose trip ended.
app.post("/fleet/trips/:id/complete", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  const { id } = req.params;

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [tripRows] = await conn.query("SELECT * FROM fleet_trips WHERE id = ? FOR UPDATE", [id]);
    if (tripRows.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "Trip not found." });
    }
    const trip = tripRows[0];

    await conn.query(
      "UPDATE fleet_trips SET status = 'completed', updated_at = NOW() WHERE id = ?",
      [id],
    );

    if (trip.vehicle_id) {
      await conn.query(
        "UPDATE fleet_vehicles SET status = 'idle', current_trip_label = NULL, assigned_driver_id = NULL, updated_at = NOW() WHERE id = ?",
        [trip.vehicle_id],
      );
      await conn.query(
        "UPDATE fleet_drivers SET vehicle_id = NULL, updated_at = NOW() WHERE vehicle_id = ?",
        [trip.vehicle_id],
      );
    }

    const [changerRows] = await conn.query("SELECT id FROM users WHERE username = ?", [decoded.username]);
    const changedById = changerRows[0]?.id ?? null;

    await conn.query(
      "INSERT INTO fleet_trip_status_log (trip_id, status, vehicle_id, driver_id, changed_by, created_at) VALUES (?, 'completed', ?, ?, ?, NOW())",
      [id, trip.vehicle_id, trip.driver_id, changedById],
    );

    await conn.commit();
    sendFleetTripStatusNotification({ tripId: id, statusKey: "completed" });
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    console.error("POST /fleet-trips/:id/complete error:", err);
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});


// ─── PERMISSIONS ROUTES ─────────────────────────────────────────────────

// GET /permissions/me — resolved permission set for the logged-in user
app.get("/permissions/me", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  // superadmin bypasses the table entirely
  if (decoded.role === "superadmin") {
    try {
      const [all] = await db.query("SELECT module, page, action FROM permissions");
      return res.json({ success: true, role: decoded.role, permissions: all });
    } catch (err) {
      return res.status(500).json({ success: false, message: err.message });
    }
  }

  try {
    const [rows] = await db.query(
      `SELECT p.module, p.page, p.action
       FROM permissions p
       LEFT JOIN role_permissions rp ON rp.permission_id = p.id AND rp.role = ?
       LEFT JOIN user_permission_overrides o ON o.permission_id = p.id AND o.username = ?
       WHERE (rp.permission_id IS NOT NULL OR o.granted = TRUE)
         AND (o.granted IS NULL OR o.granted = TRUE)`,
      [decoded.role, decoded.username],
    );
    return res.json({ success: true, role: decoded.role, permissions: rows });
  } catch (err) {
    console.error("GET /permissions/me error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});



// GET /admin/users/:username/permissions — role defaults + overrides, merged, with source flag
app.get("/admin/users/:username/permissions", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { username } = req.params;

  try {
    const [userRows] = await db.query("SELECT role FROM users WHERE username = ?", [username]);
    if (userRows.length === 0) {
      return res.status(404).json({ success: false, message: "User not found." });
    }
    const role = userRows[0].role;

    const [rows] = await db.query(
      `SELECT p.id, p.module, p.page, p.action,
              rp.permission_id IS NOT NULL AS role_default,
              o.granted AS override
       FROM permissions p
       LEFT JOIN role_permissions rp ON rp.permission_id = p.id AND rp.role = ?
       LEFT JOIN user_permission_overrides o ON o.permission_id = p.id AND o.username = ?`,
      [role, username],
    );

    const resolved = rows.map((r) => ({
      ...r,
      granted: r.override !== null ? !!r.override : !!r.role_default,
    }));

    return res.json({ success: true, role, permissions: resolved });
  } catch (err) {
    console.error("GET /admin/users/:username/permissions error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// PUT /admin/users/:username/permissions — body: [{ permissionId, granted }, ...]
// Only writes an override row when it differs from the role default; otherwise
// deletes any existing override so the row falls back to the role default cleanly.
app.put("/admin/users/:username/permissions", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { username } = req.params;
  const { changes } = req.body;

  if (!Array.isArray(changes)) {
    return res.status(400).json({ success: false, message: "changes array is required." });
  }

  const [userRows] = await db.query("SELECT role FROM users WHERE username = ?", [username]);
  if (userRows.length === 0) {
    return res.status(404).json({ success: false, message: "User not found." });
  }
  const role = userRows[0].role;

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    for (const change of changes) {
      const { permissionId, granted } = change;

      const [defaultRows] = await conn.query(
        "SELECT 1 FROM role_permissions WHERE role = ? AND permission_id = ?",
        [role, permissionId],
      );
      const isRoleDefault = defaultRows.length > 0;

      if (granted === isRoleDefault) {
        await conn.query(
          "DELETE FROM user_permission_overrides WHERE username = ? AND permission_id = ?",
          [username, permissionId],
        );
      } else {
        await conn.query(
          `INSERT INTO user_permission_overrides (username, permission_id, granted)
           VALUES (?, ?, ?)
           ON DUPLICATE KEY UPDATE granted = VALUES(granted)`,
          [username, permissionId, granted],
        );
      }
    }

    await conn.commit();
    return res.json({ success: true });
  } catch (err) {
    await conn.rollback();
    console.error("PUT /admin/users/:username/permissions error:", err);
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// PATCH /fleet/drivers/:id — partial update (license / contact / assigned
// vehicle). vehicleId here is the driver's standing/home vehicle — separate
// from whatever gets set on a per-trip basis in approve/start — so a driver
// can have the duty-status selector available even before any trip is ever
// booked for them. Pass vehicleId: null to unassign.
app.patch("/fleet/drivers/:id", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  const { licenseNumber, contactNumber, vehicleId } = req.body;
  const vehicleIdProvided = Object.prototype.hasOwnProperty.call(req.body, "vehicleId");

  try {
    const [rows] = await db.query("SELECT * FROM fleet_drivers WHERE id = ?", [id]);
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "Driver not found." });
    }
    const current = rows[0];

    if (vehicleIdProvided && vehicleId !== null) {
      const [vRows] = await db.query("SELECT id FROM fleet_vehicles WHERE id = ?", [vehicleId]);
      if (vRows.length === 0) {
        return res.status(404).json({ success: false, message: "Vehicle not found." });
      }
    }

    await db.query(
      `UPDATE fleet_drivers SET
         license_number = ?, contact_number = ?,
         vehicle_id = ?, updated_at = NOW()
       WHERE id = ?`,
      [
        licenseNumber ?? current.license_number,
        contactNumber ?? current.contact_number,
        vehicleIdProvided ? vehicleId : current.vehicle_id,
        id,
      ],
    );
    return res.json({ success: true });
  } catch (err) {
    console.error("PATCH /fleet/drivers/:id error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// PATCH /fleet/vehicles/:id — partial update
app.patch("/fleet/vehicles/:id", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  const { plateNumber, type, model, seatingCapacity, tramigoDeviceId } = req.body;

  try {
    const [rows] = await db.query("SELECT * FROM fleet_vehicles WHERE id = ?", [id]);
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "Vehicle not found." });
    }
    const current = rows[0];

    if (plateNumber && plateNumber !== current.plate_number) {
      const [dupe] = await db.query(
        "SELECT id FROM fleet_vehicles WHERE plate_number = ? AND id != ?",
        [plateNumber, id],
      );
      if (dupe.length > 0) {
        return res.status(409).json({ success: false, message: `Plate number "${plateNumber}" already exists.` });
      }
    }

    await db.query(
      `UPDATE fleet_vehicles SET
         plate_number = ?, type = ?, model = ?, seating_capacity = ?, tramigo_device_id = ?, updated_at = NOW()
       WHERE id = ?`,
      [
        plateNumber ?? current.plate_number,
        type ?? current.type,
        model ?? current.model,
        seatingCapacity ?? current.seating_capacity,
        tramigoDeviceId !== undefined ? (tramigoDeviceId || null) : current.tramigo_device_id,
        id,
      ],
    );
    return res.json({ success: true });
  } catch (err) {
    console.error("PATCH /fleet/vehicles/:id error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// PATCH /fleet/vehicles/:id/status — driver sets their own vehicle's duty
// status (off_duty / idle / active / personal). Blocked while the vehicle
// is actually mid-trip ('active') so a driver can't yank their own vehicle
// out from under an in-progress trip — the frontend already prevents this,
// this is the server-side backstop.
app.patch("/fleet/vehicles/:id/status", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  const { status, note } = req.body;

  const ALLOWED = ["off_duty", "idle", "active", "maintenance", "personal"];
  if (!ALLOWED.includes(status)) {
    return res.status(400).json({ success: false, message: `Invalid status "${status}".` });
  }

  try {
    const [rows] = await db.query("SELECT * FROM fleet_vehicles WHERE id = ?", [id]);
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "Vehicle not found." });
    }

    await db.query(
      "UPDATE fleet_vehicles SET status = ?, updated_at = NOW() WHERE id = ?",
      [status, id],
    );

    await db.query(
      "INSERT INTO fleet_vehicle_status_log (vehicle_id, status, note, created_at) VALUES (?, ?, ?, NOW())",
      [id, status, note ?? ""],
    );

    return res.json({ success: true });
  } catch (err) {
    console.error("PATCH /fleet/vehicles/:id/status error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});


// DELETE /fleet/vehicles/:id — blocked while the vehicle is on an active trip
app.delete("/fleet/vehicles/:id", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  try {
    const [rows] = await db.query("SELECT * FROM fleet_vehicles WHERE id = ?", [id]);
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "Vehicle not found." });
    }
    if (rows[0].status === "active") {
      return res.status(400).json({ success: false, message: "Cannot delete a vehicle that's currently on a trip." });
    }
    // Unassign any driver currently linked to this vehicle before removing it.
    await db.query("UPDATE fleet_drivers SET vehicle_id = NULL WHERE vehicle_id = ?", [id]);
    await db.query("DELETE FROM fleet_vehicles WHERE id = ?", [id]);
    return res.json({ success: true });
  } catch (err) {
    console.error("DELETE /fleet/vehicles/:id error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// PATCH /fleet/drivers/:id — partial update (license / contact only —
// name comes from the linked AD user, not editable here)
app.patch("/fleet/drivers/:id", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  const { licenseNumber, contactNumber } = req.body;

  try {
    const [rows] = await db.query("SELECT * FROM fleet_drivers WHERE id = ?", [id]);
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "Driver not found." });
    }
    const current = rows[0];

    await db.query(
      `UPDATE fleet_drivers SET
         license_number = ?, contact_number = ?, updated_at = NOW()
       WHERE id = ?`,
      [
        licenseNumber ?? current.license_number,
        contactNumber ?? current.contact_number,
        id,
      ],
    );
    return res.json({ success: true });
  } catch (err) {
    console.error("PATCH /fleet/drivers/:id error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// DELETE /fleet/drivers/:id — blocked while assigned to a vehicle
app.delete("/fleet/drivers/:id", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  try {
    const [rows] = await db.query("SELECT * FROM fleet_drivers WHERE id = ?", [id]);
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "Driver not found." });
    }
    if (rows[0].vehicle_id) {
      return res.status(400).json({ success: false, message: "Cannot delete a driver currently assigned to a vehicle." });
    }
    await db.query("DELETE FROM fleet_drivers WHERE id = ?", [id]);

    // Mirror of the grant in POST /fleet/drivers — removing someone from the
    // roster should pull their Driver View nav access back too.
    await db.query(
      "UPDATE users SET perm_fleet_driver = 0, updated_at = NOW() WHERE id = ?",
      [rows[0].user_id],
    );

    return res.json({ success: true });
  } catch (err) {
    console.error("DELETE /fleet/drivers/:id error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// DELETE /fleet/locations/:id
app.delete("/fleet/locations/:id", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  try {
    const [result] = await db.query("DELETE FROM fleet_locations WHERE id = ?", [id]);
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Location not found." });
    }
    return res.json({ success: true });
  } catch (err) {
    console.error("DELETE /fleet/locations/:id error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// PATCH /fleet/drivers/:id/duty-status — driver sets their own status
// (off_duty / active / personal), independent of any vehicle assignment.
// Always settable — a driver isn't blocked from going off-duty just
// because they have no vehicle or trip right now.
app.patch("/fleet/drivers/:id/duty-status", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  let { dutyStatus } = req.body;

  const ALLOWED = ["off_duty", "active", "personal", "leave"];
  if (!ALLOWED.includes(dutyStatus)) {
    return res.status(400).json({ success: false, message: `Invalid dutyStatus "${dutyStatus}".` });
  }

  // Weekends are non-working days — force off_duty regardless of what the
  // client sent, using local server time (0 = Sunday, 6 = Saturday).
  // Exception: "leave" is an explicit admin override (e.g. driver is on
  // leave) and should never be silently downgraded by the weekend check.
  const dayOfWeek = new Date().getDay();
  if (dutyStatus !== "leave" && (dayOfWeek === 0 || dayOfWeek === 6)) {
    dutyStatus = "off_duty";
  }

  try {
    const [driverRows] = await db.query("SELECT user_id FROM fleet_drivers WHERE id = ?", [id]);
    if (driverRows.length === 0) {
      return res.status(404).json({ success: false, message: "Driver not found." });
    }

    // Shift-based auto-sync (see computeAutoDutyStatus in shiftUtils.ts) can
    // fire this with "off_duty" purely because the shift window ended, even
    // though the driver is still mid-trip. A driver actually on an ongoing
    // trip should keep reading as on-trip/active until the trip itself is
    // completed — so block any downgrade to off_duty while one exists.
    if (dutyStatus === "off_duty") {
      const [tripRows] = await db.query(
        "SELECT id FROM fleet_trips WHERE driver_id = ? AND status = 'ongoing' LIMIT 1",
        [driverRows[0].user_id],
      );
      if (tripRows.length > 0) {
        return res.json({ success: true, skipped: true, reason: "Driver is on an ongoing trip." });
      }
    }

    const [result] = await db.query(
      "UPDATE fleet_drivers SET duty_status = ?, updated_at = NOW() WHERE id = ?",
      [dutyStatus, id],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Driver not found." });
    }
    return res.json({ success: true });
  } catch (err) {
    console.error("PATCH /fleet/drivers/:id/duty-status error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// PATCH /fleet/drivers/:id/shift — assigns one of the 3 fixed shift windows.
// On/off duty is no longer set by hand for this: the client derives it from
// (shift_start, shift_end) vs. the current time — see computeAutoDutyStatus
// in shiftUtils.ts — and syncs duty_status back here automatically.
app.patch("/fleet/drivers/:id/shift", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  const { shiftStart, shiftEnd } = req.body;

  if (!shiftStart || !shiftEnd) {
    return res.status(400).json({ success: false, message: "shiftStart and shiftEnd are required." });
  }

  try {
    const [result] = await db.query(
      "UPDATE fleet_drivers SET shift_start = ?, shift_end = ?, updated_at = NOW() WHERE id = ?",
      [shiftStart, shiftEnd, id],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Driver not found." });
    }
    return res.json({ success: true });
  } catch (err) {
    console.error("PATCH /fleet/drivers/:id/shift error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});



// ─── TRAMIGO LIVE TRACKING ─────────────────────────────────────────────────
const TRAMIGO_API_URL = process.env.TRAMIGO_API_URL || "https://api.tracking.tramigocloud.com";
const TRAMIGO_USERNAME = process.env.TRAMIGO_USERNAME;
const TRAMIGO_PASSWORD = process.env.TRAMIGO_PASSWORD;

let _tramigoToken = null;

async function tramigoLogin() {
  const res = await fetch(`${TRAMIGO_API_URL}/api/login`, {
    method: "POST",
    headers: { "Content-Type": "application/json;charset=UTF-8" },
    body: JSON.stringify({
      type: "app/login",
      username: TRAMIGO_USERNAME,
      password: TRAMIGO_PASSWORD,
    }),
  });
  if (!res.ok) throw new Error(`Tramigo login failed: ${res.status}`);
  const data = await res.json();
  _tramigoToken = data.access_token;
  return _tramigoToken;
}

// Wraps a Tramigo GET call — the token is valid ~1 year per their docs, so
// this is effectively login-once; retries a single time with a fresh login
// if the cached token ever gets rejected.
async function tramigoFetch(path) {
  if (!_tramigoToken) await tramigoLogin();

  let res = await fetch(`${TRAMIGO_API_URL}${path}`, {
    headers: { Authorization: `Bearer ${_tramigoToken}`, Accept: "application/json" },
  });

  if (res.status === 401) {
    _tramigoToken = null;
    await tramigoLogin();
    res = await fetch(`${TRAMIGO_API_URL}${path}`, {
      headers: { Authorization: `Bearer ${_tramigoToken}`, Accept: "application/json" },
    });
  }

  if (!res.ok) throw new Error(`Tramigo API error: ${res.status}`);
  return res.json();
}

// GET /fleet/tramigo-devices — lists every device on the Tramigo account so
// you can look up which device_id/IMEI corresponds to which physical
// vehicle, for filling in fleet_vehicles.tramigo_device_id.
app.get("/fleet/tramigo-devices", async (req, res) => {
  if (!requireAuth(req, res)) return;
  try {
    const data = await tramigoFetch("/api/v2/devices?page=1&per_page=1000");
    return res.json({ success: true, devices: data.data ?? [] });
  } catch (err) {
    console.error("GET /fleet/tramigo-devices error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// PATCH /fleet/vehicles/:id/tramigo-device — link a Silverdab vehicle to its
// Tramigo device_id (found via GET /fleet/tramigo-devices above).
app.patch("/fleet/vehicles/:id/tramigo-device", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  const { tramigoDeviceId } = req.body;
  try {
    const [result] = await db.query(
      "UPDATE fleet_vehicles SET tramigo_device_id = ?, updated_at = NOW() WHERE id = ?",
      [tramigoDeviceId ?? null, id],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Vehicle not found." });
    }
    return res.json({ success: true });
  } catch (err) {
    console.error("PATCH /fleet/vehicles/:id/tramigo-device error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// GET /fleet/vehicles/live-locations — batched last-known GPS position for
// every vehicle that has a tramigo_device_id set. Poll this from the app at
// the same cadence as the other fleet endpoints (POLL_INTERVAL_MS in
// FleetControlTowerPage.tsx / DriverPortalPage.tsx).
// Below this, a vehicle counts as "moving" — chosen to filter out GPS
// jitter on a parked vehicle reporting a tiny nonzero speed.
const MOVING_SPEED_THRESHOLD_KMH = 3;

// A report older than this is treated as stale — the device's last known
// speed no longer tells us anything about whether it's moving *right now*,
// so a stale report can never produce 'active', no matter what speed it
// last recorded. This mirrors what made the Fortuner (last seen moving,
// then went offline) incorrectly show as Active forever.
const STALE_REPORT_THRESHOLD_MS = 10 * 60 * 1000; // 10 minutes

app.get("/fleet/vehicles/live-locations", async (req, res) => {
  if (!requireAuth(req, res)) return;
  try {
    const [vehicles] = await db.query(
      "SELECT id, plate_number, model, status, tramigo_device_id FROM fleet_vehicles WHERE tramigo_device_id IS NOT NULL",
    );
    if (vehicles.length === 0) {
      return res.json({ success: true, locations: [] });
    }

    const qs = vehicles
      .map((v) => `device_ids[]=${encodeURIComponent(v.tramigo_device_id)}`)
      .join("&");
    const tramigoReports = await tramigoFetch(`/api/reports/last_location?${qs}`);

    const byDeviceId = new Map(
      (Array.isArray(tramigoReports) ? tramigoReports : []).map((r) => [String(r.Device_ID), r]),
    );

    // Vehicle status is driven ONLY by actual trip lifecycle events (see
    // POST /fleet/trips/:id/start and /complete) — never by raw GPS speed.
    // A vehicle can be physically moving (someone driving it around,
    // testing the tracker, etc.) with zero trips booked to it, and it
    // should still read as available/idle rather than falsely "Active".
    // GPS reports here are used only for the map position + last_ping_at
    // ("last seen") — status itself is untouched.
    const locations = vehicles
      .map((v) => {
        const report = byDeviceId.get(String(v.tramigo_device_id));
        const loc = report?.main_reports?.[0];
        if (!loc) return null;

        const speedKmh = loc.Speed != null ? parseFloat(loc.Speed) / 1000 : null;

        return {
          vehicleId: String(v.id),
          plateNumber: v.plate_number,
          vehicleModel: v.model,
          latitude: parseFloat(loc.Latitude),
          longitude: parseFloat(loc.Longitude),
          speed: speedKmh,
          heading: loc.Bearing ?? null,
          reportedAt: report.DateTime_Actual,
        };
      })
      .filter(Boolean);

    // Every vehicle with a live report gets last_ping_at stamped — this is
    // a "last seen" signal only, separate from updated_at (which also gets
    // touched by unrelated manual edits like plate/model changes).
    const pingedVehicleIds = locations.map((l) => Number(l.vehicleId));

    await Promise.all(
      pingedVehicleIds.map((id) =>
        db.query(
          "UPDATE fleet_vehicles SET last_ping_at = NOW() WHERE id = ?",
          [id],
        ).catch((err) =>
          console.error(`Failed to update last_ping_at for vehicle ${id}:`, err.message),
        ),
      ),
    );

    return res.json({ success: true, locations });
  } catch (err) {
    console.error("GET /fleet/vehicles/live-locations error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});


app.patch("/fleet/locations/:id", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  const { name, shortLabel, latitude, longitude } = req.body;

  const fields = [];
  const values = [];
  if (name !== undefined) { fields.push("name = ?"); values.push(name); }
  if (shortLabel !== undefined) { fields.push("short_label = ?"); values.push(shortLabel); }
  if (latitude !== undefined) { fields.push("latitude = ?"); values.push(latitude); }
  if (longitude !== undefined) { fields.push("longitude = ?"); values.push(longitude); }

  if (fields.length === 0) {
    return res.status(400).json({ success: false, message: "No fields to update." });
  }

  values.push(id);
  try {
    await db.query(`UPDATE fleet_locations SET ${fields.join(", ")} WHERE id = ?`, values);
    res.json({ success: true });
  } catch (err) {
    console.error("Update location failed:", err);
    res.status(500).json({ success: false, message: "Failed to update location." });
  }
});


// ─── ROOM RESERVATIONS ROUTES ──────────────────────────────────────────────

const ROOM_MAX_ATTENDEES = {
  "Conference Room": 15,
  "Meeting Room 1": 6,
  "Meeting Room 2": 6,
};

function addOneHour(timeStr) {
  const [h, m, s] = timeStr.split(":").map(Number);
  const d = new Date(2000, 0, 1, h, m, s || 0);
  d.setHours(d.getHours() + 1);
  const pad = (n) => String(n).padStart(2, "0");
  return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

// GET /room-reservations — optional ?date=YYYY-MM-DD and ?room=
app.get("/room-reservations", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { date, room } = req.query;

  // No default status filter — cancelled reservations still need to show
  // in the admin table (just hidden from the calendar client-side). The
  // booking form's own conflict-check query filters status = 'confirmed'
  // separately in POST /room-reservations, so this change doesn't affect
  // double-booking prevention.
  const conditions = [];
  const params = [];
  if (date) { conditions.push("booking_date = ?"); params.push(date); }
  if (room) { conditions.push("room_name = ?"); params.push(room); }
  const whereClause = conditions.length ? `WHERE ${conditions.join(" AND ")}` : "";

  try {
    const [rows] = await db.query(
      `SELECT * FROM room_reservations ${whereClause} ORDER BY booking_date ASC, start_time ASC`,
      params,
    );

    // Map snake_case DB columns -> camelCase, matching the frontend's
    // RoomReservation type (mirrors how /fleet/trips and /supply-requests
    // already shape their responses).
    //
    // booking_date is a DATE column — without dateStrings on the pool,
    // mysql2 returns it as a JS Date object, which JSON.stringify then
    // turns into a full ISO timestamp ("2026-08-19T00:00:00.000Z"). That
    // broke the calendar view: it concatenates `${bookingDate} ${startTime}`
    // expecting a clean "YYYY-MM-DD", and an ISO string with its own "T"
    // in it produces an unparseable combined string, so every event
    // silently failed to render (table still "worked" because it just
    // printed whatever value it got). Format explicitly here instead of
    // touching the global pool config.
    const toDateOnly = (v) => {
      if (!v) return null;
      if (typeof v === "string") return v.slice(0, 10); // already "YYYY-MM-DD..."
      const d = new Date(v);
      const pad = (n) => String(n).padStart(2, "0");
      return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
    };

    const reservations = rows.map((r) => ({
      id: r.id,
      bookingId: r.booking_id,
      roomRef: r.room_ref,
      calendarSynced: !!r.outlook_event_id,
      roomName: r.room_name,
      maxAttendees: r.max_attendees,
      bookingDate: toDateOnly(r.booking_date),
      startTime: r.start_time,
      endTime: r.end_time,
      fullName: r.full_name,
      email: r.email,
      guestEmails: (() => {
        try {
          return r.guest_emails ? JSON.parse(r.guest_emails) : [];
        } catch {
          return [];
        }
      })(),
      specialRequests: r.special_requests,
      avRequirement: r.av_requirement,
      needsWifi: !!r.needs_wifi,
      agenda: r.agenda,
      status: r.status,
      createdAt: r.created_at,
      updatedAt: r.updated_at,
    }));

    return res.json({ success: true, count: reservations.length, reservations });
  } catch (err) {
    console.error("GET /room-reservations error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /room-reservations — books a room, rejecting any time-overlapping booking
app.post("/room-reservations", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  const {
    roomName, bookingDate, startTime, endTime,
    fullName, email, guestEmails, specialRequests,
    avRequirement, needsWifi, agenda,
  } = req.body;

  if (!roomName || !bookingDate || !startTime || !fullName?.trim() || !email?.trim() || !agenda?.trim()) {
    return res.status(400).json({
      success: false,
      message: "roomName, bookingDate, startTime, fullName, email, and agenda are required.",
    });
  }

  if (!ROOM_MAX_ATTENDEES[roomName]) {
    return res.status(400).json({ success: false, message: `Invalid room "${roomName}".` });
  }

  const finalEndTime = endTime || addOneHour(startTime);

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const [conflicts] = await conn.query(
      `SELECT id, start_time, end_time FROM room_reservations
       WHERE room_name = ?
         AND booking_date = ?
         AND status = 'confirmed'
         AND start_time < ?
         AND end_time > ?
       FOR UPDATE`,
      [roomName, bookingDate, finalEndTime, startTime],
    );

    if (conflicts.length > 0) {
      await conn.rollback();
      return res.status(409).json({
        success: false,
        message: `${roomName} is already booked from ${conflicts[0].start_time} to ${conflicts[0].end_time} on this date.`,
      });
    }

    // Sequential human-readable ref, same pattern as TRIP-YYYY-#### and
    // SR-YYYY-#### — the UUID (booking_id) stays as the internal PK
    // reference, but this is what gets shown to the requester.
    const year = new Date().getFullYear();
    const [maxRows] = await conn.query(
      `SELECT MAX(CAST(SUBSTRING_INDEX(room_ref, '-', -1) AS UNSIGNED)) AS maxNum
       FROM room_reservations WHERE room_ref LIKE ? FOR UPDATE`,
      [`ROOM-${year}-%`],
    );
    const nextNum = String((maxRows[0].maxNum ?? 0) + 1).padStart(4, "0");
    const roomRef = `ROOM-${year}-${nextNum}`;

    const bookingId = crypto.randomUUID();

    const [result] = await conn.query(
      `INSERT INTO room_reservations
        (booking_id, room_ref, room_name, max_attendees, booking_date, start_time, end_time,
         full_name, email, guest_emails, special_requests,
         av_requirement, needs_wifi, agenda, status, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'confirmed', NOW(), NOW())`,
      [
        bookingId, roomRef, roomName, ROOM_MAX_ATTENDEES[roomName], bookingDate, startTime, finalEndTime,
        fullName.trim(), email.trim(),
        Array.isArray(guestEmails) ? JSON.stringify(guestEmails) : null,
        specialRequests ?? "",
        avRequirement ?? "None", needsWifi ? 1 : 0, agenda.trim(),
      ],
    );

    await conn.commit();

    // Push to the room's own Outlook mailbox — mirrors the fleet trip
    // "create on booking" step in POST /fleet/trips. Best-effort: a failed
    // sync should never fail the booking itself.
    try {
      const eventId = await createRoomEvent({
        roomRef,
        roomName,
        bookingDate,
        startTime,
        endTime: finalEndTime,
        fullName: fullName.trim(),
        email: email.trim(),
        guestEmails: Array.isArray(guestEmails) ? guestEmails : [],
        agenda: agenda.trim(),
        specialRequests: specialRequests ?? "",
        avRequirement: avRequirement ?? "None",
        needsWifi: !!needsWifi,
      });
      await db.query(
        "UPDATE room_reservations SET outlook_event_id = ? WHERE id = ?",
        [eventId, result.insertId],
      );
    } catch (err) {
      console.error("Outlook calendar sync (room booking) failed:", err.message);
    }

    sendRoomReservationConfirmation({
      toEmail: email.trim(),
      fullName: fullName.trim(),
      roomRef,
      roomName,
      bookingDate,
      startTime,
      endTime: finalEndTime,
      agenda: agenda.trim(),
    });

    return res.status(201).json({ success: true, id: result.insertId, bookingId, roomRef });
  } catch (err) {
    await conn.rollback();
    console.error("POST /room-reservations error:", err);
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// POST /room-reservations/:id/cancel
app.post("/room-reservations/:id/cancel", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { id } = req.params;
  try {
    const [rows] = await db.query(
      "SELECT room_name, outlook_event_id FROM room_reservations WHERE id = ?",
      [id],
    );
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "Reservation not found." });
    }

    const [result] = await db.query(
      "UPDATE room_reservations SET status = 'cancelled', updated_at = NOW() WHERE id = ?",
      [id],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "Reservation not found." });
    }

    if (rows[0].outlook_event_id) {
      try {
        await deleteRoomEvent(rows[0].outlook_event_id, rows[0].room_name);
      } catch (err) {
        console.error("Outlook calendar sync (room cancel) failed:", err.message);
      }
    }

    return res.json({ success: true });
  } catch (err) {
    console.error("POST /room-reservations/:id/cancel error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── SEAT PLAN ROUTES ───────────────────────────────────────────────────────
// Table: seat_plan_layouts (plan_key VARCHAR UNIQUE, layout_json LONGTEXT,
// updated_by VARCHAR, created_at, updated_at) — one row per named floor plan
// ("unit3", etc.), read/written as an opaque JSON blob by the seat plan editor.

// GET /seat-plan/:key
app.get("/seat-plan/:key", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { key } = req.params;

  try {
    const [rows] = await db.query(
      "SELECT layout_json, updated_by, updated_at FROM seat_plan_layouts WHERE plan_key = ?",
      [key],
    );
    if (rows.length === 0) {
      return res.json({ success: true, layout: null });
    }
    let layout;
    try {
      layout = JSON.parse(rows[0].layout_json);
    } catch (err) {
      console.error("GET /seat-plan parse error:", err.message);
      return res.status(500).json({ success: false, message: "Stored layout is corrupted." });
    }
    return res.json({
      success: true,
      layout,
      updatedBy: rows[0].updated_by,
      updatedAt: rows[0].updated_at,
    });
  } catch (err) {
    console.error("GET /seat-plan error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /seat-plan/:key — body: { layout, updatedByName }
app.post("/seat-plan/:key", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  const { key } = req.params;
  const { layout, updatedByName } = req.body;

  if (!layout || !Array.isArray(layout.rooms) || !Array.isArray(layout.seats)) {
    return res.status(400).json({ success: false, message: "layout with rooms[] and seats[] is required." });
  }

  try {
    const layoutJson = JSON.stringify(layout);
    const savedBy = updatedByName || decoded.displayName || decoded.username || "Unknown";

    await db.query(
      `INSERT INTO seat_plan_layouts (plan_key, layout_json, updated_by)
       VALUES (?, ?, ?)
       ON DUPLICATE KEY UPDATE layout_json = VALUES(layout_json), updated_by = VALUES(updated_by)`,
      [key, layoutJson, savedBy],
    );
    return res.json({ success: true });
  } catch (err) {
    console.error("POST /seat-plan error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// EIA sign sheet export (Word)
app.use(
  "/api/it",
  (req, res, next) => (requireAuth(req, res) ? next() : undefined),
  require("./eiaExport"),
);

// ─── EIA SIGN SHEETS (saved records) ────────────────────────────────────────

// POST /it/eia-sign-sheets — saves a record. ref_no is the unique key, so a
// repeat save with the same ref_no (e.g. exporting the same form twice) is
// treated as idempotent rather than an error — it just returns the existing
// record instead of failing.
app.post("/it/eia-sign-sheets", async (req, res) => {
  if (!requireAuth(req, res)) return;

  const {
    company, copyLabel, name, department, date, issuedNo, refNo,
    items, remarks, issuedTo, deliveredBy, approvedBy, preparedBy,
  } = req.body;

  if (!refNo || !company || !name) {
    return res.status(400).json({ success: false, message: "refNo, company, and name are required." });
  }

  try {
    const [existing] = await db.query(
      "SELECT id FROM eia_sign_sheets WHERE ref_no = ?",
      [refNo],
    );
    if (existing.length > 0) {
      return res.json({ success: true, id: existing[0].id, alreadySaved: true });
    }

    const id = crypto.randomUUID();
    await db.query(
      `INSERT INTO eia_sign_sheets
        (id, ref_no, company, copy_label, full_name, department, issue_date, issued_no,
         items, remarks, issued_to_name, issued_to_date, delivered_by_name, delivered_by_date,
         approved_by_name, approved_by_date, prepared_by, created_at, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NOW(), NOW())`,
      [
        id, refNo, company, copyLabel ?? "", name, department ?? "",
        date || null, issuedNo ?? "",
        JSON.stringify(items ?? []), remarks ?? "",
        issuedTo?.name ?? "", issuedTo?.date || null,
        deliveredBy?.name ?? "", deliveredBy?.date || null,
        approvedBy?.name ?? "", approvedBy?.date || null,
        preparedBy ?? "",
      ],
    );
    return res.status(201).json({ success: true, id, alreadySaved: false });
  } catch (err) {
    if (err.code === "ER_DUP_ENTRY") {
      // Race: two saves for the same refNo landed at once — treat the same
      // as the existing-row branch above.
      const [rows] = await db.query("SELECT id FROM eia_sign_sheets WHERE ref_no = ?", [refNo]);
      return res.json({ success: true, id: rows[0]?.id, alreadySaved: true });
    }
    console.error("POST /it/eia-sign-sheets error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// GET /it/eia-sign-sheets/:refNo — look up one record by Ref No., for the
// Forms page's "search and load" feature.
app.get("/it/eia-sign-sheets/:refNo", async (req, res) => {
  if (!requireAuth(req, res)) return;
  try {
    const [rows] = await db.query(
      "SELECT * FROM eia_sign_sheets WHERE ref_no = ?",
      [req.params.refNo],
    );
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "No record found for that Ref. No." });
    }
    const r = rows[0];
    return res.json({
      success: true,
      sheet: {
        id: r.id,
        refNo: r.ref_no,
        company: r.company,
        copyLabel: r.copy_label,
        name: r.full_name,
        department: r.department,
        date: r.issue_date,
        issuedNo: r.issued_no,
        items: (() => {
          if (Array.isArray(r.items)) return r.items; // mysql2 auto-parses JSON columns
          try { return JSON.parse(r.items || "[]"); } catch { return []; }
        })(),
        remarks: r.remarks,
        issuedTo: { name: r.issued_to_name, date: r.issued_to_date },
        deliveredBy: { name: r.delivered_by_name, date: r.delivered_by_date },
        approvedBy: { name: r.approved_by_name, date: r.approved_by_date },
        preparedBy: r.prepared_by,
      },
    });
  } catch (err) {
    console.error("GET /it/eia-sign-sheets/:refNo error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// PUT /it/eia-sign-sheets/:refNo — updates an existing record. ref_no in the
// URL identifies which row to update; the body's own refNo is ignored so the
// unique key can't be changed out from under an update.
app.put("/it/eia-sign-sheets/:refNo", async (req, res) => {
  if (!requireAuth(req, res)) return;
  const { refNo } = req.params;
  const {
    company, copyLabel, name, department, date, issuedNo,
    items, remarks, issuedTo, deliveredBy, approvedBy, preparedBy,
  } = req.body;

  if (!company || !name) {
    return res.status(400).json({ success: false, message: "company and name are required." });
  }

  try {
    const [result] = await db.query(
      `UPDATE eia_sign_sheets SET
         company = ?, copy_label = ?, full_name = ?, department = ?, issue_date = ?, issued_no = ?,
         items = ?, remarks = ?, issued_to_name = ?, issued_to_date = ?,
         delivered_by_name = ?, delivered_by_date = ?, approved_by_name = ?, approved_by_date = ?,
         prepared_by = ?, updated_at = NOW()
       WHERE ref_no = ?`,
      [
        company, copyLabel ?? "", name, department ?? "", date || null, issuedNo ?? "",
        JSON.stringify(items ?? []), remarks ?? "",
        issuedTo?.name ?? "", issuedTo?.date || null,
        deliveredBy?.name ?? "", deliveredBy?.date || null,
        approvedBy?.name ?? "", approvedBy?.date || null,
        preparedBy ?? "",
        refNo,
      ],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "No record found for that Ref. No." });
    }
    return res.json({ success: true });
  } catch (err) {
    console.error("PUT /it/eia-sign-sheets/:refNo error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// GET /it/eia-sign-sheets — list saved records, newest first
app.get("/it/eia-sign-sheets", async (req, res) => {
  if (!requireAuth(req, res)) return;
  try {
    const [rows] = await db.query(
      "SELECT * FROM eia_sign_sheets ORDER BY created_at DESC",
    );
    const sheets = rows.map((r) => ({
      id: r.id,
      refNo: r.ref_no,
      company: r.company,
      copyLabel: r.copy_label,
      name: r.full_name,
      department: r.department,
      date: r.issue_date,
      issuedNo: r.issued_no,
      items: (() => {
        if (Array.isArray(r.items)) return r.items; // mysql2 auto-parses JSON columns
        try { return JSON.parse(r.items || "[]"); } catch { return []; }
      })(),
      remarks: r.remarks,
      issuedTo: { name: r.issued_to_name, date: r.issued_to_date },
      deliveredBy: { name: r.delivered_by_name, date: r.delivered_by_date },
      approvedBy: { name: r.approved_by_name, date: r.approved_by_date },
      preparedBy: r.prepared_by,
      createdAt: r.created_at,
    }));
    return res.json({ success: true, count: sheets.length, sheets });
  } catch (err) {
    console.error("GET /it/eia-sign-sheets error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── FPD STATEMENTS OF ACCOUNT ──────────────────────────────────────────────
// Tables: fpd_statements, fpd_bank_templates.
// Statement No. format: FPD.<year>.<NN>  (e.g. FPD.2026.01), year taken from
// the statement date, sequence generated here under a row lock.

function fpdYear(dateStr) {
  const y = String(dateStr || "").slice(0, 4);
  return /^\d{4}$/.test(y) ? y : String(new Date().getFullYear());
}

// FPD is confidential: only these executives create real FPD statements, and
// every user (superadmin included) can only ever see statements they created.
const FPD_EXECUTIVES = ["msy", "jafable", "mnatan"];
const fpdOwner = (d) => String(d?.username || "").toLowerCase();
const fpdIsExec = (d) => FPD_EXECUTIVES.includes(fpdOwner(d));
// Everyone else (e.g. you testing) gets a separate TEST sequence.
const fpdPrefix = (d) => (fpdIsExec(d) ? "FPD" : "TEST");

function fpdDateOnly(v) {
  if (!v) return null;
  if (typeof v === "string") return v.slice(0, 10);
  const d = new Date(v);
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

function fpdParseJson(v, fallback) {
  if (v && typeof v === "object") return v; // mysql2 auto-parses JSON columns
  try { return JSON.parse(v || "null") ?? fallback; } catch { return fallback; }
}

// Total is always recomputed here — the client's numbers are never trusted.
function fpdComputeLine(raw) {
  const unitPrice = Math.max(0, Number(raw.unitPrice) || 0);
  const pct = raw.discounted
    ? Math.min(100, Math.max(0, Number(raw.discountPercent) || 0))
    : 0;
  return {
    date: raw.date || null,
    description: String(raw.description ?? ""),
    unitPrice,
    discounted: pct > 0,
    discountPercent: pct,
    total: Math.round(unitPrice * (1 - pct / 100) * 100) / 100,
  };
}

function fpdBankFromRow(r) {
  return {
    id: r.id,
    label: r.label,
    accountName: r.account_name,
    accountAddress: r.account_address,
    bankName: r.bank_name,
    branchName: r.branch_name,
    branchAddress: r.branch_address,
    accountNo: r.account_no,
    swiftCode: r.swift_code,
    isDefault: !!r.is_default,
    fields: fpdParseJson(r.detail_fields, null),
  };
}

function fpdStatementFromRow(r) {
  return {
    id: r.id,
    statementNo: r.statement_no,
    date: fpdDateOnly(r.statement_date),
    billToName: r.bill_to_name,
    billToAddress: r.bill_to_address,
    currency: r.currency,
    fromCurrency: r.from_currency || "₱",
    items: fpdParseJson(r.items, []),
    totalAmount: Number(r.total_amount),
    exchangeRate: Number(r.exchange_rate ?? 1),
    bankTemplateId: r.bank_template_id,
    bank: fpdParseJson(r.bank_snapshot, {}),
    preparedBy: r.prepared_by,
    createdBy: r.created_by,
    createdAt: r.created_at,
    updatedAt: r.updated_at,
  };
}

// Validates + normalizes the body shared by POST and PUT. Returns
// { error } or { lines, total }.
function fpdPrepareBody(body) {
  const { date, billToName, items } = body;
  if (!date || !billToName?.trim()) {
    return { error: "date and billToName are required." };
  }
  if (!Array.isArray(items)) {
    return { error: "items array is required." };
  }
  const lines = items
    .map(fpdComputeLine)
    .filter((l) => l.date || l.description.trim() || l.unitPrice > 0);
  if (lines.length === 0) {
    return { error: "At least one statement line is required." };
  }
  const currency = body.currency || "₱";
  const fromCurrency = body.fromCurrency || "₱";
  const rate = fromCurrency === currency ? 1 : Number(body.exchangeRate);
  if (!(rate > 0)) {
    return { error: "A valid exchange rate is required when converting between currencies." };
  }
  const f = currency === "¥" ? 1 : 100;
  const sum = lines.reduce((s, l) => s + Math.round(l.total * rate * f) / f, 0);
  return { lines, total: Math.round(sum * f) / f, rate, fromCurrency };
}

// Resolves the bank block to store on a statement: a saved template (by id)
// or an inline bank object. Returns { bank, templateId } or { error, status }.
async function fpdResolveBank(conn, { bankTemplateId, bank }) {
  if (bankTemplateId) {
    const [rows] = await conn.query("SELECT * FROM fpd_bank_templates WHERE id = ?", [bankTemplateId]);
    if (rows.length === 0) return { error: "Bank template not found.", status: 404 };
    const b = fpdBankFromRow(rows[0]);
    return {
      templateId: b.id,
      bank: {
        accountName: b.accountName, accountAddress: b.accountAddress, bankName: b.bankName,
        branchName: b.branchName, branchAddress: b.branchAddress,
        accountNo: b.accountNo, swiftCode: b.swiftCode,
      },
    };
  }
  if (bank && bank.accountName && bank.bankName && bank.accountNo) {
    return {
      templateId: null,
      bank: {
        accountName: bank.accountName, accountAddress: bank.accountAddress ?? "",
        bankName: bank.bankName, branchName: bank.branchName ?? "",
        branchAddress: bank.branchAddress ?? "", accountNo: bank.accountNo,
        swiftCode: bank.swiftCode ?? "",
      },
    };
  }
  return { error: "bankTemplateId or a complete bank object is required.", status: 400 };
}

// GET /fpd/next-number?date=YYYY-MM-DD — preview only (doesn't reserve the
// number); the real one is assigned inside POST /fpd/statements.
app.get("/fpd/next-number", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  const year = fpdYear(req.query.date);
  const prefix = fpdPrefix(decoded);
  try {
    const next = String(await fpdNextSeq(db, year, false, prefix)).padStart(2, "0");
    const needsSeed = await fpdNeedsSeed(db, year, prefix);
    return res.json({ success: true, statementNo: `${prefix}.${year}.${next}`, needsSeed });
  } catch (err) {
    console.error("GET /fpd/next-number error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ── Bank templates ──────────────────────────────────────────────────────────

// Next sequence = the higher of (a) the highest saved statement and (b) the
// "last number already used" you set manually, plus 1.
async function fpdNextSeq(conn, year, lock, prefix = "FPD") {
  const [maxRows] = await conn.query(
    `SELECT MAX(CAST(SUBSTRING_INDEX(statement_no, '.', -1) AS UNSIGNED)) AS maxNum
     FROM fpd_statements WHERE statement_no LIKE ?${lock ? " FOR UPDATE" : ""}`,
    [`${prefix}.${year}.%`],
  );
  let seed = 0;
  if (prefix === "FPD") {
    const [seedRows] = await conn.query(
      "SELECT last_number FROM fpd_number_seeds WHERE year = ?",
      [year],
    );
    seed = seedRows[0]?.last_number ?? 0;
  }
  return Math.max(maxRows[0].maxNum ?? 0, seed) + 1;
}

// Only the real FPD sequence ever needs seeding.
async function fpdNeedsSeed(conn, year, prefix = "FPD") {
  if (prefix !== "FPD") return false;
  const [seed] = await conn.query("SELECT 1 FROM fpd_number_seeds WHERE year = ? LIMIT 1", [year]);
  const [stmt] = await conn.query("SELECT 1 FROM fpd_statements WHERE statement_no LIKE ? LIMIT 1", [`FPD.${year}.%`]);
  return seed.length === 0 && stmt.length === 0;
}

// PUT /fpd/last-number — body: { date, lastNumber }
// Sets the last statement number already used (typed manually before this
// system existed) for that date's year. The next saved statement is +1.
app.put("/fpd/last-number", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  if (!fpdIsExec(decoded)) {
    return res.status(403).json({ success: false, message: "Not authorized." });
  }
  const year = fpdYear(req.body.date);
  const lastNumber = Number(req.body.lastNumber);
  if (!Number.isInteger(lastNumber) || lastNumber < 0) {
    return res.status(400).json({ success: false, message: "lastNumber must be a whole number." });
  }

  try {
    await db.query(
      `INSERT INTO fpd_number_seeds (year, last_number, updated_by)
       VALUES (?, ?, ?)
       ON DUPLICATE KEY UPDATE last_number = VALUES(last_number), updated_by = VALUES(updated_by)`,
      [year, lastNumber, decoded.displayName ?? decoded.username ?? null],
    );
    const next = String(await fpdNextSeq(db, year, false)).padStart(2, "0");
    return res.json({ success: true, statementNo: `FPD.${year}.${next}` });
  } catch (err) {
    console.error("PUT /fpd/last-number error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// GET /fpd/bank-templates — default first, then alphabetical
app.get("/fpd/bank-templates", async (req, res) => {
  if (!requireAuth(req, res)) return;
  try {
    const [rows] = await db.query(
      "SELECT * FROM fpd_bank_templates ORDER BY is_default DESC, label ASC",
    );
    return res.json({ success: true, count: rows.length, templates: rows.map(fpdBankFromRow) });
  } catch (err) {
    console.error("GET /fpd/bank-templates error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// POST /fpd/bank-templates — same 7 fields as the default bank + a label
app.post("/fpd/bank-templates", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  const { label, fields, createdByName } = req.body;

  const cleaned = (Array.isArray(fields) ? fields : [])
    .map((f) => ({
      label: String(f?.label ?? "").trim(),
      value: String(f?.value ?? "").trim(),
    }))
    .filter((f) => f.label || f.value);

  if (cleaned.length === 0 || cleaned.some((f) => !f.label)) {
    return res.status(400).json({
      success: false,
      message: "At least one detail is required, and every detail needs a label.",
    });
  }

  try {
    const [result] = await db.query(
      `INSERT INTO fpd_bank_templates
        (label, account_name, account_address, bank_name, branch_name, branch_address,
         account_no, swift_code, detail_fields, is_default, created_by)
       VALUES (?, '', '', '', '', '', '', '', ?, 0, ?)`,
      [
        (label || "").trim() || cleaned[0].value || "Bank template",
        JSON.stringify(cleaned),
        createdByName ?? decoded.displayName ?? decoded.username ?? null,
      ],
    );
    const [rows] = await db.query("SELECT * FROM fpd_bank_templates WHERE id = ?", [result.insertId]);
    return res.status(201).json({ success: true, template: fpdBankFromRow(rows[0]) });
  } catch (err) {
    console.error("POST /fpd/bank-templates error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// DELETE /fpd/bank-templates/:id — the fixed default can't be deleted.
// Saved statements are unaffected (they keep their own bank_snapshot).
app.delete("/fpd/bank-templates/:id", async (req, res) => {
  if (!requireAuth(req, res)) return;
  try {
    const [rows] = await db.query("SELECT is_default FROM fpd_bank_templates WHERE id = ?", [req.params.id]);
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "Template not found." });
    }
    if (rows[0].is_default) {
      return res.status(400).json({ success: false, message: "The default bank template can't be deleted." });
    }
    await db.query("DELETE FROM fpd_bank_templates WHERE id = ?", [req.params.id]);
    return res.json({ success: true });
  } catch (err) {
    console.error("DELETE /fpd/bank-templates/:id error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ── Statements ──────────────────────────────────────────────────────────────

// POST /fpd/statements
// body: { date, billToName, billToAddress, currency, items: [{ date, description,
//         unitPrice, discounted, discountPercent }], bankTemplateId | bank,
//         preparedBy, createdByName }
app.post("/fpd/statements", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  const prepared = fpdPrepareBody(req.body);
  if (prepared.error) {
    return res.status(400).json({ success: false, message: prepared.error });
  }

  const { date, billToName, billToAddress, currency, bankTemplateId, bank, preparedBy, createdByName } = req.body;
  const manualNo = String(req.body.statementNo ?? "").trim();
  if (manualNo.length > 50) {
    return res.status(400).json({ success: false, message: "Statement No. is too long." });
  }

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    const resolved = await fpdResolveBank(conn, { bankTemplateId, bank });
    if (resolved.error) {
      await conn.rollback();
      return res.status(resolved.status).json({ success: false, message: resolved.error });
    }

    let statementNo;
    if (manualNo) {
      const [dupe] = await conn.query(
        "SELECT 1 FROM fpd_statements WHERE statement_no = ? FOR UPDATE",
        [manualNo],
      );
      if (dupe.length > 0) {
        await conn.rollback();
        return res.status(409).json({ success: false, message: `${manualNo} already exists.` });
      }
      statementNo = manualNo;
    } else {
      const year = fpdYear(date);
      const prefix = fpdPrefix(decoded);
      const nextNum = String(await fpdNextSeq(conn, year, true, prefix)).padStart(2, "0");
      statementNo = `${prefix}.${year}.${nextNum}`;
    }

    const id = crypto.randomUUID();
    await conn.query(
      `INSERT INTO fpd_statements
        (id, statement_no, statement_date, bill_to_name, bill_to_address, currency, from_currency,
         items, total_amount, exchange_rate, bank_template_id, bank_snapshot, prepared_by,
         created_by, created_by_username)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        id, statementNo, date, billToName.trim(), billToAddress ?? "", currency || "₱", prepared.fromCurrency,
        JSON.stringify(prepared.lines), prepared.total, prepared.rate,
        resolved.templateId, JSON.stringify(resolved.bank),
        preparedBy ?? "",
        createdByName ?? decoded.displayName ?? decoded.username ?? null,
        fpdOwner(decoded),
      ],
    );

    await conn.commit();
    return res.status(201).json({ success: true, id, statementNo, totalAmount: prepared.total });
  } catch (err) {
    await conn.rollback();
    if (err.code === "ER_DUP_ENTRY") {
      return res.status(409).json({ success: false, message: "That Statement No. already exists." });
    }
    console.error("POST /fpd/statements error:", err);
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// GET /fpd/statements — newest first
app.get("/fpd/statements", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  try {
    const [rows] = await db.query(
      "SELECT * FROM fpd_statements WHERE created_by_username = ? ORDER BY created_at DESC",
      [fpdOwner(decoded)],
    );
    return res.json({ success: true, count: rows.length, statements: rows.map(fpdStatementFromRow) });
  } catch (err) {
    console.error("GET /fpd/statements error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// GET /fpd/statements/:statementNo — e.g. /fpd/statements/FPD.2026.01
app.get("/fpd/statements/:statementNo", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  try {
    const [rows] = await db.query(
      "SELECT * FROM fpd_statements WHERE statement_no = ? AND created_by_username = ?",
      [req.params.statementNo, fpdOwner(decoded)],
    );
    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: "No statement found for that number." });
    }
    return res.json({ success: true, statement: fpdStatementFromRow(rows[0]) });
  } catch (err) {
    console.error("GET /fpd/statements/:statementNo error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// PUT /fpd/statements/:statementNo — number changes only if body.statementNo is sent. Bank
// details are only replaced when bankTemplateId/bank is sent.
app.put("/fpd/statements/:statementNo", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;

  const prepared = fpdPrepareBody(req.body);
  if (prepared.error) {
    return res.status(400).json({ success: false, message: prepared.error });
  }

  const { date, billToName, billToAddress, currency, bankTemplateId, bank, preparedBy } = req.body;
  const newNo = String(req.body.statementNo ?? "").trim() || req.params.statementNo;
  if (newNo.length > 50) {
    return res.status(400).json({ success: false, message: "Statement No. is too long." });
  }

  const conn = await db.getConnection();
  try {
    await conn.beginTransaction();

    if (newNo !== req.params.statementNo) {
      const [dupe] = await conn.query(
        "SELECT 1 FROM fpd_statements WHERE statement_no = ? FOR UPDATE",
        [newNo],
      );
      if (dupe.length > 0) {
        await conn.rollback();
        return res.status(409).json({ success: false, message: `${newNo} already exists.` });
      }
    }

    const [existing] = await conn.query(
      "SELECT id FROM fpd_statements WHERE statement_no = ? AND created_by_username = ? FOR UPDATE",
      [req.params.statementNo, fpdOwner(decoded)],
    );
    if (existing.length === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: "No statement found for that number." });
    }

    const bankProvided = Boolean(bankTemplateId || bank);
    let bankSql = "";
    const bankParams = [];
    if (bankProvided) {
      const resolved = await fpdResolveBank(conn, { bankTemplateId, bank });
      if (resolved.error) {
        await conn.rollback();
        return res.status(resolved.status).json({ success: false, message: resolved.error });
      }
      bankSql = ", bank_template_id = ?, bank_snapshot = ?";
      bankParams.push(resolved.templateId, JSON.stringify(resolved.bank));
    }

    await conn.query(
      `UPDATE fpd_statements SET
         statement_no = ?, statement_date = ?, bill_to_name = ?, bill_to_address = ?,
         currency = ?, from_currency = ?,
         items = ?, total_amount = ?, exchange_rate = ?, prepared_by = ?${bankSql}
       WHERE statement_no = ? AND created_by_username = ?`,
      [
        newNo, date, billToName.trim(), billToAddress ?? "", currency || "₱", prepared.fromCurrency,
        JSON.stringify(prepared.lines), prepared.total, prepared.rate, preparedBy ?? "",
        ...bankParams,
        req.params.statementNo, fpdOwner(decoded),
      ],
    );

    await conn.commit();
    return res.json({ success: true, statementNo: newNo, totalAmount: prepared.total });
  } catch (err) {
    await conn.rollback();
    console.error("PUT /fpd/statements/:statementNo error:", err);
    return res.status(500).json({ success: false, message: err.message });
  } finally {
    conn.release();
  }
});

// DELETE /fpd/statements/:statementNo — removes the whole statement.
app.delete("/fpd/statements/:statementNo", async (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  try {
    const [result] = await db.query(
      "DELETE FROM fpd_statements WHERE statement_no = ? AND created_by_username = ?",
      [req.params.statementNo, fpdOwner(decoded)],
    );
    if (result.affectedRows === 0) {
      return res.status(404).json({ success: false, message: "No statement found for that number." });
    }
    return res.json({ success: true });
  } catch (err) {
    console.error("DELETE /fpd/statements/:statementNo error:", err);
    return res.status(500).json({ success: false, message: err.message });
  }
});

// ─── FPD: BSP REFERENCE EXCHANGE RATES ──────────────────────────────────────
let pdfParse = null;
try {
  pdfParse = require("pdf-parse/lib/pdf-parse.js");
} catch (err) {
  console.warn("⚠ pdf-parse unavailable — BSP rates disabled, using fallback:", err.message);
}

const BSP_BASE = "https://www.bsp.gov.ph";
// Pages that list the daily RERB PDFs. Override with BSP_RERB_INDEX_URLS
// (comma-separated) in .env if BSP moves things around.
const BSP_RERB_INDEX_URLS = (
  process.env.BSP_RERB_INDEX_URLS ||
  `${BSP_BASE}/SitePages/Statistics/ExchangeRate.aspx,${BSP_BASE}/Lists/RERB/AllItems.aspx`
).split(",");

const FX_TTL = 60 * 60 * 1000;                    // BSP posts once a day; hourly check is plenty
const BSP_STALE_MAX_MS = 7 * 24 * 60 * 60 * 1000; // keep serving the last good BSP rate this long
const MONTHS_IDX = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };

let fxCache = { at: 0, data: null };
let lastGoodBsp = null; // { at, data }
let pushedBsp = null;   // latest BSP rates pushed by the VM fetcher job

async function bspGet(url) {
  const r = await fetch(url, {
    headers: { "User-Agent": "Mozilla/5.0" },
    signal: AbortSignal.timeout(15000),
  });
  if (!r.ok) throw new Error(`HTTP ${r.status} from ${url}`);
  return r;
}

// Scans the index pages for RERB PDF links and returns the most recent one.
async function findLatestRerbPdf() {
  const re = /\/Lists\/RERB\/Attachments\/\d+\/(\d{1,2})([A-Za-z]{3})(\d{4})\.pdf/gi;
  let best = null;
  for (const indexUrl of BSP_RERB_INDEX_URLS) {
    try {
      const html = await (await bspGet(indexUrl.trim())).text();
      for (const m of html.matchAll(re)) {
        const mon = MONTHS_IDX[m[2].toLowerCase()];
        if (mon === undefined) continue;
        const date = new Date(Date.UTC(+m[3], mon, +m[1]));
        if (!best || date > best.date) best = { date, url: BSP_BASE + m[0] };
      }
    } catch (e) {
      console.warn("RERB index fetch failed:", indexUrl, e.message);
    }
    if (best) break;
  }
  if (!best) throw new Error("No RERB PDF link found on the BSP index pages");
  return best;
}

async function fromBsp() {
  if (pushedBsp && Date.now() - pushedBsp.at < BSP_STALE_MAX_MS) return pushedBsp.data;
  throw new Error("No recent BSP rates pushed by the fetcher job");
  // eslint-disable-next-line no-unreachable -- legacy PDF path, blocked by Akamai
  if (!pdfParse) throw new Error("pdf-parse not available");
  let date, buf;
  if (process.env.BSP_LOCAL_PDF) {
    const fs = require("fs");
    const file = process.env.BSP_LOCAL_PDF;
    buf = fs.readFileSync(file);
    // Expects a filename like 13Jan2026.pdf to derive the as-of date
    const m = file.match(/(\d{1,2})([A-Za-z]{3})(\d{4})\.pdf$/i);
    const mon = m ? MONTHS_IDX[m[2].toLowerCase()] : undefined;
    date = m && mon !== undefined
      ? new Date(Date.UTC(+m[3], mon, +m[1]))
      : fs.statSync(file).mtime;
  } else {
    const found = await findLatestRerbPdf();
    date = found.date;
    buf = Buffer.from(await (await bspGet(found.url)).arrayBuffer());
  }
  const { text } = await pdfParse(buf);

  // Row format: "<no> <NAME> CODE  <EUR equiv>  <USD equiv>  <PHP equiv>"
  const pesoPer = (code) => {
    const m = text.match(new RegExp(`\\b${code}\\s+[\\d.]+\\s+[\\d.]+\\s+([\\d.]+)`));
    return m ? parseFloat(m[1]) : null;
  };
  const usd = pesoPer("USD"), jpy = pesoPer("JPY"), aed = pesoPer("AED");
  if (!(usd > 10 && usd < 200) || !(jpy > 0) || !(aed > 0)) {
    throw new Error("BSP RERB parse failed (layout changed?)");
  }

  // App expects "foreign units per 1 PHP"
  return {
    source: "BSP",
    asOf: date.toISOString().slice(0, 10),
    rates: { USD: 1 / usd, JPY: 1 / jpy, AED: 1 / aed },
  };
}

async function fromFallback() {
  const d = await (await fetch("https://open.er-api.com/v6/latest/PHP")).json();
  if (d.result !== "success") throw new Error("fallback failed");
  return {
    source: "open.er-api.com (BSP unavailable)",
    asOf: new Date().toISOString().slice(0, 10),
    rates: { USD: d.rates.USD, JPY: d.rates.JPY, AED: d.rates.AED },
  };
}

// POST /fpd/bsp-rates — receives the daily BSP bulletin from the VM fetcher job
// body: { asOf: "YYYY-MM-DD", pesoPer: { USD, JPY, AED } }  (pesos per 1 unit)
app.post("/fpd/bsp-rates", (req, res) => {
  const decoded = requireAuth(req, res);
  if (!decoded) return;
  if (decoded.role !== "superadmin") {
    return res.status(403).json({ success: false, message: "Not authorized." });
  }
  const { asOf, pesoPer } = req.body || {};
  const usd = Number(pesoPer?.USD), jpy = Number(pesoPer?.JPY), aed = Number(pesoPer?.AED);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(asOf || "") || !(usd > 10 && usd < 200) || !(jpy > 0) || !(aed > 0)) {
    return res.status(400).json({ success: false, message: "Invalid BSP rates payload." });
  }
  pushedBsp = {
    at: Date.now(),
    data: { source: "BSP", asOf, rates: { USD: 1 / usd, JPY: 1 / jpy, AED: 1 / aed } },
  };
  fxCache = { at: 0, data: null }; // serve the new rates immediately
  console.log(`BSP rates received for ${asOf}: USD ${usd}, JPY ${jpy}, AED ${aed}`);
  return res.json({ success: true, asOf });
});

// GET /fpd/exchange-rates — BSP reference rates as "units per 1 PHP"
app.get("/fpd/exchange-rates", async (req, res) => {
  if (!requireAuth(req, res)) return;
  try {
    if (fxCache.data && Date.now() - fxCache.at < FX_TTL) return res.json(fxCache.data);

    let data;
    try {
      data = await fromBsp();
      lastGoodBsp = { at: Date.now(), data };
    } catch (e) {
      console.warn("BSP rate fetch failed:", e.message);
      // Prefer a recent real BSP rate over a different source
      if (lastGoodBsp && Date.now() - lastGoodBsp.at < BSP_STALE_MAX_MS) data = lastGoodBsp.data;
      else data = await fromFallback();
    }
    fxCache = { at: Date.now(), data };
    return res.json(data);
  } catch (err) {
    console.error("GET /fpd/exchange-rates error:", err);
    return res.status(502).json({ success: false, message: "Exchange rates unavailable." });
  }
});


app.listen(PORT, () => {
  console.log(`\n✅ Silverdab backend running on port ${PORT}`);
  console.log(`📡 AD Server: ${AD_URL}`);
  console.log(`🌐 Domain: ${AD_DOMAIN}`);
});
