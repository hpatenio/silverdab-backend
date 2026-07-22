const express = require("express");
const ldap = require("ldapjs");
const jwt = require("jsonwebtoken");
const crypto = require("crypto");
const cors = require("cors");
const nodemailer = require("nodemailer");
const webpush = require("web-push");
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
async function sendWebPushToAdmins({ title, body, url }) {
  if (!process.env.VAPID_PUBLIC_KEY) {
    console.warn("🔔 sendWebPushToAdmins: VAPID not configured, skipping");
    return;
  }
  try {
    const [subs] = await db.query("SELECT * FROM push_subscriptions");
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
                <strong>Silverdab Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:silverdab-logo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
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
      from: `"Silverdab Requests" <${process.env.EMAIL_USER}>`,
      to: toEmail,
      subject: `Supply Request ${ticketNumber} Received`,
      text: `Dear ${requestedByName},\n\nYour supply request ${ticketNumber} has been submitted and is pending review.\n\nItems:\n${itemListText}`,
      html: htmlBody,
      attachments: [
        {
          filename: "silverdab-logo.png",
          path: "./assets/silverdab-logo.png",
          cid: "silverdab-logo",
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
                <strong>Silverdab Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:silverdab-logo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
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
      from: `"Silverdab Requests" <${process.env.EMAIL_USER}>`,
      to: SUPPLY_REQUEST_NOTIFY_EMAIL,
      subject: `New Supply Request ${ticketNumber} — Pending Your Review`,
      text: `${requestedByName} has submitted a new supply request awaiting your approval.\n\nTicket: ${ticketNumber}\n\nItems:\n${itemListText}\n\nPlease log in to Silverdab UMS to review.`,
      html: htmlBody,
      attachments: [
        {
          filename: "silverdab-logo.png",
          path: "./assets/silverdab-logo.png",
          cid: "silverdab-logo",
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
                <strong>Silverdab Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:silverdab-logo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
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
      from: `"Silverdab Requests" <${process.env.EMAIL_USER}>`,
      to: toEmail,
      subject: `Re: Supply Request ${request.ticket_number} Received`,
      text: `Dear ${request.requested_by_name},\n\nYour supply request ${request.ticket_number} status has been updated to: ${statusLabel}.${updatedByName ? ` (by ${updatedByName})` : ""}${extraMessage ? `\n\n${extraMessage}` : ""}`,
      html: htmlBody,
      attachments: [
        {
          filename: "silverdab-logo.png",
          path: "./assets/silverdab-logo.png",
          cid: "silverdab-logo",
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
                <strong>Silverdab Unified Management System</strong>
              </td>
              <td style="vertical-align: middle; text-align: right;">
                <img src="cid:silverdab-logo" alt="Silverdab" style="height: 32px; opacity: 0.85;" />
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
      from: `"Silverdab Requests" <${process.env.EMAIL_USER}>`,
      to: SUPPLY_REQUEST_NOTIFY_EMAIL,
      subject: `Supply Request ${request.ticket_number} — Status Changed to ${statusLabel}`,
      text: `Ticket ${request.ticket_number} (requested by ${request.requested_by_name}) status changed to: ${statusLabel}.${updatedByName ? ` Updated by ${updatedByName}.` : ""}${extraMessage ? `\n\n${extraMessage}` : ""}`,
      html: htmlBody,
      attachments: [
        {
          filename: "silverdab-logo.png",
          path: "./assets/silverdab-logo.png",
          cid: "silverdab-logo",
        },
      ],
    });

    console.log(`📧 Admin status update (${statusLabel}) sent to ${SUPPLY_REQUEST_NOTIFY_EMAIL} for ${request.ticket_number}`);
  } catch (err) {
    console.error("Admin status update email failed:", err.message);
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
            created_at, updated_at)
         VALUES (?, ?, ?, ?, ?, ?, 'employee', 0, 0, 0, 0, 0, NOW(), NOW())
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

  const { itAccess, itInventory, consumables, tickets, officeSupplies } = req.body;
  const username = req.params.username.toLowerCase().trim();

  try {
    await db.query(
      `UPDATE users SET
         perm_it_access = ?, perm_it_inventory = ?, perm_consumables = ?,
         perm_tickets = ?, perm_office_supplies = ?, updated_at = NOW()
       WHERE username = ?`,
      [
        itAccess ? 1 : 0,
        itInventory ? 1 : 0,
        consumables ? 1 : 0,
        tickets ? 1 : 0,
        officeSupplies ? 1 : 0,
        username,
      ]
    );
    return res.json({ success: true });
  } catch (err) {
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
  const { username } = req.params;
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
      { expiresIn: "8h" }
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
      [decoded.username, subscription.endpoint, subscription.keys.p256dh, subscription.keys.auth],
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
      [decoded.username, subscription.endpoint, subscription.keys.p256dh, subscription.keys.auth],
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

// PATCH /it-inventory/:assetTag — partial update, only sends fields present in body
app.patch("/it-inventory/:assetTag", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader?.startsWith("Bearer "))
    return res.status(401).json({ success: false, message: "No token provided." });
  try { jwt.verify(authHeader.split(" ")[1], JWT_SECRET); }
  catch { return res.status(401).json({ success: false, message: "Invalid token." }); }

  const { assetTag } = req.params;

  const FIELD_MAP = {
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

  const setClause = entries.map(([key]) => `${FIELD_MAP[key]} = ?`).join(", ");
  const values = entries.map(([key, value]) =>
    key === "datePurchased" && value === "" ? null : value,
  );

  try {
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




app.listen(PORT, () => {
  console.log(`\n✅ Silverdab backend running on port ${PORT}`);
  console.log(`📡 AD Server: ${AD_URL}`);
  console.log(`🌐 Domain: ${AD_DOMAIN}`);
});