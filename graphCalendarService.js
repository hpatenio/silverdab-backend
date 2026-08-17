// graphCalendarService.js
//
// Pushes Fleet trip events to the fleet admin's Outlook calendar via
// Microsoft Graph (application permissions, client-credentials flow).
//
// Required env vars (add to .env):
//   AZURE_TENANT_ID
//   AZURE_CLIENT_ID
//   AZURE_CLIENT_SECRET
//   FLEET_ADMIN_UPN   (mailbox that receives the events, e.g. hpatenio@silverdab.com)

const TENANT_ID = process.env.AZURE_TENANT_ID;
const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const FLEET_ADMIN_UPN = process.env.FLEET_ADMIN_UPN;

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// ─── Token cache (client-credentials grant, ~1hr expiry) ───────────────────

let cachedToken = null; // { value, expiresAt }

async function getAccessToken() {
  if (cachedToken && cachedToken.expiresAt > Date.now() + 60_000) {
    return cachedToken.value;
  }

  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Failed to acquire Graph token: ${res.status} ${text}`);
  }

  const json = await res.json();
  cachedToken = {
    value: json.access_token,
    expiresAt: Date.now() + json.expires_in * 1000,
  };
  return cachedToken.value;
}

// ─── Trip → Graph event payload ─────────────────────────────────────────────
//
// Expects a trip shape matching what your fleet_trips queries already
// return, e.g. the mapped object from GET /fleet/trips:
//   { tripRef, pickupLabel, dropoffLabel, requestorName, purpose,
//     vehiclePlate, driverName, departureDatetime, returnDatetime }

// Graph's start.dateTime/end.dateTime require ISO 8601 with a "T" separator
// ("YYYY-MM-DDTHH:MM:SS"), no trailing "Z"/milliseconds, so it reads as a
// plain local wall-clock string consistent with the "timeZone" field below.
//
// This pool does NOT set dateStrings:true, so mysql2 hands back
// departure_datetime/return_datetime as native JS Date objects (already
// parsed as local time) — not the "YYYY-MM-DD HH:MM:SS" strings some other
// parts of this codebase assume. Handle both shapes defensively: a Date
// object gets formatted from its local getters; a string just gets its
// space swapped for "T" (in case a future caller passes a raw string).
function toGraphDateTime(value) {
  if (!value) return null;

  if (value instanceof Date) {
    const pad = (n) => String(n).padStart(2, "0");
    return (
      `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}` +
      `T${pad(value.getHours())}:${pad(value.getMinutes())}:${pad(value.getSeconds())}`
    );
  }

  if (typeof value === "string") {
    const normalized = value.includes("T") ? value : value.replace(" ", "T");
    // Strip any trailing UTC offset/Z and milliseconds. Graph rejects a
    // dateTime that carries its own offset (e.g. "+08:00") alongside a
    // separate "timeZone" field — the two are contradictory. Booking-time
    // callers send an explicit "+08:00" string (see TripBookingModal);
    // approval-time Date objects never carry one, so this is a no-op there.
    return normalized.replace(/\.\d+/, "").replace(/(Z|[+-]\d{2}:\d{2})$/, "");
  }

  return null;
}

function buildEventBody(trip) {
  const tz = trip.timeZone || "Asia/Manila";
  const start = toGraphDateTime(trip.departureDatetime);

  let end = trip.returnDatetime ? toGraphDateTime(trip.returnDatetime) : null;

  // Graph rejects events where end <= start (this is the most common cause
  // of ErrorPropertyValidationFailure here) — a bad/stale return_datetime on
  // the row (earlier than departure, or identical to it) would otherwise get
  // forwarded to Graph as-is. Fall back to departure + 1hr whenever end is
  // missing OR doesn't actually come after start.
  // Compute the fallback end from the resolved wall-clock `start` string, not
  // from trip.departureDatetime via toISOString(). Booking sends an explicit
  // "+08:00" string; toISOString() re-applied that offset and made end ~7hrs
  // BEFORE start, so pending-time sync always failed while approve-time
  // (Date-object) sync worked. Appending "Z" does the +1hr math in UTC, then
  // we slice the wall-clock back out so start and end stay in the same tz.
  if (start && (!end || end <= start)) {
    end = new Date(new Date(`${start}Z`).getTime() + 60 * 60 * 1000)
      .toISOString()
      .slice(0, 19);
  }

  console.log("Graph event body dates:", {
    tripRef: trip.tripRef,
    rawDeparture: trip.departureDatetime,
    rawReturn: trip.returnDatetime,
    resolvedStart: start,
    resolvedEnd: end,
  });

  return {
    subject: trip.dropoffLabel,
    body: {
      contentType: "HTML",
      content: [
        `<b>Trip #:</b> ${trip.tripRef}`,
        `<b>Requestor:</b> ${trip.requestorName}`,
        trip.departureDatetime ? `<b>Departure:</b> ${toGraphDateTime(trip.departureDatetime)}` : null,
        trip.passengerCount != null ? `<b>Passengers:</b> ${trip.passengerCount}` : null,
        trip.purpose ? `<b>Purpose:</b> ${trip.purpose}` : null,
        trip.vehiclePlate ? `<b>Vehicle:</b> ${trip.vehiclePlate}` : null,
        trip.driverName ? `<b>Driver:</b> ${trip.driverName}` : null,
        trip.approvedByName ? `<b>Approved by:</b> ${trip.approvedByName}` : null,
        trip.approvedAt ? `<b>Approved at:</b> ${toGraphDateTime(trip.approvedAt)}` : null,
      ]
        .filter(Boolean)
        .join("<br/>"),
    },
    start: { dateTime: start, timeZone: tz },
    end: { dateTime: end, timeZone: tz },
    location: { displayName: trip.pickupLabel },
  };
}

// ─── Public API ──────────────────────────────────────────────────────────

/**
 * Creates an Outlook calendar event for a trip on the fleet admin mailbox.
 * Returns the Graph event ID — store this as fleet_trips.outlook_event_id.
 */
async function createTripEvent(trip, targetUpn = FLEET_ADMIN_UPN) {
  const token = await getAccessToken();
  const eventBody = buildEventBody(trip);
  console.log("Graph event payload:", JSON.stringify(eventBody, null, 2)); // TEMP — remove after debugging
  const res = await fetch(`${GRAPH_BASE}/users/${targetUpn}/events`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(eventBody),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph createTripEvent failed: ${res.status} ${text}`);
  }

  const json = await res.json();
  return json.id;
}

/**
 * Updates an existing event — call after reassignment or reschedule.
 */
async function updateTripEvent(eventId, trip) {
  const token = await getAccessToken();
  const res = await fetch(`${GRAPH_BASE}/users/${FLEET_ADMIN_UPN}/events/${eventId}`, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(buildEventBody(trip)),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph updateTripEvent failed: ${res.status} ${text}`);
  }
}

/**
 * Deletes an event — call on rejection or cancellation.
 * Swallows 404s (event already gone) so callers don't need to special-case it.
 */
async function deleteTripEvent(eventId, targetUpn = FLEET_ADMIN_UPN) {
  const token = await getAccessToken();
  const res = await fetch(`${GRAPH_BASE}/users/${targetUpn}/events/${eventId}`, {
    method: "DELETE",
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok && res.status !== 404) {
    const text = await res.text();
    throw new Error(`Graph deleteTripEvent failed: ${res.status} ${text}`);
  }
}

module.exports = { createTripEvent, updateTripEvent, deleteTripEvent };