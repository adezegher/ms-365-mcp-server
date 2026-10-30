/**
 * No-attendees mode (MS365_MCP_NO_ATTENDEES=1 or --no-attendees): event tools that declare
 * `attendeeGuard` in endpoints.json cannot make Exchange send mail. A body with attendees
 * sends meeting requests, and a PATCH on an organized event that has attendees sends every
 * attendee an update, so with the mode on:
 * - a body carrying `attendees` (any casing) or `responseRequested: true` is refused;
 * - an `update` first reads the event and is refused when it has an attendee other than
 *   the organizer.
 * Off by default; nothing changes unless an operator turns it on.
 */

export type AttendeeGuardKind = 'create' | 'update';

export function isAttendeeGuardKind(value: unknown): value is AttendeeGuardKind {
  return value === 'create' || value === 'update';
}

export function isNoAttendeesMode(): boolean {
  const value = process.env.MS365_MCP_NO_ATTENDEES?.trim().toLowerCase();
  return value === '1' || value === 'true';
}

function asObject(body: unknown): Record<string, unknown> | null {
  if (typeof body === 'string') {
    try {
      body = JSON.parse(body);
    } catch {
      return null;
    }
  }
  return typeof body === 'object' && body !== null && !Array.isArray(body)
    ? (body as Record<string, unknown>)
    : null;
}

/** True when the request body would add or change attendees or ask for responses. */
export function bodyCarriesAttendees(body: unknown): boolean {
  if (typeof body === 'string' && asObject(body) === null) {
    // Unparseable text: refuse if it so much as names the fields
    return /attendees|responserequested/i.test(body);
  }
  const object = asObject(body);
  if (!object) return false;
  return Object.entries(object).some(([key, value]) => {
    const lower = key.toLowerCase();
    return lower === 'attendees' || (lower === 'responserequested' && value === true);
  });
}

function addressOf(entry: unknown): string {
  const email = (entry as { emailAddress?: { address?: unknown } } | null)?.emailAddress;
  return typeof email?.address === 'string' ? email.address.trim().toLowerCase() : '';
}

/** True when the event lists an attendee who is not its organizer. */
export function hasOtherAttendees(event: unknown): boolean {
  const object = asObject(event);
  const attendees = object?.attendees;
  if (!Array.isArray(attendees) || attendees.length === 0) return false;
  const organizer = addressOf(object?.organizer);
  return attendees.some((attendee) => {
    const address = addressOf(attendee);
    return address === '' || address !== organizer;
  });
}

export const ATTENDEES_NOT_ALLOWED = {
  error: 'attendees_not_allowed',
  message:
    'Nothing was written: this server runs in no-attendees mode, so events cannot carry attendees or ask for responses (Exchange would mail them).',
};

export const EVENT_HAS_ATTENDEES = {
  error: 'event_has_attendees',
  message:
    'Nothing was written: this event has attendees, and in no-attendees mode an update that Exchange would mail to them is refused.',
};
