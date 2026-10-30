/**
 * One-shot read of the signed-in user's calendars and To Do lists for a window of days,
 * for schedulers that want plain JSON without an MCP client (`--sweep`) and for MCP
 * clients that want the same picture in one call (the `sweep` utility tool).
 *
 * Requests: one GET round for the calendars and the task lists (as one $batch), then one
 * $batch round (20 sub-requests each) with a calendarView per calendar and the tasks of
 * every list. Pages beyond the first are followed with plain GETs. A throttled
 * sub-response (429/503/504) is retried on its own, through the client's retry logic.
 */
import { cutText } from './lib/mail-text.js';

export const SWEEP_BODY_MAX_CHARS = 300;
export const SWEEP_MAX_DAYS = 366;
const BATCH_LIMIT = 20;
const MAX_PAGES = 100;
const RETRY_STATUSES = new Set([429, 503, 504]);
const EVENT_SELECT =
  'id,subject,start,end,isAllDay,showAs,responseStatus,location,organizer,isCancelled,seriesMasterId,webLink,lastModifiedDateTime';

export interface SweepGraph {
  /** Request a version-relative Graph path; resolves the parsed JSON body, throws on HTTP errors. */
  request(
    path: string,
    init?: { method?: string; headers?: Record<string, string>; body?: string }
  ): Promise<unknown>;
}

export interface SweepOptions {
  days: number;
  /** 'all' (default) or calendar ids */
  calendars?: 'all' | string[];
  tasks?: boolean;
  /** IANA zone for the window's midnight and for the returned times (Prefer: outlook.timezone) */
  timezone?: string;
  now?: Date;
}

export type SweepErrorCode = 'not_signed_in' | 'graph_error' | 'invalid_arguments';

export class SweepError extends Error {
  constructor(
    readonly code: SweepErrorCode,
    message: string
  ) {
    super(message);
    this.name = 'SweepError';
  }
}

type GraphObject = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any

export interface SweepResult {
  generated: string;
  account: string | null;
  calendars: GraphObject[];
  events: GraphObject[];
  lists: GraphObject[];
  tasks: GraphObject[];
}

// ---------------------------------------------------------------- arguments and window

export function systemTimeZone(): string {
  return Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
}

export function assertTimeZone(timeZone: string): void {
  try {
    new Intl.DateTimeFormat('en-US', { timeZone });
  } catch {
    throw new SweepError('invalid_arguments', `Unknown time zone: ${timeZone}`);
  }
}

export function parseDays(value: unknown): number {
  const days = typeof value === 'number' ? value : Number(String(value ?? '').trim());
  if (!Number.isInteger(days) || days < 1 || days > SWEEP_MAX_DAYS) {
    throw new SweepError(
      'invalid_arguments',
      `--days must be a whole number from 1 to ${SWEEP_MAX_DAYS}`
    );
  }
  return days;
}

export function parseCalendars(value: unknown): 'all' | string[] {
  if (value === undefined || value === null || value === true) return 'all';
  const ids = (Array.isArray(value) ? value.map(String) : String(value).split(','))
    .map((id) => id.trim())
    .filter((id) => id.length > 0);
  if (ids.length === 0 || (ids.length === 1 && ids[0].toLowerCase() === 'all')) return 'all';
  return [...new Set(ids)];
}

function zoneParts(timeZone: string, utcMs: number): number[] {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone,
    hourCycle: 'h23',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  }).formatToParts(new Date(utcMs));
  const get = (type: string) => Number(parts.find((p) => p.type === type)?.value);
  return [get('year'), get('month'), get('day'), get('hour'), get('minute'), get('second')];
}

/** UTC instant of local midnight on the given calendar day in timeZone (DST-safe). */
function localMidnight(timeZone: string, year: number, month: number, day: number): number {
  const wall = Date.UTC(year, month - 1, day);
  let guess = wall;
  for (let i = 0; i < 3; i++) {
    const [y, mo, d, h, mi, s] = zoneParts(timeZone, guess);
    const offset = Date.UTC(y, mo - 1, d, h, mi, s) - guess;
    guess = wall - offset;
  }
  return guess;
}

/** Today 00:00 in timeZone through the same midnight `days` later, as UTC RFC 3339. */
export function sweepWindow(
  days: number,
  timeZone: string,
  now: Date = new Date()
): { start: string; end: string } {
  const [y, m, d] = zoneParts(timeZone, now.getTime());
  const endDay = new Date(Date.UTC(y, m - 1, d + days));
  return {
    start: new Date(localMidnight(timeZone, y, m, d)).toISOString(),
    end: new Date(
      localMidnight(
        timeZone,
        endDay.getUTCFullYear(),
        endDay.getUTCMonth() + 1,
        endDay.getUTCDate()
      )
    ).toISOString(),
  };
}

// ---------------------------------------------------------------- shaping

const ENTITIES: Record<string, string> = {
  nbsp: ' ',
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

/** Body text as plain text, whitespace collapsed, at most `max` characters. */
export function plainText(
  content: unknown,
  contentType: unknown,
  max = SWEEP_BODY_MAX_CHARS
): string {
  let text = typeof content === 'string' ? content : '';
  if (String(contentType).toLowerCase() === 'html' || /<\/?[a-z][^>]*>/i.test(text)) {
    text = text
      .replace(/<(style|script|head)\b[^>]*>[\s\S]*?<\/\1>/gi, ' ')
      .replace(/<br\s*\/?>|<\/(p|div|li|tr|h\d)>/gi, '\n')
      .replace(/<[^>]*>/g, ' ')
      .replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e: string) => {
        if (e[0] === '#') {
          const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : Number(e.slice(1));
          return Number.isFinite(code) && code > 0 && code <= 0x10ffff
            ? String.fromCodePoint(code)
            : m;
        }
        return ENTITIES[e.toLowerCase()] ?? m;
      });
  }
  const collapsed = text
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.replace(/[ \t\u00a0]+/g, ' ').trim())
    .filter((line) => line.length > 0)
    .join('\n');
  return cutText(collapsed, max);
}

function str(value: unknown): string | null {
  return typeof value === 'string' && value.length > 0 ? value : null;
}

function mapAddress(address: unknown): GraphObject | null {
  if (!address || typeof address !== 'object') return null;
  const a = address as GraphObject;
  const out = {
    street: str(a.street),
    city: str(a.city),
    state: str(a.state),
    postalCode: str(a.postalCode),
    countryOrRegion: str(a.countryOrRegion),
  };
  return Object.values(out).some((v) => v !== null) ? out : null;
}

function mapLocation(location: unknown): GraphObject | null {
  if (!location || typeof location !== 'object') return null;
  const l = location as GraphObject;
  const c = l.coordinates as GraphObject | undefined;
  const coordinates =
    c && typeof c.latitude === 'number' && typeof c.longitude === 'number'
      ? { latitude: c.latitude, longitude: c.longitude }
      : null;
  const displayName = str(l.displayName);
  const address = mapAddress(l.address);
  if (!displayName && !address && !coordinates) return null;
  return { displayName, address, coordinates };
}

function mapCalendar(c: GraphObject): GraphObject {
  const owner = c.owner as GraphObject | undefined;
  return {
    id: c.id,
    name: c.name ?? null,
    isDefault: c.isDefaultCalendar === true,
    ...(owner && (owner.name || owner.address)
      ? { owner: { name: owner.name ?? null, address: owner.address ?? null } }
      : {}),
    canEdit: c.canEdit === true,
  };
}

function mapEvent(e: GraphObject, calendar: GraphObject): GraphObject {
  const organizer = (e.organizer as GraphObject | undefined)?.emailAddress as
    | GraphObject
    | undefined;
  return {
    id: e.id,
    calendarId: calendar.id,
    calendarName: calendar.name ?? null,
    subject: e.subject ?? null,
    start: e.start ?? null,
    end: e.end ?? null,
    isAllDay: e.isAllDay === true,
    showAs: e.showAs ?? null,
    responseStatus: (e.responseStatus as GraphObject | undefined)?.response ?? null,
    location: mapLocation(e.location),
    organizer: organizer
      ? { name: organizer.name ?? null, address: organizer.address ?? null }
      : null,
    isCancelled: e.isCancelled === true,
    seriesMasterId: e.seriesMasterId ?? null,
    webLink: e.webLink ?? null,
    lastModified: e.lastModifiedDateTime ?? null,
  };
}

function mapList(l: GraphObject): GraphObject {
  return {
    id: l.id,
    displayName: l.displayName ?? null,
    wellknownListName: l.wellknownListName ?? null,
  };
}

function mapTask(t: GraphObject, list: GraphObject): GraphObject {
  const body = t.body as GraphObject | undefined;
  const text = body ? plainText(body.content, body.contentType) : '';
  return {
    id: t.id,
    listId: list.id,
    listName: list.displayName ?? null,
    title: t.title ?? null,
    status: t.status ?? null,
    importance: t.importance ?? null,
    dueDateTime: t.dueDateTime ?? null,
    createdDateTime: t.createdDateTime ?? null,
    lastModifiedDateTime: t.lastModifiedDateTime ?? null,
    completedDateTime: t.completedDateTime ?? null,
    ...(text ? { body: { content: text } } : {}),
  };
}

// ---------------------------------------------------------------- Graph access

function encodeId(id: string): string {
  return encodeURIComponent(id).replace(/%3D/g, '=');
}

function relativeLink(link: string): string {
  const url = new URL(link);
  return url.pathname.replace(/^\/(v1\.0|beta)/, '') + url.search;
}

interface GetRequest {
  id: string;
  url: string;
  headers?: Record<string, string>;
}

function graphErrorMessage(status: number, body: unknown): string {
  const error = (body as GraphObject | undefined)?.error as GraphObject | undefined;
  const detail = error ? [error.code, error.message].filter(Boolean).join(': ') : '';
  return `Microsoft Graph API error: ${status}${detail ? ` - ${detail}` : ''}`;
}

/** GET every request through $batch, 20 at a time. Resolves first-page bodies by id. */
async function batchGet(graph: SweepGraph, requests: GetRequest[]): Promise<Map<string, unknown>> {
  const bodies = new Map<string, unknown>();
  for (let i = 0; i < requests.length; i += BATCH_LIMIT) {
    const chunk = requests.slice(i, i + BATCH_LIMIT);
    const payload = {
      requests: chunk.map((r) => ({
        id: r.id,
        method: 'GET',
        url: r.url,
        ...(r.headers && Object.keys(r.headers).length > 0 ? { headers: r.headers } : {}),
      })),
    };
    const result = (await graph.request('/$batch', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    })) as { responses?: Array<{ id: string; status: number; body?: unknown }> } | null;
    const responses = new Map((result?.responses ?? []).map((r) => [String(r.id), r]));
    for (const r of chunk) {
      const response = responses.get(r.id);
      if (response && response.status >= 200 && response.status < 300) {
        bodies.set(r.id, response.body ?? {});
      } else if (!response || RETRY_STATUSES.has(response.status)) {
        // Throttled or missing: one direct GET, which carries the client's retry logic
        bodies.set(r.id, await graph.request(r.url, { headers: r.headers }));
      } else {
        throw new SweepError('graph_error', graphErrorMessage(response.status, response.body));
      }
    }
  }
  return bodies;
}

/** The `value` items of a first page plus every page after it. */
async function allItems(
  graph: SweepGraph,
  first: unknown,
  headers?: Record<string, string>
): Promise<GraphObject[]> {
  const items: GraphObject[] = [];
  let page = first as GraphObject | null;
  for (let n = 0; page && n < MAX_PAGES; n++) {
    if (Array.isArray(page.value)) items.push(...page.value);
    const next = page['@odata.nextLink'];
    if (typeof next !== 'string') break;
    page = (await graph.request(relativeLink(next), { headers })) as GraphObject | null;
  }
  return items;
}

function byStart(a: GraphObject, b: GraphObject): number {
  const sa = String(a.start?.dateTime ?? '');
  const sb = String(b.start?.dateTime ?? '');
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}

// ---------------------------------------------------------------- the sweep

export async function runSweep(
  graph: SweepGraph,
  options: SweepOptions,
  account: string | null
): Promise<SweepResult> {
  const days = parseDays(options.days);
  const timeZone = options.timezone || systemTimeZone();
  assertTimeZone(timeZone);
  const wanted = options.calendars ?? 'all';
  const withTasks = options.tasks === true;
  const window = sweepWindow(days, timeZone, options.now);
  // Without --timezone, Graph writes times in UTC; with it, in that zone
  const prefer: Record<string, string> = options.timezone
    ? { Prefer: `outlook.timezone="${options.timezone}"` }
    : {};

  // Round 1: what is there
  const first = await batchGet(graph, [
    { id: 'calendars', url: '/me/calendars?$select=id,name,isDefaultCalendar,owner,canEdit' },
    ...(withTasks ? [{ id: 'lists', url: '/me/todo/lists' }] : []),
  ]);
  const allCalendars = await allItems(graph, first.get('calendars'));
  let calendars = allCalendars;
  if (wanted !== 'all') {
    const missing = wanted.filter((id) => !allCalendars.some((c) => c.id === id));
    if (missing.length > 0) {
      throw new SweepError('graph_error', `Calendar not found: ${missing.join(', ')}`);
    }
    calendars = allCalendars.filter((c) => wanted.includes(c.id));
  }
  let lists: GraphObject[] = [];
  if (withTasks) {
    lists = await allItems(graph, first.get('lists'));
    // Some accounts get only "Flagged Emails" from /me/todo/lists; the delta query
    // returns the real set
    if (lists.length > 0 && lists.every((l) => l.wellknownListName === 'flaggedEmails')) {
      const delta = await graph.request('/me/todo/lists/delta()');
      const fromDelta = (await allItems(graph, delta)).filter((l) => !l['@removed']);
      if (fromDelta.length > lists.length) lists = fromDelta;
    }
  }

  // Round 2: events per calendar, tasks per list
  const query = [
    `startDateTime=${encodeURIComponent(window.start)}`,
    `endDateTime=${encodeURIComponent(window.end)}`,
    `$select=${EVENT_SELECT}`,
    '$top=250',
  ].join('&');
  const requests: GetRequest[] = [
    ...calendars.map((c, i) => ({
      id: `c${i}`,
      url: `/me/calendars/${encodeId(String(c.id))}/calendarView?${query}`,
      headers: prefer,
    })),
    ...lists.map((l, i) => ({
      id: `l${i}`,
      url: `/me/todo/lists/${encodeId(String(l.id))}/tasks`,
      headers: prefer,
    })),
  ];
  const second = requests.length > 0 ? await batchGet(graph, requests) : new Map();

  const events: GraphObject[] = [];
  for (const [i, calendar] of calendars.entries()) {
    for (const e of await allItems(graph, second.get(`c${i}`), prefer)) {
      events.push(mapEvent(e, calendar));
    }
  }
  events.sort(byStart);

  const tasks: GraphObject[] = [];
  for (const [i, list] of lists.entries()) {
    for (const t of await allItems(graph, second.get(`l${i}`), prefer)) {
      tasks.push(mapTask(t, list));
    }
  }

  return {
    generated: new Date().toISOString(),
    account,
    calendars: calendars.map(mapCalendar),
    events,
    lists: lists.map(mapList),
    tasks,
  };
}
