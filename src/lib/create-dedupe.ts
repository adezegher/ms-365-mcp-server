/**
 * Duplicate check for create tools that declare `dedupe` in endpoints.json.
 *
 * Before the POST, the target collection is read around the new item's start (an event)
 * or due time (a task). An existing item with the same normalized subject/title whose
 * start/due is within a minute of the new one is a duplicate, and the create is refused
 * unless the caller passes `allowDuplicate: true`.
 *
 * - `event`: the POST goes to an `.../events` collection; the check reads the sibling
 *   `.../calendarView` for [start - 1 min, start + 1 min]. Cancelled events never match.
 * - `task`: the POST goes to a To Do `.../tasks` collection; the check reads that list.
 *   With a due time, any task with the same title and due time matches (a completed one
 *   too: re-creating a done task brings it back). Without one, only an open task with the
 *   same title and no due time matches.
 */

export type DedupeKind = 'event' | 'task';

export interface DateTimeTimeZone {
  dateTime: string;
  timeZone?: string;
}

export interface DuplicateResult {
  duplicate: true;
  existing: Record<string, unknown>;
}

/** GET a Graph path (version-relative, may carry a query string) and return the parsed body. */
export type DedupeFetch = (path: string, headers: Record<string, string>) => Promise<unknown>;

const MINUTE_MS = 60_000;
// Widest UTC offsets are -12:00 and +14:00, so a window this wide around a wall-clock time
// read as UTC always covers the real instant, whatever zone it was written in.
const UNKNOWN_ZONE_SLACK_MS = 15 * 60 * MINUTE_MS;
const MAX_LOOKUP_PAGES = 50;

export function isDedupeKind(value: unknown): value is DedupeKind {
  return value === 'event' || value === 'task';
}

/** Case-, width- and whitespace-insensitive comparison key for a subject or title. */
export function normalizeTitle(value: unknown): string {
  if (typeof value !== 'string') return '';
  return value.normalize('NFKC').toLowerCase().replace(/\s+/g, ' ').trim();
}

function isDateTimeTimeZone(value: unknown): value is DateTimeTimeZone {
  return (
    typeof value === 'object' &&
    value !== null &&
    typeof (value as DateTimeTimeZone).dateTime === 'string' &&
    (value as DateTimeTimeZone).dateTime.length > 0
  );
}

/** Milliseconds of a wall-clock dateTime read as if it were UTC (offset or Z honoured). */
function naiveMs(dateTime: string): number | null {
  const hasZone = /(Z|[+-]\d{2}:?\d{2})$/i.test(dateTime);
  const ms = Date.parse(hasZone ? dateTime : `${dateTime.replace(/\.\d+$/, '')}Z`);
  return Number.isNaN(ms) ? null : ms;
}

function zoneOffsetMs(timeZone: string, utcMs: number): number {
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
  const wall = Date.UTC(
    get('year'),
    get('month') - 1,
    get('day'),
    get('hour'),
    get('minute'),
    get('second')
  );
  return wall - Math.floor(utcMs / 1000) * 1000;
}

function isUtcName(timeZone: string): boolean {
  return /^(utc|gmt|z|etc\/utc|etc\/gmt|coordinated universal time)$/i.test(timeZone.trim());
}

/**
 * The UTC instant of a Graph dateTimeTimeZone, or null when the zone is not one this
 * runtime knows (Windows zone names such as "Pacific Standard Time" are not IANA names).
 */
export function instantOf(value: DateTimeTimeZone): number | null {
  const wall = naiveMs(value.dateTime);
  if (wall === null) return null;
  if (/(Z|[+-]\d{2}:?\d{2})$/i.test(value.dateTime)) return wall;
  const timeZone = value.timeZone?.trim() || 'UTC';
  if (isUtcName(timeZone)) return wall;
  try {
    // Two passes settle the offset across a DST change.
    let guess = wall - zoneOffsetMs(timeZone, wall);
    guess = wall - zoneOffsetMs(timeZone, guess);
    return guess;
  } catch {
    return null;
  }
}

/** True when two dateTimeTimeZone values are within a minute of each other. */
export function sameInstant(a: DateTimeTimeZone, b: DateTimeTimeZone): boolean {
  const ia = instantOf(a);
  const ib = instantOf(b);
  if (ia !== null && ib !== null) return Math.abs(ia - ib) <= MINUTE_MS;
  // Unknown zone names: comparable only when both sides are written in the same zone
  const za = (a.timeZone || 'UTC').trim().toLowerCase();
  const zb = (b.timeZone || 'UTC').trim().toLowerCase();
  if (za !== zb) return false;
  const na = naiveMs(a.dateTime);
  const nb = naiveMs(b.dateTime);
  return na !== null && nb !== null && Math.abs(na - nb) <= MINUTE_MS;
}

function parseBody(body: unknown): Record<string, unknown> | null {
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

function relativeNextLink(link: string): string {
  const url = new URL(link);
  return url.pathname.replace(/^\/(v1\.0|beta)/, '') + url.search;
}

async function readAll(
  fetchPage: DedupeFetch,
  firstPath: string,
  headers: Record<string, string>
): Promise<Record<string, unknown>[]> {
  const items: Record<string, unknown>[] = [];
  let path: string | undefined = firstPath;
  for (let page = 0; path && page < MAX_LOOKUP_PAGES; page++) {
    const data = (await fetchPage(path, headers)) as {
      value?: Record<string, unknown>[];
      '@odata.nextLink'?: string;
    } | null;
    items.push(...(Array.isArray(data?.value) ? data.value : []));
    const next = data?.['@odata.nextLink'];
    path = typeof next === 'string' ? relativeNextLink(next) : undefined;
  }
  return items;
}

function preferZone(timeZone: string): Record<string, string> {
  return { Prefer: `outlook.timezone="${timeZone.replace(/"/g, '')}"` };
}

async function findDuplicateEvent(
  collectionPath: string,
  body: Record<string, unknown>,
  fetchPage: DedupeFetch
): Promise<DuplicateResult | null> {
  const subject = normalizeTitle(body.subject);
  if (!subject || !isDateTimeTimeZone(body.start)) return null;
  const start = body.start;
  if (!/\/events$/.test(collectionPath)) return null;
  const viewPath = collectionPath.replace(/\/events$/, '/calendarView');

  const instant = instantOf(start);
  let from: number;
  let to: number;
  let headers: Record<string, string>;
  if (instant !== null) {
    from = instant - MINUTE_MS;
    to = instant + MINUTE_MS;
    headers = preferZone('UTC');
  } else {
    // A zone this runtime cannot resolve: read a wide window and let Graph write the
    // existing starts in the same zone, so the wall-clock times compare directly.
    const wall = naiveMs(start.dateTime);
    if (wall === null) return null;
    from = wall - UNKNOWN_ZONE_SLACK_MS;
    to = wall + UNKNOWN_ZONE_SLACK_MS;
    headers = preferZone(start.timeZone || 'UTC');
  }
  const query = [
    `startDateTime=${encodeURIComponent(new Date(from).toISOString())}`,
    `endDateTime=${encodeURIComponent(new Date(to).toISOString())}`,
    '$select=id,subject,start,isCancelled',
    '$top=100',
  ].join('&');

  const events = await readAll(fetchPage, `${viewPath}?${query}`, headers);
  const match = events.find(
    (e) =>
      e.isCancelled !== true &&
      normalizeTitle(e.subject) === subject &&
      isDateTimeTimeZone(e.start) &&
      sameInstant(e.start, start)
  );
  return match
    ? { duplicate: true, existing: { id: match.id, subject: match.subject, start: match.start } }
    : null;
}

async function findDuplicateTask(
  collectionPath: string,
  body: Record<string, unknown>,
  fetchPage: DedupeFetch
): Promise<DuplicateResult | null> {
  const title = normalizeTitle(body.title);
  if (!title) return null;
  const due = isDateTimeTimeZone(body.dueDateTime) ? body.dueDateTime : null;
  const headers = due ? preferZone(due.timeZone || 'UTC') : {};

  const tasks = await readAll(fetchPage, collectionPath, headers);
  const match = tasks.find((t) => {
    if (normalizeTitle(t.title) !== title) return false;
    const existingDue = isDateTimeTimeZone(t.dueDateTime) ? t.dueDateTime : null;
    if (due === null) return existingDue === null && t.status !== 'completed';
    return existingDue !== null && sameInstant(existingDue, due);
  });
  return match
    ? {
        duplicate: true,
        existing: {
          id: match.id,
          title: match.title,
          dueDateTime: match.dueDateTime ?? null,
          status: match.status,
        },
      }
    : null;
}

/**
 * Look for an existing item the create would duplicate. `collectionPath` is the resolved
 * POST path (no query string). Returns null when there is none, or when the body carries
 * nothing to compare (no subject/title, or an event without a start). Graph errors throw.
 */
export async function findDuplicate(
  kind: DedupeKind,
  collectionPath: string,
  body: unknown,
  fetchPage: DedupeFetch
): Promise<DuplicateResult | null> {
  const parsed = parseBody(body);
  if (!parsed) return null;
  const path = collectionPath.split('?')[0];
  return kind === 'event'
    ? findDuplicateEvent(path, parsed, fetchPage)
    : findDuplicateTask(path, parsed, fetchPage);
}

export const ALLOW_DUPLICATE_PARAM_DESCRIPTION =
  'Set true to create the item even though one with the same subject/title and the same start/due time already exists. ' +
  'Without it the server checks first and answers { duplicate: true, existing: {...} } instead of creating a second copy.';
