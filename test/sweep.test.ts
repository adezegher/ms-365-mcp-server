import { describe, expect, it, vi } from 'vitest';
import type AuthManager from '../src/auth.js';
import type GraphClient from '../src/graph-client.js';
import {
  parseCalendars,
  parseDays,
  plainText,
  runSweep,
  SweepError,
  sweepWindow,
  type SweepGraph,
} from '../src/sweep.js';
import { isSignInError, sweepCommand } from '../src/sweep-command.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

type Init = { method?: string; headers?: Record<string, string>; body?: string };
type SubRequest = { id: string; method: string; url: string; headers?: Record<string, string> };

/**
 * A fake Graph: `routes` answers a GET path (query string included, prefix match on the
 * part before '?'), and $batch answers each sub-request from the same table unless
 * `batchStatus` overrides its status.
 */
function fakeGraph(
  routes: Record<string, unknown>,
  batchStatus: Record<string, { status: number; body?: unknown }> = {}
) {
  const calls: Array<{ path: string; init?: Init }> = [];
  const batches: SubRequest[][] = [];
  const lookup = (path: string): unknown => {
    if (path in routes) return routes[path];
    const base = path.split('?')[0];
    if (base in routes) return routes[base];
    throw new Error(`unexpected path ${path}`);
  };
  const graph: SweepGraph = {
    request: async (path, init) => {
      calls.push({ path, init });
      if (path === '/$batch') {
        const { requests } = JSON.parse(init!.body!) as { requests: SubRequest[] };
        batches.push(requests);
        return {
          responses: requests.map((r) => {
            const base = r.url.split('?')[0];
            const forced = batchStatus[base];
            if (forced) return { id: r.id, status: forced.status, body: forced.body ?? {} };
            return { id: r.id, status: 200, body: lookup(r.url) };
          }),
        };
      }
      return lookup(path);
    },
  };
  return { graph, calls, batches };
}

const CALENDARS = {
  value: [
    { id: 'cal-1', name: 'Calendar', isDefaultCalendar: true, canEdit: true },
    {
      id: 'cal-2',
      name: 'Family',
      isDefaultCalendar: false,
      canEdit: true,
      owner: { name: 'Sam', address: 'sam@example.com' },
    },
  ],
};

const EVENT_A = {
  id: 'ev-a',
  subject: 'Dentist',
  start: { dateTime: '2026-10-29T16:00:00.0000000', timeZone: 'UTC' },
  end: { dateTime: '2026-10-29T17:00:00.0000000', timeZone: 'UTC' },
  isAllDay: false,
  showAs: 'busy',
  responseStatus: { response: 'organizer', time: '0001-01-01T00:00:00Z' },
  location: {
    displayName: 'Smile Dental',
    address: { street: '1 Main St', city: 'Torrance', state: 'CA', countryOrRegion: 'US' },
    coordinates: { latitude: 33.8, longitude: -118.3 },
  },
  organizer: { emailAddress: { name: 'Alec', address: 'alec@example.com' } },
  isCancelled: false,
  seriesMasterId: null,
  webLink: 'https://outlook.live.com/a',
  lastModifiedDateTime: '2026-10-20T10:00:00Z',
};

const EVENT_B = {
  id: 'ev-b',
  subject: 'Soccer',
  start: { dateTime: '2026-10-29T09:00:00.0000000', timeZone: 'UTC' },
  end: { dateTime: '2026-10-29T10:00:00.0000000', timeZone: 'UTC' },
  isAllDay: false,
  showAs: 'busy',
  responseStatus: { response: 'accepted' },
  location: { displayName: '', address: {}, coordinates: {} },
  isCancelled: false,
  seriesMasterId: 'series-1',
  lastModifiedDateTime: '2026-10-21T10:00:00Z',
};

const NOW = new Date('2026-10-29T10:00:00Z');

describe('sweep arguments', () => {
  it('accepts whole days from 1 to 366', () => {
    expect(parseDays('2')).toBe(2);
    expect(parseDays(366)).toBe(366);
    for (const bad of ['0', '-1', '1.5', 'x', '', '367', undefined]) {
      expect(() => parseDays(bad), String(bad)).toThrow(SweepError);
    }
  });

  it('reads --calendars as all or a list of ids', () => {
    expect(parseCalendars(undefined)).toBe('all');
    expect(parseCalendars('all')).toBe('all');
    expect(parseCalendars('ALL')).toBe('all');
    expect(parseCalendars('a, b,,a')).toEqual(['a', 'b']);
    expect(parseCalendars(['x'])).toEqual(['x']);
    expect(parseCalendars([])).toBe('all');
  });
});

describe('sweepWindow', () => {
  it('runs from local midnight today for N days', () => {
    expect(sweepWindow(2, 'Europe/Brussels', NOW)).toEqual({
      start: '2026-10-28T23:00:00.000Z',
      end: '2026-10-30T23:00:00.000Z',
    });
    expect(sweepWindow(1, 'UTC', NOW)).toEqual({
      start: '2026-10-29T00:00:00.000Z',
      end: '2026-10-30T00:00:00.000Z',
    });
  });

  it('ends at local midnight across a DST change', () => {
    // Los Angeles leaves daylight time on 2026-11-01
    expect(sweepWindow(2, 'America/Los_Angeles', new Date('2026-10-31T12:00:00Z'))).toEqual({
      start: '2026-10-31T07:00:00.000Z',
      end: '2026-11-02T08:00:00.000Z',
    });
  });

  it('uses the local date, not the UTC one', () => {
    // 2026-10-30 03:00 UTC is still 29 October in Los Angeles
    expect(sweepWindow(1, 'America/Los_Angeles', new Date('2026-10-30T03:00:00Z')).start).toBe(
      '2026-10-29T07:00:00.000Z'
    );
  });
});

describe('plainText', () => {
  it('turns html into collapsed text', () => {
    expect(
      plainText(
        '<html><head><style>p{}</style></head><body><p>Call&nbsp;Tom &amp; Ann</p><div>re: <b>car</b></div></body></html>',
        'html'
      )
    ).toBe('Call Tom & Ann\nre: car');
  });

  it('keeps text bodies and cuts to 300 characters', () => {
    expect(plainText('  a   b  \r\n\r\n c ', 'text')).toBe('a b\nc');
    expect(plainText('x'.repeat(500), 'text')).toHaveLength(300);
    expect(plainText('', 'text')).toBe('');
    expect(plainText(undefined, 'text')).toBe('');
  });
});

describe('runSweep', () => {
  it('reads calendars and events in two $batch rounds and shapes them', async () => {
    const { graph, batches } = fakeGraph({
      '/me/calendars': CALENDARS,
      '/me/calendars/cal-1/calendarView': { value: [EVENT_A] },
      '/me/calendars/cal-2/calendarView': { value: [EVENT_B] },
    });
    const result = await runSweep(graph, { days: 2, timezone: 'UTC', now: NOW }, 'alec@x.com');

    expect(batches).toHaveLength(2);
    expect(batches[0].map((r) => r.url)).toEqual([
      '/me/calendars?$select=id,name,isDefaultCalendar,owner,canEdit',
    ]);
    const view = batches[1][0];
    expect(view.url).toContain('/me/calendars/cal-1/calendarView?');
    expect(view.url).toContain('startDateTime=2026-10-29T00%3A00%3A00.000Z');
    expect(view.url).toContain('endDateTime=2026-10-31T00%3A00%3A00.000Z');
    expect(view.headers).toEqual({ Prefer: 'outlook.timezone="UTC"' });

    expect(result.account).toBe('alec@x.com');
    expect(result.generated).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/);
    expect(result.calendars).toEqual([
      { id: 'cal-1', name: 'Calendar', isDefault: true, canEdit: true },
      {
        id: 'cal-2',
        name: 'Family',
        isDefault: false,
        owner: { name: 'Sam', address: 'sam@example.com' },
        canEdit: true,
      },
    ]);
    // Sorted by start
    expect(result.events.map((e) => e.id)).toEqual(['ev-b', 'ev-a']);
    expect(result.events[1]).toEqual({
      id: 'ev-a',
      calendarId: 'cal-1',
      calendarName: 'Calendar',
      subject: 'Dentist',
      start: EVENT_A.start,
      end: EVENT_A.end,
      isAllDay: false,
      showAs: 'busy',
      responseStatus: 'organizer',
      location: {
        displayName: 'Smile Dental',
        address: {
          street: '1 Main St',
          city: 'Torrance',
          state: 'CA',
          postalCode: null,
          countryOrRegion: 'US',
        },
        coordinates: { latitude: 33.8, longitude: -118.3 },
      },
      organizer: { name: 'Alec', address: 'alec@example.com' },
      isCancelled: false,
      seriesMasterId: null,
      webLink: 'https://outlook.live.com/a',
      lastModified: '2026-10-20T10:00:00Z',
    });
    expect(result.events[0]).toMatchObject({
      calendarName: 'Family',
      location: null,
      organizer: null,
      seriesMasterId: 'series-1',
      responseStatus: 'accepted',
    });
    expect(result.lists).toEqual([]);
    expect(result.tasks).toEqual([]);
  });

  it('sends no Prefer header without a timezone', async () => {
    const { graph, batches } = fakeGraph({
      '/me/calendars': { value: [CALENDARS.value[0]] },
      '/me/calendars/cal-1/calendarView': { value: [] },
    });
    await runSweep(graph, { days: 1, now: NOW }, null);
    expect(batches[1][0].headers).toBeUndefined();
  });

  it('reads only the requested calendars and fails on an unknown id', async () => {
    const { graph, batches } = fakeGraph({
      '/me/calendars': CALENDARS,
      '/me/calendars/cal-2/calendarView': { value: [EVENT_B] },
    });
    const result = await runSweep(graph, { days: 1, calendars: ['cal-2'], now: NOW }, null);
    expect(batches[1].map((r) => r.url.split('?')[0])).toEqual([
      '/me/calendars/cal-2/calendarView',
    ]);
    expect(result.calendars.map((c) => c.id)).toEqual(['cal-2']);

    const again = fakeGraph({ '/me/calendars': CALENDARS });
    await expect(
      runSweep(again.graph, { days: 1, calendars: ['nope'], now: NOW }, null)
    ).rejects.toMatchObject({ code: 'graph_error' });
  });

  it('reads every To Do list and its tasks with --tasks', async () => {
    const { graph, batches } = fakeGraph({
      '/me/calendars': { value: [] },
      '/me/todo/lists': {
        value: [
          { id: 'l1', displayName: 'Tasks', wellknownListName: 'defaultList', isOwner: true },
          { id: 'l2', displayName: 'Groceries', wellknownListName: 'none' },
        ],
      },
      '/me/todo/lists/l1/tasks': {
        value: [
          {
            id: 't1',
            title: 'Call Tom',
            status: 'notStarted',
            importance: 'normal',
            dueDateTime: { dateTime: '2026-10-30T00:00:00.0000000', timeZone: 'UTC' },
            createdDateTime: '2026-10-01T10:00:00Z',
            lastModifiedDateTime: '2026-10-02T10:00:00Z',
            body: { content: '<p>About the car</p>', contentType: 'html' },
          },
        ],
      },
      '/me/todo/lists/l2/tasks': {
        value: [
          {
            id: 't2',
            title: 'Milk',
            status: 'completed',
            importance: 'low',
            createdDateTime: '2026-10-01T10:00:00Z',
            lastModifiedDateTime: '2026-10-03T10:00:00Z',
            completedDateTime: { dateTime: '2026-10-03T00:00:00.0000000', timeZone: 'UTC' },
            body: { content: '', contentType: 'text' },
          },
        ],
      },
    });
    const result = await runSweep(graph, { days: 1, tasks: true, now: NOW }, null);
    expect(batches[0].map((r) => r.id)).toEqual(['calendars', 'lists']);
    expect(result.lists).toEqual([
      { id: 'l1', displayName: 'Tasks', wellknownListName: 'defaultList' },
      { id: 'l2', displayName: 'Groceries', wellknownListName: 'none' },
    ]);
    expect(result.tasks).toEqual([
      {
        id: 't1',
        listId: 'l1',
        listName: 'Tasks',
        title: 'Call Tom',
        status: 'notStarted',
        importance: 'normal',
        dueDateTime: { dateTime: '2026-10-30T00:00:00.0000000', timeZone: 'UTC' },
        createdDateTime: '2026-10-01T10:00:00Z',
        lastModifiedDateTime: '2026-10-02T10:00:00Z',
        completedDateTime: null,
        body: { content: 'About the car' },
      },
      {
        id: 't2',
        listId: 'l2',
        listName: 'Groceries',
        title: 'Milk',
        status: 'completed',
        importance: 'low',
        dueDateTime: null,
        createdDateTime: '2026-10-01T10:00:00Z',
        lastModifiedDateTime: '2026-10-03T10:00:00Z',
        completedDateTime: { dateTime: '2026-10-03T00:00:00.0000000', timeZone: 'UTC' },
      },
    ]);
  });

  it('falls back to the lists delta when only Flagged Emails comes back', async () => {
    const { graph } = fakeGraph({
      '/me/calendars': { value: [] },
      '/me/todo/lists': {
        value: [{ id: 'f', displayName: 'Flagged Emails', wellknownListName: 'flaggedEmails' }],
      },
      '/me/todo/lists/delta()': {
        value: [
          { id: 'f', displayName: 'Flagged Emails', wellknownListName: 'flaggedEmails' },
          { id: 'l1', displayName: 'Tasks', wellknownListName: 'defaultList' },
        ],
        '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/todo/lists/delta()?$deltatoken=x',
      },
      '/me/todo/lists/f/tasks': { value: [] },
      '/me/todo/lists/l1/tasks': { value: [] },
    });
    const result = await runSweep(graph, { days: 1, tasks: true, now: NOW }, null);
    expect(result.lists.map((l) => l.id)).toEqual(['f', 'l1']);
  });

  it('follows nextLink pages', async () => {
    const { graph, calls } = fakeGraph({
      '/me/calendars': {
        value: [CALENDARS.value[0]],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/calendars?$skiptoken=p2',
      },
      '/me/calendars?$skiptoken=p2': { value: [CALENDARS.value[1]] },
      '/me/calendars/cal-1/calendarView': {
        value: [EVENT_A],
        '@odata.nextLink':
          'https://graph.microsoft.com/v1.0/me/calendars/cal-1/calendarView?$skiptoken=e2',
      },
      '/me/calendars/cal-1/calendarView?$skiptoken=e2': { value: [EVENT_B] },
      '/me/calendars/cal-2/calendarView': { value: [] },
    });
    const result = await runSweep(graph, { days: 1, timezone: 'UTC', now: NOW }, null);
    expect(result.calendars).toHaveLength(2);
    expect(result.events.map((e) => e.id)).toEqual(['ev-b', 'ev-a']);
    const page = calls.find((c) => c.path === '/me/calendars/cal-1/calendarView?$skiptoken=e2');
    expect(page?.init?.headers).toEqual({ Prefer: 'outlook.timezone="UTC"' });
  });

  it('retries a throttled sub-request on its own and fails on any other error', async () => {
    const routes = {
      '/me/calendars': { value: [CALENDARS.value[0]] },
      '/me/calendars/cal-1/calendarView': { value: [EVENT_A] },
    };
    const throttled = fakeGraph(routes, { '/me/calendars/cal-1/calendarView': { status: 429 } });
    const result = await runSweep(throttled.graph, { days: 1, now: NOW }, null);
    expect(result.events.map((e) => e.id)).toEqual(['ev-a']);
    expect(throttled.calls.some((c) => c.path.startsWith('/me/calendars/cal-1/calendarView'))).toBe(
      true
    );

    const denied = fakeGraph(routes, {
      '/me/calendars/cal-1/calendarView': {
        status: 403,
        body: { error: { code: 'ErrorAccessDenied', message: 'Access is denied.' } },
      },
    });
    await expect(runSweep(denied.graph, { days: 1, now: NOW }, null)).rejects.toThrow(
      'Microsoft Graph API error: 403 - ErrorAccessDenied: Access is denied.'
    );
  });

  it('splits more than 20 sub-requests over several batches', async () => {
    const calendars = Array.from({ length: 25 }, (_, i) => ({ id: `c${i}`, name: `C${i}` }));
    const routes: Record<string, unknown> = { '/me/calendars': { value: calendars } };
    for (const c of calendars) routes[`/me/calendars/${c.id}/calendarView`] = { value: [] };
    const { graph, batches } = fakeGraph(routes);
    await runSweep(graph, { days: 1, now: NOW }, null);
    expect(batches.map((b) => b.length)).toEqual([1, 20, 5]);
  });

  it('rejects a bad time zone', async () => {
    const { graph } = fakeGraph({});
    await expect(
      runSweep(graph, { days: 1, timezone: 'Mars/Base', now: NOW }, null)
    ).rejects.toMatchObject({ code: 'invalid_arguments' });
  });
});

describe('sweepCommand', () => {
  const graphClient = (routes: Record<string, unknown>) =>
    ({
      makeRequest: vi.fn(async (path: string) => {
        if (path === '/$batch') return { responses: [] };
        if (path.split('?')[0] in routes) return routes[path.split('?')[0]];
        throw new Error(`Microsoft Graph API error: 404 Not Found - ${path}`);
      }),
    }) as unknown as GraphClient;

  const auth = (account: unknown, getToken: () => Promise<string | null>) =>
    ({
      getCurrentAccount: vi.fn(async () => account),
      getToken: vi.fn(getToken),
    }) as unknown as AuthManager;

  it('exits 3 with exactly {"error":"not_signed_in"} when no account is cached', async () => {
    const result = await sweepCommand(
      auth(null, async () => 'x'),
      async () => graphClient({}),
      { days: '2' }
    );
    expect(result).toEqual({ code: 3, output: { error: 'not_signed_in' } });
  });

  it('exits 3 when the refresh needs the user again, 4 when it is the network', async () => {
    const interaction = Object.assign(new Error('Silent token acquisition failed'), {
      cause: Object.assign(new Error('AADSTS70000'), {
        name: 'InteractionRequiredAuthError',
        errorCode: 'invalid_grant',
      }),
    });
    const expired = await sweepCommand(
      auth({ username: 'a@x.com' }, async () => {
        throw interaction;
      }),
      async () => graphClient({}),
      { days: '2' }
    );
    expect(expired.code).toBe(3);
    expect(expired.output.error).toBe('not_signed_in');

    const offline = Object.assign(new Error('Silent token acquisition failed'), {
      cause: Object.assign(new Error('fetch failed'), { errorCode: 'network_error' }),
    });
    const down = await sweepCommand(
      auth({ username: 'a@x.com' }, async () => {
        throw offline;
      }),
      async () => graphClient({}),
      { days: '2' }
    );
    expect(down.code).toBe(4);
    expect(down.output.error).toBe('graph_error');
  });

  it('exits 2 on bad arguments before touching the account', async () => {
    const manager = auth(null, async () => 'x');
    expect((await sweepCommand(manager, async () => graphClient({}), { days: 'soon' })).code).toBe(
      2
    );
    expect(
      (await sweepCommand(manager, async () => graphClient({}), { days: '1', timezone: 'Nope/No' }))
        .code
    ).toBe(2);
    expect(manager.getCurrentAccount).not.toHaveBeenCalled();
  });

  it('exits 4 on a Graph error and 0 with the sweep otherwise', async () => {
    const manager = auth({ username: 'alec@x.com' }, async () => 'token-1');
    const failing = await sweepCommand(
      manager,
      async () =>
        ({
          makeRequest: vi.fn(async () => {
            throw new Error('Microsoft Graph API error: 500');
          }),
        }) as unknown as GraphClient,
      { days: '1' }
    );
    expect(failing).toEqual({
      code: 4,
      output: { error: 'graph_error', message: 'Microsoft Graph API error: 500' },
    });

    const client = graphClient({});
    (client.makeRequest as ReturnType<typeof vi.fn>).mockImplementation(
      async (path: string, init: { body?: string; accessToken?: string }) => {
        expect(init.accessToken).toBe('token-1');
        if (path !== '/$batch') throw new Error(path);
        const { requests } = JSON.parse(init.body!) as { requests: SubRequest[] };
        return {
          responses: requests.map((r) => ({ id: r.id, status: 200, body: { value: [] } })),
        };
      }
    );
    const ok = await sweepCommand(manager, async () => client, { days: '1', tasks: true });
    expect(ok.code).toBe(0);
    expect(ok.output).toMatchObject({
      account: 'alec@x.com',
      calendars: [],
      events: [],
      lists: [],
      tasks: [],
    });
  });

  it('classifies sign-in errors', () => {
    expect(isSignInError(new Error('No valid token found'))).toBe(true);
    expect(isSignInError(Object.assign(new Error('x'), { errorCode: 'consent_required' }))).toBe(
      true
    );
    expect(isSignInError(new Error('ECONNRESET'))).toBe(false);
  });
});
