import { beforeEach, describe, expect, it, vi } from 'vitest';
import { z } from 'zod';
import type { GraphClient } from '../src/graph-client.js';
import { registerGraphTools } from '../src/graph-tools.js';
import { findDuplicate, instantOf, normalizeTitle, sameInstant } from '../src/lib/create-dedupe.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

vi.mock('../src/generated/client-beta.js', () => ({ api: { endpoints: [] } }));
vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [
      {
        alias: 'create-calendar-event',
        method: 'post',
        path: '/me/events',
        description: 'Create an event.',
        parameters: [{ name: 'body', type: 'Body', schema: z.object({}).passthrough() }],
      },
      {
        alias: 'create-specific-calendar-event',
        method: 'post',
        path: '/me/calendars/:calendarId/events',
        description: 'Create an event on a calendar.',
        parameters: [{ name: 'body', type: 'Body', schema: z.object({}).passthrough() }],
      },
      {
        alias: 'create-todo-task',
        method: 'post',
        path: '/me/todo/lists/:todoTaskListId/tasks',
        description: 'Create a task.',
        parameters: [{ name: 'body', type: 'Body', schema: z.object({}).passthrough() }],
      },
      {
        alias: 'update-calendar-event',
        method: 'patch',
        path: '/me/events/:eventId',
        description: 'Update an event.',
        parameters: [{ name: 'body', type: 'Body', schema: z.object({}).passthrough() }],
      },
    ],
  },
}));

describe('dedupe helpers', () => {
  it('normalizes titles', () => {
    expect(normalizeTitle('  Dinner   with\tTom ')).toBe('dinner with tom');
    expect(normalizeTitle('ＤＩＮＮＥＲ')).toBe('dinner');
    expect(normalizeTitle(undefined)).toBe('');
  });

  it('resolves IANA and UTC zones, not Windows names', () => {
    const utc = Date.parse('2026-11-01T17:00:00Z');
    expect(instantOf({ dateTime: '2026-11-01T17:00:00.0000000', timeZone: 'UTC' })).toBe(utc);
    expect(instantOf({ dateTime: '2026-11-01T17:00:00' })).toBe(utc);
    expect(instantOf({ dateTime: '2026-11-01T18:00:00', timeZone: 'Europe/Brussels' })).toBe(utc);
    expect(instantOf({ dateTime: '2026-11-01T17:00:00Z', timeZone: 'Asia/Tokyo' })).toBe(utc);
    expect(instantOf({ dateTime: '2026-11-01T09:00:00', timeZone: 'Pacific Standard Time' })).toBe(
      null
    );
  });

  it('compares instants within a minute', () => {
    const a = { dateTime: '2026-11-01T18:00:00', timeZone: 'Europe/Brussels' };
    expect(sameInstant(a, { dateTime: '2026-11-01T17:00:30.0000000', timeZone: 'UTC' })).toBe(true);
    expect(sameInstant(a, { dateTime: '2026-11-01T17:02:00', timeZone: 'UTC' })).toBe(false);
    const w = { dateTime: '2026-11-01T09:00:00', timeZone: 'Pacific Standard Time' };
    expect(sameInstant(w, { dateTime: '2026-11-01T09:00:00.0000000', timeZone: w.timeZone })).toBe(
      true
    );
    expect(sameInstant(w, { dateTime: '2026-11-01T09:00:00', timeZone: 'UTC' })).toBe(false);
  });
});

describe('findDuplicate', () => {
  const event = {
    subject: 'Saturday soccer',
    start: { dateTime: '2026-10-31T10:00:00', timeZone: 'America/Los_Angeles' },
    end: { dateTime: '2026-10-31T11:00:00', timeZone: 'America/Los_Angeles' },
  };

  it('reads the calendarView a minute either side of the start', async () => {
    const fetchPage = vi.fn(async () => ({
      value: [
        {
          id: 'other',
          subject: 'Saturday soccer',
          start: { dateTime: '2026-10-31T18:00:00.0000000', timeZone: 'UTC' },
        },
        {
          id: 'same',
          subject: 'saturday  SOCCER',
          start: { dateTime: '2026-10-31T17:00:00.0000000', timeZone: 'UTC' },
        },
      ],
    }));
    const result = await findDuplicate('event', '/me/calendars/fam/events', event, fetchPage);
    expect(result).toEqual({
      duplicate: true,
      existing: {
        id: 'same',
        subject: 'saturday  SOCCER',
        start: { dateTime: '2026-10-31T17:00:00.0000000', timeZone: 'UTC' },
      },
    });
    const [path, headers] = fetchPage.mock.calls[0] as unknown as [string, Record<string, string>];
    expect(path.split('?')[0]).toBe('/me/calendars/fam/calendarView');
    expect(path).toContain('startDateTime=2026-10-31T16%3A59%3A00.000Z');
    expect(path).toContain('endDateTime=2026-10-31T17%3A01%3A00.000Z');
    expect(headers).toEqual({ Prefer: 'outlook.timezone="UTC"' });
  });

  it('checks the default calendar for /me/events and ignores cancelled events', async () => {
    const fetchPage = vi.fn(async () => ({
      value: [
        {
          id: 'gone',
          subject: 'Saturday soccer',
          isCancelled: true,
          start: { dateTime: '2026-10-31T17:00:00', timeZone: 'UTC' },
        },
      ],
    }));
    expect(await findDuplicate('event', '/me/events', event, fetchPage)).toBeNull();
    expect((fetchPage.mock.calls[0] as unknown as [string])[0]).toMatch(/^\/me\/calendarView\?/);
  });

  it('lets Graph write starts in a zone it cannot resolve itself', async () => {
    const fetchPage = vi.fn(async () => ({
      value: [
        {
          id: 'same',
          subject: 'Saturday soccer',
          start: { dateTime: '2026-10-31T10:00:00.0000000', timeZone: 'Pacific Standard Time' },
        },
      ],
    }));
    const body = {
      subject: 'Saturday soccer',
      start: { dateTime: '2026-10-31T10:00:00', timeZone: 'Pacific Standard Time' },
    };
    const result = await findDuplicate('event', '/me/events', body, fetchPage);
    expect(result?.existing.id).toBe('same');
    const [path, headers] = fetchPage.mock.calls[0] as unknown as [string, Record<string, string>];
    expect(path).toContain('startDateTime=2026-10-30T19%3A00%3A00.000Z');
    expect(headers).toEqual({ Prefer: 'outlook.timezone="Pacific Standard Time"' });
  });

  it('follows pages and skips bodies with nothing to compare', async () => {
    const fetchPage = vi
      .fn()
      .mockResolvedValueOnce({
        value: [],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=2',
      })
      .mockResolvedValueOnce({
        value: [
          {
            id: 'p2',
            subject: 'Saturday soccer',
            start: { dateTime: '2026-10-31T17:00:00', timeZone: 'UTC' },
          },
        ],
      });
    expect((await findDuplicate('event', '/me/events', event, fetchPage))?.existing.id).toBe('p2');
    expect(fetchPage.mock.calls[1][0]).toBe('/me/calendarView?$skiptoken=2');

    const none = vi.fn();
    expect(await findDuplicate('event', '/me/events', { subject: 'x' }, none)).toBeNull();
    expect(await findDuplicate('task', '/me/todo/lists/l/tasks', {}, none)).toBeNull();
    expect(none).not.toHaveBeenCalled();
  });

  it('matches a task on title and due time, completed or not', async () => {
    const fetchPage = vi.fn(async () => ({
      value: [
        {
          id: 't1',
          title: 'Pay rent',
          status: 'completed',
          dueDateTime: { dateTime: '2026-10-31T23:00:00.0000000', timeZone: 'UTC' },
        },
      ],
    }));
    const result = await findDuplicate(
      'task',
      '/me/todo/lists/l1/tasks',
      {
        title: 'Pay Rent',
        dueDateTime: { dateTime: '2026-11-01T00:00:00', timeZone: 'Europe/Brussels' },
      },
      fetchPage
    );
    expect(result).toEqual({
      duplicate: true,
      existing: {
        id: 't1',
        title: 'Pay rent',
        dueDateTime: { dateTime: '2026-10-31T23:00:00.0000000', timeZone: 'UTC' },
        status: 'completed',
      },
    });
    expect(fetchPage.mock.calls[0]).toEqual([
      '/me/todo/lists/l1/tasks',
      { Prefer: 'outlook.timezone="Europe/Brussels"' },
    ]);

    expect(
      await findDuplicate(
        'task',
        '/me/todo/lists/l1/tasks',
        { title: 'Pay rent', dueDateTime: { dateTime: '2026-11-02T00:00:00', timeZone: 'UTC' } },
        fetchPage
      )
    ).toBeNull();
  });

  it('matches a task without a due time only against open ones without one', async () => {
    const fetchPage = vi.fn(async () => ({
      value: [
        { id: 'done', title: 'Buy milk', status: 'completed' },
        {
          id: 'dated',
          title: 'Buy milk',
          status: 'notStarted',
          dueDateTime: { dateTime: '2026-11-01T00:00:00', timeZone: 'UTC' },
        },
      ],
    }));
    const path = '/me/todo/lists/l1/tasks';
    expect(await findDuplicate('task', path, { title: 'Buy milk' }, fetchPage)).toBeNull();
    fetchPage.mockResolvedValueOnce({
      value: [{ id: 'open', title: 'Buy milk', status: 'notStarted' }],
    });
    expect((await findDuplicate('task', path, { title: 'buy milk' }, fetchPage))?.existing.id).toBe(
      'open'
    );
  });
});

describe('create tools with dedupe in endpoints.json', () => {
  let mockServer: { tool: ReturnType<typeof vi.fn>; registerTool: ReturnType<typeof vi.fn> };
  let makeRequest: ReturnType<typeof vi.fn>;
  let graphRequest: ReturnType<typeof vi.fn>;

  const existing = {
    id: 'ev-1',
    subject: 'Dentist',
    start: { dateTime: '2026-11-02T16:00:00.0000000', timeZone: 'UTC' },
  };
  const body = {
    subject: 'Dentist',
    start: { dateTime: '2026-11-02T16:00:00', timeZone: 'UTC' },
    end: { dateTime: '2026-11-02T17:00:00', timeZone: 'UTC' },
  };

  beforeEach(() => {
    vi.clearAllMocks();
    mockServer = { tool: vi.fn(), registerTool: vi.fn() };
    makeRequest = vi.fn(async () => ({ value: [existing] }));
    graphRequest = vi.fn(async () => ({
      content: [{ type: 'text', text: JSON.stringify({ id: 'new' }) }],
    }));
    registerGraphTools(
      mockServer as never,
      { makeRequest, graphRequest } as unknown as GraphClient
    );
  });

  function tool(name: string) {
    const call = mockServer.registerTool.mock.calls.find((c: unknown[]) => c[0] === name);
    expect(call).toBeDefined();
    return {
      schema: (call![1] as { inputSchema: z.ZodObject }).inputSchema.shape as Record<
        string,
        unknown
      >,
      run: call![call!.length - 1] as (
        p: Record<string, unknown>
      ) => Promise<{ content: Array<{ text: string }>; isError?: boolean }>,
    };
  }

  it('offers allowDuplicate only on the tools that declare dedupe', () => {
    expect(tool('create-calendar-event').schema).toHaveProperty('allowDuplicate');
    expect(tool('create-specific-calendar-event').schema).toHaveProperty('allowDuplicate');
    expect(tool('create-todo-task').schema).toHaveProperty('allowDuplicate');
    expect(tool('update-calendar-event').schema).not.toHaveProperty('allowDuplicate');
  });

  it('refuses a duplicate without creating anything', async () => {
    const result = await tool('create-specific-calendar-event').run({ calendarId: 'fam', body });
    expect(JSON.parse(result.content[0].text)).toEqual({ duplicate: true, existing });
    expect(result.isError).toBeUndefined();
    expect(graphRequest).not.toHaveBeenCalled();
    expect(makeRequest.mock.calls[0][0]).toMatch(/^\/me\/calendars\/fam\/calendarView\?/);
    expect(makeRequest.mock.calls[0][1]).toMatchObject({ method: 'GET' });
  });

  it('creates when allowDuplicate is true, without sending the flag to Graph', async () => {
    const result = await tool('create-calendar-event').run({ body, allowDuplicate: true });
    expect(JSON.parse(result.content[0].text)).toEqual({ id: 'new' });
    expect(makeRequest).not.toHaveBeenCalled();
    const [path, options] = graphRequest.mock.calls[0] as [string, { body: string }];
    expect(path).toBe('/me/events');
    expect(JSON.parse(options.body)).toEqual(body);
  });

  it('creates when nothing matches', async () => {
    makeRequest.mockResolvedValueOnce({ value: [] });
    await tool('create-calendar-event').run({ body });
    expect(graphRequest).toHaveBeenCalledTimes(1);
  });

  it('creates nothing when the check cannot read the calendar', async () => {
    makeRequest.mockRejectedValueOnce(new Error('Microsoft Graph API error: 503'));
    const result = await tool('create-calendar-event').run({ body });
    expect(result.isError).toBe(true);
    expect(JSON.parse(result.content[0].text).error).toBe('duplicate_check_failed');
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('checks the target To Do list for create-todo-task', async () => {
    makeRequest.mockResolvedValueOnce({
      value: [{ id: 't1', title: 'Call Tom', status: 'notStarted' }],
    });
    const result = await tool('create-todo-task').run({
      todoTaskListId: 'l1',
      body: { title: 'call tom' },
    });
    expect(JSON.parse(result.content[0].text)).toMatchObject({
      duplicate: true,
      existing: { id: 't1' },
    });
    expect(makeRequest.mock.calls[0][0]).toBe('/me/todo/lists/l1/tasks');
    expect(graphRequest).not.toHaveBeenCalled();
  });
});
