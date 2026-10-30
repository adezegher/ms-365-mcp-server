import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { z } from 'zod';
import type { GraphClient } from '../src/graph-client.js';
import { registerGraphTools } from '../src/graph-tools.js';
import { bodyCarriesAttendees, hasOtherAttendees } from '../src/lib/attendee-guard.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

vi.mock('../src/generated/client-beta.js', () => ({ api: { endpoints: [] } }));
vi.mock('../src/generated/client.js', () => {
  const body = [{ name: 'body', type: 'Body', schema: z.object({}).passthrough() }];
  return {
    api: {
      endpoints: [
        { alias: 'create-calendar-event', method: 'post', path: '/me/events', parameters: body },
        {
          alias: 'create-specific-calendar-event',
          method: 'post',
          path: '/me/calendars/:calendarId/events',
          parameters: body,
        },
        {
          alias: 'update-calendar-event',
          method: 'patch',
          path: '/me/events/:eventId',
          parameters: body,
        },
        {
          alias: 'update-specific-calendar-event',
          method: 'patch',
          path: '/me/calendars/:calendarId/events/:eventId',
          parameters: body,
        },
        {
          alias: 'create-todo-task',
          method: 'post',
          path: '/me/todo/lists/:todoTaskListId/tasks',
          parameters: body,
        },
      ],
    },
  };
});

const ALEC = { emailAddress: { name: 'Alec', address: 'Alec@Example.com' } };
const TOM = { emailAddress: { name: 'Tom', address: 'tom@example.com' } };

describe('attendee guard helpers', () => {
  it('spots attendees in any casing and responseRequested: true', () => {
    expect(bodyCarriesAttendees({ subject: 'x', attendees: [] })).toBe(true);
    expect(bodyCarriesAttendees({ Attendees: [TOM] })).toBe(true);
    expect(bodyCarriesAttendees(JSON.stringify({ ATTENDEES: [TOM] }))).toBe(true);
    expect(bodyCarriesAttendees({ responseRequested: true })).toBe(true);
    expect(bodyCarriesAttendees({ ResponseRequested: false, subject: 'x' })).toBe(false);
    expect(bodyCarriesAttendees({ subject: 'attendees' })).toBe(false);
    expect(bodyCarriesAttendees('not json, but attendees')).toBe(true);
    expect(bodyCarriesAttendees(null)).toBe(false);
  });

  it('counts attendees other than the organizer', () => {
    expect(hasOtherAttendees({ organizer: ALEC, attendees: [] })).toBe(false);
    expect(hasOtherAttendees({ organizer: ALEC })).toBe(false);
    expect(
      hasOtherAttendees({
        organizer: ALEC,
        attendees: [{ emailAddress: { address: 'alec@example.com ' } }],
      })
    ).toBe(false);
    expect(hasOtherAttendees({ organizer: ALEC, attendees: [ALEC, TOM] })).toBe(true);
    expect(hasOtherAttendees({ organizer: ALEC, attendees: [{ emailAddress: {} }] })).toBe(true);
  });
});

describe('no-attendees mode on event tools', () => {
  let mockServer: { tool: ReturnType<typeof vi.fn>; registerTool: ReturnType<typeof vi.fn> };
  let makeRequest: ReturnType<typeof vi.fn>;
  let graphRequest: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    vi.clearAllMocks();
    process.env.MS365_MCP_NO_ATTENDEES = '1';
    mockServer = { tool: vi.fn(), registerTool: vi.fn() };
    makeRequest = vi.fn(async () => ({ value: [] }));
    graphRequest = vi.fn(async () => ({
      content: [{ type: 'text', text: JSON.stringify({ id: 'written' }) }],
    }));
    registerGraphTools(
      mockServer as never,
      { makeRequest, graphRequest } as unknown as GraphClient
    );
  });

  afterEach(() => {
    delete process.env.MS365_MCP_NO_ATTENDEES;
  });

  function run(name: string, params: Record<string, unknown>) {
    const call = mockServer.registerTool.mock.calls.find((c: unknown[]) => c[0] === name);
    expect(call).toBeDefined();
    const handler = call![call!.length - 1] as (
      p: Record<string, unknown>
    ) => Promise<{ content: Array<{ text: string }>; isError?: boolean }>;
    return handler(params);
  }

  const event = {
    subject: 'Dinner',
    start: { dateTime: '2026-11-02T19:00:00', timeZone: 'UTC' },
    end: { dateTime: '2026-11-02T21:00:00', timeZone: 'UTC' },
  };

  it.each([
    ['create-calendar-event', { body: { ...event, attendees: [TOM] } }],
    ['create-specific-calendar-event', { calendarId: 'fam', body: { ...event, Attendees: [] } }],
    ['create-calendar-event', { body: { ...event, responseRequested: true } }],
    ['create-calendar-event', { body: event, attendees: [TOM] }],
    ['update-calendar-event', { eventId: 'e1', body: { attendees: [TOM] } }],
    ['update-specific-calendar-event', { calendarId: 'c', eventId: 'e1', body: { ATTENDEES: [] } }],
  ])('%s refuses a body with attendees and writes nothing', async (name, params) => {
    const result = await run(name, params);
    expect(result.isError).toBe(true);
    expect(JSON.parse(result.content[0].text).error).toBe('attendees_not_allowed');
    expect(graphRequest).not.toHaveBeenCalled();
    expect(makeRequest).not.toHaveBeenCalled();
  });

  it('creates an event without attendees', async () => {
    const result = await run('create-calendar-event', { body: event });
    expect(JSON.parse(result.content[0].text)).toEqual({ id: 'written' });
    expect(graphRequest).toHaveBeenCalledTimes(1);
  });

  it('refuses to update an event that has attendees besides the organizer', async () => {
    makeRequest.mockResolvedValueOnce({ organizer: ALEC, attendees: [ALEC, TOM] });
    const result = await run('update-specific-calendar-event', {
      calendarId: 'fam',
      eventId: 'e1',
      body: { subject: 'Dinner moved' },
    });
    expect(result.isError).toBe(true);
    expect(JSON.parse(result.content[0].text).error).toBe('event_has_attendees');
    expect(makeRequest.mock.calls[0][0]).toBe(
      '/me/calendars/fam/events/e1?$select=attendees,organizer'
    );
    expect(makeRequest.mock.calls[0][1]).toMatchObject({ method: 'GET' });
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('updates an event whose only attendee is the organizer', async () => {
    makeRequest.mockResolvedValueOnce({ organizer: ALEC, attendees: [ALEC] });
    const result = await run('update-calendar-event', { eventId: 'e1', body: { subject: 'x' } });
    expect(JSON.parse(result.content[0].text)).toEqual({ id: 'written' });
    const [path, options] = graphRequest.mock.calls[0] as [string, { method: string }];
    expect(path).toBe('/me/events/e1');
    expect(options.method).toBe('PATCH');
  });

  it('writes nothing when the event cannot be read', async () => {
    makeRequest.mockRejectedValueOnce(new Error('Microsoft Graph API error: 404'));
    const result = await run('update-calendar-event', { eventId: 'e1', body: { subject: 'x' } });
    expect(JSON.parse(result.content[0].text).error).toBe('attendee_check_failed');
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('leaves tools without attendeeGuard alone', async () => {
    const result = await run('create-todo-task', {
      todoTaskListId: 'l1',
      body: { title: 'x', attendees: 'not an event field' },
    });
    expect(JSON.parse(result.content[0].text)).toEqual({ id: 'written' });
  });

  it('does nothing with the mode off', async () => {
    delete process.env.MS365_MCP_NO_ATTENDEES;
    const result = await run('update-calendar-event', {
      eventId: 'e1',
      body: { attendees: [TOM] },
    });
    expect(JSON.parse(result.content[0].text)).toEqual({ id: 'written' });
    expect(makeRequest).not.toHaveBeenCalled();
  });
});
