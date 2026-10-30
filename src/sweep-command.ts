import type AuthManager from './auth.js';
import type GraphClient from './graph-client.js';
import {
  assertTimeZone,
  parseCalendars,
  parseDays,
  runSweep,
  SweepError,
  type SweepGraph,
  type SweepOptions,
} from './sweep.js';

/** Exit codes of `--sweep`: 0 ok, 2 bad arguments, 3 not signed in, 4 Graph or network error. */
export const SWEEP_EXIT = { ok: 0, invalidArguments: 2, notSignedIn: 3, graphError: 4 } as const;

export interface SweepCommandArgs {
  days?: unknown;
  calendars?: unknown;
  tasks?: unknown;
  timezone?: unknown;
}

export interface SweepCommandResult {
  code: number;
  output: Record<string, unknown>;
}

// MSAL error codes that mean the cached sign-in cannot be used without the user
const SIGN_IN_ERROR_CODES = new Set([
  'invalid_grant',
  'interaction_required',
  'consent_required',
  'login_required',
  'no_tokens_found',
  'no_account_in_silent_request',
  'token_refresh_required',
  'refresh_token_expired',
  'bad_token',
]);

export function isSignInError(error: unknown): boolean {
  const err = error as { cause?: unknown; message?: string } | undefined;
  if (err?.message === 'No valid token found') return true;
  const cause = (err?.cause ?? error) as { errorCode?: string; name?: string } | undefined;
  return (
    cause?.name === 'InteractionRequiredAuthError' ||
    (typeof cause?.errorCode === 'string' && SIGN_IN_ERROR_CODES.has(cause.errorCode))
  );
}

/** A SweepGraph over GraphClient.makeRequest with a fixed access token. */
export function graphClientAdapter(graphClient: GraphClient, accessToken?: string): SweepGraph {
  return {
    request: (path, init = {}) =>
      graphClient.makeRequest(path, {
        method: init.method ?? 'GET',
        headers: init.headers ?? {},
        ...(init.body !== undefined ? { body: init.body } : {}),
        ...(accessToken ? { accessToken } : {}),
      }),
  };
}

function sweepOptions(args: SweepCommandArgs): SweepOptions {
  if (typeof args.timezone === 'string' && args.timezone.trim() !== '') {
    assertTimeZone(args.timezone.trim());
  }
  return {
    days: parseDays(args.days),
    calendars: parseCalendars(args.calendars),
    tasks: args.tasks === true,
    ...(typeof args.timezone === 'string' && args.timezone.trim() !== ''
      ? { timezone: args.timezone.trim() }
      : {}),
  };
}

/**
 * The `--sweep` one-shot: never throws, never prompts. The caller prints `output` as one
 * JSON line and exits with `code`.
 */
export async function sweepCommand(
  authManager: AuthManager,
  makeGraphClient: () => Promise<GraphClient>,
  args: SweepCommandArgs
): Promise<SweepCommandResult> {
  let options: SweepOptions;
  try {
    options = sweepOptions(args);
  } catch (error) {
    return {
      code: SWEEP_EXIT.invalidArguments,
      output: { error: 'invalid_arguments', message: (error as Error).message },
    };
  }

  let account: string | null;
  let token: string | null;
  try {
    const current = await authManager.getCurrentAccount();
    if (!current) return { code: SWEEP_EXIT.notSignedIn, output: { error: 'not_signed_in' } };
    account = current.username ?? null;
    token = await authManager.getToken();
    if (!token) return { code: SWEEP_EXIT.notSignedIn, output: { error: 'not_signed_in' } };
  } catch (error) {
    if (isSignInError(error)) {
      return {
        code: SWEEP_EXIT.notSignedIn,
        output: { error: 'not_signed_in', message: (error as Error).message },
      };
    }
    return {
      code: SWEEP_EXIT.graphError,
      output: { error: 'graph_error', message: (error as Error).message },
    };
  }

  try {
    const graph = graphClientAdapter(await makeGraphClient(), token);
    const result = await runSweep(graph, options, account);
    return { code: SWEEP_EXIT.ok, output: result as unknown as Record<string, unknown> };
  } catch (error) {
    if (error instanceof SweepError && error.code === 'invalid_arguments') {
      return {
        code: SWEEP_EXIT.invalidArguments,
        output: { error: 'invalid_arguments', message: error.message },
      };
    }
    if (isSignInError(error)) {
      return {
        code: SWEEP_EXIT.notSignedIn,
        output: { error: 'not_signed_in', message: (error as Error).message },
      };
    }
    return {
      code: SWEEP_EXIT.graphError,
      output: { error: 'graph_error', message: (error as Error).message },
    };
  }
}
