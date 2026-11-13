import { vi } from "vitest";

/** JSON response as returned by Nextcloud OCS / Microsoft Graph. */
export const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });

/** OCS envelope ({ ocs: { meta, data } }). */
export const ocs = (data: unknown, status = 200, message = "OK") =>
  json({ ocs: { meta: { status: status < 400 ? "ok" : "failure", statuscode: status, message }, data } }, status);

/** Replaces global fetch with a mock that answers the queued responses in order and records the requests. */
export function mockFetch(...responses: Response[]) {
  const fetchMock = vi.fn(async (_input: RequestInfo | URL, _init?: RequestInit) => {
    const next = responses.shift();
    if (!next) throw new Error("Unexpected fetch call");
    return next;
  });
  vi.stubGlobal("fetch", fetchMock);
  return {
    fetchMock,
    /** URL and init of the n-th call. */
    call: (n: number) => ({
      url: String(fetchMock.mock.calls[n][0]),
      init: fetchMock.mock.calls[n][1] ?? {},
    }),
  };
}

export const quietLog = () => ({ info: vi.fn(), warn: vi.fn() });
