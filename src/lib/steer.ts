/**
 * Steer file search to the places an organisation wants searched.
 *
 * Two optional settings, both from the environment so no path lives in code:
 *
 *   MS365_MCP_EXCLUDED_PATHS  folder names or paths, separated by "|". Any item whose URL
 *                             or path contains one is dropped from every response, unless
 *                             the request itself names that folder.
 *   MS365_MCP_SEARCH_SCOPE    site or folder URLs, separated by "|". A search-query over
 *                             files, list items or sites is limited to these paths, unless
 *                             the request asks for OneDrive explicitly.
 *
 * With neither set, nothing here changes anything.
 */

const split = (value: string | undefined): string[] =>
  (value || '')
    .split('|')
    .map((s) => s.trim())
    .filter(Boolean);

export function excludedPaths(): string[] {
  return split(process.env.MS365_MCP_EXCLUDED_PATHS);
}

export function searchScope(): string[] {
  return split(process.env.MS365_MCP_SEARCH_SCOPE);
}

export function active(): boolean {
  return excludedPaths().length > 0 || searchScope().length > 0;
}

/** Whether the request itself names an excluded folder, which is the explicit ask. */
export function named(requestText: string): boolean {
  const low = requestText.toLowerCase();
  return excludedPaths().some(
    (p) => low.includes(p.toLowerCase()) || low.includes(encodeURIComponent(p).toLowerCase())
  );
}

function forms(path: string): string[] {
  const lower = path.toLowerCase();
  return [lower, encodeURIComponent(path).toLowerCase(), lower.replace(/ /g, '%20')];
}

function hits(item: unknown, paths: string[]): boolean {
  if (item === null || typeof item !== 'object') return false;
  const text = JSON.stringify(item).toLowerCase();
  return paths.some((p) => forms(p).some((f) => text.includes(f)));
}

/**
 * The response with every item under an excluded path taken out of every list in it.
 * Returns the filtered value and how many items were dropped.
 */
export function filterResponse(json: unknown, requestText: string): { value: unknown; dropped: number } {
  const paths = excludedPaths();
  if (!paths.length || named(requestText)) return { value: json, dropped: 0 };
  let dropped = 0;
  const walk = (node: unknown): unknown => {
    if (Array.isArray(node)) {
      // Inside out: the lists within an item are cleaned first, so a search container that
      // held one excluded hit keeps its other hits rather than being dropped whole.
      return node.map(walk).filter((x) => {
        const drop = hits(x, paths);
        if (drop) dropped++;
        return !drop;
      });
    }
    if (node && typeof node === 'object') {
      const out: Record<string, unknown> = {};
      for (const [k, v] of Object.entries(node as Record<string, unknown>)) out[k] = walk(v);
      return out;
    }
    return node;
  };
  return { value: walk(json), dropped };
}

const FILE_TYPES = new Set(['driveitem', 'listitem', 'site', 'list', 'drive']);
const ONEDRIVE = /onedrive|-my\.sharepoint\.com/i;

/** A search-query body limited to the search scope, unless it asks for OneDrive. */
export function scopeSearch(body: string, requestText: string): string {
  const scope = searchScope();
  if (!scope.length || ONEDRIVE.test(requestText)) return body;
  let parsed: { requests?: Array<{ entityTypes?: string[]; query?: { queryString?: string } }> };
  try {
    parsed = JSON.parse(body);
  } catch {
    return body;
  }
  const filter = scope.map((s) => `path:"${s}"`).join(' OR ');
  for (const r of parsed.requests || []) {
    const types = (r.entityTypes || []).map((t) => String(t).toLowerCase());
    if (!types.some((t) => FILE_TYPES.has(t)) || !r.query?.queryString) continue;
    r.query.queryString = `(${r.query.queryString}) AND (${filter})`;
  }
  return JSON.stringify(parsed);
}
