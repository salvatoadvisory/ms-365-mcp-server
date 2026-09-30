import { afterEach, describe, expect, it } from 'vitest';
import { active, filterResponse, named, scopeSearch } from '../lib/steer.js';

const EX = 'Folder A - Private Work|Workspace B Code';
const SCOPE = 'https://contoso.sharepoint.com/sites/Home';

afterEach(() => {
  delete process.env.MS365_MCP_EXCLUDED_PATHS;
  delete process.env.MS365_MCP_SEARCH_SCOPE;
});

describe('steer', () => {
  it('changes nothing when neither setting is set', () => {
    expect(active()).toBe(false);
    const json = { value: [{ name: 'x', webUrl: 'https://contoso-my.sharepoint.com/Folder A - Private Work/x' }] };
    expect(filterResponse(json, '/me/drive').value).toEqual(json);
    expect(scopeSearch('{"requests":[]}', '')).toBe('{"requests":[]}');
  });

  it('drops items under an excluded folder, in plain and encoded paths', () => {
    process.env.MS365_MCP_EXCLUDED_PATHS = EX;
    const json = {
      value: [
        { name: 'keep.docx', webUrl: 'https://contoso-my.sharepoint.com/Documents/Firm/keep.docx' },
        { name: 'a.docx', parentReference: { path: '/drive/root:/Folder A - Private Work' } },
        { name: 'b.ts', webUrl: 'https://contoso-my.sharepoint.com/Workspace%20B%20Code/b.ts' },
      ],
    };
    const out = filterResponse(json, '/me/drive/root/children');
    expect((out.value as { value: unknown[] }).value).toHaveLength(1);
    expect(out.dropped).toBe(2);
  });

  it('drops search hits nested inside hitsContainers', () => {
    process.env.MS365_MCP_EXCLUDED_PATHS = EX;
    const json = {
      value: [
        {
          hitsContainers: [
            {
              hits: [
                { resource: { webUrl: 'https://contoso-my.sharepoint.com/Folder A - Private Work/z.pdf' } },
                { resource: { webUrl: 'https://contoso.sharepoint.com/sites/Home/y.pdf' } },
              ],
            },
          ],
        },
      ],
    };
    const out = filterResponse(json, '/search/query');
    const hitsLeft = (out.value as { value: Array<{ hitsContainers: Array<{ hits: unknown[] }> }> }).value[0]
      .hitsContainers[0].hits;
    expect(hitsLeft).toHaveLength(1);
  });

  it('leaves an excluded folder alone when the request names it', () => {
    process.env.MS365_MCP_EXCLUDED_PATHS = EX;
    expect(named('list the files in Folder A - Private Work please')).toBe(true);
    const json = { value: [{ webUrl: 'https://x/Folder A - Private Work/a' }] };
    expect(filterResponse(json, '/me/drive/root:/Folder A - Private Work:/children').dropped).toBe(0);
  });

  it('limits a file search to the scope unless OneDrive is asked for', () => {
    process.env.MS365_MCP_SEARCH_SCOPE = SCOPE;
    const body = JSON.stringify({ requests: [{ entityTypes: ['driveItem'], query: { queryString: 'proposal' } }] });
    const scoped = JSON.parse(scopeSearch(body, 'proposal'));
    expect(scoped.requests[0].query.queryString).toBe(`(proposal) AND (path:"${SCOPE}")`);
    expect(scopeSearch(body, 'proposal in my OneDrive')).toBe(body);
    const mail = JSON.stringify({ requests: [{ entityTypes: ['message'], query: { queryString: 'proposal' } }] });
    expect(scopeSearch(mail, 'proposal')).toBe(mail);
  });
});
