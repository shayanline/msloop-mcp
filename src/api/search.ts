/**
 * Search for Loop pages/components via the Microsoft Graph Search API.
 *
 * Graph can surface `.loop` / `.fluid` files across the SharePoint Embedded
 * containers a user has access to. This is the most reliable cross-workspace
 * search available without the Fluid runtime.
 */

import { GRAPH_BASE } from '../constants.js';
import { graphPost } from './client.js';

export interface SearchHit {
  id?: string;
  name?: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  summary?: string;
  driveId?: string;
}

interface GraphSearchResponse {
  value?: Array<{
    hitsContainers?: Array<{
      total?: number;
      hits?: Array<{
        summary?: string;
        resource?: {
          id?: string;
          name?: string;
          webUrl?: string;
          lastModifiedDateTime?: string;
          parentReference?: { driveId?: string };
        };
      }>;
    }>;
  }>;
}

/** Search the user's accessible Loop files by keyword. */
export async function searchLoopFiles(query: string, size = 25): Promise<SearchHit[]> {
  const queryString = `(${query}) AND (filetype:loop OR filetype:fluid)`;
  const body = {
    requests: [
      {
        entityTypes: ['driveItem'],
        query: { queryString },
        from: 0,
        size,
      },
    ],
  };

  const res = await graphPost<GraphSearchResponse>(`${GRAPH_BASE}/search/query`, body);
  const hits = res.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
  return hits.map(h => ({
    id: h.resource?.id,
    name: h.resource?.name,
    webUrl: h.resource?.webUrl,
    lastModifiedDateTime: h.resource?.lastModifiedDateTime,
    summary: h.summary,
    driveId: h.resource?.parentReference?.driveId,
  }));
}
