/**
 * Shared interfaces for the internal Loop / Substrate API shapes.
 *
 * These mirror the fields the Loop web client receives from
 * substrate.office.com/recommended/api/v1.1/loop/*. Only the fields this
 * server actually uses are typed; everything else is ignored.
 */

export interface LoopWorkspace {
  id: string;
  title?: string;
  /** Base64 pointer to the SharePoint backing store: "…|{host}|{driveId}|{itemId}". */
  mfs_info?: { pod_id?: string };
  isPersonal?: boolean;
}

export interface LoopPage {
  id: string;
  title?: string;
  type?: string;
  workspace_id?: string;
  is_deleted?: boolean;
  onedrive_info?: { drive_id?: string };
  sharepoint_info?: { site_url?: string };
}

export interface LoopData {
  workspaces?: LoopWorkspace[];
  pages?: LoopPage[];
  next_page_link?: string;
}

/** Decoded SharePoint coordinates for a Loop workspace or page. */
export interface SpoCoordinates {
  host: string;
  driveId: string;
  itemId: string;
}
