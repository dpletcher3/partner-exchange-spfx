// Domain models for the Audience Hero web part.

// One configured hero tile. allowedDivisions is the set of Partner Profiles
// Division values whose members may see this tile (empty set = shown to no one).
export interface IAudienceTile {
  header: string;
  // Background image URL (server-relative or absolute). '' renders a neutral fill.
  imageUrl: string;
  // Click target for the whole tile. '' renders a non-navigating tile.
  linkUrl: string;
  // Division values this tile is shown to (matched against the viewer's single
  // Division). Empty = no audience (hidden from everyone) — fail-closed.
  allowedDivisions: string[];
}

// Inputs for the cross-site viewer-division lookup.
export interface IProfileLookupOptions {
  // Absolute URL of the site hosting Partner Profiles (cross-site from The Hub).
  siteUrl: string;
  listTitle: string;
  // Internal name of the Person column joined on the viewer (default LinkedUser).
  personField: string;
  // Internal name of the single-value Choice column (default Division).
  divisionField: string;
  // The viewer's email from pageContext (NO Graph) — the join key.
  viewerEmail: string;
}
