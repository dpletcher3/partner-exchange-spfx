// Fixtures for the extractor unit tests. These mirror real SharePoint REST
// payloads (captured from the live endpoint) plus crafted edge cases. SP REST
// legitimately returns `null` for empty typed fields, so the null literals here
// are intentional (the rig's no-new-null rule only flags null in type positions,
// not value positions, so no exemption is needed).

// --- extractUrl ---------------------------------------------------------------

export const urlAsObject = {
  Url: 'https://phillipscorp.com/article',
  Description: 'Article'
};
export const urlAsString = 'https://phillipscorp.com/article';
export const urlEmpty = '';
export const urlNull = null;
export const urlObjectWithoutUrl = { Description: 'no url here' };

// --- extractChoices -----------------------------------------------------------

export const categoriesAsArray = ['Phillips Loop', 'Phillips In The News'];
export const categoriesAsWrapped = { results: ['Phillips Loop', 'Phillips In The News'] };
export const categoriesAsString = 'Phillips Loop';
export const categoriesNull = null;

// --- extractThumbnailUrl ------------------------------------------------------

const MATCH_FILENAME =
  'Reserved_ImageAttachment_[14]_[ThumbnailImage][32]_[abc123def456][1]_[1].png';
const MATCH_URL =
  '/sites/PartnerExchange-Sandbox/Lists/News Repository/Attachments/5/' + MATCH_FILENAME;

export const thumbnailValidWithMatch = {
  rawThumbnail: JSON.stringify({
    fileName: MATCH_FILENAME,
    originalImageName: 'Screenshot 2026-05-26'
  }),
  attachmentFiles: [{ FileName: MATCH_FILENAME, ServerRelativeUrl: MATCH_URL }],
  expectedUrl: MATCH_URL
};

export const thumbnailNullField = {
  rawThumbnail: null,
  attachmentFiles: [] as unknown[]
};

export const thumbnailEmptyStringField = {
  rawThumbnail: '',
  attachmentFiles: [] as unknown[]
};

export const thumbnailFilenameNotInAttachments = {
  rawThumbnail: JSON.stringify({ fileName: MATCH_FILENAME, originalImageName: 'x' }),
  attachmentFiles: [
    { FileName: 'some-other-file.png', ServerRelativeUrl: '/x/some-other-file.png' }
  ],
  expectedFileName: MATCH_FILENAME
};

export const thumbnailMalformedJson = {
  rawThumbnail: 'not json',
  attachmentFiles: [] as unknown[]
};

export const thumbnailMissingFileNameKey = {
  rawThumbnail: '{"otherKey":"x"}',
  attachmentFiles: [] as unknown[],
  parsed: { otherKey: 'x' }
};

export const thumbnailAttachmentFilesUndefined = {
  rawThumbnail: JSON.stringify({ fileName: MATCH_FILENAME }),
  attachmentFiles: undefined as unknown,
  expectedFileName: MATCH_FILENAME
};
