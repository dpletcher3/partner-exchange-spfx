# News Repository — content guidelines

How to author items in the News Repository list so they render cleanly in the Phillips News web part on the hub home page and any other surface that pulls from this list.

## Title

- Keep it under 60 characters. Longer titles wrap to a second line in the card and start crowding the description.
- Avoid all-caps unless the source uses it (e.g., acronyms). The card's typography handles emphasis; ALL CAPS reads as shouting.
- Title sentence case is recommended over title case for readability ("How the additive team cut prototype cycles" not "How The Additive Team Cut Prototype Cycles").

## ShortDescription

- Keep it under 120 characters. The card truncates with an ellipsis at roughly that point.
- One sentence. If you need two, the article is the place for the second one; the card description is a teaser.
- Lead with the most newsworthy fact, not the setup ("Phillips named to Defense News Top 100 for the seventh year" not "Earlier this month, an industry publication announced...").
- Don't repeat the title. The card shows both; the description should add information.
- Plain text only. No formatting, links, or HTML.
- Optional — items without a description still render, with a slightly shorter card. Prefer including one for visual consistency in a row.

## ThumbnailImage

- Recommended size: at least 800×600 (landscape orientation). Smaller images get upscaled and lose sharpness.
- Aspect ratio: 4:3 is ideal. Other ratios get center-cropped to fit the 4:3 card frame.
- Composition: keep important detail (text, faces, logos) within the **center 80% of the image**. Edges may be cropped to fit a 4:3 frame at various card widths.
- Avoid heavy text overlays on the image; the card has its own text below the thumbnail. Branded headers or watermarks at the image edges are likely to be cropped — re-export with the brand in the center, or rely on the card title.
- Optional — items without a thumbnail render with a solid Phillips-red placeholder block. Prefer including a thumbnail; the placeholder dominates the visual grid when several items lack images.

## Category

Multi-select. Pick one or more from the configured choices:
- Phillips Loop — internal culture, partnership, milestones, behind-the-scenes
- Phillips In The News — external coverage of Phillips by third-party publications
- (other categories as added)

An item can belong to multiple categories. It will surface in any web part instance filtered to a category it belongs to (e.g., the Markforged partnership announcement tagged both "Phillips Loop" and "Phillips In The News" appears in both hub-home instances).

## ItemType

Single value. Picks the kind of item:
- Internal Story — points at a SharePoint or Phillips internal page
- External Link — points at a phillipscorp.com or other external site
- (other types as added)

This drives where clicking a card takes the user. The LinkUrl field controls the actual destination; ItemType is a categorical tag for filtering and (in future) different rendering treatments.

## LinkUrl

- Always populate this. Cards are clickable; without a URL the card has nothing to navigate to.
- Use the canonical URL from phillipscorp.com or the intranet, not a redirected or tracking link.
- The card opens this URL in a new tab. Same-tab navigation isn't currently configurable.

## PublishedDate

- Use the date the news was originally published, not the date you added it to the list.
- The card displays this date prominently. Backdating to "make news look fresh" is a poor practice — surface real news, or remove items that are no longer timely.
- Items are ordered by PublishedDate descending in the web part; the most recent news appears first.

## Editorial cadence

- The web part shows up to N items per instance (default 6, configurable in the property pane).
- Aim for at least 6 active items in the list at any time so the grid is always full.
- Archive or remove items more than ~6 months old unless they are still being referenced. Old items pushed off the visible grid are still in the list and searchable; they just don't appear in the default rendering.
