import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType,
  ICustomCollectionField
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'PhillipsAudienceHeroWebPartStrings';
import { PhillipsAudienceHero, IPhillipsAudienceHeroProps } from './components/PhillipsAudienceHero';
import { AudienceProfileService } from './services/AudienceProfileService';
import { IAudienceTile } from './services/models';
import {
  FieldMappingController,
  IFieldSlot,
  resolveListIdByTitle,
  fetchChoiceFieldValues
} from '../../shared/fieldMapping';

const LOG = '[AudienceHero]';
// Tiles are effectively uncapped — authors add as many as they want. The high
// backstop (30) only guards the property-pane Add button against runaway input.
const MAX_TILES = 30;

// Defaults (confirmed by the 2026-06-29 survey). The Partner Profiles site is a
// pane setting so production can repoint it without a code change.
const DEFAULT_PP_SITE = 'https://phillipscorp.sharepoint.com/sites/PartnerExchange-OurPartners';
const DEFAULT_LIST_TITLE = 'Partner Profiles';
const DEFAULT_PERSON_FIELD = 'LinkedUser';
const DEFAULT_DIVISION_FIELD = 'Division';

// Tile row keys that are NOT division checkboxes (so the rest, when boolean-true,
// are the selected divisions). Includes PropertyFieldCollectionData internals.
const STRUCTURAL_TILE_KEYS = ['header', 'imageUrl', 'linkUrl', 'allowedDivisionsCsv', 'uniqueId', 'sortIdx'];

export interface IPhillipsAudienceHeroWebPartProps {
  partnerProfilesSiteUrl: string;
  partnerProfilesListTitle: string;
  personField: string;
  divisionField: string;
  // Each row: { header, imageUrl, linkUrl, [divisionValue]: boolean, ... }.
  tiles: Array<Record<string, unknown>>;
}

export default class PhillipsAudienceHeroWebPart extends BaseClientSideWebPart<IPhillipsAudienceHeroWebPartProps> {
  private _service!: AudienceProfileService;
  private _mapping!: FieldMappingController;
  private _ppListId = '';
  private _divisionChoices: string[] = [];
  // Lifecycle of the cross-site Division-choice load. The pane needs to tell
  // "not loaded yet" (PENDING → show a loading gate, NEVER the CSV fallback)
  // apart from "load failed" (→ CSV fallback). The earlier code had only the
  // choices array, so a pending load looked identical to a failed one and the
  // pane raced to the CSV field before the live choices arrived (the bug).
  private _divisionChoicesStatus: 'idle' | 'loading' | 'loaded' | 'failed' = 'idle';
  private _resolving = false;

  private get _ppSiteUrl(): string {
    return (this.properties.partnerProfilesSiteUrl || DEFAULT_PP_SITE).trim().replace(/\/+$/, '');
  }
  private get _ppListTitle(): string {
    return (this.properties.partnerProfilesListTitle || DEFAULT_LIST_TITLE).trim();
  }

  private get _slots(): IFieldSlot[] {
    return [
      { property: 'personField', label: strings.PersonFieldLabel, defaultInternalName: DEFAULT_PERSON_FIELD, typeFilter: 'person' },
      { property: 'divisionField', label: strings.DivisionFieldLabel, defaultInternalName: DEFAULT_DIVISION_FIELD, typeFilter: 'choice' }
    ];
  }

  protected onInit(): Promise<void> {
    this._service = new AudienceProfileService(this.context.spHttpClient);
    this._rebuildController();
    return super.onInit();
  }

  // The shared FieldMappingController fixes siteUrl at construction; the Partner
  // Profiles site is editable, so the controller is rebuilt whenever the source
  // is (re)resolved. getListId reads the resolved cross-site list id.
  private _rebuildController(): void {
    this._mapping = new FieldMappingController({
      spHttpClient: this.context.spHttpClient,
      siteUrl: this._ppSiteUrl,
      slots: this._slots,
      getListId: () => this._ppListId,
      properties: this.properties as unknown as Record<string, unknown>,
      refresh: () => this.context.propertyPane.refresh(),
      hintLabel: strings.FieldMappingEmptyLabel
    });
  }

  public render(): void {
    const tiles: IAudienceTile[] = (this.properties.tiles || []).map((row) => this._toTile(row));

    const props: IPhillipsAudienceHeroProps = {
      service: this._service,
      partnerProfilesSiteUrl: this._ppSiteUrl,
      listTitle: this._ppListTitle,
      personField: this.properties.personField || DEFAULT_PERSON_FIELD,
      divisionField: this.properties.divisionField || DEFAULT_DIVISION_FIELD,
      // Identity from pageContext only — NO Microsoft Graph (CLAUDE.md constraint).
      viewerEmail: this.context.pageContext.user.email || '',
      tiles,
      isEditMode: this.displayMode === DisplayMode.Edit
    };

    ReactDom.render(React.createElement(PhillipsAudienceHero, props), this.domElement);
  }

  // Reconstruct a tile's allowedDivisions from its row: any boolean-true key that
  // isn't a structural field is a checked division; an optional comma-separated
  // fallback field is also honored (used only when live choices weren't available).
  private _toTile(row: Record<string, unknown>): IAudienceTile {
    const fromBooleans = Object.keys(row).filter(
      (k) => row[k] === true && STRUCTURAL_TILE_KEYS.indexOf(k) === -1
    );
    const csv = typeof row.allowedDivisionsCsv === 'string' ? row.allowedDivisionsCsv : '';
    const fromCsv = csv.split(',').map((s) => s.trim()).filter((s) => !!s);

    const allowed: string[] = [];
    for (const d of fromBooleans.concat(fromCsv)) {
      if (allowed.indexOf(d) === -1) {
        allowed.push(d);
      }
    }
    return {
      header: typeof row.header === 'string' ? row.header : '',
      imageUrl: typeof row.imageUrl === 'string' ? row.imageUrl : '',
      linkUrl: typeof row.linkUrl === 'string' ? row.linkUrl : '',
      allowedDivisions: allowed
    };
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._resolvePartnerProfiles().catch(() => {
      /* errors are handled & logged inside _resolvePartnerProfiles */
    });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (
      (propertyPath === 'partnerProfilesSiteUrl' ||
        propertyPath === 'partnerProfilesListTitle' ||
        propertyPath === 'divisionField') &&
      oldValue !== newValue
    ) {
      console.log(`${LOG} source changed (${propertyPath}) — re-resolving Partner Profiles`);
      this._resolvePartnerProfiles().catch(() => {
        /* handled inside */
      });
    }

    this.render();
  }

  // Resolve the (cross-site) Partner Profiles list id from its title, load its
  // columns for the mapping dropdowns (shared FieldMappingController), and fetch
  // the Division column's choices to drive the per-tile multi-select.
  private async _resolvePartnerProfiles(): Promise<void> {
    if (this._resolving) {
      return;
    }
    this._resolving = true;
    // Enter the PENDING state and re-render the pane immediately, so the tile
    // schema shows the "Loading divisions…" gate (not the CSV fallback) while
    // the two cross-site calls run. This is the fix for the race: the pane no
    // longer commits to CSV before the live choices have had a chance to load.
    this._divisionChoicesStatus = 'loading';
    this.context.propertyPane.refresh();
    const site = this._ppSiteUrl;
    const title = this._ppListTitle;
    this._rebuildController(); // controller siteUrl follows the (editable) PP site

    try {
      this._ppListId = await resolveListIdByTitle(this.context.spHttpClient, site, title);
      console.log(`${LOG} resolved list '${title}' on ${site} → id=${this._ppListId || '(empty)'}`);
    } catch (err) {
      console.warn(`${LOG} could not resolve list '${title}' on ${site} — mapping/choices unavailable`, err);
      this._ppListId = '';
    }

    this._mapping.load();
    await this._loadDivisionChoices();
    this._resolving = false;
    this.context.propertyPane.refresh();
  }

  private async _loadDivisionChoices(): Promise<void> {
    const divField = this.properties.divisionField || DEFAULT_DIVISION_FIELD;
    if (!this._ppListId) {
      // The list never resolved — a genuine failure, not a pending state.
      this._divisionChoices = [];
      this._divisionChoicesStatus = 'failed';
      console.warn(`${LOG} division choices path=fallback (no resolved list) — comma-separated field shown`);
      return;
    }
    try {
      this._divisionChoices = await fetchChoiceFieldValues(
        this.context.spHttpClient,
        this._ppSiteUrl,
        this._ppListId,
        divField
      );
      // A successful fetch with at least one choice is the only path to the
      // checkbox control. A successful fetch returning zero choices is treated
      // as a failure so authors still get the CSV fallback (a Division field
      // with no choices is degenerate and checkboxes would be empty).
      this._divisionChoicesStatus = this._divisionChoices.length > 0 ? 'loaded' : 'failed';
      console.log(`${LOG} division choices path=live (${this._divisionChoices.length}): ${this._divisionChoices.join(', ')}`);
    } catch (err) {
      this._divisionChoices = [];
      this._divisionChoicesStatus = 'failed';
      console.warn(`${LOG} division choices path=fallback (live fetch failed) — comma-separated field shown`, err);
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const tiles = this.properties.tiles || [];

    // Tile schema: header / image / link, then the division control. The
    // division control has THREE shapes, picked from the choice-load lifecycle:
    //   • loaded (live choices) → ONE boolean checkbox per Division choice. Each
    //     field id IS the division value, so it binds to the boolean row key
    //     (`row["Hanover Operations"]` etc.) — existing selections render
    //     pre-checked and are written back in the same shape (no clobber).
    //   • loading (cross-site fetch in flight) → a single disabled placeholder
    //     so the pane never shows CSV merely because choices haven't arrived.
    //   • failed/idle (list not resolved, fetch threw, or zero choices) → the
    //     comma-separated fallback so authors can still target divisions.
    const tileFields: ICustomCollectionField[] = [
      { id: 'header', title: strings.TileHeaderLabel, type: CustomCollectionFieldType.string, required: true },
      { id: 'imageUrl', title: strings.TileImageUrlLabel, type: CustomCollectionFieldType.string },
      { id: 'linkUrl', title: strings.TileLinkUrlLabel, type: CustomCollectionFieldType.string }
    ];

    // The collection-control key encodes the division-control shape. PnP's
    // PropertyFieldCollectionData caches its column schema per key; bumping the
    // key on the loading→loaded transition forces React to REMOUNT the control
    // with the checkbox columns. A plain propertyPane.refresh() alone did not
    // reliably swap the columns — this remount is the load-bearing piece.
    let collectionKey: string;
    if (this._divisionChoices.length > 0) {
      for (const choice of this._divisionChoices) {
        tileFields.push({ id: choice, title: choice, type: CustomCollectionFieldType.boolean });
      }
      collectionKey = `audienceHeroTiles-choices-${this._divisionChoices.length}`;
    } else if (this._divisionChoicesStatus === 'loading') {
      tileFields.push({
        id: '__divisionsLoading',
        title: strings.DivisionsLoadingLabel,
        type: CustomCollectionFieldType.string,
        disableEdit: true
      });
      collectionKey = 'audienceHeroTiles-loading';
    } else {
      tileFields.push({
        id: 'allowedDivisionsCsv',
        title: strings.TileAllowedCsvLabel,
        type: CustomCollectionFieldType.string
      });
      collectionKey = 'audienceHeroTiles-csv';
    }

    const sourceFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneTextField('partnerProfilesSiteUrl', {
        label: strings.SiteUrlFieldLabel,
        placeholder: DEFAULT_PP_SITE
      }),
      PropertyPaneTextField('partnerProfilesListTitle', {
        label: strings.ListTitleFieldLabel,
        placeholder: DEFAULT_LIST_TITLE
      })
    ];

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.SourceGroupName,
              groupFields: sourceFields
            },
            {
              groupName: strings.FieldMappingGroupName,
              // Expanded — a collapsed mapping group reads as empty (lessons-learned).
              isCollapsed: false,
              groupFields: this._mapping.buildFields()
            },
            {
              groupName: strings.TilesGroupName,
              groupFields: [
                PropertyFieldCollectionData('tiles', {
                  key: collectionKey,
                  label: strings.TilesFieldLabel,
                  panelHeader: strings.TilesPanelHeader,
                  manageBtnLabel: strings.TilesManageButtonLabel,
                  value: tiles,
                  enableSorting: true,
                  // Effectively uncapped — disable the Add row only at the high backstop (MAX_TILES).
                  disableItemCreation: tiles.length >= MAX_TILES,
                  fields: tileFields
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
