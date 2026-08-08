import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';

export interface ToolCategory {
  name: string;
  pattern: RegExp;
  description: string;
  requiresOrgMode?: boolean;
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const endpointEntries = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as Array<{ toolName: string; presets?: string[] }>;

// Preset metadata. Membership lives in endpoints.json: each endpoint declares
// which presets it belongs to via its `presets` array, so presets are exact
// tool-name allow-lists that can't over-match across apps the way the old
// loose name regexes could (e.g. "mail" also matching shared-mailbox tools).
const PRESET_META: Record<string, { description: string; requiresOrgMode?: boolean }> = {
  mail: {
    description: 'Email operations (read, send, manage folders, attachments)',
  },
  calendar: {
    description: 'Calendar and event management',
  },
  files: {
    description: 'OneDrive file and folder operations',
  },
  personal: {
    description:
      'Personal productivity tools (mail, calendar, files, contacts, tasks, notes, search)',
  },
  work: {
    description: 'Organization/work tools (Teams, SharePoint, shared mailboxes, search)',
    requiresOrgMode: true,
  },
  excel: {
    description: 'Excel spreadsheet operations',
  },
  contacts: {
    description: 'Outlook contacts management',
  },
  tasks: {
    description: 'Task and planning tools (To Do, Planner)',
  },
  onenote: {
    description: 'OneNote notebook operations',
  },
  search: {
    description: 'Microsoft Search capabilities',
  },
  users: {
    description: 'User directory access',
    requiresOrgMode: true,
  },
  outlook: {
    description: 'Outlook app only: mail, calendar and contacts',
  },
  onedrive: {
    description: 'OneDrive app only: drive and file operations, excluding Excel',
  },
  teams: {
    description: 'Teams app only: chats, channels, meetings and presence',
    requiresOrgMode: true,
  },
};

// Utility tools (graph-tools.ts UTILITY_TOOLS) are code-defined, not in endpoints.json, so they
// carry no `presets` there and every --preset filter dropped them - e.g. `--preset files` had
// get-drive-item but no downloader, though its llmTip tells the model to call one. Declare their
// preset membership here and fold it into each preset pattern.
//
// download-bytes and download-bytes-to-file both read ANY relative Graph binary path (drive files,
// attachments, photos, Teams content, recordings, OneNote resources) - one returns base64, the other
// streams the bytes to a local file - so they belong in EVERY preset. Universal rather than an
// enumerated list because a list would silently miss apps and every future preset.
const UNIVERSAL_UTILITY_TOOLS = ['download-bytes', 'download-bytes-to-file'];

// Scoped utilities are only meaningful where the resources they act on appear.
//
// get-download-url has TWO reach settings, and this list is only the smaller one. With no flag it
// resolves Graph's own @microsoft.graph.downloadUrl, which exists for drive/SharePoint item content
// and nothing else, so it rides with the drive-backed presets; where it is absent the universal
// download-bytes still reads the bytes. Under --enable-attachment-urls it stops being a drive tool
// (see FLAG_UNIVERSAL_UTILITY_TOOLS below) - do not read this list as a statement about what the
// tool can do, only about what it can do unflagged.
//
// parse-teams-url turns any Teams meeting URL (short /meet/, full /meetup-join/, recap ?threadId=)
// into the joinWebUrl that list-online-meetings/get-online-meeting want. Every onlineMeetings
// endpoint in endpoints.json is presets ["teams","work"], and no flag adds a caller for it, so
// teams/work is its whole reach. Rechecked with this change.
const SCOPED_UTILITY_TOOLS: Record<string, string[]> = {
  'get-download-url': ['files', 'onedrive', 'personal', 'work', 'search'],
  'parse-teams-url': ['teams', 'work'],
};

// Utilities whose reach widens to EVERY preset once a runtime flag is on.
//
// --enable-attachment-urls exists precisely for the byte endpoints Graph publishes no
// pre-authenticated URL for: mail and event attachments (mail/outlook/personal), meeting recordings
// and transcripts (teams/work), profile photos, Teams hosted content, and - by the tool's own
// implementation - any other authenticated `/$value` path (users, onenote, contacts, ...). Under the
// flag get-download-url mints for exactly the set download-bytes reads, so it is universal for the
// same reason download-bytes is, and enumerated the same way for the same reason: a list would
// silently miss apps and every future preset.
//
// Scoping it to the drive-backed presets while the flag was on made the feature unreachable on the
// deployments it was built for - `--preset mail,calendar --enable-attachment-urls` started cleanly,
// validated the URL base and key, served the route, and registered no tool that could mint.
export const FLAG_UNIVERSAL_UTILITY_TOOLS: Record<string, keyof PresetToolOptions> = {
  'get-download-url': 'attachmentUrls',
  // read-document (the tool --attachment-proxy exists to serve) is deliberately NOT mapped here
  // yet: that tool does not exist until the task that registers it. Adding the name now would
  // create the exact failure this table exists to prevent, just aimed at itself -- a preset could
  // claim a tool no build of the server actually has. The mapping entry and the tool it gates are
  // added together, in the same change, once the tool exists.
};

/**
 * Runtime flags that change which tools a preset contains. A preset is not a pure function of
 * endpoints.json: a flag can widen a tool's reachable resource set, and the preset has to follow.
 */
export interface PresetToolOptions {
  /** --enable-attachment-urls: get-download-url can mint URLs for any authenticated byte endpoint. */
  attachmentUrls?: boolean;
  /**
   * --attachment-proxy. Declared here so the CLI can validate and thread the flag through ahead of
   * the tool it will gate; it does not yet appear as a value in FLAG_UNIVERSAL_UTILITY_TOOLS, so it
   * currently widens no preset. Deliberately present-but-inert, not dead: a field on an options
   * type is not a claim that a tool exists, the way a FLAG_UNIVERSAL_UTILITY_TOOLS entry would be.
   */
  attachmentProxy?: boolean;
}

// Fail fast if a scoped utility references a preset that does not exist (e.g. a typo like
// 'serach'): otherwise the tool would silently never join the intended preset, with no error.
for (const [tool, presets] of Object.entries(SCOPED_UTILITY_TOOLS)) {
  for (const preset of presets) {
    if (!Object.prototype.hasOwnProperty.call(PRESET_META, preset)) {
      throw new Error(
        `SCOPED_UTILITY_TOOLS["${tool}"] references unknown preset "${preset}" (not in PRESET_META)`
      );
    }
  }
}

function presetPattern(preset: string, options: PresetToolOptions): RegExp {
  const endpointNames = [
    ...new Set(endpointEntries.filter((e) => e.presets?.includes(preset)).map((e) => e.toolName)),
  ];
  // Guard on endpoint membership, not the final `names` list: the universal utility spread below
  // would otherwise mask a preset that no endpoint declares (a typo'd or unwired preset name).
  if (endpointNames.length === 0) {
    throw new Error(`Preset "${preset}" matches no endpoints in endpoints.json`);
  }
  const flagUniversal = Object.entries(FLAG_UNIVERSAL_UTILITY_TOOLS)
    .filter(([, flag]) => options[flag])
    .map(([name]) => name);
  const names = [
    ...endpointNames,
    ...UNIVERSAL_UTILITY_TOOLS,
    ...flagUniversal,
    ...Object.entries(SCOPED_UTILITY_TOOLS)
      .filter(([name, presets]) => !flagUniversal.includes(name) && presets.includes(preset))
      .map(([name]) => name),
  ];
  return new RegExp(`^(?:${names.join('|')})$`);
}

/**
 * Pattern for one preset (or `all`) under a given flag set. Returns undefined for an unknown name
 * so callers that take a user-supplied category can fall through to "no filter" instead of throwing.
 */
export function getCategoryPattern(
  name: string,
  options: PresetToolOptions = {}
): RegExp | undefined {
  if (name === 'all') return /.*/;
  if (!Object.prototype.hasOwnProperty.call(PRESET_META, name)) return undefined;
  return presetPattern(name, options);
}

// The default (no-flag) view of every preset. Used for descriptions, for preset-name validation and
// as the base pattern anywhere the runtime flags are not in scope; anything that knows the flags
// should call getCategoryPattern/getCombinedPresetPattern with them instead.
export const TOOL_CATEGORIES: Record<string, ToolCategory> = {
  ...Object.fromEntries(
    Object.entries(PRESET_META).map(([name, meta]) => [
      name,
      { name, pattern: presetPattern(name, {}), ...meta },
    ])
  ),
  all: {
    name: 'all',
    pattern: /.*/,
    description: 'All available tools',
  },
};

export function getCombinedPresetPattern(
  presets: string[],
  options: PresetToolOptions = {}
): string {
  const patterns = presets.map((preset) => {
    const pattern = getCategoryPattern(preset, options);
    if (!pattern) {
      throw new Error(
        `Unknown preset: ${preset}. Available presets: ${Object.keys(TOOL_CATEGORIES).join(', ')}`
      );
    }
    return pattern.source;
  });
  return patterns.join('|');
}

export function listPresets(): Array<{
  name: string;
  description: string;
  requiresOrgMode?: boolean;
}> {
  return Object.values(TOOL_CATEGORIES).map((category) => ({
    name: category.name,
    description: category.description,
    requiresOrgMode: category.requiresOrgMode,
  }));
}

export function presetRequiresOrgMode(preset: string): boolean {
  const category = TOOL_CATEGORIES[preset];
  return category?.requiresOrgMode || false;
}
