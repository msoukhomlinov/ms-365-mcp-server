import { describe, expect, it } from 'vitest';
import { buildMcpServerInstructions } from '../src/mcp-instructions.js';

describe('buildMcpServerInstructions', () => {
  // A stdio, non-proxy server: every byte tool registered. Stated explicitly rather than derived,
  // because these tests are about the wording; the derivation is covered in
  // proxy-mode-guidance.test.ts.
  const STDIO_TOOLS = new Set([
    'get-download-url',
    'download-bytes',
    'download-bytes-to-file',
    'list-users',
    'upload-file-content',
    'create-upload-session',
  ]);
  const baseCtx = {
    orgMode: true,
    readOnly: false,
    multiAccount: false,
    registeredTools: STDIO_TOOLS,
  };

  it('includes general Graph guidance for standard mode', () => {
    const s = buildMcpServerInstructions({ ...baseCtx, discovery: false });
    expect(s).toContain('Microsoft Graph');
    expect(s).toContain('$filter');
    expect(s).not.toContain('DISCOVERY MODE ADD-ON');
  });

  it('appends discovery addon when discovery is true', () => {
    const s = buildMcpServerInstructions({ ...baseCtx, discovery: true });
    expect(s).toContain('DISCOVERY MODE ADD-ON');
    expect(s).toContain('search-tools');
    expect(s).toContain('$filter');
  });

  it('adds read-only line when readOnly', () => {
    const s = buildMcpServerInstructions({ ...baseCtx, discovery: false, readOnly: true });
    expect(s).toContain('read-only');
  });

  it('does not suggest account switching when multiAccount is false', () => {
    const s = buildMcpServerInstructions({ ...baseCtx, discovery: false, multiAccount: false });
    expect(s).not.toContain('Multiple accounts');
    expect(s).not.toContain('account parameter');
  });

  it('routes drive file downloads to get-download-url and authenticated byte reads to download-bytes', () => {
    const s = buildMcpServerInstructions({ ...baseCtx, discovery: false });
    expect(s).toContain('large drive/SharePoint file content');
    expect(s).toContain('prefer get-download-url');
    expect(s).toContain('download-bytes for authenticated byte reads');
    expect(s).toContain(
      'mail attachments, profile photos, Teams hosted content, and meeting recordings'
    );
    expect(s).toContain('relative Microsoft Graph paths, not absolute URLs');
  });

  it('leaves the text unchanged when every byte tool is registered, or when no set is given', () => {
    const filtered = buildMcpServerInstructions({ ...baseCtx, discovery: false });
    const unfiltered = buildMcpServerInstructions({
      orgMode: true,
      readOnly: false,
      multiAccount: false,
      discovery: false,
    });
    expect(filtered).toBe(unfiltered);
    expect(filtered).toContain('For large mail attachments and meeting recordings that is one');
    expect(filtered).toContain('use create-upload-session for anything but small files');
  });

  it('names read-document and not the byte tools under the proxy', () => {
    const s = buildMcpServerInstructions({
      ...baseCtx,
      discovery: false,
      registeredTools: new Set(['read-document', 'list-users']),
    });
    expect(s).toContain('read-document');
    expect(s).toContain('identity_not_supported');
    expect(s).not.toContain('download-bytes');
    expect(s).not.toContain('get-download-url');
  });

  it('drops the byte paragraph entirely when no byte or document tool is registered', () => {
    const s = buildMcpServerInstructions({
      ...baseCtx,
      discovery: false,
      registeredTools: new Set(['list-users']),
    });
    expect(s).not.toContain('Files / binary content');
    expect(s).not.toContain('upload-file-content');
  });
});
