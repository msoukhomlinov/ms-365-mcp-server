import { describe, expect, it } from 'vitest';
import { buildMcpServerInstructions } from '../src/mcp-instructions.js';

describe('buildMcpServerInstructions', () => {
  // A stdio, non-proxy server: every byte tool registered. Stated explicitly
  // rather than derived, because these tests are about the wording; the
  // derivation itself is covered in proxy-mode-guidance.test.ts.
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

  // These instructions are emitted in every mode, so a flat claim about what
  // get-download-url can reach is false in one of them. Same class rule as
  // attachment-mint-identity: byte endpoints tied to get-download-url must name
  // the flag they depend on.
  it('qualifies get-download-url coverage of authenticated byte endpoints by the flag', () => {
    const s = buildMcpServerInstructions({ ...baseCtx, discovery: false });
    expect(s).toContain('--enable-attachment-urls');
    expect(s).toContain('OAuth, OBO, or bearer mode');
    // The old text asserted an unconditional impossibility.
    expect(s).not.toContain('get-download-url cannot handle');
  });
});
