/**
 * Playwright-based browser login flow for the Microsoft Loop web app.
 *
 * Ported closely from msoutlook-mcp / msteams-mcp:
 * - Uses the system default browser (Edge on Windows, Chrome on macOS)
 * - Imports Microsoft SSO cookies from the user's real browser profile,
 *   enabling silent authentication without typing credentials
 * - Headless first, visible browser only as last resort
 * - Handles stale SingletonLock files from crashed sessions
 *
 * Loop-specific detail: the Loop SPA may keep MSAL tokens in localStorage OR
 * sessionStorage depending on its cacheLocation, so we gather entries from BOTH
 * stores in-browser before extracting tokens.
 */

import { chromium, type BrowserContext, type Page } from 'playwright';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { execSync } from 'node:child_process';
import { logger } from '../utils/logger.js';
import { LOOP_URL, LOGIN_TIMEOUT_MS } from '../constants.js';
import {
  getBrowserProfileDir,
  writeSessionState,
  writeTokenCache,
  clearSession,
  type TokenCache,
} from './session-store.js';
import { extractTokensFromEntries, type StorageEntry } from './token-extractor.js';
import { importMicrosoftCookies } from '../browser/cookie-import.js';

// ─────────────────────────────────────────────────────────────────────────────
// Browser channel
// ─────────────────────────────────────────────────────────────────────────────

/** Microsoft login domains — redirect to these means we're not authenticated. */
const LOGIN_DOMAINS = [
  'login.microsoftonline.com',
  'login.live.com',
  'login.microsoft.com',
];

/**
 * In-browser snippet (as a string body) that scans BOTH localStorage and
 * sessionStorage for a Loop-relevant MSAL access token. This is the signal that
 * authentication has completed and tokens are available to extract.
 */
function loopAccessTokenPresent(): boolean {
  const stores = [localStorage, sessionStorage];
  for (const store of stores) {
    for (let i = 0; i < store.length; i++) {
      const key = store.key(i) ?? '';
      const lower = key.toLowerCase();
      if (!lower.includes('accesstoken')) continue;
      if (lower.includes('substrate.office.com') || lower.includes('.sharepoint.com') || lower.includes('graph.microsoft.com')) {
        return true;
      }
    }
  }
  return false;
}

/** Gather every localStorage + sessionStorage entry as {name,value} pairs. */
function gatherStorageEntriesInBrowser(): Array<{ name: string; value: string }> {
  const out: Array<{ name: string; value: string }> = [];
  for (const store of [localStorage, sessionStorage]) {
    for (let i = 0; i < store.length; i++) {
      const name = store.key(i);
      if (name === null) continue;
      const value = store.getItem(name);
      if (value !== null) out.push({ name, value });
    }
  }
  return out;
}

function getBrowserChannel(): string {
  const override = process.env.MSLOOP_BROWSER?.trim().toLowerCase();
  if (override && override !== 'chromium' && override !== 'bundled') return override;

  if (process.platform === 'darwin') {
    const detected = getMacOSDefaultBrowser();
    if (detected) {
      logger.debug(`macOS default browser detected: ${detected}`);
      return detected;
    }
    return 'chrome';
  }

  return process.platform === 'win32' ? 'msedge' : 'chrome';
}

function getMacOSDefaultBrowser(): string | undefined {
  try {
    const plistPath = path.join(
      process.env.HOME ?? '',
      'Library/Preferences/com.apple.LaunchServices/com.apple.launchservices.secure.plist',
    );
    const json = execSync(`plutil -convert json -o - "${plistPath}"`, {
      encoding: 'utf8',
      timeout: 3000,
    });
    const data = JSON.parse(json) as { LSHandlers?: Array<{ LSHandlerURLScheme?: string; LSHandlerRoleAll?: string }> };
    const handlers = data.LSHandlers ?? [];
    const httpsHandler = handlers.find(h => h.LSHandlerURLScheme === 'https');
    const bundleId = (httpsHandler?.LSHandlerRoleAll ?? '').toLowerCase();

    if (bundleId.includes('microsoft.edgemac') || bundleId.includes('edge')) return 'msedge';
    if (bundleId.includes('google.chrome') || bundleId.includes('chrome')) return 'chrome';
    return undefined;
  } catch {
    return undefined;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// SingletonLock cleanup
// ─────────────────────────────────────────────────────────────────────────────

function cleanupStaleSingletonLock(profileDir: string): void {
  const lockPath = path.join(profileDir, 'SingletonLock');
  if (!fs.existsSync(lockPath)) return;

  try {
    const linkTarget = fs.readlinkSync(lockPath);
    const match = linkTarget.match(/-(\d+)$/);
    if (match) {
      const pid = parseInt(match[1], 10);
      try {
        process.kill(pid, 0);
        return; // still running — don't remove
      } catch {
        // process is gone — remove stale lock
      }
    }
    fs.unlinkSync(lockPath);
    logger.debug('Removed stale SingletonLock');
  } catch {
    // ignore
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Shared helpers
// ─────────────────────────────────────────────────────────────────────────────

function hasSavedBrowserProfile(): boolean {
  const dir = getBrowserProfileDir();
  return fs.existsSync(dir) && fs.readdirSync(dir).length > 0;
}

async function extractAndCacheTokens(context: BrowserContext, page: Page): Promise<string | null> {
  // Persist the full Playwright storage state for the persistent-profile refresh path.
  const state = await context.storageState();
  writeSessionState(state);

  // Gather tokens from BOTH localStorage and sessionStorage in the live page.
  const entries = await page.evaluate(gatherStorageEntriesInBrowser).catch(() => [] as StorageEntry[]);
  if (entries.length === 0) {
    logger.debug('No storage entries gathered from the Loop page');
    return null;
  }

  const tokens = extractTokensFromEntries(entries);
  if (!tokens) {
    logger.debug('Could not extract Loop tokens from page storage');
    return null;
  }

  const cache: TokenCache = {
    substrateToken: tokens.substrateToken,
    substrateTokenExpiry: tokens.substrateTokenExpiry?.getTime(),
    sharePointToken: tokens.sharePointToken,
    sharePointTokenExpiry: tokens.sharePointTokenExpiry?.getTime(),
    sharePointResource: tokens.sharePointResource,
    graphToken: tokens.graphToken,
    graphTokenExpiry: tokens.graphTokenExpiry?.getTime(),
    refreshToken: tokens.refreshToken,
    tenantId: tokens.tenantId,
    upn: tokens.upn,
    extractedAt: Date.now(),
  };
  writeTokenCache(cache);
  return tokens.upn ?? 'unknown';
}

async function launchContext(profileDir: string, headless: boolean, channel: string): Promise<BrowserContext> {
  cleanupStaleSingletonLock(profileDir);

  const launch = () => chromium.launchPersistentContext(profileDir, {
    headless,
    channel,
    viewport: { width: 1280, height: 800 },
    acceptDownloads: false,
  });

  try {
    return await launch();
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    if (msg.includes('ProcessSingleton') || msg.includes('SingletonLock')) {
      logger.debug('Profile lock conflict — removing lock and retrying');
      const lockPath = path.join(profileDir, 'SingletonLock');
      try { fs.unlinkSync(lockPath); } catch { /* ignore */ }
      return launch();
    }
    throw err;
  }
}

/**
 * Wait for Loop to authenticate AND for MSAL tokens to be present in storage.
 * Uses redirect-detection for fast failure on unauthenticated sessions.
 *
 * @returns true if authenticated with tokens, false if redirected to login
 */
async function waitForLoopAuth(context: BrowserContext, timeoutMs: number): Promise<boolean> {
  const page = context.pages()[0] ?? await context.newPage();

  let redirectedToLogin = false;
  const onNavigation = (frame: { url: () => string }) => {
    if (frame === page.mainFrame() && LOGIN_DOMAINS.some(d => frame.url().includes(d))) {
      redirectedToLogin = true;
    }
  };
  page.on('framenavigated', onNavigation);

  try {
    await page.goto(LOOP_URL, { waitUntil: 'domcontentloaded' });

    const deadline = Date.now() + Math.min(timeoutMs, 5_000);
    while (Date.now() < deadline) {
      if (redirectedToLogin) return false;
      await page.waitForTimeout(100);
    }
    if (redirectedToLogin) return false;

    await page.waitForFunction(loopAccessTokenPresent, { timeout: timeoutMs });
    return true;
  } finally {
    page.off('framenavigated', onNavigation);
  }
}

/**
 * Poll storage in-browser until Loop-relevant MSAL tokens appear. Handles the
 * case where session cookies are alive but the short-lived access token expired
 * and the Loop JS is silently re-acquiring it.
 */
async function waitForMsalTokens(context: BrowserContext, timeoutMs = 20_000): Promise<boolean> {
  const page = context.pages()[0];
  if (!page) return false;

  logger.debug('Waiting for MSAL to acquire tokens...');
  const interval = 1_000;
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    const ready = await page.evaluate(loopAccessTokenPresent).catch(() => false);
    if (ready) return true;
    await page.waitForTimeout(interval);
  }

  logger.debug('MSAL token wait timed out');
  return false;
}

// ─────────────────────────────────────────────────────────────────────────────
// Headless login (primary)
// ─────────────────────────────────────────────────────────────────────────────

export async function headlessLogin(): Promise<LoginResult | null> {
  logger.debug('Attempting headless (silent) login...');

  const profileDir = getBrowserProfileDir();
  const channel = getBrowserChannel();
  let context: BrowserContext | null = null;

  try {
    context = await launchContext(profileDir, true, channel);

    const authenticated = await waitForLoopAuth(context, 5_000);
    if (!authenticated) {
      logger.debug('Headless: session invalid or expired (login redirect detected)');
      return null;
    }

    const tokensReady = await waitForMsalTokens(context, 20_000);
    if (!tokensReady) {
      logger.debug('Headless: MSAL tokens did not appear — falling back to headed login');
      return null;
    }

    const page = context.pages()[0];
    if (!page) return null;
    const upn = await extractAndCacheTokens(context, page);
    if (!upn) return null;
    logger.info(`Headless login succeeded (${upn})`);
    return { upn, method: 'headless-sso' };
  } catch (err) {
    logger.debug('Headless login failed', err instanceof Error ? err.message : String(err));
    return null;
  } finally {
    await context?.close().catch(() => {});
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Headed login (fallback)
// ─────────────────────────────────────────────────────────────────────────────

async function showOverlay(page: Page, phase: 'pending' | 'saving' | 'done' | 'error'): Promise<void> {
  const phases = {
    pending: { icon: '⋯', title: "You're signed in!", detail: 'Setting up your Loop connection...',  bg: '#5b5fc7' },
    saving:  { icon: '⋯', title: 'Saving your session...',   detail: "So you won't need to log in again.", bg: '#5b5fc7' },
    done:    { icon: '✓', title: 'All done!',                 detail: 'This window will close automatically.', bg: '#107c10' },
    error:   { icon: '✕', title: 'Something went wrong',      detail: 'Please try again.',                 bg: '#c42b1c' },
  };
  const p = phases[phase];
  try {
    await page.evaluate(({ icon, title, detail, bg }) => {
      const existing = document.getElementById('msloop-mcp-overlay');
      if (existing) existing.remove();
      const overlay = document.createElement('div');
      overlay.id = 'msloop-mcp-overlay';
      Object.assign(overlay.style, { position: 'fixed', inset: '0', background: 'rgba(0,0,0,.65)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: '999999', fontFamily: "'Segoe UI',system-ui,sans-serif" });
      overlay.innerHTML = `<div style="background:white;border-radius:12px;padding:40px 48px;max-width:420px;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,.3)"><div style="width:64px;height:64px;border-radius:50%;background:${bg};color:white;font-size:32px;display:flex;align-items:center;justify-content:center;margin:0 auto 24px">${icon}</div><h2 style="margin:0 0 12px;font-size:20px;font-weight:600;color:#242424">${title}</h2><p style="margin:0;font-size:14px;color:#616161;line-height:1.5">${detail}</p></div>`;
      document.body.appendChild(overlay);
    }, p);
  } catch { /* cosmetic — ignore */ }
}

export async function headedLogin(clearCookiesFirst = false): Promise<LoginResult | null> {
  logger.info('Opening browser for interactive login...');

  const profileDir = getBrowserProfileDir();
  const channel = getBrowserChannel();
  let context: BrowserContext | null = null;
  let browserClosed = false;

  try {
    context = await launchContext(profileDir, false, channel);
    context.on('close', () => { browserClosed = true; });

    if (clearCookiesFirst) {
      await context.clearCookies();
      logger.debug('Browser cookies cleared for force_new login');
    }

    await importMicrosoftCookies(context, channel);

    const authenticated = await waitForLoopAuth(context, LOGIN_TIMEOUT_MS);
    const page = context.pages()[0];

    if (!authenticated) {
      logger.info('Waiting for you to complete sign-in in the browser...');
      if (page) {
        await page.waitForFunction(loopAccessTokenPresent, { timeout: LOGIN_TIMEOUT_MS });
      }
    }

    if (page) {
      await showOverlay(page, 'pending');
      await page.waitForTimeout(800);
      await showOverlay(page, 'saving');
    }

    if (!page) return null;
    const upn = await extractAndCacheTokens(context, page);
    if (!upn) {
      await showOverlay(page, 'error');
      return null;
    }

    await showOverlay(page, 'done');
    await page.waitForTimeout(1200);

    logger.info(`Headed login succeeded (${upn}). Browser closing.`);
    return { upn, method: 'headed-browser' };
  } catch (err) {
    if (browserClosed) {
      logger.error('Login aborted — browser was closed before authentication completed.');
    } else {
      logger.error('Headed login failed', err instanceof Error ? err.message : String(err));
    }
    return null;
  } finally {
    if (context && !browserClosed) {
      await context.close().catch(() => {});
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Result type
// ─────────────────────────────────────────────────────────────────────────────

export interface LoginResult {
  upn: string;
  method: 'token-cache' | 'headless-sso' | 'headed-browser';
}

// ─────────────────────────────────────────────────────────────────────────────
// Main entry point
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Login to Microsoft Loop.
 *
 * @param forceNew - When true, clear the saved session and force a full re-auth.
 */
export async function browserLogin(forceNew = false): Promise<LoginResult | null> {
  if (forceNew) {
    clearSession();
    logger.info('Forced re-login — cleared previous session and token cache.');
  }

  if (!forceNew && hasSavedBrowserProfile()) {
    const result = await headlessLogin();
    if (result) return result;
    logger.info('Headless login failed — falling back to visible browser...');
  } else if (!forceNew) {
    logger.info('No saved browser profile — opening visible browser for first-time setup...');
  }

  return headedLogin(forceNew);
}

// ─────────────────────────────────────────────────────────────────────────────
// Token refresh alias
// ─────────────────────────────────────────────────────────────────────────────

export async function headlessTokenRefresh(): Promise<boolean> {
  return (await headlessLogin()) !== null;
}
