// tokenManager.js
// Bluebeam new Developer Portal (developers.bluebeam.com)
// Token endpoint: https://api.bluebeam.com/oauth2/token
// Auth method: Basic (base64 client_id:client_secret in header)
// Required scopes: jobs full_user offline_access
// Do NOT request full_prime scope

const sqlite3 = require('sqlite3').verbose();
const qs      = require('querystring');
const path    = require('path');

class TokenManager {
  constructor(dbPath) {
    const isRender = !!process.env.RENDER;

    const defaultDbPath = isRender
      ? '/tmp/tokens.db'
      : path.join(process.cwd(), 'tokens.db');

    this.dbPath = dbPath || process.env.TOKEN_DB_PATH || defaultDbPath;

    console.log('🧭 Token DB path:', this.dbPath);
    console.log('🧭 isRender:', isRender, '| NODE_ENV:', process.env.NODE_ENV);

    this.clientId     = process.env.BB_CLIENT_ID;
    this.clientSecret = process.env.BB_CLIENT_SECRET;

    if (!this.clientId || !this.clientSecret) {
      throw new Error('❌ Missing BB_CLIENT_ID or BB_CLIENT_SECRET');
    }

    // New Developer Portal endpoint (developers.bluebeam.com apps).
    // Legacy Developer Network apps used authserver.bluebeam.com/auth/token —
    // that endpoint returns Cloudflare 403 from Render's IP range.
    this.TOKEN_URL = 'https://api.bluebeam.com/oauth2/token';

    this.db          = null;
    this.initPromise = this._initDb();
  }

  fetch(...args) {
    return import('node-fetch').then(({ default: fetch }) => fetch(...args));
  }

  // ---------------------------------------------------------------------------
  // SQLite init
  // ---------------------------------------------------------------------------
  async _initDb() {
    return new Promise((resolve, reject) => {
      this.db = new sqlite3.Database(this.dbPath, (err) => {
        if (err) {
          console.error('❌ SQLite failed to open:', this.dbPath);
          return reject(err);
        }
        this.db.run(
          `CREATE TABLE IF NOT EXISTS tokens (
            id            INTEGER PRIMARY KEY,
            refresh_token TEXT,
            access_token  TEXT,
            expires_at    INTEGER
          )`,
          (err) => {
            if (err) reject(err);
            else {
              console.log('✅ Token database ready:', this.dbPath);
              resolve();
            }
          }
        );
      });
    });
  }

  // ---------------------------------------------------------------------------
  // Persist tokens
  // ---------------------------------------------------------------------------
  async saveTokens(accessToken, refreshToken, expiresIn) {
    await this.initPromise;
    const expiresAt = Math.floor(Date.now() / 1000) + expiresIn;

    return new Promise((resolve, reject) => {
      this.db.serialize(() => {
        this.db.run('DELETE FROM tokens');
        this.db.run(
          'INSERT INTO tokens (refresh_token, access_token, expires_at) VALUES (?, ?, ?)',
          [refreshToken, accessToken, expiresAt],
          (err) => {
            if (err) return reject(err);
            console.log(`💾 Tokens saved — expires in ${expiresIn}s`);
            resolve();
          }
        );
      });
    });
  }

  // ---------------------------------------------------------------------------
  // Read stored tokens
  // ---------------------------------------------------------------------------
  async getTokens() {
    await this.initPromise;
    return new Promise((resolve, reject) => {
      this.db.get(
        'SELECT access_token, refresh_token, expires_at FROM tokens LIMIT 1',
        [],
        (err, row) => {
          if (err) reject(err);
          else resolve(row || null);
        }
      );
    });
  }

  // ---------------------------------------------------------------------------
  // Refresh via refresh_token grant
  //
  // New portal requires credentials in a Basic Authorization header
  // (base64 of "client_id:client_secret"), NOT in the request body.
  // Sending credentials in the body causes 401/403 on the new platform.
  // ---------------------------------------------------------------------------
  async refreshAccessToken(refreshToken) {
    console.log(`🔄 Refreshing token via ${this.TOKEN_URL}`);

    // Base64-encode "client_id:client_secret" for Basic auth
    const credentials = Buffer
      .from(`${this.clientId}:${this.clientSecret}`)
      .toString('base64');

    const fetchFn  = await this.fetch;
    const response = await fetchFn(this.TOKEN_URL, {
      method:  'POST',
      headers: {
        'Content-Type':  'application/x-www-form-urlencoded',
        'Authorization': `Basic ${credentials}`
        // Do NOT include client_id / client_secret in the body —
        // credentials belong in the Basic header only.
      },
      body: qs.stringify({
        grant_type:    'refresh_token',
        refresh_token: refreshToken
      })
    });

    const text = await response.text();

    if (!response.ok) {
      console.error('❌ Token refresh failed');
      console.error('   Status:', response.status);
      console.error('   Body:',   text.slice(0, 400)); // truncate HTML error pages
      throw new Error(`Token refresh failed: ${response.status} - ${text.slice(0, 200)}`);
    }

    const data = JSON.parse(text);
    console.log('🔁 Token refreshed successfully');
    return data;
  }

  // ---------------------------------------------------------------------------
  // Get a valid access token — handles bootstrap, cache, and refresh
  // ---------------------------------------------------------------------------
  async getValidAccessToken() {
    await this.initPromise;

    const tokens  = await this.getTokens();
    const nowUnix = Math.floor(Date.now() / 1000);

    // Case 1: No tokens stored — bootstrap from BB_REFRESH_TOKEN env var
    if (!tokens) {
      console.log('⚠️  No tokens stored — bootstrapping from BB_REFRESH_TOKEN...');
      const initialRefresh = process.env.BB_REFRESH_TOKEN;
      if (!initialRefresh)
        throw new Error('❌ No stored tokens and BB_REFRESH_TOKEN not set in env');

      const newTokens = await this.refreshAccessToken(initialRefresh);
      await this.saveTokens(
        newTokens.access_token,
        newTokens.refresh_token,
        newTokens.expires_in
      );
      console.log('🔐 Tokens bootstrapped');
      return newTokens.access_token;
    }

    // Case 2: Access token still valid (with 5-minute buffer)
    if (tokens.expires_at > nowUnix + 300) {
      console.log('✅ Using cached access token');
      return tokens.access_token;
    }

    // Case 3: Expired — refresh using stored refresh token
    console.log('⏳ Access token expired — refreshing...');
    try {
      const newTokens = await this.refreshAccessToken(tokens.refresh_token);
      await this.saveTokens(
        newTokens.access_token,
        newTokens.refresh_token,
        newTokens.expires_in
      );
      console.log('🔐 Token refreshed from stored refresh token');
      return newTokens.access_token;

    } catch (err) {
      // Case 4: Stored refresh token failed — fall back to BB_REFRESH_TOKEN env var
      // This happens when a deploy wipes /tmp/tokens.db and the stored token rotated
      console.error('❌ Stored refresh token failed:', err.message);
      console.log('🔁 Attempting fallback to BB_REFRESH_TOKEN env var...');

      const fallbackRefresh = process.env.BB_REFRESH_TOKEN;
      if (!fallbackRefresh)
        throw new Error('❌ Refresh failed and BB_REFRESH_TOKEN not available for fallback');

      const newTokens = await this.refreshAccessToken(fallbackRefresh);
      await this.saveTokens(
        newTokens.access_token,
        newTokens.refresh_token,
        newTokens.expires_in
      );
      console.log('🔐 Recovered using fallback refresh token');
      return newTokens.access_token;
    }
  }

  close() {
    if (this.db) this.db.close();
  }
}

module.exports = TokenManager;
