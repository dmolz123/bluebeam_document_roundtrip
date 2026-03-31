// tokenManager.js
const sqlite3 = require('sqlite3').verbose();
const qs = require('querystring');
const path = require('path');

class TokenManager {
  constructor(dbPath) {
    // Render sets this env var automatically for services
    const isRender = !!process.env.RENDER;

    // Choose a safe writable location
    const defaultDbPath = isRender
      ? '/tmp/tokens.db'
      : path.join(process.cwd(), 'tokens.db');

    // Env var still overrides, but now you can see it clearly in logs
    this.dbPath = dbPath || process.env.TOKEN_DB_PATH || defaultDbPath;

    console.log('🧭 Token DB path selected:', this.dbPath);
    console.log('🧭 isRender:', isRender, 'NODE_ENV:', process.env.NODE_ENV);

    this.clientId = process.env.BB_CLIENT_ID;
    this.clientSecret = process.env.BB_CLIENT_SECRET;

    if (!this.clientId || !this.clientSecret) {
      throw new Error(
        '❌ Missing BB_CLIENT_ID or BB_CLIENT_SECRET in environment variables.'
      );
    }

    this.db = null;
    this.initPromise = this._initDb();
  }

  // ESM-compatible fetch wrapper
  fetch(...args) {
    return import('node-fetch').then(({ default: fetch }) => fetch(...args));
  }

  // ---------------------------------------------------------
  // Initialize SQLite DB
  // ---------------------------------------------------------
  async _initDb() {
    return new Promise((resolve, reject) => {
      this.db = new sqlite3.Database(this.dbPath, (err) => {
        if (err) {
          console.error('❌ SQLite failed to open:', this.dbPath);
          return reject(err);
        }

        this.db.run(
          `
          CREATE TABLE IF NOT EXISTS tokens (
            id INTEGER PRIMARY KEY,
            refresh_token TEXT,
            access_token TEXT,
            expires_at INTEGER
          )
        `,
          (err) => {
            if (err) reject(err);
            else {
              console.log('✅ Token database initialized:', this.dbPath);
              resolve();
            }
          }
        );
      });
    });
  }

  // ---------------------------------------------------------
  // Save tokens (rotating refresh token supported)
  // ---------------------------------------------------------
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
            if (err) reject(err);
            else {
              console.log('💾 Tokens saved:');
              console.log('   • Expires in:', expiresIn, 'sec');
              console.log(
                '   • Access token preview:',
                accessToken?.substring(0, 20),
                '...'
              );
              console.log(
                '   • Refresh token preview:',
                refreshToken?.substring(0, 20),
                '...'
              );
              resolve();
            }
          }
        );
      });
    });
  }

  // ---------------------------------------------------------
  // Read tokens from DB
  // ---------------------------------------------------------
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

  // ---------------------------------------------------------
  // Refresh OAuth tokens using refresh_token
  // ---------------------------------------------------------
  async refreshAccessToken(refreshToken) {
    const tokenUrl = 'https://api.bluebeam.com/oauth2/token';

    const payload = {
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: this.clientId,
      client_secret: this.clientSecret
    };

    console.log(`🔄 Refreshing token via ${tokenUrl}`);

    const fetch = await this.fetch;
    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: qs.stringify(payload)
    });

    const text = await response.text();

    if (!response.ok) {
      console.error('❌ Token refresh failed');
      console.error('   Status:', response.status);
      console.error('   Body:', text);
      throw new Error(`Token refresh failed: ${response.status} - ${text}`);
    }

    const data = JSON.parse(text);
    console.log('🔁 Token refreshed successfully.');

    return data;
  }

  // ---------------------------------------------------------
  // Get valid access token (handles refresh + bootstrap)
  // ---------------------------------------------------------
  async getValidAccessToken() {
    await this.initPromise;

    const tokens = await this.getTokens();
    const nowUnix = Math.floor(Date.now() / 1000);

    // Case 1: No tokens in DB → bootstrap using env refresh token
    if (!tokens) {
      console.log('⚠️ No tokens stored — using BB_REFRESH_TOKEN to bootstrap...');

      const initialRefreshToken = process.env.BB_REFRESH_TOKEN;
      if (!initialRefreshToken) {
        throw new Error('❌ No stored tokens and BB_REFRESH_TOKEN not provided.');
      }

      const newTokens = await this.refreshAccessToken(initialRefreshToken);
      await this.saveTokens(
        newTokens.access_token,
        newTokens.refresh_token,
        newTokens.expires_in
      );

      console.log('🔐 Tokens bootstrapped and saved.');
      return newTokens.access_token;
    }

    // Case 2: Access token still valid
    if (tokens.expires_at > nowUnix + 300) {
      console.log('✅ Using cached access token.');
      return tokens.access_token;
    }

    // Case 3: Access token expired → refresh
    console.log('⏳ Access token expired. Refreshing...');

    try {
      const newTokens = await this.refreshAccessToken(tokens.refresh_token);

      await this.saveTokens(
        newTokens.access_token,
        newTokens.refresh_token,
        newTokens.expires_in
      );

      console.log(`🔐 Access token refreshed. Expires at: ${newTokens.expires_in}s`);
      return newTokens.access_token;
    } catch (err) {
      console.error('❌ Refresh failed. Attempt environmental bootstrap...');

      const fallbackRefresh = process.env.BB_REFRESH_TOKEN;
      if (!fallbackRefresh) {
        throw new Error('❌ Refresh failed and no BB_REFRESH_TOKEN available.');
      }

      const newTokens = await this.refreshAccessToken(fallbackRefresh);
      await this.saveTokens(
        newTokens.access_token,
        newTokens.refresh_token,
        newTokens.expires_in
      );

      console.log('🔐 Recovered from refresh failure using fallback refresh token.');
      return newTokens.access_token;
    }
  }

  close() {
    if (this.db) this.db.close();
  }
}

module.exports = TokenManager;
