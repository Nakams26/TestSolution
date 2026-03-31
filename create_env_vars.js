/**
 * Creates environment variables in Power Platform via Dataverse API
 * Uses client credentials (service principal) for auth
 */

const fetch = require('node-fetch');

const ENV_URL  = process.env.PP_ENVIRONMENT_URL || 'https://org0a2fe1f5.crm.dynamics.com';
const APP_ID   = process.env.PP_APP_ID;
const SECRET   = process.env.PP_CLIENT_SECRET;
const TENANT   = process.env.PP_TENANT_ID;
const API      = `${ENV_URL}/api/data/v9.2`;
const SOL_NAME = 'AdriInvoiceManagement';

let TOKEN = '';

async function getToken() {
  const url = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: APP_ID,
    client_secret: SECRET,
    scope: `${ENV_URL}/.default`,
  });
  const res = await fetch(url, { method: 'POST', body });
  const data = await res.json();
  if (!data.access_token) throw new Error(`Auth failed: ${JSON.stringify(data)}`);
  TOKEN = data.access_token;
  console.log('Authenticated via service principal');
}

async function api(method, path, body) {
  const headers = {
    'Authorization': `Bearer ${TOKEN}`,
    'Content-Type': 'application/json',
    'OData-MaxVersion': '4.0',
    'OData-Version': '4.0',
    'Accept': 'application/json',
    'Prefer': 'return=representation',
    'MSCRM.SolutionUniqueName': SOL_NAME,
  };
  const res = await fetch(`${API}${path}`, {
    method, headers,
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  if (!res.ok) {
    let msg = text;
    try { msg = JSON.parse(text)?.error?.message || text; } catch {}
    throw new Error(`${method} ${path} → ${res.status}: ${msg.substring(0, 500)}`);
  }
  if (text) { try { return JSON.parse(text); } catch { return text; } }
  return null;
}

async function upsertEnvVar(schemaname, displayname) {
  // Check if already exists
  const existing = await api('GET', `/environmentvariabledefinitions?$filter=schemaname eq '${schemaname}'&$select=environmentvariabledefinitionid,schemaname`);
  if (existing?.value?.length > 0) {
    console.log(`  ✓ ${schemaname} already exists (${existing.value[0].environmentvariabledefinitionid})`);
    return;
  }

  await api('POST', '/environmentvariabledefinitions', {
    schemaname,
    displayname,
    type: 100000000,  // String
    isrequired: false,
    introducedversion: '1.0.0.0',
  });
  console.log(`  ✓ Created ${schemaname}`);
}

async function main() {
  await getToken();

  console.log('\nCreating environment variables...');
  await upsertEnvVar('rb_EmailIT', 'EmailIT');

  console.log('\nDone.');
}

main().catch(e => { console.error(e.message); process.exit(1); });
