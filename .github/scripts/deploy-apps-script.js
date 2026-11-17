#!/usr/bin/env node
/*
 * Pushes every file in ./apps-script/ to the configured Apps Script project
 * via the Apps Script REST API, using a service-account access token minted
 * by Workload Identity Federation (handed to us in GCP_ACCESS_TOKEN).
 *
 * If DEPLOYMENT_ID is set, also creates a new version and updates that
 * named deployment so the live web-app URL serves the new code.
 *
 * Replaces the old clasp-based deploy — no clasp install, no OAuth
 * refresh tokens, no CLASPRC_JSON secret to rotate monthly.
 */
'use strict';

const fs = require('fs');
const path = require('path');

const SRC_DIR     = 'apps-script';
const SCRIPT_ID   = process.env.APPS_SCRIPT_ID;
const TOKEN       = process.env.GCP_ACCESS_TOKEN;
const DEPLOY_ID   = process.env.DEPLOYMENT_ID || '';
const COMMIT_SHA  = process.env.COMMIT_SHA || '';
const COMMIT_MSG  = process.env.COMMIT_MSG || '';

if (!SCRIPT_ID) { console.error('::error::APPS_SCRIPT_ID secret is not set.'); process.exit(1); }
if (!TOKEN)     { console.error('::error::GCP_ACCESS_TOKEN not populated — auth step failed.'); process.exit(1); }

// Build the files payload. Apps Script REST expects:
//   { files: [{ name, type: SERVER_JS|HTML|JSON, source }, ...] }
// Name is the logical name WITHOUT extension. The single appsscript.json
// manifest is type JSON with the fixed name "appsscript".
const files = [];
for (const entry of fs.readdirSync(SRC_DIR)) {
  const full = path.join(SRC_DIR, entry);
  if (!fs.statSync(full).isFile()) continue;
  const source = fs.readFileSync(full, 'utf8');
  if (entry === 'appsscript.json') {
    files.push({ name: 'appsscript', type: 'JSON', source });
  } else if (entry.endsWith('.gs')) {
    files.push({ name: entry.replace(/\.gs$/, ''), type: 'SERVER_JS', source });
  } else if (entry.endsWith('.html')) {
    files.push({ name: entry.replace(/\.html$/, ''), type: 'HTML', source });
  } else {
    // Skip anything else (README, .gitignore, etc.)
    console.log('Skipping unknown file type: ' + entry);
  }
}
console.log('Prepared ' + files.length + ' files for deploy.');

async function apiCall(method, url, body) {
  const resp = await fetch(url, {
    method,
    headers: {
      Authorization: 'Bearer ' + TOKEN,
      'Content-Type': 'application/json'
    },
    body: body ? JSON.stringify(body) : undefined
  });
  const text = await resp.text();
  if (!resp.ok) {
    throw new Error(method + ' ' + url + ' -> HTTP ' + resp.status + '\n' + text);
  }
  return text ? JSON.parse(text) : {};
}

(async () => {
  // 1) Update the script's content (push the files)
  console.log('Pushing content to script project ' + SCRIPT_ID + '…');
  await apiCall(
    'PUT',
    'https://script.googleapis.com/v1/projects/' + encodeURIComponent(SCRIPT_ID) + '/content',
    { files }
  );
  console.log('✓ Content updated.');

  // 2) If a deployment ID is configured, cut a new version and point the
  // deployment at it. Without this, the web-app URL still serves the old
  // version until someone manually deploys from the Apps Script editor.
  if (!DEPLOY_ID) {
    console.log('ℹ No DEPLOYMENT_ID secret set — skipping deployment update. The live web-app URL will still serve the previous version until that secret is configured or you deploy from the Apps Script editor.');
    return;
  }

  const shortSha = COMMIT_SHA.substring(0, 7);
  const label = ('CI ' + shortSha + ': ' + COMMIT_MSG.split('\n')[0]).substring(0, 80);

  console.log('Creating version "' + label + '"…');
  const version = await apiCall(
    'POST',
    'https://script.googleapis.com/v1/projects/' + encodeURIComponent(SCRIPT_ID) + '/versions',
    { description: label }
  );
  console.log('✓ Created version ' + version.versionNumber);

  console.log('Pointing deployment ' + DEPLOY_ID + ' at new version…');
  await apiCall(
    'PUT',
    'https://script.googleapis.com/v1/projects/' + encodeURIComponent(SCRIPT_ID) + '/deployments/' + encodeURIComponent(DEPLOY_ID),
    {
      deploymentConfig: {
        scriptId: SCRIPT_ID,
        versionNumber: version.versionNumber,
        manifestFileName: 'appsscript',
        description: label
      }
    }
  );
  console.log('✓ Deployment updated — live web-app URL now serves version ' + version.versionNumber + '.');
})().catch(err => {
  console.error('::error::Apps Script deploy failed');
  console.error(err && err.message ? err.message : err);
  process.exit(1);
});
