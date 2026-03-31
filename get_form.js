const { DeviceCodeCredential } = require('@azure/identity');
const fetch = require('node-fetch');
const fs = require('fs');

const ENV_URL = 'https://org0a2fe1f5.crm.dynamics.com';
const API = `${ENV_URL}/api/data/v9.2`;

async function main() {
  const cred = new DeviceCodeCredential({
    tenantId: 'common',
    userPromptCallback: ({ verificationUri, userCode }) => {
      console.log(`Visit: ${verificationUri}  Code: ${userCode}\n`);
    }
  });
  const token = (await cred.getToken(`${ENV_URL}/.default`)).token;

  const res = await fetch(`${API}/systemforms?$filter=objecttypecode eq 'account' and type eq 2&$select=formxml,name&$top=1`, {
    headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/json', 'OData-Version': '4.0' }
  });
  const data = await res.json();
  const form = data.value[0];
  console.log('Form name:', form.name);
  // Write XML to file for inspection
  fs.writeFileSync('reference_form.xml', form.formxml);
  console.log('Written to reference_form.xml');
}

main().catch(console.error);
