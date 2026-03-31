/**
 * Adri Invoice Management - Full Dataverse API Setup
 */

const { DeviceCodeCredential } = require('@azure/identity');
const fetch = require('node-fetch');

const ENV_URL  = 'https://org0a2fe1f5.crm.dynamics.com';
const API      = `${ENV_URL}/api/data/v9.2`;
const SCOPE    = `${ENV_URL}/.default`;
const SOL_NAME = 'AdriInvoiceManagement';
const PUB_OPTION_PREFIX = 12608;
const STATUS_OPEN   = PUB_OPTION_PREFIX * 10000;    // 126080000
const STATUS_CLOSED = PUB_OPTION_PREFIX * 10000 + 1; // 126080001

let TOKEN = '';

async function api(method, path, body, sol, noPrefer) {
  const headers = {
    'Authorization': `Bearer ${TOKEN}`,
    'Content-Type': 'application/json',
    'OData-MaxVersion': '4.0',
    'OData-Version': '4.0',
    'Accept': 'application/json',
  };
  if (!noPrefer) headers['Prefer'] = 'return=representation';
  if (sol) headers['MSCRM.SolutionUniqueName'] = SOL_NAME;

  const res = await fetch(`${API}${path}`, {
    method, headers,
    body: body ? JSON.stringify(body) : undefined
  });

  const text = await res.text();
  if (!res.ok) {
    let msg = text;
    try { msg = JSON.parse(text)?.error?.message || text; } catch {}
    const e = new Error(`${method} ${path} → ${res.status}: ${msg.substring(0, 500)}`);
    e.serverMsg = msg;
    throw e;
  }
  if (!text) return null;
  try { return JSON.parse(text); } catch { return text; }
}

function label(text) {
  return {
    '@odata.type': 'Microsoft.Dynamics.CRM.Label',
    LocalizedLabels: [{ '@odata.type': 'Microsoft.Dynamics.CRM.LocalizedLabel', Label: text, LanguageCode: 1033 }],
    UserLocalizedLabel: { '@odata.type': 'Microsoft.Dynamics.CRM.LocalizedLabel', Label: text, LanguageCode: 1033 }
  };
}

function reqLevel(v) {
  return { Value: v, CanBeChanged: true, ManagedPropertyLogicalName: 'canmodifyrequirementlevelsettings' };
}

async function step(name, fn) {
  process.stdout.write(`  ${name}... `);
  try {
    const r = await fn();
    console.log('✓');
    return r;
  } catch (e) {
    console.log('FAILED');
    throw e;
  }
}

async function main() {
  console.log('\n=== Adri Invoice Management Setup ===\n');

  // AUTH
  const cred = new DeviceCodeCredential({
    tenantId: 'common',
    userPromptCallback: ({ verificationUri, userCode }) => {
      console.log('>>> ACTION REQUIRED <<<');
      console.log(`Visit: ${verificationUri}`);
      console.log(`Code:  ${userCode}`);
      console.log('Waiting...\n');
    }
  });
  TOKEN = (await cred.getToken(SCOPE)).token;
  console.log('Authenticated ✓\n');

  // ── 0. CLEAN SLATE: delete broken solution, keep entity ───────────────────
  console.log('[0] Cleaning up...');
  // Known managed system app module IDs — never delete these
  const MANAGED_APP_IDS = new Set([
    '3ce49995-11d9-f011-8406-0022480c4191', // PowerPlatformEnvironmentSettings
    'd530c432-17d9-f011-8406-0022480c4191', // SolutionHealthHub
    '331c1e6d-1bd9-f011-8406-0022480c4191', // PowerPageManagement
    '52f33a00-55f5-f011-8407-000d3a316fbe'  // other managed
  ]);
  // App modules created via API go into the Default solution regardless of MSCRM.SolutionUniqueName header.
  // Clean up any orphaned custom app modules that could block uniquename reuse.
  const defaultSolR = await api('GET', `/solutions?$filter=uniquename eq 'Default'&$select=solutionid`);
  const defaultSolId = defaultSolR?.value?.[0]?.solutionid;
  if (defaultSolId) {
    const orphans = await api('GET', `/solutioncomponents?$filter=_solutionid_value eq ${defaultSolId} and componenttype eq 80&$select=objectid`);
    for (const comp of orphans?.value || []) {
      if (!MANAGED_APP_IDS.has(comp.objectid)) {
        process.stdout.write(`  Deleting orphaned app module (${comp.objectid})... `);
        await api('DELETE', `/appmodules(${comp.objectid})`).then(() => console.log('✓')).catch(() => console.log('skipped'));
      }
    }
  }
  const existing = await api('GET', `/solutions?$filter=uniquename eq '${SOL_NAME}'&$select=solutionid,uniquename`);
  if (existing?.value?.length) {
    const badId = existing.value[0].solutionid;
    await step(`Delete old solution (${badId})`, () =>
      api('DELETE', `/solutions(${badId})`)
    );
  } else {
    console.log('  No existing solution found, skipping.');
  }

  // ── 1. CREATE FRESH SOLUTION ───────────────────────────────────────────────
  console.log('\n[1/6] Creating fresh solution...');
  const pub = await api('GET', `/publishers?$filter=uniquename eq 'Adrien'&$select=publisherid,customizationprefix`);
  if (!pub?.value?.length) throw new Error('Publisher "Adrien" not found');
  const publisherId = pub.value[0].publisherid;
  console.log(`  Publisher: Adrien (prefix: ${pub.value[0].customizationprefix})`);

  await step('Create solution', () => api('POST', '/solutions', {
    uniquename: SOL_NAME,
    friendlyname: 'Adri - Invoice Management',
    version: '1.0.0.0',
    'publisherid@odata.bind': `/publishers(${publisherId})`
  }));

  // ── 2. ENTITY: create if missing, then add to solution ────────────────────
  console.log('\n[2/6] Facture table...');
  let entityMetaId;
  const entityCheck = await api('GET', `/EntityDefinitions?$filter=LogicalName eq 'rb_facture'&$select=MetadataId`);

  if (entityCheck?.value?.length) {
    entityMetaId = entityCheck.value[0].MetadataId;
    console.log(`  rb_facture already exists (${entityMetaId})`);
  } else {
    const created = await step('Create rb_facture', () => api('POST', '/EntityDefinitions', {
      '@odata.type': 'Microsoft.Dynamics.CRM.EntityMetadata',
      SchemaName: 'rb_facture',
      DisplayName: label('Facture'),
      DisplayCollectionName: label('Factures'),
      Description: label('Invoice management table'),
      OwnershipType: 'UserOwned',
      HasNotes: false, HasActivities: false, IsActivity: false,
      PrimaryNameAttribute: 'rb_name',
      Attributes: [{
        '@odata.type': 'Microsoft.Dynamics.CRM.StringAttributeMetadata',
        SchemaName: 'rb_name',
        DisplayName: label('Number'),
        RequiredLevel: reqLevel('ApplicationRequired'),
        MaxLength: 100,
        FormatName: { Value: 'Text' },
        IsPrimaryName: true
      }]
    }));
    entityMetaId = created?.MetadataId;
  }

  // Add entity to the fresh solution
  await step('Add entity to solution', () => api('POST', '/AddSolutionComponent', {
    ComponentId: entityMetaId,
    ComponentType: 1,
    SolutionUniqueName: SOL_NAME,
    AddRequiredComponents: true,
    DoNotIncludeSubcomponents: false
  }));

  // ── 3. COLUMNS ─────────────────────────────────────────────────────────────
  console.log('\n[3/6] Columns...');
  const attrBase = `/EntityDefinitions(LogicalName='rb_facture')/Attributes`;
  const existingAttrs = await api('GET', `${attrBase}?$select=LogicalName&$filter=LogicalName eq 'rb_date' or LogicalName eq 'rb_amount' or LogicalName eq 'rb_status'`);
  const existingNames = new Set((existingAttrs?.value || []).map(a => a.LogicalName));

  for (const [logName, payload] of [
    ['rb_date', {
      '@odata.type': 'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata',
      SchemaName: 'rb_date', DisplayName: label('Date'),
      RequiredLevel: reqLevel('None'), Format: 'DateOnly',
      DateTimeBehavior: { Value: 'DateOnly' }
    }],
    ['rb_amount', {
      '@odata.type': 'Microsoft.Dynamics.CRM.DecimalAttributeMetadata',
      SchemaName: 'rb_amount', DisplayName: label('Amount'),
      RequiredLevel: reqLevel('None'), MinValue: 0, MaxValue: 1000000000, Precision: 2
    }],
    ['rb_status', {
      '@odata.type': 'Microsoft.Dynamics.CRM.PicklistAttributeMetadata',
      SchemaName: 'rb_status', DisplayName: label('Status'),
      RequiredLevel: reqLevel('None'),
      OptionSet: {
        '@odata.type': 'Microsoft.Dynamics.CRM.OptionSetMetadata',
        IsGlobal: false, OptionSetType: 'Picklist',
        Options: [
          { Value: STATUS_OPEN,   Label: label('Open'),   Description: label('') },
          { Value: STATUS_CLOSED, Label: label('Closed'), Description: label('') }
        ]
      }
    }]
  ]) {
    if (existingNames.has(logName)) {
      console.log(`  ${logName}: (already exists)`);
    } else {
      await step(logName, () => api('POST', attrBase, payload));
    }
  }

  await step('Publish entity', () => api('POST', '/PublishXml', {
    ParameterXml: '<importexportxml><entities><entity>rb_facture</entity></entities></importexportxml>'
  }));

  // ── 4. VIEWS ───────────────────────────────────────────────────────────────
  console.log('\n[4/6] Views...');

  // Get the entity ObjectTypeCode (integer) — required by savedquery layoutxml
  const entityMeta = await api('GET', `/EntityDefinitions(LogicalName='rb_facture')?$select=ObjectTypeCode`);
  const otc = entityMeta.ObjectTypeCode;
  console.log(`  Entity ObjectTypeCode: ${otc}`);

  const layout = `<grid name="resultset" object="${otc}" jump="rb_name" select="1" preview="1" icon="1"><row name="result" id="rb_factureid"><cell name="rb_name" width="200"/><cell name="rb_date" width="150"/><cell name="rb_amount" width="150"/><cell name="rb_status" width="150"/></row></grid>`;
  const mkFetch = (f) => `<fetch version="1.0" output-format="xml-platform" mapping="logical"><entity name="rb_facture"><attribute name="rb_name"/><attribute name="rb_date"/><attribute name="rb_amount"/><attribute name="rb_status"/><attribute name="rb_factureid"/><order attribute="rb_name" descending="false"/>${f}</entity></fetch>`;

  const views = [
    ['All Invoices',    true,  ''],
    ['Open Invoices',   false, `<filter type="and"><condition attribute="rb_status" operator="eq" value="${STATUS_OPEN}"/></filter>`],
    ['Closed Invoices', false, `<filter type="and"><condition attribute="rb_status" operator="eq" value="${STATUS_CLOSED}"/></filter>`]
  ];

  for (const [name, isDefault, filter] of views) {
    await step(name, () => api('POST', '/savedqueries', {
      name, returnedtypecode: 'rb_facture', querytype: 0,
      isdefault: isDefault, layoutxml: layout, fetchxml: mkFetch(filter)
    }, true));
  }

  // ── 5. FORM ────────────────────────────────────────────────────────────────
  // When an entity is created via API, Dataverse auto-generates a default main form.
  // We update that form with our layout instead of creating a new one.
  console.log('\n[5/6] Form...');
  const existingForms = await api('GET', `/systemforms?$filter=objecttypecode eq 'rb_facture' and type eq 2&$select=formid,name,formxml`);
  if (existingForms?.value?.length) {
    const existingForm = existingForms.value[0];
    console.log(`  Found existing form: "${existingForm.name}" (${existingForm.formid})`);

    // Parse existing formxml to extract the root <form> wrapper (preserving any schema-required attributes)
    // then inject our two-column layout inside it
    const formXml = `<form><tabs><tab name="tab_general" id="{a1b2c3d4-0010-4000-8000-000000000001}" IsUserDefined="0" showlabel="true" expanded="true"><labels><label description="General" languagecode="1033"/></labels><columns><column width="70%"><sections><section name="section_invoice" showlabel="true" showbar="false" id="{a1b2c3d4-0011-4000-8000-000000000001}" IsUserDefined="0" layout="varwidth" columns="1" labelwidth="115" celllabelalignment="Left" celllabelposition="Left"><labels><label description="Invoice Details" languagecode="1033"/></labels><rows><row><cell id="{a1b2c3d4-0020-4000-8000-000000000001}" showlabel="true"><labels><label description="Number" languagecode="1033"/></labels><control id="rb_name" classid="{4273EDBD-AC1D-40d3-9FB2-095C621B552D}" datafieldname="rb_name" disabled="false"/></cell></row><row><cell id="{a1b2c3d4-0021-4000-8000-000000000001}" showlabel="true"><labels><label description="Date" languagecode="1033"/></labels><control id="rb_date" classid="{5B773807-9FB2-42DB-97C3-7A91EFF8ADFF}" datafieldname="rb_date" disabled="false"/></cell></row><row><cell id="{a1b2c3d4-0022-4000-8000-000000000001}" showlabel="true"><labels><label description="Amount" languagecode="1033"/></labels><control id="rb_amount" classid="{533B9E00-756B-4312-95A0-DC888637AC78}" datafieldname="rb_amount" disabled="false"/></cell></row></rows></section></sections></column><column width="30%"><sections><section name="section_status" showlabel="true" showbar="false" id="{a1b2c3d4-0012-4000-8000-000000000001}" IsUserDefined="0" layout="varwidth" columns="1" labelwidth="115" celllabelalignment="Left" celllabelposition="Left"><labels><label description="Status" languagecode="1033"/></labels><rows><row><cell id="{a1b2c3d4-0023-4000-8000-000000000001}" showlabel="true"><labels><label description="Status" languagecode="1033"/></labels><control id="rb_status" classid="{3EF39988-22BB-4f0b-BBBE-64B5A3748AEE}" datafieldname="rb_status" disabled="false"/></cell></row></rows></section></sections></column></columns></tab></tabs></form>`;

    await step('Update Invoice Information form layout', () => api('PATCH', `/systemforms(${existingForm.formid})`, {
      name: 'Invoice Information',
      formxml: formXml
    }, true));
  } else {
    console.log('  No auto-generated form found — skipping (create manually in portal)');
  }

  // ── 6. MODEL-DRIVEN APP ────────────────────────────────────────────────────
  console.log('\n[6/6] Model-driven app...');

  // Create sitemap first, then link it to the app module
  let sitemapId;
  const existingSitemap = await api('GET', `/sitemaps?$filter=sitemapnameunique eq 'rb_InvoiceManagement_SiteMap'&$select=sitemapid`);
  if (existingSitemap?.value?.length) {
    sitemapId = existingSitemap.value[0].sitemapid;
    console.log(`  Sitemap already exists (${sitemapId})`);
  } else {
    await step('Create sitemap', async () => {
      const r = await api('POST', '/sitemaps', {
        sitemapnameunique: 'rb_InvoiceManagement_SiteMap',
        sitemapname: 'Invoice Management',
        sitemapxml: '<SiteMap><Area Id="area_invoices" ShowGroups="true"><Titles><Title LCID="1033" Title="Invoices"/></Titles><Group Id="group_invoices"><Titles><Title LCID="1033" Title="Invoices"/></Titles><SubArea Id="subarea_factures" Entity="rb_facture"><Titles><Title LCID="1033" Title="Factures"/></Titles></SubArea></Group></Area></SiteMap>'
      }, true);
      sitemapId = r?.sitemapid;
    });
  }

  let appModuleId;
  const existingApp = await api('GET', `/appmodules?$filter=uniquename eq 'rb_InvMgmtApp'&$select=appmoduleid`);
  if (existingApp?.value?.length) {
    appModuleId = existingApp.value[0].appmoduleid;
    console.log(`  App module already exists (${appModuleId})`);
  } else {
    // webresourceid is required as the app icon (SVG) — use the known GearGeneral.svg from this environment
    const webResourceId = '05fea1a2-b17a-ed11-81ad-00224803d8ec';

    await step('Create app module', async () => {
      // Use noPrefer=true: app modules are created in the Default solution and
      // are not visible via standard GET /appmodules collection query
      await api('POST', '/appmodules', {
        name: 'Invoice Management',
        uniquename: 'rb_InvMgmtApp',
        clienttype: 4,
        formfactor: 1,
        isdefault: false,
        isfeatured: false,
        navigationtype: 0,
        webresourceid: webResourceId
      }, true, true);

      // Retrieve by querying Default solution components (type 80 = AppModule), most recent
      const defSol = await api('GET', `/solutions?$filter=uniquename eq 'Default'&$select=solutionid`);
      const defSolId = defSol?.value?.[0]?.solutionid;
      if (!defSolId) throw new Error('Default solution not found');
      const comps = await api('GET', `/solutioncomponents?$filter=_solutionid_value eq ${defSolId} and componenttype eq 80&$select=objectid&$orderby=versionnumber desc&$top=1`);
      appModuleId = comps?.value?.[0]?.objectid;
      if (!appModuleId) throw new Error('App module created but could not find it in Default solution components');
    });
  }

  if (appModuleId && sitemapId) {
    await step('Link sitemap to app', () => api('PUT', `/appmodules(${appModuleId})/sitemap/$ref`, {
      '@odata.id': `${API}/sitemaps(${sitemapId})`
    }));
  }

  console.log(`
╔══════════════════════════════════════════════════╗
║  All done! Open https://make.powerapps.com       ║
╠══════════════════════════════════════════════════╣
║  ✓ Solution : Adri - Invoice Management          ║
║  ✓ Table    : rb_facture (Facture)               ║
║  ✓ Columns  : Number, Date, Amount, Status       ║
║  ✓ Views    : All / Open / Closed Invoices       ║
║  ✓ Form     : details left | status top-right    ║
║  ✓ App      : Invoice Management (model-driven)  ║
╠══════════════════════════════════════════════════╣
║  FLOW (add in Power Automate):                   ║
║  Trigger: Recurrence – daily 07:00 AM            ║
║  List rows: rb_factures                          ║
║    Filter: rb_status eq ${STATUS_OPEN}           ║
║            and rb_date le \@{utcNow()}            ║
║  Apply each → Update row rb_status=${STATUS_CLOSED}║
╚══════════════════════════════════════════════════╝
`);
}

main().catch(e => { console.error('\nFatal:', e.message); process.exit(1); });
