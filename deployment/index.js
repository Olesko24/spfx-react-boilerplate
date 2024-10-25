import { SPDefault } from '@pnp/nodejs';
import { spfi } from '@pnp/sp';
import '@pnp/sp/webs/index.js';
import '@pnp/sp/appcatalog/index.js';
import '@pnp/sp/lists/index.js';
import '@pnp/sp/items/index.js';
import '@pnp/sp/folders/index.js';
import { readdir, readFile } from 'fs/promises';

const keyBuffer = await readFile('./key.pem');

const sp = spfi().using(
  SPDefault({
    baseUrl: 'https://tenant.sharepoint.com',
    msal: {
      config: {
        auth: {
          clientId: '<clientId>>',
          authority:
            'https://login.microsoftonline.com/<tenantId>',
          clientCertificate: {
            thumbprint: '{thumbprint from Azure AD}',
            privateKey: keyBuffer.toString()
          }
        }
      },
      scopes: ['https://tenant.sharepoint.com/.default']
    }
  })
);

// get solution file
const solutionFolder = await readdir('../sharepoint/solution');
const solutionFileName = solutionFolder.find(v => v.includes('.sppkg'));
const solutionBuffer = await readFile(
  `../sharepoint/solution/${solutionFileName}`
);
// get tenant app catalog
const tenantCatalogWeb = await sp.getTenantAppCatalogWeb();
// upload app file
const addResult = await tenantCatalogWeb.appcatalog.add(
  solutionFileName,
  solutionBuffer,
  true // overwrite existing app
);

if (addResult) {
  console.log(`Upload app ${solutionFileName} successfull!`);
}