import axios from 'axios';

export async function getToken() {
  try {
    const azureTenantId = process.env.AZURE_TENANT_ID;
    const url = `https://login.microsoftonline.com/${azureTenantId}/oauth2/token`;

    const body = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      resource: 'https://management.azure.com',
    });

    const {
      data: { access_token },
    } = await axios.post(url, body);

    return access_token;
  } catch (error) {
    console.error('Error:', error);
  }
}
