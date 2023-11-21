import dotenv from 'dotenv';
dotenv.config();

import { getBilling, getToken, exportToExcel, sendEmailWithAttachment, findLatestExcelFile } from './helpers/index.js';

export async function main() {
  const subscriptionId = process.env.AZURE_SUBSCRIPTION_ID;
  const bearerToken = await getToken();

  try {
    const response = await getBilling(subscriptionId, bearerToken);
    const billingDetails = response.value;
    const costByResource = {};
    const insignificantCostsByTag = {};
    let totalCost = 0;
    let subscriptionName = '';

    billingDetails.forEach(detail => {
      if (detail.properties) {
        subscriptionName = detail.properties.subscriptionName.toUpperCase();
        const resourceGroup = detail.properties.resourceGroup.toUpperCase();
        const tags = detail.tags
          ? Object.entries(detail.tags)
              .map(([key, value]) => `${key.toUpperCase()}:${value.toUpperCase()}`)
              .join(', ')
          : 'SIN ETIQUETA';
        const instancePath = detail.properties.instanceName.split('/');
        const resourceType = instancePath[instancePath.length - 2].toUpperCase();
        const resourceName = instancePath[instancePath.length - 1].toUpperCase();

        const uniqueKey = `${subscriptionName}|${resourceGroup}|${tags}|${resourceType}|${resourceName}`;
        const cost = detail.properties.costInBillingCurrency || 0;

        totalCost += cost;

        if (resourceType !== 'METRICALERTS' && cost > 0 && cost < 0.01) {
          if (!insignificantCostsByTag[tags]) {
            insignificantCostsByTag[tags] = 0;
          }
          insignificantCostsByTag[tags] += cost;
        } else if (resourceType !== 'METRICALERTS' && cost >= 0.01) {
          if (costByResource[uniqueKey]) {
            costByResource[uniqueKey].Costo += cost;
          } else {
            costByResource[uniqueKey] = {
              Suscripción: subscriptionName,
              'Grupo de recursos': resourceGroup,
              Etiqueta: tags,
              'Tipo de recurso': resourceType,
              Recurso: resourceName,
              Costo: cost,
            };
          }
        }
      }
    });

    const dataForExcel = Object.values(costByResource);

    // Add insignificant expenses per tag
    for (const [tag, cost] of Object.entries(insignificantCostsByTag)) {
      dataForExcel.push({
        Suscripción: subscriptionName,
        'Grupo de recursos': 'VARIOS',
        Etiqueta: tag,
        'Tipo de recurso': 'OTROS',
        Recurso: 'SUMA DE GASTOS INSIGNIFICANTES (< $0.01)',
        Costo: cost,
      });
    }

    // Add the total row
    dataForExcel.push({
      Suscripción: 'TOTAL',
      'Grupo de recursos': '',
      Etiqueta: '',
      'Tipo de recurso': '',
      Recurso: '',
      Costo: totalCost,
    });

    await exportToExcel(dataForExcel, subscriptionName);
    const fileExcelPath = findLatestExcelFile();
    sendEmailWithAttachment(fileExcelPath, subscriptionName);
  } catch (error) {
    console.error('Error fetching usage data:', error);
  }
}

main();
