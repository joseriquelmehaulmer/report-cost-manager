import { getBilling } from './index.js';

export async function processBillingData(previousMonths, subscription, bearerToken) {
  const subscriptionId = subscription.subscriptionId;
  const subscriptionName = subscription.displayName;

  try {
    const response = await getBilling(previousMonths, subscriptionId, bearerToken);
    const billingDetails = response.value;
    const costByResource = {};
    const insignificantCostsByTag = {};
    let totalCost = 0;

    billingDetails.forEach(detail => {
      if (detail.properties) {
        subscriptionName.toUpperCase();
        const resourceGroup = detail.properties.resourceGroup.toUpperCase();
        const tags = formatTags(detail.tags);
        const instancePath = detail.properties.instanceName.split('/');
        const resourceType = instancePath[instancePath.length - 2].toUpperCase();
        const resourceName = instancePath[instancePath.length - 1].toUpperCase();

        const uniqueKey = `${subscriptionName}|${resourceGroup}|${tags}|${resourceType}|${resourceName}`;
        const cost = detail.properties.costInBillingCurrency || 0;

        totalCost += cost;

        if (resourceType !== 'METRICALERTS') {
          const resourceData = {
            Suscripción: subscriptionName,
            'Grupo de recursos': resourceGroup,
            Etiqueta: tags,
            'Tipo de recurso': resourceType,
            Recurso: resourceName,
            Costo: cost,
          };

          if (cost > 0 && cost < 0.01) {
            handleInsignificantCosts(tags, cost, insignificantCostsByTag);
          } else if (cost >= 0.01) {
            handleRegularCosts(uniqueKey, cost, costByResource, resourceData);
          }
        }
      }
    });

    const dataForExcel = Object.values(costByResource);

    // Add insignificant expenses per tag
    calculateInsignificantCostsByTag(dataForExcel, subscriptionName, insignificantCostsByTag);

    // Add the total row
    dataForExcel.push({
      Suscripción: 'TOTAL',
      'Grupo de recursos': '',
      Etiqueta: '',
      'Tipo de recurso': '',
      Recurso: '',
      Costo: totalCost,
    });
    return dataForExcel;
  } catch (error) {
    console.error('Error fetching usage data:', error);
  }
}

function formatTags(tags) {
  if (!tags || Object.keys(tags).length === 0) {
    return 'SIN ETIQUETA';
  }

  return Object.entries(tags)
    .map(([key, value]) => `${key.toUpperCase()}:${value.toUpperCase()}`)
    .join(', ');
}

function handleInsignificantCosts(tags, cost, insignificantCostsByTag) {
  if (!insignificantCostsByTag[tags]) {
    insignificantCostsByTag[tags] = 0;
  }
  insignificantCostsByTag[tags] += cost;
}

function handleRegularCosts(uniqueKey, cost, costByResource, resourceData) {
  if (costByResource[uniqueKey]) {
    costByResource[uniqueKey].Costo += cost;
  } else {
    costByResource[uniqueKey] = resourceData;
  }
}

function calculateInsignificantCostsByTag(data, subscriptionName, insignificantCostsByTag) {
  for (const [tag, cost] of Object.entries(insignificantCostsByTag)) {
    data.push({
      Suscripción: subscriptionName,
      'Grupo de recursos': 'VARIOS',
      Etiqueta: tag,
      'Tipo de recurso': 'OTROS',
      Recurso: 'SUMA DE GASTOS INSIGNIFICANTES (< $0.01)',
      Costo: cost,
    });
  }
}
