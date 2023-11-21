import axios from 'axios';
import { getPreviousMonthDates } from '../helpers/dates.js';

export async function getBilling(subscription, token) {
  const { startDate, endDate } = getPreviousMonthDates();

  const baseUrl = `https://management.azure.com/subscriptions/${subscription}/providers/Microsoft.Consumption/usagedetails`;
  const params = `api-version=2023-03-01&startDate=${startDate}&endDate=${endDate}`;
  const url = `${baseUrl}?${params}`;

  // Get all billing details including pagination
  return { value: await getAllBillingDetails(url, token) };
}

async function getAllBillingDetails(url, token, aggregatedData = []) {
  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    // Add the current details to the rollup list
    const currentData = response.data.value;
    aggregatedData = aggregatedData.concat(currentData);

    // Check if a nextLink exists and make a recursive call if necessary
    if (response.data.nextLink) {
      return await getAllBillingDetails(response.data.nextLink, token, aggregatedData);
    } else {
      return aggregatedData;
    }
  } catch (error) {
    console.error('Error al realizar la solicitud:', error.message);
    throw error;
  }
}
