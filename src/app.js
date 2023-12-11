import dotenv from 'dotenv';
dotenv.config();

import {
  deleteExcelFile,
  findExcelFilePath,
  getAllSubscriptions,
  getToken,
  processBillingData,
  sendEmailWithAttachment,
} from './helpers/index.js';

export async function main() {
  const bearerToken = await getToken();
  const subscriptions = await getAllSubscriptions(bearerToken);
  const subscriptionNames = subscriptions.map(subscription => subscription.displayName).join(' - ');

  try {
    for (const subscription of subscriptions) {
      await processBillingData(subscription.subscriptionId, subscription.displayName, bearerToken);
    }
    const fileExcel = findExcelFilePath();
    await sendEmailWithAttachment(fileExcel, subscriptionNames);
    deleteExcelFile(fileExcel);
  } catch (error) {
    console.error('Error: ', error);
  }
}

main();
