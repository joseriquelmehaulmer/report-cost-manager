import dotenv from 'dotenv';
dotenv.config();

import {
  deleteExcelFile,
  findExcelFilePath,
  getAllSubscriptions,
  getToken,
  processBillingData,
  exportToExcel,
  sendEmailWithAttachment,
  compareArraysByResourceAndType,
  updateExcelWithChangesIfExists,
} from './helpers/index.js';

export async function main() {
  const bearerToken = await getToken();
  const subscriptions = await getAllSubscriptions(bearerToken);
  const subscriptionNames = subscriptions.map(subscription => subscription.displayName).join(' - ');

  const dataLastMonthArray = [];
  const dataBeforeLastMonthArray = [];

  const fileExcel = findExcelFilePath();
  if (fileExcel) {
    console.log('Deleting previous excel file...');
    deleteExcelFile(fileExcel);
  }

  try {
    for (const subscription of subscriptions) {
      // Get data from the last month
      console.log(`Processing data for subscription: ${subscription.displayName} - last month`);
      let previousMonth = 1;
      const dataLastMonth = await processBillingData(previousMonth, subscription, bearerToken);
      dataLastMonthArray.push(dataLastMonth);
      await exportToExcel(dataLastMonth, subscription.displayName);

      // Get data from the month before
      console.log(`Processing data for subscription: ${subscription.displayName} - before last month`);
      previousMonth = 2;
      const dataBeforeLastMonth = await processBillingData(previousMonth, subscription, bearerToken);
      dataBeforeLastMonthArray.push(dataBeforeLastMonth);

      previousMonth = 1;
    }
    // Compare data from last month with data from the month before
    const result = compareArraysByResourceAndType(dataBeforeLastMonthArray, dataLastMonthArray);
    await updateExcelWithChangesIfExists(result.deletedElements, result.newElements);

    const fileExcel = findExcelFilePath();
    await sendEmailWithAttachment(fileExcel, subscriptionNames);
    deleteExcelFile(fileExcel);
  } catch (error) {
    console.error('Error: ', error);
  }
}

main();
