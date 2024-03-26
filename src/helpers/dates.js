import moment from 'moment-timezone';

export function getPreviousMonthDates(previousMonths) {
  const timeZone = 'America/Santiago';
  const todayDate = moment.tz(timeZone);

  // Set to the first day of the previous month
  const startDate = todayDate.subtract(previousMonths, 'months').startOf('month').format('YYYY-MM-DD');

  // Set to the last day of the previous month
  const endDate = todayDate.endOf('month').format('YYYY-MM-DD');

  return { startDate, endDate };
}

export function getPreviousMonthAndYear() {
  const previousMonth = moment().subtract(1, 'months');
  return {
    month: previousMonth.format('MMMM'),
    year: previousMonth.format('YYYY')
  };
}