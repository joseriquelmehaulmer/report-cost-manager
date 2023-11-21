import { app } from '@azure/functions';
import { main } from '../app.js';

app.timer('timerTrigger1', {
  schedule: '0 * * * * *',
  handler: async (myTimer, context) => {
    try {
      await main();
      context.log('Correo electrónico enviado exitosamente');
    } catch (error) {
      context.log('Error al enviar correo electrónico:', error);
    }
  },
});
