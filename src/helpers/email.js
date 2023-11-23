import nodemailer from 'nodemailer';
import { getPreviousMonthAndYear } from './dates.js';

export const sendEmailWithAttachment = async (filePath, subscriptionName) => {
  const { month, year } = getPreviousMonthAndYear();
  const transport = nodemailer.createTransport({
    host: process.env.EMAIL_HOST,
    port: process.env.EMAIL_PORT,
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS,
    },
  });

  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: 'jose.riquelme@haulmer.com, rodrigo.verdugo@haulmer.com, frivas@haulmer.com',
    subject: `Reporte de costos suscripción: ${subscriptionName} periodo ${month} ${year}`,
    text: `Adjunto encontrarás el reporte de costos de la suscripción: ${subscriptionName} correspondiente al periodo ${month} ${year}. Este reporte detalla los costos asociados y el uso durante el periodo mencionado. Por favor, revisa el archivo adjunto para obtener información detallada.`,
    html: `
        [CORREO PRUEBA]
        <p>Estimado usuario,</p>
        <p>Te enviamos el reporte de costos de la suscripción: <strong>${subscriptionName}</strong> correspondiente al periodo <strong>${month} ${year}</strong>.</p>
        <p>Este reporte incluye un desglose detallado de los costos y el uso durante el periodo mencionado. Es importante revisar estos datos para comprender mejor el consumo y la distribución de los costos.</p>
        <p>Adjunto a este correo, encontrarás el archivo Excel con toda la información relevante.</p>
        <p>Saludos cordiales,</p>`,
    attachments: [
      {
        path: filePath,
      },
    ],
  };

  try {
    await transport.sendMail(mailOptions);
    console.log('Email sent with attachment');
  } catch (error) {
    console.error('Error sending email:', error);
  }
};
