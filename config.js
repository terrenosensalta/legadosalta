/**
 * Configuración compartida de la tienda (index.html) y del panel de pedidos (pedidos.html).
 *
 * whatsapp:   número que recibe los pedidos, en formato internacional sin "+" ni espacios.
 *             Para celulares de Argentina va 54 + 9 + característica + número.
 * pedidosUrl: dirección de la aplicación web de Google Apps Script que guarda los pedidos
 *             en la planilla (ver apps-script/Pedidos.gs). Vacío = solo WhatsApp.
 */
window.LEGADO_CONFIG = {
  whatsapp: '5493874121739',
  pedidosUrl: ''
};
