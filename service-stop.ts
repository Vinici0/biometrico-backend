import { Service } from 'node-windows';

// Crea un nuevo objeto de servicio
const svc = new Service({
  name: 'Junta de Agua v2',
  description: 'Junta de Agua v2',
  script: 'C:\\Users\\VINICIO BORJA\\Desktop\\biometrico\\dist\\app.js' // Ruta al archivo compilado
});

// Escucha el evento "install", que indica que el proceso está disponible como servicio
svc.on('uninstall', () => {
  console.log('Encendido');
  svc.start();
});

// Instala el servicio
svc.uninstall();
