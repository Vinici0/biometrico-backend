"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const node_windows_1 = require("node-windows");
// Crea un nuevo objeto de servicio
const svc = new node_windows_1.Service({
    name: 'Junta de Agua v2..........',
    description: 'Junta de Agua v2........',
    script: "C:\\Users\\VINICIO BORJA\\Desktop\\biometrico\\src\\app.ts", // Ruta al archivo compilado
});
// Escucha el evento "install", que indica que el proceso está disponible como servicio
svc.on('install', () => {
    console.log('Encendido');
    svc.start();
});
// Instala el servicio
svc.install();
