// Genera los iconos de la PWA (icon-192.png, icon-512.png, icon-maskable-512.png)
// a partir de logo-cobra.png, recortando solo el simbolo (sin el texto "cobra")
// y centrandolo sobre un fondo cuadrado blanco.

const sharp = require('sharp');
const path = require('path');

const carpeta = __dirname;
const origen = path.join(carpeta, 'logo-cobra.png');

async function icono(nombre, lienzo, escalaMarca) {
  const marca = await sharp(origen)
    .extract({ left: 0, top: 0, width: 118, height: 124 })
    .resize(Math.round(lienzo * escalaMarca))
    .toBuffer();
  const meta = await sharp(marca).metadata();
  await sharp({
    create: {
      width: lienzo,
      height: lienzo,
      channels: 4,
      background: { r: 255, g: 255, b: 255, alpha: 1 },
    },
  })
    .composite([{ input: marca, left: Math.round((lienzo - meta.width) / 2), top: Math.round((lienzo - meta.height) / 2) }])
    .png()
    .toFile(path.join(carpeta, nombre));
  console.log('Generado:', nombre);
}

async function main() {
  await icono('icon-192.png', 192, 0.68);
  await icono('icon-512.png', 512, 0.68);
  await icono('icon-maskable-512.png', 512, 0.5); // mas padding: zona segura para iconos adaptativos
}

main().catch((err) => {
  console.error('ERROR:', err.message);
  process.exitCode = 1;
});
