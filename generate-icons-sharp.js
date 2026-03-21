const sharp = require("sharp");
const path = require("path");

const sizes = [16, 32, 64, 80, 128];
const assetsDir = path.join(__dirname, "assets");

async function generate() {
  for (const size of sizes) {
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">
  <rect width="${size}" height="${size}" rx="${Math.round(size * 0.15)}" fill="#0078d4"/>
  <text x="50%" y="55%" dominant-baseline="middle" text-anchor="middle"
        font-family="Segoe UI, Arial, sans-serif" font-weight="700"
        font-size="${Math.round(size * 0.38)}" fill="white">BB</text>
</svg>`;

    const outFile = path.join(assetsDir, `icon-${size}.png`);
    await sharp(Buffer.from(svg)).png().toFile(outFile);
    console.log(`✅ icon-${size}.png`);
  }
  console.log("\n🎨 Fertig – echte PNGs in ./assets/");
}

generate().catch(console.error);
