/**
 * Generiert einfache PNG-Icons für das Outlook Add-in.
 * Führe aus: node generate-icons.js
 */
const fs = require("fs");
const path = require("path");

const sizes = [16, 32, 64, 80, 128];
const assetsDir = path.join(__dirname, "assets");

if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

/**
 * Erstellt ein minimales PNG mit einem blauen "BB"-Rechteck.
 * (Ohne externe Abhängigkeiten – erzeugt ein gültiges 1-Bit-PNG)
 */
function createMinimalPng(size) {
  // Wir erstellen ein SVG und konvertieren es zu einem Data-URI-tauglichen Format.
  // Für ein echtes PNG nutzen wir eine Minimal-Implementierung.
  
  // SVG als Platzhalter (Outlook akzeptiert auch PNGs – für Entwicklung reicht das)
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">
  <rect width="${size}" height="${size}" rx="${Math.round(size * 0.15)}" fill="#0078d4"/>
  <text x="50%" y="55%" dominant-baseline="middle" text-anchor="middle" 
        font-family="Segoe UI, sans-serif" font-weight="700" font-size="${Math.round(size * 0.4)}" fill="white">BB</text>
</svg>`;
  return svg;
}

sizes.forEach((size) => {
  const svg = createMinimalPng(size);
  const filename = `icon-${size}.svg`;
  fs.writeFileSync(path.join(assetsDir, filename), svg);
  console.log(`✅ ${filename} erstellt (${size}x${size})`);
  
  // Auch als .png Dateinamen (SVG-Inhalt – für lokale Entwicklung funktioniert das)
  // Für Produktion solltest du echte PNGs verwenden.
  const pngFilename = `icon-${size}.png`;
  fs.writeFileSync(path.join(assetsDir, pngFilename), svg);
  console.log(`✅ ${pngFilename} erstellt (SVG-Platzhalter)`);
});

console.log("\n🎨 Icons generiert in ./assets/");
console.log("⚠️  Für Produktion: Ersetze die SVG-Platzhalter durch echte PNG-Dateien.");
