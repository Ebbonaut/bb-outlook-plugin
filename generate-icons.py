"""
Generiert echte PNG-Icons für das Outlook Add-in.
Verwendet nur Python-Standardbibliothek (struct + zlib).
"""
import struct
import zlib
import os

ASSETS_DIR = os.path.join(os.path.dirname(__file__), "assets")
os.makedirs(ASSETS_DIR, exist_ok=True)

# BlockBrain-Blau: #0078d4
BG_COLOR = (0, 120, 212)
TEXT_COLOR = (255, 255, 255)

# Minimale 5x7 Bitmap-Font für "B" (1 = weiß, 0 = transparent)
LETTER_B = [
    [1,1,1,0,0],
    [1,0,0,1,0],
    [1,1,1,0,0],
    [1,0,0,1,0],
    [1,0,0,1,0],
    [1,1,1,0,0],
    [0,0,0,0,0],
]

def create_png(width, height, bg, letter_data=None):
    """Erstellt ein PNG als bytes."""
    # Pixel-Daten erzeugen (RGBA)
    rows = []
    
    # Berechne Position für "BB" Text
    letter_w = len(letter_data[0]) if letter_data else 0
    letter_h = len(letter_data) if letter_data else 0
    
    # Skalierung der Buchstaben
    scale = max(1, width // 16)
    total_text_w = (letter_w * 2 + 1) * scale  # "BB" mit 1px Lücke
    total_text_h = letter_h * scale
    
    start_x = (width - total_text_w) // 2
    start_y = (height - total_text_h) // 2
    
    # Abgerundete Ecken (Radius)
    radius = max(2, width // 6)
    
    for y in range(height):
        row = bytearray()
        row.append(0)  # Filter byte: None
        for x in range(width):
            # Abgerundete Ecken prüfen
            in_rect = True
            for (cx, cy) in [(radius, radius), (width-1-radius, radius), 
                             (radius, height-1-radius), (width-1-radius, height-1-radius)]:
                dx = abs(x - cx)
                dy = abs(y - cy)
                if (x < radius or x > width-1-radius) and (y < radius or y > height-1-radius):
                    if dx*dx + dy*dy > radius*radius:
                        in_rect = False
                        break
            
            if not in_rect:
                row.extend([0, 0, 0, 0])  # Transparent
                continue
            
            # Prüfe ob Pixel zum Text gehört
            is_text = False
            if letter_data and scale > 0:
                # Erster Buchstabe "B"
                lx1 = (x - start_x) // scale
                ly = (y - start_y) // scale
                # Zweiter Buchstabe "B"  
                lx2 = (x - start_x - (letter_w + 1) * scale) // scale
                
                if 0 <= ly < letter_h:
                    if 0 <= lx1 < letter_w and letter_data[ly][lx1]:
                        is_text = True
                    elif 0 <= lx2 < letter_w and letter_data[ly][lx2]:
                        is_text = True
            
            if is_text:
                row.extend([*TEXT_COLOR, 255])
            else:
                row.extend([*bg, 255])
        rows.append(bytes(row))
    
    raw_data = b''.join(rows)
    
    # PNG zusammenbauen
    def png_chunk(chunk_type, data):
        chunk = chunk_type + data
        crc = struct.pack('>I', zlib.crc32(chunk) & 0xffffffff)
        return struct.pack('>I', len(data)) + chunk + crc
    
    signature = b'\x89PNG\r\n\x1a\n'
    ihdr = struct.pack('>IIBBBBB', width, height, 8, 6, 0, 0, 0)  # 8bit RGBA
    idat = zlib.compress(raw_data)
    
    return signature + png_chunk(b'IHDR', ihdr) + png_chunk(b'IDAT', idat) + png_chunk(b'IEND', b'')


sizes = [16, 32, 64, 80, 128]

for size in sizes:
    png_data = create_png(size, size, BG_COLOR, LETTER_B)
    filepath = os.path.join(ASSETS_DIR, f"icon-{size}.png")
    with open(filepath, 'wb') as f:
        f.write(png_data)
    print(f"✅ icon-{size}.png ({len(png_data)} bytes)")

print("\n🎨 Echte PNG-Icons generiert!")
