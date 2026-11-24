const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const candidatesDir = 'C:/Users/USER/.gemini/antigravity/scratch/municipio-escolar/public/assets/candidates';

function createCropHtml(canvasId, src, sx, sy, sSize, outName) {
  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    * { margin: 0; padding: 0; }
    body { background: transparent; }
    canvas { display: block; }
  </style>
</head>
<body>
  <canvas id="c" width="320" height="320"></canvas>
  <script>
    const img = new Image();
    img.onload = () => {
      const c = document.getElementById("c");
      const ctx = c.getContext("2d");
      ctx.beginPath();
      ctx.arc(160, 160, 156, 0, Math.PI * 2);
      ctx.closePath();
      ctx.clip();
      ctx.drawImage(img, ${sx}, ${sy}, ${sSize}, ${sSize}, 0, 0, 320, 320);
      window.__DONE__ = true;
    };
    img.src = "${src}";
  </script>
</body>
</html>`;

  const htmlPath = path.join(candidatesDir, `crop_${outName}.html`);
  fs.writeFileSync(htmlPath, html, 'utf8');

  const edge = '"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"';
  const outPng = path.join(candidatesDir, `${outName}.png`);
  const fileUrl = 'file:///' + htmlPath.replace(/\\/g, '/');

  execSync(`${edge} --headless --disable-gpu --window-size=320,320 --screenshot="${outPng}" "${fileUrl}"`);
  console.log(`Saved ${outName}.png, size:`, fs.statSync(outPng).size);
}

// Lista 1: 768x1024 -> logo circle is around x: 254, y: 185, size: 260
createCropHtml(
  'c1',
  'file:///' + path.join(candidatesDir, 'card_lista_1.jpg').replace(/\\/g, '/'),
  254,
  185,
  260,
  'logo_lista_1'
);

// Lista 2: 914x1024 -> compass logo is around x: 305, y: 230, size: 304
createCropHtml(
  'c2',
  'file:///' + path.join(candidatesDir, 'card_lista_2.jpg').replace(/\\/g, '/'),
  305,
  230,
  304,
  'logo_lista_2'
);

// Lista 3: 720x731 -> round logo is around x: 240, y: 205, size: 240
createCropHtml(
  'c3',
  'file:///' + path.join(candidatesDir, 'card_lista_3.jpg').replace(/\\/g, '/'),
  240,
  205,
  240,
  'logo_lista_3'
);

console.log('All 3 logos cropped successfully!');
