// Photographie chaque page HTML préparée par make_photos.py et l'enregistre en JPEG.
// Usage : node shoot.js jobs.json largeur hauteur échelle
const fs = require('fs');
const { chromium } = require('playwright');

(async () => {
  const [jobsFile, width, height, scale] = process.argv.slice(2);
  const jobs = JSON.parse(fs.readFileSync(jobsFile, 'utf8'));
  const browser = await chromium.launch();
  const page = await browser.newPage({
    viewport: { width: Number(width), height: Number(height) },
    deviceScaleFactor: Number(scale),
  });
  for (const job of jobs) {
    await page.goto(job.html, { waitUntil: 'load' });
    await page.evaluate(() => document.fonts.ready);
    await page.screenshot({ path: job.out, type: 'jpeg', quality: 88 });
    console.log('photo :', job.out);
  }
  await browser.close();
})();
