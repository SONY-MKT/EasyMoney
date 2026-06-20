const fs = require('fs');
const https = require('https');
const path = require('path');

const dir = path.join(__dirname, 'public', 'icons');
if (!fs.existsSync(dir)) {
  fs.mkdirSync(dir, { recursive: true });
}

const download = (url, dest) => {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(dest);
    https.get(url, function(response) {
      if (response.statusCode !== 200) {
        reject(new Error(`Failed to get '${url}' (${response.statusCode})`));
        return;
      }
      response.pipe(file);
      file.on('finish', function() {
        file.close(() => resolve());
      });
    }).on('error', function(err) {
      fs.unlink(dest, () => reject(err));
    });
  });
};

async function run() {
  try {
    await download('https://dummyimage.com/192x192/059669/ffffff.png&text=EM', path.join(dir, 'icon-192.png'));
    await download('https://dummyimage.com/512x512/059669/ffffff.png&text=EM', path.join(dir, 'icon-512.png'));
    await download('https://dummyimage.com/100x100/059669/ffffff.png&text=EM', path.join(dir, 'easymoney-logo.png'));
    await download('https://dummyimage.com/32x32/059669/ffffff.png&text=EM', path.join(dir, 'favicon.png'));
    console.log('Icons written successfully.');
  } catch (err) {
    console.error('Error writing icons:', err);
  }
}

run();
