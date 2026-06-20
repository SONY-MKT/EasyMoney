const fs = require('fs');
const path = require('path');
function walk(dir) {
  let results = [];
  const list = fs.readdirSync(dir);
  list.forEach(function(file) {
    file = path.join(dir, file);
    const stat = fs.statSync(file);
    if (stat && stat.isDirectory() && !file.includes('node_modules') && !file.includes('.git')) { 
      results = results.concat(walk(file));
    } else if (stat && stat.isFile()) {
      results.push({
         file,
         mtime: stat.mtime
      });
    }
  });
  return results;
}
const all = walk(__dirname);
const recent = all.filter(f => Date.now() - f.mtime.getTime() < 15 * 60 * 1000);
console.log(recent.map(f => f.file));
