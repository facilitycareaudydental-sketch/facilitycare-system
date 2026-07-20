const fs = require('fs');
const html = fs.readFileSync('htmlview_full.html', 'utf8');
const menuStart = html.indexOf('sheet-menu');
if (menuStart > -1) {
  const menuChunk = html.substring(menuStart, menuStart + 2000);
  console.log(menuChunk);
} else {
  // Let's just find anything with gid=
  const gidRegex = /gid=\d+/g;
  const matches = html.match(gidRegex);
  if (matches) {
    console.log("Found gids:", new Set(matches));
  } else {
    console.log("No gids found.");
  }
}
