var fs = require('fs');
var src = fs.readFileSync('schedule_intelligence/build.js').toString();
var lines = src.split('\n');
var sq = "'";
var bs = '\\';
var pat = sq + bs + 'n';
lines.forEach(function(l, i) {
  if (i >= 1170 && l.indexOf(pat) >= 0) console.log(i+1, l.trim().slice(0, 120));
});
