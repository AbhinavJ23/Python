var http = require('http');
var myMod = require('./MyModule.js');

http.createServer(function (req, res) {
  res.writeHead(200, {'Content-Type': 'text/html'});
  res.write("The date and time are currently: " + myMod.myDateTime());
  res.end();
}).listen(8080);