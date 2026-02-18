var config = {
	dirListing: false,
	dir: WScript.Arguments.Named.Item("dir"),
	defaultPages: ["index.html", "index.htm", "index.txt"]
};

var mainRouter = new Router(config);

http.createServer(function(request, response) {
	mainRouter.route(request, response);
}).listen(8080);
