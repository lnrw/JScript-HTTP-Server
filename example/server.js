// Main object
var http = {};

// Status code corresponding names
var STATUS_NAME = {
	100: "Continue",
	200: "OK",
	206: "Partial Content",
	301: "Moved Permanently",
	302: "Found",
	304: "Not Modified",
	400: "Bad Request",
	401: "Unauthorized",
	403: "Forbidden",
	404: "Not Found",
	500: "Internal Server Error",
	503: "Service Unavailable"
};

var global = global || this;
var WshShell = WScript.CreateObject("WScript.Shell");
var fso = WScript.CreateObject("Scripting.FileSystemObject");

var appDataPath = WshShell.ExpandEnvironmentStrings("%AppData%");
var logFolderPath = fso.BuildPath(appDataPath, "JScript HTTP Server");

if (!fso.FolderExists(logFolderPath)) {
	fso.CreateFolder(logFolderPath);
}

var logFilePath = fso.BuildPath(logFolderPath, "http.log");
var logFile = fso.OpenTextFile(logFilePath, 8, true, -1);

function timeStamp() {
	// Get the current date and time
	var currentDate = new Date();

	// Extract individual components
	var year = currentDate.getFullYear();
	var month = ("0" + currentDate.getMonth() + 1).slice(-2); // Months are 0-based
	var day = ("0" + currentDate.getDate()).slice(-2);
	var hour = ("0" + currentDate.getHours()).slice(-2);
	var minute = ("0" + currentDate.getMinutes()).slice(-2);
	var second = ("0" + currentDate.getSeconds()).slice(-2);
	var millisecond = ("000" + currentDate.getMilliseconds()).slice(-3);

	return year + "-" + month + "-" + day + " " + hour + ":" + minute + ":" + second + "." + millisecond;
}

// Used for debugging
function log(msg) {
	var dateString = timeStamp();

	try {
		logFile.Write("[" + dateString + "] " + msg + "\r\n");
	} catch (err) {}

	WScript.StdOut.Write("[" + dateString + "] " + msg + "\r\n");
}

function logErr(msg) {
	var dateString = timeStamp();

	try {
		logFile.Write("[" + dateString + "] [ERROR] " + msg + "\r\n");
	} catch (err) {}

	WScript.StdErr.Write("[" + dateString + "] [ERROR] " + msg + "\r\n");
}

function logErrObj(err) {
	var msg = err.name + ": " + err.message;
	logErr(msg);
}

// var wmiService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2");
// var oss = wmiService.ExecQuery("SELECT * FROM Win32_OperatingSystem");
// var osVer = oss.ItemIndex(0).Caption + " " + oss.ItemIndex(0).Version + "." + oss.ItemIndex(0).BuildNumber;
// wmiService = null;
// oss = null;

var hostVer = WSH.Name + " " + WSH.Version;
var scriptEngineVer = ScriptEngine() + " " + ScriptEngineMajorVersion() + "." + ScriptEngineMinorVersion() + "." + ScriptEngineBuildVersion();

// Main socket listening process
var listener;

function listen(port) {
	// log("OS: " + osVer);
	log("Script Host: " + hostVer);
	log("Script Engine: " + scriptEngineVer);

	listener = WScript.CreateObject("MSWinsock.Winsock.1", "listener_");

	listener.localPort = port;
	listener.bind();
	listener.listen();

	log("Listening on Port " + port);
}

// Requested connections
var connections = {};

/* Event when a client requests a connection */
// Event[1/n]: WSH event listeners can only be registered globally
var listener_ConnectionRequest = global.listener_ConnectionRequest = function(requestID) {
	log("Connection request " + requestID);
	connections[requestID] = {};
	var connection = connections[requestID].listener =
		WScript.CreateObject("MSWinsock.Winsock", "listener_");

	connection.accept(requestID);
}

/* Event after client data arrives */
var listener_DataArrival = global.listener_DataArrival = function(length) {
	log("Accepted " + length + " bytes " + this.socketHandle);

	//var str = "\0\0\0\0\0";

	//global.listener000 = this;

	var reqData = getSckData(this);

	//this.getData(str, vbString);

	//log("" + reqData);

	//this.sendData("HTTP/1.1 200 OK\r\n\r\n<h1>Hello JScript!</h1>");

	// Process one request
	var request = connections[this.socketHandle].request = new Request(reqData);
	// Create a new response
	var response = connections[this.socketHandle].response = new Response(this);

	log(request.head.join(" "));

	// Process and send response
	dataFactory(request, response);
}

/* Event triggered after each sendData completes */
var listener_SendComplete = global.listener_SendComplete = function() {
	log("Send complete, closing socket " + this.socketHandle);
	this.close();
	delete connections[this.socketHandle];
}

/* Keep running */
function idle() {
	var ret = 0;

	while (ret !== 1) {
		//ret = wshShell.Popup("Running...", 36000);
		WScript.Sleep(1);
	}
}

function getSckData(objWinsock) {
	var sc = new ActiveXObject("MSScriptControl.ScriptControl");
	sc.Language = "VBScript";
	sc.AddObject("objWinsock", objWinsock);
	sc.ExecuteStatement("objWinsock.GetData data, 8");
	var data = sc.CodeObject.data;
	return data;
}

/** Response class: instantiated during data exchange */

// response is the object that contains write
var Response = function(_conn) {
	this.conn = _conn;
	this.headers = "";
	this.body = "";
};

Response.prototype.writeHead = function(statCode, headers) {
	var headerStr = "HTTP/1.1 " + statCode + " " + STATUS_NAME[statCode] + "\r\n";
	log(statCode.toString() + " " + STATUS_NAME[statCode]);

	for (var n in headers) {
		headerStr += n + ": " + headers[n] + "\r\n";
	}

	headerStr += "Connection: close\r\n"; // FORCE CLOSE
	headerStr += "\r\n";

	this.headers = headerStr;
};

Response.prototype.write = function(data) {
	this.body = data;
};

Response.prototype.end = function() {
	var fullResponse;

	if (typeof this.body === "string") {
		fullResponse = this.headers + this.body;
		this.conn.sendData(fullResponse);
	} else {
		// binary case
		this.conn.sendData(this.headers);
		this.conn.sendData(this.body);
	}
};

/** Request class: parses request headers and builds object */
// Class constructor
var Request = function(__) {
	// Private members
	var reqData = {},
		head;

	// Constructor function
	var _Request = function(_reqData) {

		// Split by lines
		_reqData = _reqData.split("\n");
		// Split first line by spaces, remove first line
		head = _reqData.shift().split(" ");
		// Remaining lines become KV array
		var _key, _val;
		for (var n in _reqData) {
			// Check for blank line; below blank line is request body
			if (/^\s*$/.test(_reqData[n])) {
				// Call static method to handle POST data
				//_Request.PostHandle.call(this, n);
				break;
			}
			_reqData[n] = _reqData[n].split(": ");
			_key = _reqData[n].shift();
			_val = _reqData[n].join(": ");
			// Assign to private member reqData
			reqData[_key] = _val;
		}

		this.head = head;
	};

	// Public methods
	// Get request header object
	_Request.prototype.getHeaders = function() {
		return reqData;
	};

	_Request.prototype.path = function() {
		return head[1] || "";
	};

	// exports
	return new _Request(__);
};

/** Create data handling factory and start listening methods */
var dataFactory = function(response) {
	response.end()
};

var createServer = http.createServer = function(_dataFactory) {
	if (typeof dataFactory !== "function") {
		throw new TypeError("dataFactory must be a function.");
	}

	dataFactory = _dataFactory;

	return http;
}

http.listen = function(port) {
	listen(port);
	idle();
	return http;
}
// =============================
// Constants
// =============================

var MEDIA_TYPES = {
	"html": "text/html",
	"htm": "text/html",
	"xhtml": "application/xhtml+xml",
	"xml": "application/xml",
	"css": "text/css",
	"jpg": "image/jpeg",
	"jpeg": "image/jpeg",
	"png": "image/png",
	"gif": "image/gif",
	"ico": "image/vnd.microsoft.icon",
	"webp": "image/webp",
	"bin": "application/octet-stream",
	"exe": "application/vnd.microsoft.portable-executable",
	"txt": "text/plain",
	"js": "text/javascript",
	"json": "application/json",
	"*": "text/plain"
};

var SERVER_STRING = "JScript HTTP Server/0.0.1 (WSH/" + WSH.Version + ")";

var fso = new ActiveXObject("Scripting.FileSystemObject");
var WshShell = new ActiveXObject("WScript.Shell");

var respGenSuccess;

// =============================
// Helper Functions
// =============================

function getFileData(path) {
	var content = "";
	var size = 0;
	try {
		var adoStream = new ActiveXObject("ADODB.Stream");
		adoStream.Open();
		adoStream.Type = 1;
		adoStream.LoadFromFile(path);
		size = adoStream.Size;
		content = adoStream.Read();
	} catch (err) {
		respGenSuccess = false;
		logErrObj(err);
	} finally {
		adoStream.Close();
	}

	return {
		content: content,
		size: size,
		extension: fso.GetExtensionName(path)
	};
}

function countUTF8Bytes(str) {
	var stream = new ActiveXObject("ADODB.Stream");
	stream.Type = 2; // adTypeText
	stream.Charset = "utf-8";
	stream.Open();
	stream.WriteText(str);

	var res = stream.Size - 3;

	stream.Close();
	return res;
}

function ensureCharset(mediaType) {
	if (mediaType.indexOf("text/") === 0 ||
		mediaType === "application/json" ||
		mediaType === "application/xml" ||
		mediaType === "application/xhtml+xml") {

		if (mediaType.indexOf("charset=") === -1) {
			return mediaType + "; charset=utf-8";
		}
	}
	return mediaType;
}

function formatBytes(bytes, decimals) {
	decimals = decimals !== undefined ? decimals : 1;
	if (!+bytes) return "0 B";

	var k = 1024;
	var dm = decimals < 0 ? 0 : decimals;
	var sizes = ["B", "KiB", "MiB", "GiB", "TiB", "PiB", "EiB", "ZiB", "YiB"];
	var i = Math.floor(Math.log(bytes) / Math.log(k));

	return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + " " + sizes[i];
}

function generateErrorPage(status) {
	return "<!DOCTYPE html>\n" +
		"<html lang=\"en-US\" style=\"background-color: #eef; color: #222; font-family: 'Courier New', Courier, monospace; text-align: center;\">\n" +
		"<head>\n" +
		"<meta charset=\"utf-8\" />\n" +
		"<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />\n" +
		"<title>" + status + " " + STATUS_NAME[status] + "</title>\n" +
		"</head>\n" +
		"<body>\n" +
		"<h1>" + status + " " + STATUS_NAME[status] + "</h1>\n" +
		"<hr />\n" +
		"JScript HTTP Server/0.0.1 (WSH/" + WSH.Version + ")\n" +
		"</body>\n" +
		"</html>\n";
}

function generateErrorPageResponse(status) {
	if (this.config.errorPages[status] &&
		fso.FileExists(this.config.errorPages[status])) {

		var fileData = getFileData(this.config.errorPages[status]);

		return {
			body: fileData.content,
			mediaType: "text/html",
			status: status
		};
	}

	return {
		content: generateErrorPage(status),
		mediaType: "text/html",
		status: status
	};
}

function generateDirectoryListing(path, baseDir) {
	var result;
	try {
		var folder = fso.GetFolder(path);
		var files = new Enumerator(folder.Files);
		var subFolders = new Enumerator(folder.SubFolders);

		var relativePath = path.replace(baseDir, "").replace(/\\/g, "");

		result = "<!DOCTYPE html>\n" +
			"<html lang=\"en-US\" style=\"background-color: #eef; color: #222; font-family: 'Courier New', Courier, monospace;\">\n" +
			"<head>\n" +
			"<meta charset=\"utf-8\" />\n" +
			"<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />\n" +
			"<title>Directory Listing of /" + relativePath + "</title>\n" +
			"</head>\n" +
			"<body>\n" +
			"<h2>Directory Listing of /" + relativePath + "</h2>\n" +
			"<table>\n" +
			"<tr><th colspan=\"5\"><hr></th></tr>\n" +
			"<tr><th></th><th>Name</th><th>Last modified</th><th>Size</th><th>Description</th></tr>\n" +
			"<tr><th colspan=\"5\"><hr></th></tr>\n";

		for (; !files.atEnd(); files.moveNext()) {
			var file = files.item();
			result += "<tr><td></td><td><a href=\"" + file.Name + "\">" + file.Name + "</a></td>" +
				"<td>" + new Date(file.DateLastModified).toUTCString() + "</td>" +
				"<td>" + formatBytes(file.Size) + "</td>" +
				"<td>" + file.Type + "</td></tr>\n";
		}

		for (; !subFolders.atEnd(); subFolders.moveNext()) {
			var folderItem = subFolders.item();
			result += "<tr><td></td><td><a href=\"" + folderItem.Name + "/\">" + folderItem.Name + "</a></td>" +
				"<td>" + new Date(folderItem.DateLastModified).toUTCString() + "</td>" +
				"<td>" + formatBytes(folderItem.Size) + "</td>" +
				"<td>" + folderItem.Type + "</td></tr>\n";
		}

		result += "<tr><th colspan=\"5\"><hr></th></tr>\n" +
			"</table>\n" +
			"</body>\n" +
			"</html>\n";
	} catch (err) {
		result = "";
		respGenSuccess = false;
		logErrObj(err);
	} finally {
		return result;
	}
}

// =============================
// Router
// =============================

function Router(config) {
	if (typeof config !== "object" || config === null) {
		this.config = {
			errorPages: {}
		};
	} else {
		this.config = config;
		if (!this.config.errorPages) {
			this.config.errorPages = {};
		}
	}

	this.config.dir = fso.GetAbsolutePathName(
		this.config.dir || WshShell.CurrentDirectory
	);
}

Router.prototype.route = function(request, response) {
	var responseData;

	respGenSuccess = true;

	var realPath = fso.GetAbsolutePathName(
		fso.BuildPath(this.config.dir, request.path().replace(/^\/+/g, ""))
	);

	if (realPath.indexOf(this.config.dir) !== 0 || request.path() === "") {

		responseData = generateErrorPageResponse(400);

	} else if (fso.FileExists(realPath) && request.path().slice(-1) !== "/") {

		var fileData = getFileData(realPath);
		var fSize = fileData.size;

		responseData = {
			content: fSize === 0 ? "" : fileData.content,
			mediaType: MEDIA_TYPES[fileData.extension] || MEDIA_TYPES["*"],
			contentLength: fSize,
			status: 200
		};

	} else if (fso.FolderExists(realPath) && request.path().slice(-1) === "/") {

		var defaultFiles = this.config.defaultPages || ["index.html", "index.htm", "index.txt"];
		var found = false;
		var initialPath = realPath;

		for (var i = 0; i < defaultFiles.length; i++) {
			var testPath = fso.BuildPath(realPath, defaultFiles[i]);
			if (fso.FileExists(testPath)) {
				realPath = testPath;
				found = true;
				break;
			}
		}

		if (found) {

			var fileData = getFileData(realPath);
			var fSize = fileData.size;

			responseData = {
				content: fSize === 0 ? "" : fileData.content,
				mediaType: MEDIA_TYPES[fileData.extension] || MEDIA_TYPES["*"],
				contentLength: fSize,
				status: 200
			};

		} else if (this.config.dirListing !== false) {

			responseData = {
				content: generateDirectoryListing(initialPath, this.config.dir),
				mediaType: "text/html",
				status: 200
			};

		} else {

			responseData = generateErrorPageResponse(403);
		}

	} else {

		responseData = generateErrorPageResponse(404);
	}

	if (typeof responseData.content === "string") {
		responseData.contentLength = countUTF8Bytes(responseData.content);
	}

	if (!respGenSuccess) {

		responseData = {
			content: generateErrorPage(500),
			mediaType: "text/html",
			status: 500
		};

		responseData.contentLength = countUTF8Bytes(responseData.content);

		log("Error during response generation, sending status 500");

	}

	response.writeHead(responseData.status, {
		"Content-Type": ensureCharset(responseData.mediaType),
		"Content-Length": responseData.contentLength,
		"Server": SERVER_STRING
	});

	response.write(responseData.content);
	response.end();
};
var config = {
	dirListing: false,
	dir: WScript.Arguments.Named.Item("dir"),
	defaultPages: ["index.html", "index.htm", "index.txt"]
};

var mainRouter = new Router(config);

http.createServer(function(request, response) {
	mainRouter.route(request, response);
}).listen(8080);
