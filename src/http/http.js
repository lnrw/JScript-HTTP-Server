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
