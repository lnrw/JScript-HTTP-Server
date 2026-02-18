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
