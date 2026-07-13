// Web.cs -- part of the shared Homer toolkit (namespace Homer).
//
// A small, dependency-free web client that brings the practical ideas of the
// durl.py downloader to C#: a modern TLS handshake, a realistic User-Agent,
// automatic redirect following, filenames taken from the Content-Disposition
// header (including RFC 5987 filename*), an extension guessed from the MIME type
// when the URL has none, sanitized and uniquified output names, and link
// extraction from a page's HTML (so EdSharp no longer needs Internet Explorer to
// gather links). It uses only the .NET base class library.

using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace Homer {

public static class Web {

private static bool bConfigured = false;

public static void configure() {
// Enable modern TLS once. Older .NET defaults negotiate only TLS 1.0, which
// many sites now reject; this is the single change that fixes most HTTPS
// download failures. Certificate validation is intentionally left on.
if (bConfigured) return;
try {
ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls | (SecurityProtocolType)12288;
}
catch {
try { ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12; } catch {}
}
ServicePointManager.DefaultConnectionLimit = 16;
bConfigured = true;
} // configure method

public static string userAgent() {
// A current desktop-browser User-Agent. Some servers serve interstitials or
// block requests that arrive with no agent or an obvious script agent.
return "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36";
} // userAgent method

private static HttpWebRequest makeRequest(string sUrl, string sMethod) {
configure();
HttpWebRequest request = (HttpWebRequest) WebRequest.Create(sUrl);
request.Method = sMethod;
request.UserAgent = userAgent();
request.Accept = "*/*";
request.AllowAutoRedirect = true;
request.MaximumAutomaticRedirections = 16;
request.Timeout = 60000;
request.ReadWriteTimeout = 120000;
request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
return request;
} // makeRequest method

public static string getPage(string sUrl, out string sFinalUrl) {
// Fetch a page as text, following redirects. Returns "" on failure and sets
// sFinalUrl to the post-redirect address (or the original on failure).
sFinalUrl = sUrl;
try {
HttpWebRequest request = makeRequest(sUrl, "GET");
using (HttpWebResponse response = (HttpWebResponse) request.GetResponse()) {
sFinalUrl = response.ResponseUri.ToString();
Encoding en = Encoding.UTF8;
try { if (!string.IsNullOrEmpty(response.CharacterSet)) en = Encoding.GetEncoding(response.CharacterSet); }
catch {}
using (StreamReader reader = new StreamReader(response.GetResponseStream(), en)) {
return reader.ReadToEnd();
}
}
}
catch { return ""; }
} // getPage method

public static List<string[]> getLinks(string sUrl) {
// Parse a page for links the way durl.py does: every href and src attribute,
// resolved against the page's address, plus bare http(s) and www URLs that
// appear in the visible text. Each result is { absoluteUrl, "" } to match the
// shape EdSharp expects (the caller fills in display text from the URL).
List<string[]> lsLinks = new List<string[]>();
HashSet<string> hsSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
string sFinalUrl;
string sHtml = getPage(sUrl, out sFinalUrl);
if (sHtml.Length == 0) return lsLinks;

Uri uriBase = null;
try { uriBase = new Uri(sFinalUrl); } catch {}

MatchCollection mcAttr = Regex.Matches(sHtml, "(?:href|src)\\s*=\\s*(?:\"([^\"]*)\"|'([^']*)'|([^\\s>]+))", RegexOptions.IgnoreCase);
foreach (Match m in mcAttr) {
string sRaw = m.Groups[1].Value;
if (sRaw.Length == 0) sRaw = m.Groups[2].Value;
if (sRaw.Length == 0) sRaw = m.Groups[3].Value;
addLink(lsLinks, hsSeen, uriBase, sRaw, sFinalUrl);
}

foreach (Match m in Regex.Matches(sHtml, "https?://[^\\s\"'<>)\\]]+", RegexOptions.IgnoreCase))
addLink(lsLinks, hsSeen, uriBase, m.Value, sFinalUrl);
foreach (Match m in Regex.Matches(sHtml, "(?<![\\w/])www\\.[^\\s\"'<>)\\]]+", RegexOptions.IgnoreCase))
addLink(lsLinks, hsSeen, uriBase, "http://" + m.Value, sFinalUrl);

return lsLinks;
} // getLinks method

private static void addLink(List<string[]> lsLinks, HashSet<string> hsSeen, Uri uriBase, string sRaw, string sSource) {
if (sRaw == null) return;
sRaw = WebUtility.HtmlDecode(sRaw).Trim();
if (sRaw.Length == 0) return;
if (sRaw.StartsWith("#") || sRaw.StartsWith("javascript:") || sRaw.StartsWith("mailto:") || sRaw.StartsWith("data:")) return;
sRaw = sRaw.TrimEnd('.', ',', ';', ')', ']', '>', '`', '\'', '"');

string sAbs = sRaw;
if (sRaw.IndexOf("://") < 0) {
if (uriBase == null) return;
try { sAbs = new Uri(uriBase, sRaw).ToString(); } catch { return; }
}
if (sAbs == sSource) return;
if (hsSeen.Add(sAbs)) lsLinks.Add(new string[] {sAbs, ""});
} // addLink method

public static string download(string sUrl, string sDir, string sSuggestedName) {
// Download one URL into sDir, following redirects. The saved name comes from
// Content-Disposition when present, otherwise from the suggested name or the
// final URL, and is given an extension guessed from the MIME type when it has
// none. The name is sanitized and made unique. Returns the saved path, or ""
// on failure.
try {
HttpWebRequest request = makeRequest(sUrl, "GET");
using (HttpWebResponse response = (HttpWebResponse) request.GetResponse()) {
string sName = fileFromDisposition(response.Headers["Content-Disposition"]);
if (sName.Length == 0) sName = (sSuggestedName == null) ? "" : sSuggestedName.Trim();
if (sName.Length == 0) {
try { sName = Path.GetFileName(response.ResponseUri.LocalPath); } catch {}
}
if (sName.Length == 0) sName = "download";

string sExt = Path.GetExtension(sName);
if (sExt.Length == 0) {
string sMimeExt = mimeToExt(response.ContentType);
if (sMimeExt.Length > 0) sName = sName + "." + sMimeExt;
}

sName = sanitizeName(sName);
string sPath = uniquePath(Path.Combine(sDir, sName));
using (Stream streamIn = response.GetResponseStream())
using (FileStream streamOut = new FileStream(sPath, FileMode.Create, FileAccess.Write)) {
byte[] abBuffer = new byte[65536];
int iRead;
while ((iRead = streamIn.Read(abBuffer, 0, abBuffer.Length)) > 0) streamOut.Write(abBuffer, 0, iRead);
}
return sPath;
}
}
catch { return ""; }
} // download method

public static string fileFromDisposition(string sHeader) {
// Pull a filename out of a Content-Disposition header. Prefers the RFC 5987
// filename* form (which may be percent-encoded with a charset prefix) and
// falls back to a plain filename. Returns "" when none is present.
if (string.IsNullOrEmpty(sHeader)) return "";
Match mStar = Regex.Match(sHeader, "filename\\*\\s*=\\s*([^;]+)", RegexOptions.IgnoreCase);
if (mStar.Success) {
string sVal = mStar.Groups[1].Value.Trim().Trim('"');
int iTick = sVal.IndexOf("''");
if (iTick >= 0) sVal = sVal.Substring(iTick + 2);
try { sVal = Uri.UnescapeDataString(sVal); } catch {}
sVal = Path.GetFileName(sVal.Trim());
if (sVal.Length > 0) return sVal;
}
Match mPlain = Regex.Match(sHeader, "filename\\s*=\\s*\"?([^\";]+)\"?", RegexOptions.IgnoreCase);
if (mPlain.Success) return Path.GetFileName(mPlain.Groups[1].Value.Trim());
return "";
} // fileFromDisposition method

public static string mimeToExt(string sMime) {
// Map a content type to a file extension for URLs that carry none. Covers the
// common document, archive, image, audio, and video types; returns "" when
// unknown so the caller can leave the name as-is.
if (string.IsNullOrEmpty(sMime)) return "";
int iSemi = sMime.IndexOf(';');
if (iSemi >= 0) sMime = sMime.Substring(0, iSemi);
sMime = sMime.Trim().ToLower();
switch (sMime) {
case "text/html": case "application/xhtml+xml": return "html";
case "text/plain": return "txt";
case "text/markdown": return "md";
case "text/css": return "css";
case "text/csv": return "csv";
case "text/xml": case "application/xml": return "xml";
case "application/json": return "json";
case "application/javascript": case "text/javascript": return "js";
case "application/pdf": return "pdf";
case "application/rtf": case "text/rtf": return "rtf";
case "application/zip": return "zip";
case "application/x-7z-compressed": return "7z";
case "application/x-rar-compressed": case "application/vnd.rar": return "rar";
case "application/gzip": case "application/x-gzip": return "gz";
case "application/x-tar": return "tar";
case "application/epub+zip": return "epub";
case "application/msword": return "doc";
case "application/vnd.openxmlformats-officedocument.wordprocessingml.document": return "docx";
case "application/vnd.ms-excel": return "xls";
case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": return "xlsx";
case "application/vnd.ms-powerpoint": return "ppt";
case "application/vnd.openxmlformats-officedocument.presentationml.presentation": return "pptx";
case "image/jpeg": return "jpg";
case "image/png": return "png";
case "image/gif": return "gif";
case "image/bmp": return "bmp";
case "image/svg+xml": return "svg";
case "image/webp": return "webp";
case "image/tiff": return "tiff";
case "image/x-icon": case "image/vnd.microsoft.icon": return "ico";
case "audio/mpeg": return "mp3";
case "audio/wav": case "audio/x-wav": return "wav";
case "audio/ogg": return "ogg";
case "audio/mp4": case "audio/x-m4a": return "m4a";
case "video/mp4": return "mp4";
case "video/x-msvideo": return "avi";
case "video/quicktime": return "mov";
case "video/webm": return "webm";
case "application/x-msdownload": case "application/octet-stream": return "";
default: return "";
}
} // mimeToExt method

public static string sanitizeName(string sName) {
// Remove characters that are illegal in a Windows filename, collapse the
// resulting separators, and guard against an empty result. Mirrors the
// cleanup durl.py applies before saving.
if (string.IsNullOrEmpty(sName)) return "download";
sName = WebUtility.HtmlDecode(sName);
foreach (char c in Path.GetInvalidFileNameChars()) sName = sName.Replace(c, '_');
foreach (char c in new char[] {'&', '=', '@', '%', '+', '\''}) sName = sName.Replace(c, '_');
while (sName.Contains("__")) sName = sName.Replace("__", "_");
sName = sName.Trim().Trim('_', '.', ' ');
if (sName.Length == 0) sName = "download";
return sName;
} // sanitizeName method

public static string uniquePath(string sPath) {
// Return sPath if free, else insert -001, -002, ... before the extension.
if (!File.Exists(sPath)) return sPath;
string sDir = Path.GetDirectoryName(sPath);
string sRoot = Path.GetFileNameWithoutExtension(sPath);
string sExt = Path.GetExtension(sPath);
for (int i = 1; i < 1000; i++) {
string sTry = Path.Combine(sDir, sRoot + "-" + i.ToString("000") + sExt);
if (!File.Exists(sTry)) return sTry;
}
return sPath;
} // uniquePath method

public static bool downloadTo(string sUrl, string sFile, string sUserName, string sPassword) {
// Download sUrl to the exact path sFile, over HTTP, HTTPS, or FTP.  This replaces
// Microsoft.VisualBasic.Devices.Network.DownloadFile, so no Visual Basic runtime
// is needed and every app in the suite transfers files the same way.  Credentials
// are optional: pass "" for anonymous access.  Returns true on success.
try {
configure();
WebRequest request = WebRequest.Create(sUrl);
if (sUserName != null && sUserName.Length > 0) request.Credentials = new NetworkCredential(sUserName, sPassword == null ? "" : sPassword);
FtpWebRequest ftpRequest = request as FtpWebRequest;
if (ftpRequest != null) {
ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;
ftpRequest.UseBinary = true;
ftpRequest.UsePassive = true;
ftpRequest.KeepAlive = false;
ftpRequest.Timeout = 60000;
ftpRequest.ReadWriteTimeout = 120000;
}
else {
HttpWebRequest httpRequest = request as HttpWebRequest;
if (httpRequest != null) {
httpRequest.UserAgent = userAgent();
httpRequest.Accept = "*/*";
httpRequest.AllowAutoRedirect = true;
httpRequest.MaximumAutomaticRedirections = 16;
httpRequest.Timeout = 60000;
httpRequest.ReadWriteTimeout = 120000;
httpRequest.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
}
}
using (WebResponse response = request.GetResponse())
using (Stream streamIn = response.GetResponseStream())
using (FileStream streamOut = new FileStream(sFile, FileMode.Create, FileAccess.Write)) {
byte[] abBuffer = new byte[65536];
int iRead;
while ((iRead = streamIn.Read(abBuffer, 0, abBuffer.Length)) > 0) streamOut.Write(abBuffer, 0, iRead);
}
return true;
}
catch { return false; }
} // downloadTo method

public static bool uploadFrom(string sFile, string sUrl, string sUserName, string sPassword) {
// Upload sFile to sUrl over FTP (or HTTP PUT for an http/https address).  This
// replaces Microsoft.VisualBasic.Devices.Network.UploadFile.  Returns true on
// success.
try {
configure();
if (!File.Exists(sFile)) return false;
WebRequest request = WebRequest.Create(sUrl);
if (sUserName != null && sUserName.Length > 0) request.Credentials = new NetworkCredential(sUserName, sPassword == null ? "" : sPassword);
FtpWebRequest ftpRequest = request as FtpWebRequest;
if (ftpRequest != null) {
ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;
ftpRequest.UseBinary = true;
ftpRequest.UsePassive = true;
ftpRequest.KeepAlive = false;
ftpRequest.Timeout = 60000;
ftpRequest.ReadWriteTimeout = 120000;
ftpRequest.ContentLength = new FileInfo(sFile).Length;
}
else {
request.Method = "PUT";
HttpWebRequest httpRequest = request as HttpWebRequest;
if (httpRequest != null) {
httpRequest.UserAgent = userAgent();
httpRequest.Timeout = 60000;
httpRequest.ReadWriteTimeout = 120000;
}
}
using (FileStream streamIn = new FileStream(sFile, FileMode.Open, FileAccess.Read))
using (Stream streamOut = request.GetRequestStream()) {
byte[] abBuffer = new byte[65536];
int iRead;
while ((iRead = streamIn.Read(abBuffer, 0, abBuffer.Length)) > 0) streamOut.Write(abBuffer, 0, iRead);
}
using (WebResponse response = request.GetResponse()) { }
return true;
}
catch { return false; }
} // uploadFrom method

public static string nameFromUrl(string sUrl) {
// The file name a URL implies, before any network call: the last path segment,
// percent-decoded.  Returns "" when the URL carries no usable name.
try {
Uri uri = new Uri(sUrl);
string sName = Path.GetFileName(uri.LocalPath);
if (sName.Length == 0) return "";
try { sName = Uri.UnescapeDataString(sName); } catch {}
return sName.Trim();
}
catch { return ""; }
} // nameFromUrl method

public static string suggestedName(string sUrl) {
// The name a download would be given on disk, worked out WITHOUT downloading it.
// This is the durl.py rule, and it is what lets the caller show a target file name
// beside each link, and let the user filter the list by extension:
//
//   1. If the URL already ends in a file name with an extension, use it.  No
//      network call is made, so a page of ordinary links costs nothing.
//   2. Otherwise ask the server with a HEAD request: Content-Disposition gives the
//      name the server recommends, and failing that the MIME type in Content-Type
//      supplies the extension.  This is what rescues links like
//      "example.com/download?id=42", which carry no extension at all.
//   3. If the server will not say, fall back to the bare URL name.
//
// The result is sanitized but NOT made unique: uniqueness is applied at save time,
// against the actual target folder.
string sName = nameFromUrl(sUrl);
if (sName.Length > 0 && Path.GetExtension(sName).Length > 1) return sanitizeName(sName);

try {
configure();
HttpWebRequest request = makeRequest(sUrl, "HEAD");
using (HttpWebResponse response = (HttpWebResponse) request.GetResponse()) {
string sHeaderName = fileFromDisposition(response.Headers["Content-Disposition"]);
if (sHeaderName.Length > 0) return sanitizeName(sHeaderName);
if (sName.Length == 0) {
try { sName = Path.GetFileName(response.ResponseUri.LocalPath); } catch {}
}
if (sName.Length == 0) sName = "download";
if (Path.GetExtension(sName).Length <= 1) {
string sExt = mimeToExt(response.ContentType);
if (sExt.Length > 0) sName = sName + "." + sExt;
}
return sanitizeName(sName);
}
}
catch { }

if (sName.Length == 0) sName = "download";
return sanitizeName(sName);
} // suggestedName method

public static string extensionOf(string sUrl) {
// The extension a download would end up with, lower case and without the dot, or
// "" if it cannot be determined.  Callers use this to offer the user a choice of
// extensions to download, the way durl.py filters links by type.
string sName = suggestedName(sUrl);
string sExt = Path.GetExtension(sName);
if (sExt.Length <= 1) return "";
return sExt.TrimStart('.').ToLower();
} // extensionOf method

} // Web class
} // namespace Homer
