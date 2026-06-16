using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

class Program {
static void Main() {
foreach (string sFile in Directory.GetFiles(@"c:\edsharp", "Unicode*.txt")) {
Console.WriteLine(sFile);
Console.WriteLine(GetFileEncoding(sFile).EncodingName);
}
return;

Dictionary<string, int> codes = new Dictionary<string, int>();
codes.Add("Unicode (Big-Endian)", 1201);
codes.Add("Unicode (UTF-32 Big-Endian)", 12001);
codes.Add("Unicode (UTF-32)", 12000);
codes.Add("Unicode (UTF-7)", 65000);
codes.Add("Unicode (UTF-8)", 65001);
codes.Add("Unicode", 1200);
string sBody = "";
foreach (string sKey in codes.Keys) {
int iValue = codes[sKey];
Encoding en = Encoding.GetEncoding(iValue);
// Encoding en = Encoding.GetEncoding(sKey);
string sFile = sKey + ".txt";
// Console.WriteLine(sFile);
File.WriteAllText(sFile, sBody, en);
} // foreach
GetBomDictionary();
} // Main method

public static int GetBomHash(string sFile) {
FileStream file = new FileStream(sFile, FileMode.Open, FileAccess.Read, FileShare.Read);
byte[] aBom = new byte[4];
int iCount = file.Read(aBom, 0, 4);
file.Close();
return aBom.GetHashCode();
} // GetBom method

public static Dictionary<int, int> GetBomDictionary() {
Dictionary<string, int> dCodes = new Dictionary<string, int>();
Dictionary<int, int> dBoms = new Dictionary<int, int>();
dCodes.Add("Unicode (Big-Endian)", 1201);
dCodes.Add("Unicode (UTF-32 Big-Endian)", 12001);
dCodes.Add("Unicode (UTF-32)", 12000);
dCodes.Add("Unicode (UTF-7)", 65000);
dCodes.Add("Unicode (UTF-8)", 65001);
dCodes.Add("Unicode", 1200);

string sBody = "";
foreach (string sKey in dCodes.Keys) {
int iValue = dCodes[sKey];
Encoding en = Encoding.GetEncoding(iValue);
// Encoding en = Encoding.GetEncoding(sKey);
// string sFile = sKey + ".txt";
string sFile = Path.GetTempFileName();
// Console.WriteLine(sFile);
// Console.WriteLine(sKey);
File.WriteAllText(sFile, sBody, en);

int iHash = GetBomHash(sFile);
// Console.WriteLine(iHash);
dBoms.Add(iHash, iValue);
File.Delete(sFile);
}
return dBoms;
} // GetBomDictionary method


public static Encoding GetFileEncoding(string sFile) {
Dictionary<int, int> dBom = GetBomDictionary();
foreach (int iKey in dBom.Keys) Console.WriteLine(iKey);
int iHash = GetBomHash(sFile);
Console.WriteLine("hash=" + iHash);
Encoding en = Encoding.Default;
if (dBom.ContainsKey(iHash)) en = Encoding.GetEncoding(dBom[iHash]);
return en;
} // GetFileEncoding method
} // Program class
