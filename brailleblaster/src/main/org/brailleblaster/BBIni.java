package org.brailleblaster;

import org.eclipse.swt.*;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.brailleblaster.localization.LocaleHandler;
import org.brailleblaster.util.BrailleblasterPath;
import org.liblouis.liblouisutdml;
import java.lang.UnsatisfiedLinkError;
import java.util.logging.Logger;
import java.util.logging.Level;
import java.util.logging.FileHandler;
import java.io.IOException;
import java.io.File;

/**
* Determine and set initial conditions.
*/

public final class BBIni {

private static BBIni bbini;

public static BBIni getInstance () {
if (bbini == null)
bbini = new BBIni();
  return bbini;
}

private static Logger logger;
private static Display display = null;
private static String brailleblasterPath;
private static String osName;
private static String osVersion;
private static String fileSep;
private static String nativeCommandPath;
private static String nativeLibraryPath;
private static String programDataPath;
private static String nativeCommandSuffix;
private static String nativeLibrarySuffix;
private static String settingsPath;
private static String tempFilesPath;
private static String platformName;
private static boolean hLiblouisutdml = false;
private static FileHandler logFile;

  private BBIni() {
Main m = new Main();
brailleblasterPath = BrailleblasterPath.getPath (m);
osName = System.getProperty ("os.name");
osVersion = System.getProperty ("os.version");
fileSep = System.getProperty ("file.separator");
platformName = SWT.getPlatform();
nativeLibraryPath = brailleblasterPath + fileSep + "native" + fileSep + 
"lib";
if (platformName.equals("win32"))
nativeLibrarySuffix = ".dll";
else if (platformName.equals ("cocoa"))
nativeLibrarySuffix = ".dylib";
else nativeLibrarySuffix = ".so";
programDataPath = brailleblasterPath + fileSep + "programData";
String userHome = System.getProperty ("user.home");
settingsPath = userHome + fileSep + "bbsettings";
File settings = new File (settingsPath);
if (!settings.exists())
settings.mkdir();
tempFilesPath = userHome + fileSep + "bbtemp";
File temps = new File (tempFilesPath);
if (!temps.exists())
temps.mkdir();
logger = Logger.getLogger ("org.brailleblaster");
try {
logFile = new FileHandler 
(tempFilesPath + fileSep + "bblog.xml");
} catch (IOException e) {
logger.log (Level.SEVERE, "cannot open logfile", e);
}
if (logFile != null) {
logger.addHandler (logFile);
}
try {
display = new Display();
} catch (SWTError e) {
logger.log (Level.SEVERE, "Can't find GUI", e);
}
try {
liblouisutdml.loadLibrary (nativeLibraryPath + fileSep + 
"liblouisutdml" + nativeLibrarySuffix);
liblouisutdml louisutdml = liblouisutdml.getInstance();
louisutdml.setDataPath (programDataPath);
louisutdml.setWriteablePath (tempFilesPath);
hLiblouisutdml = true;
} catch (UnsatisfiedLinkError e) {
logger.log (Level.SEVERE, "Problem with liblouisutdml library", e);
}
catch (Exception e) {
logger.log (Level.WARNING, "This shouldn't happen", e);
}
}

public static Display getDisplay()
{
return display;
}

public static boolean haveLiblouisutdml()
{
return hLiblouisutdml;
}

public static String getBrailleblasterPath()
{
return brailleblasterPath;
}

public static String getFileSep()
{
return fileSep;
}

public static String getNativeCommandPath()
{
return nativeCommandPath;
}

public static String getNativeLibraryPath()
{
return nativeLibraryPath;
}

public static String getprogramDataPath()
{
return programDataPath;
}

public static String getNativeCommandSuffix()
{
return nativeCommandSuffix;
}

public static String getNativeLibrarySuffix()
{
return nativeLibrarySuffix;
}

public static String getSettingsPath()
{
return settingsPath;
}

public static String getTempFilesPath () {
return tempFilesPath;
}

public static String getplatformName() {
return platformName;
}

public static Logger getLogger() {
return logger;
}

}

