package org.brailleblaster;

import org.liblouis.liblouisutdml;
import org.brailleblaster.wordprocessor.WPManager;
import org.brailleblaster.localization.LocaleHandler;
import java.io.IOException;
import org.eclipse.swt.*;
import org.eclipse.swt.widgets.Display;

/**
* This class contains the main method. If there are no arguments it 
* passes control directly to the word processor. If there are arguments 
* it passes them to the constructor of the class Subcommands. It will do more 
* processing as the project develops.
*/

public class Main
{

public static void main (String[] args) {
BBIni initialConditions = BBIni.getInstance();
if (args.length == 0)
new WPManager ();
else
{
new Subcommands (args);
}
Display display = BBIni.getDisplay();
if (display != null)
display.dispose();
if (BBIni.haveLiblouisutdml())
{
liblouisutdml louisutdml = liblouisutdml.getInstance ();
louisutdml.free();
}
}

}
