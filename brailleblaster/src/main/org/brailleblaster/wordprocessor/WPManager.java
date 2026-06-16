package org.brailleblaster.wordprocessor;

import org.eclipse.swt.*;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.MessageBox;
import org.brailleblaster.BBIni;

/**
* This is the controller for the whole word processing operation.
*/

public class WPManager {
private Display display;
private SettingsDialogs settings;
private DocumentManager[] documents = new DocumentManager[8];
private int currentDocument = 0;

/**
* Normal word processor entry poinnt.
*/

public WPManager () {
display = BBIni.getDisplay();
if (display == null) {
System.exit (1);
}
documents[currentDocument] = new DocumentManager (display);
}

/**
* Handles text editor, etc.
*/

public WPManager (Object o)
{
TextEditor editor = new TextEditor ();
}

}

