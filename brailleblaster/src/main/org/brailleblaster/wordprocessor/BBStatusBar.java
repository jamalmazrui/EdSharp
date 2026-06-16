package org.brailleblaster.wordprocessor;

import org.eclipse.swt.*;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.layout.FormData;
import org.eclipse.swt.layout.FormAttachment;

public class BBStatusBar {

final private Label statusBar;

public BBStatusBar (Shell documentWindow) {
statusBar = new Label (documentWindow, SWT.BORDER);
statusBar.setText ("This is the status bar.");
FormData location = new FormData();
location.left = new FormAttachment(0);
location.right = new FormAttachment(100);
location.bottom = new FormAttachment(100);
statusBar.setLayoutData (location);
}

}

