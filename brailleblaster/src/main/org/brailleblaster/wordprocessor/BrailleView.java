package org.brailleblaster.wordprocessor;

import org.eclipse.swt.*;
import org.eclipse.swt.custom.StyledText;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.layout.FormData;
import org.eclipse.swt.layout.FormAttachment;

public class BrailleView {

private StyledText braille;

public BrailleView (Shell documentWindow) {
braille = new StyledText 
    (documentWindow, SWT.BORDER | SWT.H_SCROLL | SWT.V_SCROLL);
braille.setText ("braille view");
FormData location = new FormData();
location.left = new FormAttachment(0);
location.right = new FormAttachment(100);
location.top = new FormAttachment (15);
location.bottom = new FormAttachment(95);
braille.setLayoutData (location);
}

}

