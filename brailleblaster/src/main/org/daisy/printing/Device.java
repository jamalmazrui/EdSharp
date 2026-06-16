package org.daisy.printing;

import java.io.File;

import javax.print.PrintException;

public interface Device {
	
	public void transmit(File file) throws PrintException;

}
