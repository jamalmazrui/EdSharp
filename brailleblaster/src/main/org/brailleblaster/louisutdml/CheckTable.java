package org.brailleblaster.louisutdml;

import org.liblouis.liblouisutdml;

/**
* Checks a liblouis table for correctness.
*/

public class CheckTable
{

public CheckTable (String tableList, String logFile, int mode)
throws LiblouisutdmlException
{
boolean result;
result = liblouisutdml.getInstance().checkTable (tableList, logFile, 
mode);
}

}

