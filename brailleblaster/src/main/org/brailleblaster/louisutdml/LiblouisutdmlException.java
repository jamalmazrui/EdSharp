package org.brailleblaster.louisutdml;

/**
* Handles exceptions that other classes in this package throw when a 
* call to a method in liblouisutdml fails.
*/

public class LiblouisutdmlException extends RuntimeException
{

public LiblouisutdmlException ()
{
super();
}

public LiblouisutdmlException (String message)
{
super (message);
}

}

