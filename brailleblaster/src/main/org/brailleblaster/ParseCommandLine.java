package org.brailleblaster;

import org.apache.commons.cli.AlreadySelectedException;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.GnuParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.MissingArgumentException;
import org.apache.commons.cli.MissingOptionException;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.OptionBuilder;
import org.apache.commons.cli.OptionGroup;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.commons.cli.Parser;
import org.apache.commons.cli.UnrecognizedOptionException;

/** 
* Parse a command line like thyat used by subversion and Mercurial. 
* Provide mehods to handle options and print help.
*/

public class ParseCommandLine
{
private HelpFormatter help = new HelpFormatter();
private Options options = new Options();
private CommandLine parsedArgs = null;

/**
* Only a single instance is allowed.
*/

private static ParseCommandLine singleInstance = new ParseCommandLine();

public static ParseCommandLine getInstance()
{
return singleInstance;
}

private ParseCommandLine()
{
}

public ParseCommandLine parseCommand (String[] args)
throws IllegalArgumentException
{
if (parsedArgs != null)
{
throw new IllegalArgumentException 
("Attempt to parse command line more than once");
}
options.addOption ("translate", true, "translate a file to braille");
options.addOption ("emboss", true, "translate and emboss a file");
options.addOption ("help", false, "print command description");
CommandLineParser parser = new GnuParser();
try {
parsedArgs = parser.parse (options, args);
} catch (ParseException e)
{
throw new IllegalArgumentException (e.getMessage());
}
return singleInstance;
}

public void printUsage()
{
help.printHelp ("brailleblaster", options, true);
}

public void printHelp()
{
help.printHelp ("brailleblaster", options);
}

public void printSubcommandHelp (String subcommand)
{
}

}

