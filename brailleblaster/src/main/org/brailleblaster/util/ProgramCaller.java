package org.brailleblaster.util;

import org.apache.commons.exec.*;
import java.io.IOException;
import org.brailleblaster.BBIni;

/**
* This class calls any program in a safe way
*/
public class ProgramCaller
{

private CommandLine cmdLine;
private DefaultExecuteResultHandler resultHandler;

public ProgramCaller (String command, String[] args, 
int returnValue)
throws ExecuteException, IOException
{
cmdLine = new CommandLine(command + BBIni.getNativeCommandSuffix());
for (int i = 0; i < args.length; i++)
cmdLine.addArgument(args[i]);
resultHandler = new DefaultExecuteResultHandler();
ExecuteWatchdog watchdog = new ExecuteWatchdog(60*1000);
Executor executor = new DefaultExecutor();
executor.setExitValue(returnValue);
executor.setWatchdog(watchdog);
executor.execute(cmdLine, resultHandler);
}

public boolean doneYet () throws InterruptedException
{
return resultHandler.hasResult();
}

public void waitTillDone () throws InterruptedException
{
resultHandler.waitFor();
}

}

