LaTeX-access

Python scripts for processing LaTeX source into niemeth braille and
audible speech.

Scripts written by Alastair Irving: alastair.irving@sjc.ox.ac.uk

Website: http://latex-access.sourceforge.net

Overview

These scripts, written in the Python language, are designed to
process a line of LaTeX source, and translate it into Nemeth
braille on a refreshable braille display. They only handle a single
line of LaTeX at any one time, however when scrolling through a
LaTeX document the braille is refreshed on the fly.  At the present time, support
for cursor rooting facilities on braille displays is not supported,
and it is advised that the user does not try to edit documents using
the cursor rooting buttons, and should use standard arrow keys
instead. It is hoped that cursor rooting will be implemented in the future.

The scripts also translate the current line into English speech which
  is then spoken by jaws.

The Python scripts provide output in a large area of
mathematics, such as fractions, superscripts and subscripts,
calculus notation, set theory notation, and a large number of
mathematical symbols.

The scripts currently interface to Jaws for Windows, version 5 and
above.  To install the scripts on a machine running windows and Jaws
version 5 or higher, do the following.

1. Download the latest stable release of version 2 of the Python software from
www.python.org and install it.

Note: LaTeX-access has not been tested with python 3 and probably won't work with it.

2. Download the pywin32 package  from
https://sourceforge.net/projects/pywin32 ensuring you have the correct
file to match your version of python. Install this.

3. Create a directory on the c: drive named latex-access. (The name
of the directory or it's level of depth in the filesystem does not
matter, however choosing a directory with a name easy to remember and
not too deep in the filesystem makes the next step easier.

4. Obtain the latest version of the scripts, (although you've
presumably already done this).  For details of how to obtain the
latest version please consult the website.

5. Extract the files and folder in the zip file to the latex_access
folder you created earlier. To do this, either simply run the zip file
with an application such as winrar and copy everything to the folder,
or extract the files and folder to the relevant location by right
clicking (or equivalent) on the file within windows explorer.

6. Copy the files from within the jaws folder to the folder where your
jaws scripts are located. These are usually in a path such as
c:\documents and settings\%username%\application data\freedom
scientific\jaws\%jaws version%\settings\enu
(This can be reached using the explore jaws submenu of the jaws menu
in your start menu).

7. Now open a command prompt by going to run, and typing 'cmd'.
Switch to the latex-access directory by typing 'cd %directory%' for
example, type 'cd c:\latex-access'

8. Type
'latex_access_com.py' you should hear 'latex_access registered'.  Now
register the matrix processor object by typing 'matrix_processor.py'
You should hear a similar message.  Exit the command
prompt by typing exit.

Note: the object here is to run the specified python files with 
python, so the above will only work if python is the default program
associated with .py files.  If it is not then try 
python latex_access_com.py
and if this fails then use the full path to your python installation,
for example 
c:\python26\python.exe latex_access_com.py
then repeat with matrix_processor.py

9. Finally, open the confignames.ini file in your jaws scripts folder
using a text editor such as notepad. After the line which reads
[ConfigNames]
enter a new line as described below.

As many different applications can be used to read or write LaTeX, the scripts have a
generic name latex.j**. Therefore, we use this file to add an alias,
so that many different applications can be used without needing to
change the name of each Jaws script to match the executable name of
the application used to write LaTeX. So for example, as a user of
winedt, I have 'winedt=latex' If you use notepad to write and read
LaTeX, you should add 'notepad=latex' etc.  If you are one of multiple
users on the computer, you may have to search to find the
confignames.ini file. The file will be found within the same file
structure, within the all users (or something similar) directory,
within the documents and settings directory.

Note: with older versions of jaws you could just enter the lines at
the bottom of the file, but in more recent versions the .ini file
contains more sections and the lines must be inserted in the
[ConfigNames] section.

10. Load the relevant program for writing/reading LaTeX, and press
ctrl+m to initialise the scripts. Repeat this keystroke to turn them
off. You should be able to navigate the document and listen to audible
speech output, as well as reading the mathematical translation on the
braille display.

Using the scripts

As stated above, the processing of LaTeX is toggled by means of
ctrl+m. There is currently no means of toggling speech and braille
independently but this is easy to implement if required. 

It is possible to toggle the speaking and/or brailling of dollar
signs.  The keys for this are ctrl+shift+d for speech and ctrl+d for Braille.

The Preprocessor

LaTeX enables you to define custom commands.  The scripts can handle
this but they must be told what the custom commands are.  This is done
by means of the preprocessor.

To add a command press ctrl+shift+a.  In the first textbox, enter the
custom command, in the next enter the number of arguments, 0 if there
are none, and in the 3rd box enter the translation of the custom
command.  The translation is the standard LaTeX equivalent of the
command, using #n to denote places where the nth argument should be
interpolated into the translation.  The 3 textboxes correspond to the
3 arguments to the \newcommand command used to define the custom
command.

When you close jaws the preprocessor entries are lost.  However, you
can save them to a file.  The keystroke for this is ctrl+shift+w,
after which you must enter a full file name including path.  To load
saved preprocessor data the keystroke is control+shift+r.  You can in
fact reload multiple preprocessor files and their entries will be
merged, however if the result is then saved all the entries will be
saved in one file.

The Matrix Processor

Because matrices are spread over multiple lines it can be hard to
perform calculations with them using only speech and a braille
display.  For this reason the scripts include a matrix processor.  To
load a matrix into the processor, highlight its contents, (not
including any \begin and \end commands), and press ctrl+shift+m.  For
example you might highlight the following:
1 & 2 & 3\\
4 & 5 & 6\\
7 & 8 & 9\\

Jaws should say initialised m by n matrix, where m is the number of
rows and n the number of columns.  The matrix is now stored in
memory, and it can be navigated independently of navigation and
editing in the LaTeX document.  Ctrl+Shift in conjunction with j, k, l
and i act as arrow keys enabling you to navigate the matrix entry by
entry.  Ctrl+shift with a number reads that row and ctrl+alt with a
number reads that column.  




