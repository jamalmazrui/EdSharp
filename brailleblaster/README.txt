This file gives general instructions and explains the structure of this 
directory.

The files COPYRIGHT.txt and LICENSE.txt contain the information which 
their names imply. BrailleBlaster is copyrighted by ViewPlus 
Technologies, Inc. and Abilitiessoft, Inc. It is licensed under the 
Apache license, version 2.0.

For instructions on building BrailleBlaster see BUILD.txt the 
brailleblaster.jar
file will be placed in the dist directory. Note that you must compile 
the software listed in BUILD.txt as going into the native subdirectory 
to complete the build.

The src directory is patterned on the src directories of Apache 
projects. It contains the subdirectories main, resources and test. The 
main subdirectory contains the directory tree for BrailleBlaster source. 
the resources directory contains various files useful in development or 
necessary fro running. The test directory contains the directory tree 
for Junit tests.

The main directory contains the org directory which in turn contains the 
brailleblaster directory. There are two classes in this package and a 
number of subpackages. The first class is Main.java and takes care of 
starting and ending BrailleBlaster. If there are no arguments it passes 
control directly to the word processor. If there are arguments it passes 
control to the class Subcommands.java. This class processes subcommands 
like file2brl, emboss, help, etc. It seemed neater and less confusing to 
put the handling of subcommands in their own class rather than burdening 
Main.java with handling them.

The most interesting of the subpackages is wordprocessor. This was 
originally called editor, but this seemed unclear and confusing. The 
class in this package which initially receives control is 
WPManager.java. It takes care of getting things set up and co-ordinating 
word processing. DocumentManager.java handles multiple documents in an 
MDI environment. It co-ordinates BrailleView.java and DaisyView.java for 
each document. TextEditor.java is a very simple text editor intended for 
chores such as writing and editing liblouis tables. GUICommons.java 
contains methods used  by the other classes in carrying out complex GUI 
operations.

The org.brailleblaster.louisutdml package might be better named. It 
contains classes used to mediate between BrailleBlaster and the 
liblouisutdml bindings. The bindings have been made as simple as 
possible. Problems are handled by LiblouisutdmlException.java in this 
package.

The embossers package contains drivers for various ebossers.

The input package handles reading various types of input and translating 
them into jdom trees for the editor.

The util package contains classes used in various parts of 
BrailleBlaster, such as ProgramCaller.java for calling other programs.

The status of the output package is not clear at present. It may be 
redcundant.




