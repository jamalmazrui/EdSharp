JFW Script File                                                           (� |     appactivate        Wscript.Shell     objectcreate    '  %    %     appactivate '  %  '  %     	      �     appactivatehandle           pathgethomer     ForceWin.exe    
  '     %    stringquote      
     %     inttostring 
  '     %               shellrun       %     	      �     appdestroy     %                getfocus      getappmainwindow    '           '     %   %  %  %    sendmessage    	      �     appgettitle    %                getfocus      getappmainwindow    '         %     getwindowname      	      �     appquit    %                getfocus      getappmainwindow    '           '     %   %  %  %    sendmessage    	      �     appsetactive        %                getfocus      getappmainwindow    '           '     %   %  %          sendmessage    	      X    boodialogmulticheck            getfocus    '     %     stringisblank       Multi Check '           getjawssettingsdirectory     \Homer\DialogMultiCheck.boo 
  '     %    filetostring    '     %  %   %  %     %    inttostring   booeval '  %        %    setfocus          %     	      �    boodialogchoose           booinit         %   %  %    iniformdialogchoose    	           getfocus    '     %     stringisblank       Choose  '           getjawssettingsdirectory     \Homer\DialogChoose.boo 
  '     %    filetostring    '     %  %   %     %    inttostring        booeval '  %        %    setfocus          %     	      4    boodialogmultiinput           booinit         %   %  %    iniformdialogmultiinput    	           getfocus    '     %        stringcountsegment  '     %     stringisblank      %       
      Input   '       Fields  '            %    stringisblank   # L%       
  
      Text    '          getjawssettingsdirectory     \Homer\DialogMultiInput.boo 
  '     %    filetostring    '     %  %   %  %         booeval '  %        %    setfocus          %     	      �     booeval             booinit      	      $  ohomerboo     %   %  %  %  %    eval    '  %     	      �     booevalfile           %    \     stringcontains            getjawssettingsdirectory     \Homer\ 
  %   
  '         %     filetostring    '     %  %  %  %  %    booeval '  %     	      d    booinit $  ibooinitialized   ����
              	      $  ohomerboo              	          Iron.COM            createobjectex  &  ohomerboo   $  ohomerboo           &  ibooinitialized         	         		 Error initializing Boo component!     saystring        ����&  ibooinitialized          	         �     boopathgetshort         booinit      %      	           getjawssettingsdirectory     \Homer\PathGetShort.boo 
  '     %    filetostring    '     %  %                    booeval '  %     	      (    booshellgetsounddevices      booinit              	           getjawssettingsdirectory     \Homer\ShellGetSoundDevices.boo 
  '      %     filetostring    '     %                        booeval '     %         stringchopright '  %     	      L    booshellurltofile          %    filedelete       	       import Microsoft.VisualBasic.Devices from Microsoft.VisualBasic.dll
    '  %   oNetwork = Network()
   
  '  %   oNetwork.DownloadFile(s1, s2)
  
  '     %  %   %              booeval       %    fileexists  '  %     	      �     buttonclick    %             getfocus    '        �   '     %   %                sendmessage    	      �     buttonget3state    %             getfocus    '         %     getwindowstylebits  '       '  %  %  
     	      �     buttongetcheck     %             getfocus    '         %     getwindowstylebits  '       '  %  %  
     	      �     buttongetchecked       %             getfocus    '        �   '     %   %                sendmessage    	      �     buttongetdefault       %             getfocus    '         %     getwindowstylebits  '       '  %  %  
     	      �     buttongetgroup     %             getfocus    '         %     getwindowstylebits  '       '  %  %  
     	      �     buttongetpush      %             getfocus    '         %     getwindowstylebits  '        '  %  %  
     	      �     buttongetradio     %             getfocus    '         %     getwindowstylebits  '    	   '  %  %  
     	      �     buttoninvokeid      %                getfocus      getrealwindow   '            %   %    finddescendantwindow      buttonclick    	      �     buttonsetchecked        %             getfocus    '         %     buttonischecked '  %  # | %    
  " � %    # � %  
  
        %     buttonclick          �     controlclear       %             getfocus    '          '     %   %                sendmessage    	      �     controlcopy    %             getfocus    '          '     %   %                sendmessage    	      �     controlcut     %             getfocus    '           '     %   %                sendmessage    	      �     controlgettextlength       %             getfocus    '           '     %   %                sendmessage    	      �     controlisedit      %             getfocus    '            %     getwindowclass   Edit      stringequal    	      d     controlisiedoc       iegetcurrentdocument               	               	         �     controlisieserver           getfocus    '      %     getwindowclass  '     %   Internet Explorer_Server      stringequal    	      �     controlismdiframe           getfocus    '      %     getwindowclass  '     %   MDIClient     stringequal    	      �     controlisrichedit           getfocus    '         %     getwindowclass    stringlower '     %   richedit      stringcontains  " �    %   tclipdoc      stringcontains  
     	      h     controlisricheditdoc         getricheditdocument            	               	         �     controlnext    %                getfocus      getrealwindow   '        (   '        '       '     %   %  %  %    sendmessage    	      �     controlpaste       %             getfocus    '          '     %   %                sendmessage    	      �     controlprevious    %                getfocus      getrealwindow   '        (   '        '        '     %   %  %  %    sendmessage    	      |     controlprint       %             getfocus    '          '     %   %  %  %    sendmessage    	      �     controlundo    %             getfocus    '          '     %   %                sendmessage    	      �    convertexceltotext              getfocus      getrealwindow   '      Excel.Application     getobject   '  %          '  %      displayalerts   '  %      screenupdating  '         Excel.Application     objectcreate    '     %           %  !  displayalerts         %  !  screenupdating  %      workbooks   '  %    %           ����  open    '  %      sheets  ' 	 % 	     count   ' 
      '  %  % 
 
     % 	   %    item    '     %    filedelete     %    %         saveas     %      name    '  %  '  %  ' 	 %            close      %  '  %   

    
     %    filetostring    
  '  %     %       
 
            stringif    
  %  
  '  %    %           ����  open    '  %      sheets  ' 	 %       
  '   �      %  %    stringtofile       %  '  %  ' 	 %            close      %  '  %  '  %     %  %  !  displayalerts   %  %  !  screenupdating     %      quit                 Error!    ssay          %  '     %    fileexists     	      �    convertfiletotext          %    filedelete        %     pathgetextension    '  %   htm 
        %   %    converthtmltotext         %   doc 
  " � %   rtf 
  
        %   %    convertwordtotext         %   ppt 
        %   %    convertpowerpointtotext       %   xls 
        %   %    convertexceltotext              		 Cannot delete existing target file!   saystring            %    fileexists     	      �    converthtmltostring        InternetExplorer.Application      objectcreate    '        %  !  visible      %  !  silent  %    %     navigate       %      readystate       
               delay       �    %      document    '  %     %      body    '  %     %      createtextrange '  %     %      text    '  %  '     %  '     %      close      %  '  %      quit          %  '  %     	      �     converthtmltotext          %     converthtmltostring '     %  %    stringtofile          %    fileexists     	      $    convertjsstojsd       %          stringchopright  jsd 
  '     %    fileexists        %  %   .bak    
    filecopy          %    filetostring    '     %   ^\s*:(Function|Script)\s+(\S+)   [$1 $2]   regexpreplaceequiv  '     %  

 ^\s*:(Synopsis|Description)\s+(.*?)
    $1=$2
   regexpreplaceequiv  '     %   ^\s*:.*?
          regexpreplaceequiv  '       pathgettempfolder    \temp.ini   
  '     %  %    stringtofile            pause            %     fileexists                	       
  '     %     filetostring    '     %  %   
     regexpreplacecase   '     %   
     stringcountsegment  '       '  %  %  
           %   
   %    stringsegment     stringtrim  '     %   ;     stringleadequiv   # �   %   (     stringcontains  
  # �   %   )     stringcontains  
  #    %   Script    stringcontainsequiv "    %   Function      stringcontainsequiv 
  
        %   (     (    stringreplacecase   '     %              stringreplaceallcase    '     %   Function      stringleadequiv     Void    %  
  '        %              stringsegment   '     %              stringsegment   ' 	    %              stringsegment   ' 
    %   Script    stringequiv    %   :Script     
  % 	 
  %  
  '      Script  % 	 
   Synopsis         %    inireadstring   '     %    stringisblank            %   
   %       
    stringsegment     stringtrim  '        %   ;     stringleadcase        %         stringchopleft    stringtrim         stringif    '        %    stringisblank        %   :Synopsis   
  %  
  %  
  '         Script  % 	 
   Description      %    inireadstring   '     %    stringisblank      %       
  '           %   
   %    stringsegment     stringtrim   ;     stringleadcase     %              %   
   %    stringsegment     stringtrim         stringchopleft    stringtrim  
       
  '  %       
  '   �      %    stringtrim  '        %    stringisblank        %   :Description    
  %  
  %  
  '     %  %  
  '        % 	  Function      stringequiv    %   :Function   
  % 
 
  %  
  '      Function    % 
 
   Synopsis         %    inireadstring   '     %    stringisblank            %   
   %       
    stringsegment     stringtrim  '        %   ;     stringleadcase        %         stringchopleft    stringtrim         stringif    '        %    stringisblank        %   :Synopsis   
  %  
  %  
  '         Function    % 
 
   Description      %    inireadstring   '     %    stringisblank      %       
  '           %   
   %    stringsegment     stringtrim   ;     stringleadcase     %              %   
   %    stringsegment     stringtrim         stringchopleft    stringtrim  
       
  '  %       
  '   �      %    stringtrim  '        %    stringisblank        %   :Description    
  %  
  %  
  '        %   Void      stringequiv      %   :Returns    
  %  
  %  
  '        %   (     stringcontains  '     %   )     stringcontains  '     %  %       
  %  %  
       
    substring   '     %   ,     stringcountsegment  '       '  %  %  
           %   ,   %    stringsegment     stringtrim  '     %    stringisblank           %        /          stringreplaceex '  %       
     %   :Param  
  %  
  %  
  '     %   :Param  
     %    inttostring 
       
  %  
  %  
  '        %       
  '   d   %  %  
  '        %       
  '   �      %    filedelete     %     	      �	    convertpowerpointtotext             getfocus      getrealwindow   '      PowerPoint.Application    getobject   '  %          '  %      displayalerts   '         PowerPoint.Application    objectcreate    '     %          %  !  visible       %  !  displayalerts   %      presentations   '  %    %          open    '  %      name    '     %    pathgetbase ' 	 %      slides  ' 
 % 
     count   '  % 	  
  
     %    inttostring 
    Slide  
     %       
        s     stringif    
  ' 	      '  %  %  
     % 
   %    item    '  % 	  


  
   Slide   
     %    inttostring 
  ' 	 %      notespage   '  %      count   '       '       '  %  %  
     %    %    item    '  %      shapes  '  %      count   '       '  %  %  
     %    %    item    '  %      hastextframe       %      textframe   '  %      textrange   '  %      text    '  %       
     %     % 	  
  
   Notes:  
   
  
  %  
  ' 	       '     % 	  
  
  %  
  ' 	       %  '  %  '     %  '  %       
  '   �   %  '  %  '  %       
  '   $   %  '  %      shapes  '  %      count   '       '       '  %  %  
     %    %    item    '  %      hastextframe       %      textframe   '  %      textrange   '   %       text    '  %       
  # @   %   outline   stringequiv   
     %     % 	  
  
   Outline:    
   
  
  %  
  ' 	       '     % 	  
  
  %  
  ' 	       %  '   %  '     %      hastextframe      " ,%      hastextframe    # (%       
  
  
     %      alternativetext '  %       
     % 	  
  
  %  
  ' 	    %      type    ' ! % !      
     %      texteffect  ' " % "     text    '  %       
  #  %  %      alternativetext 
  
     % 	  
  
   Text Effect:    
  %  
  ' 	    %  ' "    %  ' #    %  '  %       
  '   d   %  '  %  '  %       
  '   h   %  ' 
    % 	 %    stringtofile            %  !  saved   %      close      %  '  %  '  %     % $ %  !  visible %  %  !  displayalerts      %      quit                 Error!    ssay          %  '     %    fileexists     	      �    convertwordtotext            '          getfocus      getrealwindow   '      Word.Application      getobject   '  %          '  %      displayalerts   '         Word.Application      objectcreate    '     %           %  !  displayalerts   %      documents   '  %    %           ����        open    '  %      selection   ' 	 % 	     storylength ' 
 % 	         % 
   setrange       % 	     text    '     %          regexpreplacecase   '     %       
    regexpreplacecase   '     %  %    stringtofile       %            close      %  '  %  '  %      normaltemplate  '       %  !  saved   %  '  %     %  %  !  displayalerts      %            quit                 Error!    ssay             '     %  '  %     	      �    dialogbrowseforfolder          Shell.Application     objectcreate    '  %         DialogBrowseForFolderHelper         schedulefunction       %    %  %  %  %     browseforfolder '  %     %      self    '  %      path    '            Folder Name:     Input   %     inputbox       %   '        %  '  %  '  %  '  %     	      �     dialogbrowseforfolderhelper  Browse for Folder   &  shomertitle    $  shomertitle   
     uiwaitfortitleandactivate         P    dialogconfirm                   
  '        %     stringisblank    Confirm %     stringif    '      %   Y     stringequiv    %        
  '     %       
  '        %  %   %    exmessagebox    '  %       
      Y   '     %       
      N   '     %     	      T     dialoghelpmenu                      dialoghelpmenuex       	      �    dialoghelpmenuex               %     stringisblank    Help Menu   %     stringif    '         %    stringisblank    Common Keys %    stringif    '       getjawssettingsdirectory     \   
       getactiveconfiguration  
   .jkm    
  '     %  %    inireadsectionkeys  '     %   |     stringcountsegment  '       '  %  %  
        %   |   %    stringsegment   '     %    stringisblank     #       %    stringtrim   ;     stringleadcase    
           %  %       %    inireadstring     stringtrim  '     %   UI    stringleadequiv          %         stringchopleft         (     stringpadright     %         (     stringpadright  
  '        %         (     stringpadright     %         (     stringpadright  
  '     % 	 %  
      
  ' 	 % 
          %    (     stringleft    stringtrim   	   
  %  
         P     stringpadright  
      
  ' 
    %       
  '   |      % 	        stringchopright ' 	    % 
        stringchopright ' 
 %        % 	       stringsegmentsort   ' 	    % 
       stringsegmentsort   ' 
               %   % 
         dialogpick  '     %    stringisblank      	         % 
     %    stringindexsegmentcase  '     % 	     %    stringsegment   '        %    (     stringright   stringtrim  '        % 
 %           dlgselectiteminlist '     % 	     %    stringsegment   '     %    stringisblank      	             $   %  
          schedulefunction          �    dialoginput          Internal     SuppressTitle         Homer.ini     iniwriteinteger       %     stringisblank       Input   '         %    stringisblank     #     %         stringchopright  :   
  
     %   :   
  '        %  %   %    inputbox       %  '         Internal     SuppressTitle          Homer.ini     iniwriteinteger    %     	          dialogpick          %     stringisblank       Pick    '      %        %        stringsegmentsort   '        %  %           dlgselectiteminlist '       pause      %        %      %    stringsegment   '     %     	      X    dialogopenfile         UserAccounts.CommonDialog     objectcreate    '  %        %    stringisblank       All Files (*.*)|*.* %  !  filter     %  %  !  filter     %        
          %  !  filterindex    %  %  !  filterindex    %   %  !  initialdir      DialogOpenFileHelper            schedulefunction       %      showopen       %      filename    '              %     folderexists    %    \   
         stringif    '       File Name:   Open    %     inputbox       %   '        %  '  %     	      |     dialogopenfilehelper     Open    &  shomertitle    $  shomertitle   
     uiwaitfortitleandactivate         �    dialogsavefile         SAFRCFileDlg.FileSave     objectcreate    '  %     %   %  !  filename       %    stringisblank       All Files (*.*) %  !  filetype       %  %  !  filetype           DialogSaveFileHelper            schedulefunction       %      openfilesavedlg    %      filename    '            File Name:   Save    %     inputbox       %   '        %  '  %     	      |     dialogsavefilehelper     Save As &  shomertitle    $  shomertitle   
     uiwaitfortitleandactivate         D     editgetindex          %     editgetselectionend    	      �     editgetindexcolumn      %             getfocus    '      %    ����
        %     editgetselectionend '        %   %    editgetindexline    '     %   %    editgetlinestart    '  %  %  
  '  %     	      �     editgetindexline        %             getfocus    '        �   '     %   %  %          sendmessage    	      �     editgetlinecount       %             getfocus    '        �   '     %   %                sendmessage    	      4    editgetlineend      %             getfocus    '         %   %    editgetlinestart    '  %        
       ����   	           pause         %   %       
    editgetlinestart    '       pause      %        
        %     editgetsize '     %     	      �     editgetlinelength       %             getfocus    '         %   %    editgetlinestart    '  %        
       ����   	           pause         %   %    editgetlinelengthfromindex  '  %     	      �     editgetlinelengthfromindex      %             getfocus    '        �   '     %   %  %          sendmessage    	      �     editgetlinestart        %             getfocus    '        �   '     %   %  %          sendmessage    	      �    editgetlinetext     %             getfocus    '      %    ����
        %     editgetindex    '     %   %    editgetindexline    '        %   %    editgetlinestart    '  %        
             	           pause         %   %       
    editgetlinestart    '       pause      %        
        %     editgetsize '        %   %  %    editgetrange    '  %     	      �     editgetmodified    %             getfocus    '        �   '     %   %                sendmessage    	      �     editgetmultiline       %             getfocus    '         %     getwindowstylebits  '       '  %  %  
     	      �     editgetrange         %             getfocus    '         %     editgettext '  %       
  '  %       
  '     %  %  %    stringgetrange  '  %     	      �     editgetreadonly    %             getfocus    '         %     getwindowstylebits  '       '  %  %  
     	      �     editgetselectedtext    %             getfocus    '         %     editgettext '     %     editgetselectionstart        
  '     %     editgetselectionend      
  '     %  %  %    stringgetrange  '  %     	      �     editgetselectionend    %             getfocus    '        �   '        %   %                sendmessage   mathhighword       	      �     editgetselectionstart      %             getfocus    '        �   '        %   %                sendmessage   mathlowword    	      �     editgetsize    %             getfocus    '         %     editgetlinecount         
  '     %   %    editgetlinestart    '     %   %    editgetlinelengthfromindex  '  %  %  
  '  %     	      �     editgettext    %             getfocus    '         %     ����     %    getobjectfromevent  '  %      accvalue    '  %  '  %     	      �     editgobottom       %             getfocus    '         %     editgetsize      
  '     %   %    editsetindex          h     editgotop      %             getfocus    '         %           editsetindex          �     editreplacerange          %             getfocus    '         %     editgettext '  %       
  '  %       
  '     %  %  %  %    stringreplacerange  '     %   %    editsettext       �     editscrollcaret    %             getfocus    '        �   '     %   %                sendmessage    	      �     editselectall      %             getfocus    '         %           ����  editsetselection          %     editscrollcaret       �    editselectchunk    %             getfocus    '         %     editgetindex    '  %  '     %     editgetsize '     %   %  %    editgetrange    '   \s+ '     %  %    regexpcontainscase  '        %             stringsegment     stringtoint '  %  %  
       
  '        '  %  '     %   %  %    editgetrange    '     %  %    regexpcontainslastequiv '        %             stringsegment     stringtoint '     %             stringsegment   ' 	 %     % 	   stringlength    
       
  '     %   %  %    editsetselection          �     editsetindex        %             getfocus    '         %   %  %    editsetselection            pause         %     editscrollcaret       �     editsetmodified     %             getfocus    '        �   '     %   %  %          sendmessage    	      �     editsetselection         %             getfocus    '        �   '     %   %  %  %    sendmessage    	      �     editsettext     %             getfocus    '         %     ����     %    getobjectfromevent  '       pause      %  %  !  accvalue    %  '  %     	      �     editunselectall    %             getfocus    '         %     editgetindex    '     %     ����        editsetselection          %   %    editsetindex          %     editscrollcaret       �     filecopy           %    filedelete         Scripting.FileSystemObject    objectcreate    '  %    %   %    copyfile          %    fileexists  '  %  '     %     	      �     filedelete        %     fileexists         Scripting.FileSystemObject    objectcreate    '  %    %          deletefile     %  '        %     fileexists       	      �     filegetsize        Scripting.FileSystemObject    objectcreate    '  %    %     getfile '  %      size    '  %  '  %  '  %     	      �     filemove           %    filedelete         Scripting.FileSystemObject    objectcreate    '  %    %   %    movefile          %    fileexists  '  %  '     %     	      �     filetostring           Scripting.FilesystemObject    objectcreate    '  %    %     opentextfile    '  %      readall '  %      close      %  '  %  '  %     	      �     foldercopy          Scripting.FileSystemObject    objectcreate    '  %    %   %    copyfolder        %    folderexists    '  %  '  %     	      �     foldercreate          %     folderexists               	          Scripting.FileSystemObject    objectcreate    '  %    %     createfolder       %  '     %     folderexists       	      �     folderdelete          %     folderexists           Scripting.FileSystemObject    objectcreate    '  %    %          deletefolder       %  '        %     folderexists         	      �     folderexists           Scripting.FileSystemObject    objectcreate    '  %    %     folderexists    '  %  '  %     	      �     foldergetsize          Scripting.FileSystemObject    objectcreate    '  %    %     getfolder   '  %      size    '  %  '  %  '  %     	      ,    folderisroot          %    \     stringequal "        %     stringlength           intequal    #     abcdefghijklmnopqrstuvwxyz     %          stringleft    stringcontainsequiv #       %          stringright  :\    stringequal 
  
  
     	      �     foldermove          Scripting.FileSystemObject    objectcreate    '  %    %   %    movefolder        %    folderexists    '  %  '  %     	      �     ibox                 voicesetspeech  '        %     inttostring   messagebox        %    voicesetspeech        �    inideletefile         %     fileexists               	         %     inireadsectionnames '     %   |     stringcountsegment  '       '  %  %  
        %   |   %    stringsegment   '     %  %     iniremovesection       %       
  '   �       %     iniflushfile            pause         %     filedelete          pause         X     iniflushfile                         %     iniwritestring     	      �    ieeval          getclipboardtext    '       pause       javascript:window.clipboardData.setData('Text', eval("  %   
   ").toString()); 
  '      InternetExplorer.Application      objectcreate    '        %  !  visible %    %    navigate            %  !  silent       pause      %      quit       %  '       getclipboardtext    '     %    copytoclipboard    %     	      ,    iniformdialogmultiinput          Internal     SuppressTitle         Homer.ini     iniwriteinteger            getjawssettingsdirectory     \   
       getactiveconfiguration  
   .jcf    
       getjawssettingsdirectory     \IniForm.jcf    
    filecopy            getjawssettingsdirectory     \Homer\IniForm.exe  
  '     %    pathgetshort    '       pathcreatetempfolder    '  %   \Input.ini  
  '  %       
  %  
   \   
  '     %        stringcountsegment  '     %     stringisblank      %       
      Input   '       Fields  '            %     saystring         %    stringisblank   # �%       
  
      Text    '      [   %   
   ]
Misc=NoStatus
  
  '       ' 	 % 	 %  
        %      % 	   stringsegment   ' 
    %      % 	   stringsegment   '  %   [   
  % 
 
   ]
Value=   
  %  
   
  
  '  % 	      
  ' 	  4   %   [OK]
Control=Button
  
  '  %   [Cancel]
Control=Button
  
  '     %    inideletefile         %  %    stringtofile            getfocus    '     %              shellrun            pause      %        %    setfocus            pause            %    filedelete     %   \Output.ini 
  '     %    fileexists          ' 	 % 	 %  
        %      % 	   stringsegment   ' 
 % 	 %  
     %      Results % 
      %    inireadstring   
  '     %      Results % 
      %    inireadstring   
      
  '     % 	      
  ' 	  �      %    filedelete           %    folderdelete           Internal     SuppressTitle          Homer.ini     iniwriteinteger    %     	      \     iniformdialoghelper    $  shomertitle   
     uiwaitfortitleandactivate         �    iniformdialoginfo           Internal     SuppressTitle         Homer.ini     iniwriteinteger            getjawssettingsdirectory     \   
       getactiveconfiguration  
   .jcf    
       getjawssettingsdirectory     \IniForm.jcf    
    filecopy            getjawssettingsdirectory     \Homer\IniForm.exe  
  '     %    pathgetshort    '       pathcreatetempfolder    '  %   \Input.ini  
  '  %       
  %  
   \   
  '  %   \Input.txt  
  '     %     stringisblank       Info    '       [   %   
   ]
MemoWidth=300
MemoHeight=300
Misc=NoStatus
   
  '  %   [InfoView]
Control=Memo
  
  '  %   Misc=NoLabel|ReadOnly
 
  '  %   [Close]
Control=Button
ID=2
 
  '     %  %    stringtofile        [[InfoView]]
  %  
  '     %  %    stringtofile            getfocus    '     %              shellrun            pause      %        %    setfocus            pause            %    folderdelete           Internal     SuppressTitle          Homer.ini     iniwriteinteger       \     iniformdialoginput          %   %  %    iniformdialogmultiinput    	      �    iniformdialogpick            Internal     SuppressTitle         Homer.ini     iniwriteinteger            getjawssettingsdirectory     \   
       getactiveconfiguration  
   .jcf    
       getjawssettingsdirectory     \IniForm.jcf    
    filecopy            getjawssettingsdirectory     \Homer\IniForm.exe  
  '     %    pathgetshort    '       pathcreatetempfolder    '  %   \Input.ini  
  '  %       
  %  
   \   
  '  %   \Input.txt  
  '     %     stringisblank       Pick    '         %     saystring       [   %   
   ]
ListWidth=300
ListHeight=300
Misc=NoStatus
   
  '  %   [List]
Control=List
  
  '     %       
    regexpreplacecase   '   [[List]]
  %  
  '     %  %    stringtofile       %   Selection=1 
   
  
  '  %     %   Misc=NoLabel|Sort
 
  '     %   Misc=NoLabel
  
  '     %   [OK]
Control=Button
  
  '  %   [Cancel]
Control=Button
  
  '     %  %    stringtofile            getfocus    ' 	     IniFormDialogHelper         schedulefunction          %              shellrun            pause      % 	       % 	   setfocus            pause            %    filedelete     %   \Output.ini 
  '     %    fileexists         Results  List         %    inireadstring   ' 
    % 
  |              stringreplaceex ' 
    %    filedelete           %    folderdelete           Internal     SuppressTitle          Homer.ini     iniwriteinteger    % 
    	          iniformdialogmemo           Internal     SuppressTitle         Homer.ini     iniwriteinteger            getjawssettingsdirectory     \   
       getactiveconfiguration  
   .jcf    
       getjawssettingsdirectory     \IniForm.jcf    
    filecopy            getjawssettingsdirectory     \Homer\IniForm.exe  
  '     %    pathgetshort    '       pathcreatetempfolder    '  %   \Input.ini  
  '  %       
  %  
   \   
  '  %   \Input.txt  
  '     %     stringisblank       Memo    '       [   %   
   ]
MemoWidth=300
MemoHeight=300
Misc=NoStatus
   
  '  %   [MemoEdit]
Control=Memo
  
  '  %   Misc=NoLabel
  
  '  %   [OK]
Control=Button
  
  '  %   [Cancel]
Control=Button
  
  '     %  %    stringtofile        [[MemoEdit]]
  %  
  '     %  %    stringtofile            getfocus    '     %              shellrun            pause      %        %    setfocus            pause            %    filedelete        %    filedelete     %   \Output.ini 
  '     %    fileexists     %   \Output.txt 
  '     %    filetostring    ' 	    % 	     [[MemoEdit]]
    stringlength       % 	   stringlength      substring   ' 	    %    filedelete        %    filedelete           %    folderdelete           Internal     SuppressTitle          Homer.ini     iniwriteinteger    % 	    	      p    iniformdialogmultipick           Internal     SuppressTitle         Homer.ini     iniwriteinteger            getjawssettingsdirectory     \   
       getactiveconfiguration  
   .jcf    
       getjawssettingsdirectory     \IniForm.jcf    
    filecopy            getjawssettingsdirectory     \Homer\IniForm.exe  
  '     %    pathgetshort    '       pathcreatetempfolder    '  %   \Input.ini  
  '  %       
  %  
   \   
  '  %   \Input.txt  
  '     %     stringisblank       Multi Pick  '       [   %   
   ]
MultiWidth=600
MultiHeight=600
Misc=NoStatus
 
  '  %   [Multi]
Control=Multi
    
  '     %       
    regexpreplacecase   '   [[Multi]]
 %  
  '     %  %    stringtofile       %     %   Misc=NoLabel|Sort
 
  '     %   Misc=NoLabel
  
  '     %   [OK]
Control=Button
  
  '  %   [Cancel]
Control=Button
  
  '     %  %    stringtofile            getfocus    ' 	 %   &  shomertitle     IniFormDialogHelper         schedulefunction          %              shellrun            pause      % 	       % 	   setfocus            pause            %    filedelete     %   \Output.ini 
  '     %    fileexists         Results  Multi        %    inireadstring   ' 
    % 
  |              stringreplaceex ' 
    %    filedelete           %    folderdelete           Internal     SuppressTitle          Homer.ini     iniwriteinteger    % 
    	      �    iniforminieditsection          %   %    inireadsectionkeys  '     %   |              stringreplaceex '     %        stringcountsegment  '       '  %  %  
        %      %    stringsegment   '     %   %       %    inireadstring   '  %  %  
     %  %  
  '     %  %  
      
  '     %       
  '   �        Fields  %  %    iniformdialogmultiinput '     %    stringisblank     '  %          '  %  %  
        %      %    stringsegment   '     %      %    stringsegment   '     %   %  %  %    iniwritequote      %       
  '   �      %     	      �     inireadsetting          Internal    %   %       getactiveconfiguration   .ini    
    inireadstring      	      �     iniwritequote          "   %  
   "   
  '     %   %  %  %         iniwritestring        %    iniflush          �     iniwritesetting         Internal    %   %       getactiveconfiguration   .ini    
    iniwritequote      	      0     intequal        %   %  
     	      H     intif        %      %  '     %  '     %     	      @     isay             %     inttostring   ssay          0    jsdialogmultiinput            jsinit          %   %  %    iniformdialogmultiinput    	           getfocus    '     %        stringcountsegment  '     %     stringisblank      %       
      Input   '       Fields  '            %    stringisblank   # L%       
  
      Text    '          getjawssettingsdirectory     \Homer\MultiInput.JS    
  '     %    filetostring    '     %  %   %  %        jseval  '  %        %    setfocus          %     	      �     jseval              jsinit       	      $  ohomerjs      %   %  %  %  %    eval    '  %     	      �     jsevalfile            %    \     stringcontains            getjawssettingsdirectory     \Homer\ 
  %   
  '         %     filetostring    '     %  %  %  %  %    jseval  '  %     	      d    jsinit  $  ijsinitialized    ����
              	      $  ohomerjs               	          Iron.JS         createobjectex  &  ohomerjs    $  ohomerjs            &  ijsinitialized          	         

 Error initializing JScript component!     saystring        ����&  ijsinitialized           	             jsshellurltofile           %    filedelete       	       var o = new Microsoft.VisualBasic.Devices.Network();
   '  %   o.DownloadFile(s1, s2);
    
  '     %  %   %              jseval        %    fileexists  '  %     	      �     keyboardsend           Wscript.Shell     objectcreate    '  %          pause                delay      %    %     keyboardsendkeys    '       ����'     %  '  %     	      �     keyboardtype                   %     typestring     	         %     stringlength    '       '  %  %  
        %   %    
     substring   '     %    typestring          pause      %    
   
  '   t       p    keyboardtypeandwait           speechoff           '  %        %    |   %    stringsegment   '     %    stringlength          
           '        %    typekey                    delay      %       
  '      H            getfocus      getwindowclass  '        '     %    stringlength       %    stringlength    
  # �%  %  
  
  # �%  %  
  
                     delay      %       
  '          getfocus      getwindowclass  '   P   %  %  
          '          speechon       %     	           keyboardtypelist             speechoff           pause           '  %        
 
       %   %  %    stringsegment   '  %       
           '        %    typekey         pause      %       
  '      X         speechon          �    lbcdialogmultiinput           lbcinit         %   %  %    iniformdialogmultiinput    	          Internal     SuppressTitle         Homer.ini     iniwriteinteger       %        stringcountsegment  '     %     stringisblank      %       
      Input   '       Fields  '            %    stringisblank   # �%       
  
      Text    '     $  ohomerlbc     %   %  %    fielddialog '  %        %    setfocus             %    stringisblank      	         %   	         regexpreplacecase   '      Internal     SuppressTitle          Homer.ini     iniwriteinteger    %  '  %     	      �     lbcdialoghelper              JAWS      appactivatetitle            pause            $  shomertitle   
     uiwaitfortitleandactivate         T     lbcdialoginput          %   %  %    lbcdialogmultiinput    	      �    lbcdialogpick            Internal     SuppressTitle         Homer.ini     iniwriteinteger         lbcinit         %   %  %    iniformdialogpick      	           getfocus    '     %     stringisblank       Pick    '      %   &  shomertitle    %       	     regexpreplacecase   '      LbcDialogHelper         schedulefunction       $  ohomerlbc     %   %  %    listdialog  '  %        %    setfocus             %    stringisblank      	         %   	         regexpreplacecase   '      Internal     SuppressTitle          Homer.ini     iniwriteinteger    %     	      �    lbcdialogmultipick            lbcinit         %   %  %    iniformdialogmultipick     	          Internal     SuppressTitle         Homer.ini     iniwriteinteger       %     stringisblank       Multi Pick  '         %       	     regexpreplacecase   '      LbcDialogHelper         schedulefunction       $  ohomerlbc     %   %  %    multilistdialog '  %        %    setfocus             %    stringisblank      	         %   	         regexpreplacecase   '      Internal     SuppressTitle          Homer.ini     iniwriteinteger    %  '  %     	      8    lbcinit $  ohomerlbc            Initializing .NET component   saystring          Please wait   saystring          LayoutByCode.Lbc            createobjectex  &  ohomerlbc       Done!     saystring         $  ohomerlbc           '            '      %      	      �     listgetextended    %             getfocus    '         %     getwindowstylebits  '       '  %  %  
     	      �     listgetmulti       %             getfocus    '         %     getwindowstylebits  '       '  %  %  
     	      �     listgetsort    %             getfocus    '         %     getwindowstylebits  '       '  %  %  
     	      T    mathdecinttohexstring           '   0x  '   123456789abcdef '       '  %        
 	            %    mathinttopower  '  %   %  
  '  %        %  %         substring   '        '      0   '     %       %  %  
  '     %   %  
  '   %       
  '   \    %     	      �    mathhexstringtodecint       123456789abcdef '     %     stringlower '      %          stringleft   0x  
        %          stringchopleft  '          0x0000000   %   
         stringright '        '  %        
 	       %        %  
         substring   '     %  %    stringcontains  '  %  %          %    mathinttopower  
  
  '  %       
  '   �    %     	      �     mathhighword            '  %       
 	            %    mathinttopower  '  %   %  
  '  %  %          %       
    mathinttopower  
  
  '  %   %  
  '   %       
  '   $    %     	      x     mathinttopower           '  %        
 
    %  %   
  '  %       
  '   (    %     	      h     mathlowword         '          %    mathinttopower  '  %   %  
  '  %     	      �     menuinvokeid        %                getfocus      getappmainwindow    '          '     %   %  %          sendmessage    	      �     msaagetchildcount         %     ����     %    getobjectfromevent  '  %      accchildcount   '  %  '  %     	      �     msaagetdefaultaction          %     ����     %    getobjectfromevent  '  %            accdefaultaction    '  %  '  %     	      �     msaagetdescription        %     ����     %    getobjectfromevent  '  %            accdescription  '  %  '  %     	      �     msaagetfocus          %     ����     %    getobjectfromevent  '  %            accfocus    '  %  '  %     	      �     msaagethelp       %     ����     %    getobjectfromevent  '  %            acchelp '  %  '  %     	      �    msaagetinfo         getfocus    '      %     ����     %    getobjectfromevent  '  %            accchildcount   '        '  %  %  
        %    %    accrole   getroletext '     %    stringisblank        %   Role=   
  %  
   
   
  '     %    %    accname '     %    stringisblank        %   Name=   
  %  
   
   
  '     %    %    accvalue    '     %    stringisblank        %   Value=  
  %  
   
   
  '     %    %    accfocus    '  %     %   Focus=  
     %    inttostring 
   
   
  '        %    %    accstate      msaagetstatetext    ' 	    % 	   stringisblank        %   State=  
  % 	 
   
   
  '     %    %    acchelp ' 
    % 
   stringisblank        %   Help=   
  % 
 
   
   
  '     %    %    accdescription  '     %    stringisblank        %   Description=    
  %  
   
   
  '     %    %    acckeyboardshortcut '     %    stringisblank        %   KeyboardShortcut=   
  %  
   
   
  '     %    %    accdefaultaction    '     %    stringisblank        %   DefaultAction=  
  %  
   
   
  '     %    %    accchildcount   '  %     %   Children=   
     %    inttostring 
   
   
  '     %       
  '   �    %  '  %     	      �     msaagetkeyboardshortcut       %     ����     %    getobjectfromevent  '  %            acckeyboardshortcut '  %  '  %     	      �     msaagetname       %     ����     %    getobjectfromevent  '  %            accname '  %  '  %     	      �     msaagetrole       %     ����     %    getobjectfromevent  '  %            accrole '  %  '  %     	      T     msaagetrolestring            %     msaagetrole   getroletext    	      �     msaagetstate          %     ����     %    getobjectfromevent  '  %            accstate    '  %  '  %     	      \     msaagetstatestring           %     msaagetstate      getstatetext       	      �    msaagetstatetext       %        
     %   unavailable     
  '     %        
     %   selected    
  '     %        
     %   focused     
  '     %        
     %   pressed     
  '     %        
     %   checked     
  '     %         
     %   mixed   
  '     %     @   
     %   readonly    
  '     %     �   
     %   hottracked  
  '     %        
     %   default     
  '     %        
     %   expanded    
  '     %        
     %   collapsed   
  '     %        
     %   busy    
  '     %        
     %   floating    
  '     %         
     %   marqueed    
  '     %      @  
     %   animated    
  '     %      �  
     %   invisible   
  '     %        
     %   offscreen   
  '     %        
     %   sizeable    
  '     %        
     %   moveable    
  '     %        
     %   selfvoicing     
  '     %        
     %   focusable   
  '     %         
     %   selectable  
  '     %       @ 
     %   linked  
  '     %       � 
     %   traversed   
  '     %        
     %   multiselectable     
  '     %        
     %   extselectable   
  '     %        
     %   alert_low   
  '     %        
     %   alert_medium    
  '     %        
     %   alert_high  
  '     %     ���
     %   protected   
  '     %     ���
     %   valid   
  '        %    stringtrimtrailingblanks    '  %     	      �     msaagetvalue          %     ����     %    getobjectfromevent  '  %            accvalue    '  %  '  %     	      �     msaasetselection        %             getfocus    '         %     ����%  %    getobjectfromevent  '  %    %  %    accselect      %  '  %     	      �     objectcreate          %          createobjectex  '  %          %           createobjectex  '     %     	      |     pathcombine     %    \   
  %  
  '     %   \\   \     stringreplaceallcase    '  %     	      �     pathcreatetempfolder         pathgettempfolder    \   
       pathgettempname 
  '      %     foldercreate          %     folderexists       %      	         �     pathgetbase        Scripting.FileSystemObject    objectcreate    '  %    %     pathgetbase '  %  '  %     	      �     pathgetcurrentdirectory     Wscript.Shell     objectcreate    '   %       currentdirectory    '  %  '   %     	      �    pathgetdir        %COMSPEC% /c dir /b     %  
    "  
  %   
   \   
  %  
   "   
  '       pathgettempfolder    \   
       pathgettempname 
  '  %    >  
  %  
  '        '       '     %  %  %    shellrun            pause            %    filetostring      stringtrim  '     %    filedelete     %     	      �     pathgetextension           Scripting.FileSystemObject    objectcreate    '  %    %     getextensionname    '  %  '  %     	           pathgetfolder          Scripting.FileSystemObject    objectcreate    '  %    %     getparentfoldername '     %    stringisblank     # �    %   \     stringcontains    
     %   \   
  '     %  '  %     	      T     pathgethomer         getjawssettingsdirectory     \Homer\ 
     	      �     pathgetinternetcachefolder      Shell.Application     objectcreate    '         '  %     %    namespace   '  %      self    '  %      path    '  %  '  %  '  %  '   %     	      �     pathgetlong        WScript.Shell     objectcreate    '  %     temp.lnk      createshortcut  '  %   %  !  targetpath  %      targetpath  '  %  '  %  '  %     	      �     pathgetname        Scripting.FileSystemObject    objectcreate    '  %    %     getfilename '  %  '  %     	          pathgetshort           Scripting.FileSystemObject    objectcreate    '     %     folderexists       %    %     getfolder   '  %      shortpath   '     %    %     getfile '  %      shortpath   '     %  '  %  '  %  '  %     	      �    pathgetspecialfolder           WScript.Shell     objectcreate    '  %      specialfolders  '  %      count   '       '  %  %  
     %    %    item    '      \   %   
     %     %    stringlength       %     stringlength    
     %    stringlength      substring     stringequiv    %       
  '     %       
  '      �    %  '  %  '  %     	      d     pathgettempfile      pathgettempfolder    \   
       pathgettempname 
     	      �     pathgettempfolder       Scripting.FileSystemObject    objectcreate    '        '  %     %    getspecialfolder        path    '  %  '   %     	      �     pathgettempname     Scripting.FileSystemObject    objectcreate    '   %       gettempname '  %  '   %     	      �     pathgettype       %     folderexists        Folder  '         Scripting.FileSystemObject    objectcreate    '  %    %     getfile '  %      type    '     %  '  %  '  %     	          pathgetvalidname           @%*+\|;:'"<>/?  '   !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~  '       %  
  '  %   '  %  '  %  ' 	 %  ' 
    %    stringlength    '  %  %  
        %  %         substring   '     % 
 %    stringcontains        % 
 %   _     stringreplaceallcase    ' 
    %       
  '   �       % 
             stringreplaceallcase    ' 
    % 
   _   _     stringreplaceallcase    ' 
    % 
  _    _     stringreplaceallcase    ' 
    % 
  __   _     stringreplaceallcase    ' 
    % 
   stringtrim  ' 
    % 
        stringleft   _   
        % 
        stringchopleft  ' 
  �      % 
        stringright  _   
        % 
        stringchopright ' 
        % 
   stringtrim  ' 
 % 
 '     %  %    pathcombine % 	 
  '  %  # �   %    fileexists  
      _01 '     %  %    pathcombine %  
  % 	 
  '       '     %    fileexists  # T%    c   
  
     %       
  '     %    inttostring '   _      %   0          stringpadleft   
  '     %  %    pathcombine %  
  % 	 
  '   $      %     	      |     pathsetcurrentdirectory        Wscript.Shell     objectcreate    '  %   %  !  currentdirectory    %  '          perlcombogetcontents            perlinit         	      %             getfocus    '       join '', Win32::GuiTest::GetComboContents($p1) '     %     %     inttostring                  perleval    '  %     	      �     perlcombogettext            perlinit         	      %             getfocus    '      		 Win32::GuiTest::GetComboText($p1)   '     %     %     inttostring                  perleval    '  %     	      <    perldatecalc             %     stringtrim  '      %    stringtrim  '     %    stringtrim  '     %    stringtrim  '   package Date::Calc;
    '     %    stringtoint      %   $p2 = Decode_Month($p2);
   
  '        %    stringtoint      %   $p4 = Decode_Day_of_Week($p4);
 
  '        %    stringtoint    %   ($y, $m, $d) = Nth_Weekday_of_Month_Year($p1, $p2, $p4, $p3);
  
  '     %  		 ($y, $m, $d) = ($p1, $p2, $p4);
    
  '     %   Date_to_Text_Long($y, $m, $d);
 
  '     %  %   %  %  %    perleval    '     %   st   ,     stringreplacecase   '     %   nd   ,     stringreplacecase   '     %   rd   ,     stringreplacecase   '     %   th   ,     stringreplacecase   '  %     	      �    perldialogappendfile            perlinit            %     dialogsavefile     	           getfocus    '   Win32::GUI::GetSaveFileName(-file => $p1, -overwriteprompt => 0);   '      DialogSaveFileHelper            schedulefunction       $  ohomerperl    %  %                    eval    '       pause         %    setfocus       %     	      �    perldialogbrowseforfolder           perlinit            %     dialogbrowseforfolder      	           getfocus    '     %     pathsetcurrentdirectory     Win32::GUI::BrowseForFolder(-folderonly => 1, -editbox => 1);   '      DialogBrowseForFolderHelper         schedulefunction       $  ohomerperl    %  %                    eval    '       pause         %    setfocus       %     	          perldialogmultiinput             Internal     SuppressTitle         Homer.ini     iniwriteinteger         perlinit            %   %  %    iniformdialogmultiinput    	           getfocus    '     %        stringcountsegment  '     %     stringisblank      %       
      Input   '       Fields  '            %    stringisblank   # �%       
  
      Text    '          getjawssettingsdirectory     \Homer\DialogMultiInput.pl  
  '     %    filetostring    '     %  %   %  %         perleval    '  %        %    setfocus             %    stringisblank      	          Internal     SuppressTitle          Homer.ini     iniwriteinteger    %     	      |    perldialoginfo          Internal     SuppressTitle         Homer.ini     iniwriteinteger         perlinit            %   %    iniformdialoginfo      	           getfocus    '     %     stringisblank       Info    '         %     saystring           getjawssettingsdirectory     \Homer\DialogInfo.pl    
  '     %    filetostring    '     %  %   %              perleval    '  %        %    setfocus             %    stringisblank      	          Internal     SuppressTitle          Homer.ini     iniwriteinteger    %     	      X     perldialoginput         %   %  %    perldialogmultiinput       	      �    perldialogpick           Internal     SuppressTitle         Homer.ini     iniwriteinteger         perlinit            %   %  %    iniformdialogpick      	           getfocus    '     %     stringisblank       Pick    '      %   &  shomertitle      getjawssettingsdirectory     \Homer\DialogPick.pl    
  '     %    filetostring    '     %       |     regexpreplacecase   '     %  %   %     %    inttostring    %    inttostring   perleval    '  %        %    setfocus             %    stringisblank      	          Internal     SuppressTitle          Homer.ini     iniwriteinteger    %     	      |    perldialogmemo          Internal     SuppressTitle         Homer.ini     iniwriteinteger         perlinit            %   %    iniformdialogmemo      	           getfocus    '     %     stringisblank       Memo    '         %     saystring           getjawssettingsdirectory     \Homer\DialogMemo.pl    
  '     %    filetostring    '     %  %   %              perleval    '  %        %    setfocus             %    stringisblank      	          Internal     SuppressTitle          Homer.ini     iniwriteinteger    %     	          perldialogmultipick          Internal     SuppressTitle         Homer.ini     iniwriteinteger         perlinit            %   %  %    iniformdialogmultipick     	           getfocus    '     %     stringisblank       Multi Pick  '      %   &  shomertitle      getjawssettingsdirectory     \Homer\DialogMultiPick.pl   
  '     %    filetostring    '     %       |     regexpreplacecase   '     %  %   %     %    inttostring    %    inttostring   perleval    '  %        %    setfocus             %    stringisblank      	         %   \|        regexpreplacecase   '      Internal     SuppressTitle          Homer.ini     iniwriteinteger    %     	      �    perldialogopenfile          perlinit            %     dialogopenfile     	           getfocus    '   Win32::GUI::GetOpenFileName(-directory => $p1, -filemustexist => 1, -pathmustexist => 1);   '      DialogOpenFileHelper            schedulefunction       $  ohomerperl    %  %                    eval    '       pause         %    setfocus       %     	      t    perldialogsavefile          perlinit            %     dialogsavefile     	           getfocus    '   Win32::GUI::GetSaveFileName(-file => $p1);  '      DialogSaveFileHelper            schedulefunction       $  ohomerperl    %  %                    eval    '       pause         %    setfocus       %     	          perleditgettext         perlinit            %     editgettext    	      %             getfocus    '       Win32::GuiTest::WMGetText($p1)  '     %     %     inttostring                  perleval    '  %     	           perleditsettext          perlinit            %   %    editsettext    	      %             getfocus    '      		 Win32::GuiTest::WMSetText($p1, $p2) '        %     %     inttostring %              perleval      stringtoint '  %     	      �     perleval                perlinit         	                  %          stringchopright '      $  ohomerperl    %   %  %  %  %    eval    '  %     	      �     perlgetwmiproperty           getjawssettingsdirectory     \Homer\GetWMIProperty.pl    
  '     %    filetostring    '     %  %   %              perleval    '  %     	      �    perlinit    $  iperlinitialized      ����
              	      $  ohomerperl             	           reloadallconfigs               getjawssettingsdirectory     \Homer\PerlEval.dll 
    fileexists         ����&  iperlinitialized             	          Initializing component    saystring           getfocus    '       Perl.Eval           createobjectex  &  ohomerperl  %         %     setfocus          $  ohomerperl         Done!     saystring           &  iperlinitialized            	          Error!    saystring        ����&  iperlinitialized             	         �     perllistgetcontents         perlinit         	      %             getfocus    '       join '', Win32::GuiTest::GetListContents($p1)  '     %     %     inttostring                  perleval    '  %     	      �     perllistgettext         perlinit         	      %             getfocus    '      		 Win32::GuiTest::GetListText($p1)    '     %     %     inttostring                  perleval    '  %     	          perlrunexamplecode             %  %  %  %  %    perleval    '  %      p4= %  
   
  
  %  
  '     %      p3= %  
   
  
  %  
  '     %      p2= %  
   
  
  %  
  '     %      p1= %  
   
  
  %  
  '     %        %         stringchopright '        %   
    stringcontains      
  %  
  '     %      Parameters:     %  
   

    
  '        %   ;    ;
   stringreplaceallcase    '  %        %         stringchopright '        %   
    stringcontains      
  %  
  '     %      Code:   %  
   

    
  '        %   
   
     stringreplaceallcase    '     %   
    
    stringreplaceallcase    '     %   

     
    stringreplaceallcase    '     %   
    stringcontains      
  %  
  '      Result:     %  
  '  %  %  
  %  
  '   Task:   %   
  '      %   %    perldialoginfo        �     perlrunexamplefile             %    filetostring    '     %   %  %  %  %  %    perlrunexamplecode        L    perlshellurltofile         %    filedelete       	       $UrlToFile = new Win32::API('urlmon.dll', 'URLDownloadToFileA', 'NPPNN', 'N');  '  %  

 $UrlToFile->Call(0, $p1, $p2, 0, 0);    
  '        %  %   %              perleval      stringtoint '     %    fileexists  '  %     	          perlwindowfind             perlinit         	       (Win32::GuiTest::FindWindowLike($p1, $p2, $p3, $p4))[0] '        %     %     inttostring %  %     %    inttostring   perleval      stringtohandle  '  %     	      �     perlwindowgetactive      perlinit         	      		 Win32::GuiTest::GetActiveWindow()   '         %                         perleval      stringtohandle  '  %     	      �     perlwindowgetdesktop         perlinit         	      		 Win32::GuiTest::GetDesktopWindow()  '         %                         perleval      stringtohandle  '  %     	      �     perlwindowgetforeground      perlinit         	      

 Win32::GuiTest::GetForegroundWindow()   '         %                         perleval      stringtohandle  '  %     	      �     perlwindowsetactive         perlinit         	      

 Win32::GuiTest::SetActiveWindow($p1)    '        %     %     inttostring                  perleval      stringtoint '  %     	      �     perlwindowsetforeground         perlinit         	       Win32::GuiTest::SetForegroundWindow($p1)    '        %     %     inttostring                  perleval      stringtoint '  %     	          perlwindowsetstate           perlinit            %   %    windowsetstate     	      

 Win32::GuiTest::ShowWindow($p1, $p2)    '        %     %     inttostring    %    inttostring             perleval      stringtoint '  %     	      �     pyeval              saytoolsinit         	      $  ohomerst      %   %  %  %  %    eval    '  %     	      �     pyevalfile            %    \     stringcontains            getjawssettingsdirectory     \Homer\ 
  %   
  '         %     filetostring    '     %  %  %  %  %    pyeval  '  %     	      `     regexpcontainscase         %   %             regexpcontainsex       	      `     regexpcontainsequiv        %   %              regexpcontainsex       	      �    regexpcontainsex              VBScript.RegExp   objectcreate    '  %  %  !  pattern %    %  !  ignorecase        %  !  multiline         %  !  global  %    %     execute '  %      count   '  %     %            item    '  %      firstindex       
  '  %      value   ' 	    %    inttostring %  
  % 	 
  ' 
    %  '  %  '  %  '  % 
    	      h     regexpcontainslastcase         %   %             regexpcontainslastex       	      h     regexpcontainslastequiv        %   %              regexpcontainslastex       	      �    regexpcontainslastex              VBScript.RegExp   objectcreate    '  %  %  !  pattern %    %  !  ignorecase        %  !  multiline        %  !  global  %    %     execute '  %      count   '  %     %    %       
    item    '  %      firstindex       
  '  %      value   ' 	    %    inttostring %  
  % 	 
  ' 
    %  '  %  '  %  '  % 
    	      P     regexpcountcase        %   %         regexpcountex      	      T     regexpcountequiv           %   %          regexpcountex      	          regexpcountex            VBScript.RegExp   objectcreate    '  %  %  !  pattern %    %  !  ignorecase        %  !  multiline        %  !  global  %    %     execute '  %      count   '  %  '  %  '  %     	      \     regexpextractcase          %   %             regexpextractex    	      \     regexpextractequiv         %   %              regexpextractex    	      �    regexpextractex           VBScript.RegExp   objectcreate    '  %  %  !  pattern %    %  !  ignorecase        %  !  multiline        %  !  global  %    %     execute '  %      count   '        '  %  %  
     %    %    item    '  %      value   ' 	 % 
 % 	 
  %  
  ' 
 %       
  '   �       % 
        stringchopright ' 
 %  '  %  '  %  '  % 
    	      d     regexpreplacecase           %   %  %              regexpreplaceex    	      d     regexpreplaceequiv          %   %  %               regexpreplaceex    	      �     regexpreplaceex            VBScript.RegExp   objectcreate    '  %  %  !  pattern %    %  !  ignorecase        %  !  multiline   %  %  !  global  %    %   %    replace '  %  '  %     	      �     registryread           Wscript.Shell     objectcreate    '  %    %     regread '  %    %    expandenvironmentstrings    '  %  '  %     	      �     registrywrite           Wscript.Shell     objectcreate    '  %    %   %   REG_SZ    regwrite    '  %  '  %     	      \    sayfilebyline           Scripting.FileSystemObject    objectcreate    '  %    %     opentextfile    '  %      atendofstream     # �      iskeywaiting      
     %      readline    '     %    stringisblank           %  %    sayrepeat          |    %      close      %  '  %  '     d     sayifverbose            getverbosity         
        %     saystring            �     sayrepeat       %       
 	       %     spellphonetic         %        %     spellstring          %     saystring                saystringbyline        %    
     stringcountsegment  '       '  %  %  
  # �      iskeywaiting      
        %    
   %    stringsegment   '     %    stringisblank           %  %    sayrepeat         %       
  '   T       h     saystringif      %         %    saystring            %    saystring            p     saystringifobject        %         %    saystring            %    saystring            �     saytempfile      getjawssettingsdirectory     \   
       getactiveconfiguration  
   .tmp    
  '      %     filetostring    '     %    saystring         �    saytoolsinit    $  isaytoolsinitialized      ����
              	      $  ohomerst               	          Say.Tools          createobjectex  &  ohomerst    $  ohomerst            &  isaytoolsinitialized            	         

 Error initializing SayTools component!    saystring        ����&  isaytoolsinitialized             	         �     sayvirtual          userbufferdeactivate            userbufferclear       %     userbufferaddtext           userbufferactivate          jawstopoffile           sayall        t     sbox                 voicesetspeech  '     %     messagebox        %    voicesetspeech        0    shellcreateshortcut            WScript.Shell     objectcreate    '  %    %     createshortcut  '  %  %  !  targetpath  %  %  !  workingdirectory    %  %  !  windowstyle %  %  !  hotkey  %      save       %  '  %  '     %     fileexists     	      T    shellexec          Wscript.Shell           createobjectex  '       speechoff      %    %     exec    '  %      status        
               delay       �    %      stdout  '  %      readall '  %      terminate           speechon       %  '  %  '  %  '  %     	      h    shellgetdrives      Scripting.FileSystemObject    objectcreate    '    ABCDEFGHIJKLMNOPQRSTUVWXYZ  '       '  %       
        %  %         substring   '  %     %    driveexists    %     %    getdrive    '  %      isready    %  %  
  '     %  '     %       
  '   �    %  '   %     	      �     shellgetenvironmentvariable        Wscript.Shell     objectcreate    '  %      environment '  %    %     item    '  %  '  %  '  %     	      �     shellgetshortcuttargetpath         WScript.Shell     objectcreate    '  %    %     createshortcut  '  %      targetpath  '  %     	      �     shellgetwindowsname         getwindowssystemdirectory    \win.com    
   ProductName   getversioninfostring       	      `    shellgetwindowsntname        '    SOFTWARE\Microsoft\Windows NT\CurrentVersion    '   ProductName '             getjfwversion     inttostring        stringleft   6   
 	       %   %  %    getregistryentrystring     	          HKEY_LOCAL_MACHINE\ %  
   \   
  %  
    registryread       	         �     shellrun             Wscript.Shell           createobjectex  '  %    %   %  %    run '  %  '  %     	      t     shellruncommandline        %COMSPEC% /k cd "   %   
   "   
               shellrun          L     shellrunexplorer           "   %   
   "   
    run            shellsethomerdir         '    SOFTWARE\EmpowermentZone\Homer\CurrentVersion   '   UtilDir '       getjawssettingsdirectory     \Homer  
  '      HKEY_LOCAL_MACHINE\ %  
   \   
  %  
  %    registrywrite      	      <    shellurltofile         %    filedelete                	            '       '  %   sURL = "    
  %   
   "
  
  '  %   sFile = "   
  %  
   "
  
  '  %   adTypeBinary = 1
   
  '  %  		 Const adSaveCreateOverWrite = 2
    
  '  %   Const adSaveCreateNotExist = 1
 
  '  %   Set oXML = CreateObject("MSXML2.XMLHTTP")
  
  '  %  		 Call oXML.Open("GET", sURL, False)
 
  '  %   oXML.Send()
    
  '  %   Set oStream = CreateObject("Adodb.Stream")
 
  '  %   oStream.Type = adTypeBinary
    
  '  %   oStream.open()
 
  '  %  		 oStream.Write(oXML.ResponseBody)
   
  '  %   Call oStream.SaveToFile(sFile, adSaveCreateOverWrite)
  
  '  %   oStream.Close()
    
  '  %   Set oStream = Nothing
  
  '  %   Set oXML = Nothing
 
  '       pathgettempfile '     %  %    stringtofile        cscript.exe /e:vbscript        %    stringquote 
  '     %  %  %    shellrun          %    filedelete        %    fileexists     	      ,    spellansi         %     stringlength    '       pause           '  %  %  
  # �      iskeywaiting      
        %   %         substring   '     %    getcharactervalue   '     %    sayinteger               delay      %       
  '   T       X    spellphonetic       space|alpha|bravo|charlie|delta|echo|foxtrot|golf|hotel|india|juliette|kilo|lema    '  %   |mike|november|oscar|poppa|Quebec|romeo|sierra|tango|uniform|victor|Whiskey|Xray|yankee|zulu    
  '    abcdefghijklmnopqrstuvwxyz '     %     stringlength    '       '  %  %  
  # d     iskeywaiting      
        %   %         substring   '     %     %    stringlower   stringcontains  '  %        %   |   %    stringsegment   '        %    saystring                delay      %       
  '   4      �     ssay            isspeechoff         speechon          %     saystring           speechoff            %     saystring            �    statusbargettext               getfocus      getappmainwindow     msctls_statusbar32    findwindow  '   %         %           getwindowtext   '          savecursor          invisiblecursor         routeinvisibletopc          getrestriction  '            setrestriction          jawspagedown            getline '     %    setrestriction          restorecursor         %     	      |     stringaddsegment            %     stringisblank      %  '      %   %  
  %  
  '      %      	      �     stringappendtoclipboard          getclipboardtext    '       pause         %    stringisblank        %  %  
  %   
  '         %     copytoclipboard       �     stringappendtofile          %    fileexists        %    filetostring    %  
  %   
  '         %   %    stringtofile          %    fileexists     	      x     stringconcat         |   %   
  '  %  %  
  '     %         stringchopleft  '  %     	      t     stringcontainsequiv           %     stringlower    %    stringlower   stringcontains     	      �     stringcontainslast            %     stringreverse      %    stringreverse     stringcontains  '  %        %     stringlength    %  
     %    stringlength    
       
  '     %     	      �     stringcontainslastequiv           %     stringreverse      %    stringreverse     stringcontainsequiv '  %        %     stringlength    %  
     %    stringlength    
       
  '     %     	      P     stringcountsegment         %   %    stringsegmentcount     	      @    stringdeletesegment         %   %  %    stringsegment   '     %   %  %    stringsegmentstartcase  '     %    stringlength    '  %  %  
  '     %   %    stringcountsegment  '  %  %  
     %       
  '        %   %  %         stringreplacerange  '  %     	      �     stringequal        %   %    stringcontains  # |    %     stringlength       %    stringlength    
  
     	      t     stringequiv     %   %  
  # h    %     stringlength       %    stringlength    
  
     	      �     stringgetbounds         %     stringlength    '  %        
     %  %  
       
  '     %        
     %  %  
       
  '        %   %  %  %  
       
    substring      	      \     stringgetrange          %   %  %       
    stringgetbounds    	      P    stringgetvaluecharacter       @@ 	
 !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~�������������������������������������������������������������������������������������������������������������������������������� %          substring      	      L     stringif         %      %  '     %  '     %     	      d     stringindexsegmentcase          %   %  %         stringsegmentindex     	      h     stringindexsegmentequiv         %   %  %          stringindexsegmentex       	      H    stringindexsegmentex             %   %    stringcountsegment  '       '  %  %  
        %   %  %    stringsegment   '  %  # �    %  %    stringequal 
  " � %    # �    %  %    stringequiv 
  
     %  '  %       
  '     %       
  '      `    %     	      P     stringleadcase         %   %         stringleadex       	      P     stringleadequiv        %   %          stringleadex       	      �     stringleadex         %           %      %    stringlength      stringleft  %    stringequal    	            %      %    stringlength      stringleft  %    stringequiv    	         �    stringpadleft           %         stringequal     |   LL                                                                                                                                                                                                                                                                                                                 
  '        %   0     stringequal     |   LL 000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000    
  '      |   '       '  %  %  
     %  %  
  '  %       
  '            %  %   
  %    stringright '  %     	      �    stringpadright          %         stringequal    LL                                                                                                                                                                                                                                                                                                                  |   
  '        %   0     stringequal    LL 000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000     |   
  '      |   '       '  %  %  
     %  %  
  '  %       
  '            %   %  
  %    stringleft  '  %     	      �     stringplural           %    inttostring      
  %   
  '  %       
     %   s   
  '     %     	      �     stringproper             %          stringleft    stringupper       %           %     stringlength      substring     stringlower 
     	      �    stringproperwords         %          stringcountsegment  '       '  %  %  
        %        %    stringsegment   '     %    stringlength    '        %         stringleft    stringupper       %       %       
    substring     stringlower 
  '  %  %  
     %  %  
  '     %  %  
       
  '     %       
  '   T    %     	      T     stringquote     "      %     stringunquote   
   "   
     	      d     stringreplaceallcase            %   %  %         stringreplaceallex     	      d     stringreplaceallequiv           %   %  %          stringreplaceallex     	      x    stringreplaceallex        %  # T    %  %    stringcontains  
  " � %    # �    %  %    stringcontainsequiv 
  
        %   %  %  %    stringreplaceex    	      %  # �    %   %    stringcontains  
  " 0%    # ,   %   %    stringcontainsequiv 
  
        %   %  %  %    stringreplaceex '    �    %      	          stringreplacebounds          %     stringlength    '  %        
     %  %  
       
  '     %        
     %  %  
       
  '        %   %       
    stringleft  %  
     %   %  %  
    stringright 
     	      \     stringreplacecase           %   %  %    stringreplacesubstrings    	      \     stringreplaceequiv          %   %  %          stringreplaceex    	          stringreplaceex          %    stringlength    '  %   '  %  '     %    stringlength    '       '  %        %    stringlength    ' 	    %    stringlength    '  % 	 %  
  ' 
 %        %  %    stringcontains  '           %    stringlower    %    stringlower   stringcontains  '     %        %  %  %  
  %    substring   '     %       % 
 %  
       
    substring   %  
  %  
  '           '      �    %     	      l     stringreplacerange           %   %  %       
  %    stringreplacebounds    	      �     stringreplacesegment             %   %  %    stringsegment   '     %   %  %    stringsegmentstartcase  '     %    stringlength    '  %  %  
  '     %   %  %  %    stringreplacerange  '  %     	      <    stringreplacetokens    %   '   \r|\n|\t|\f '   
	    '     %   |     stringcountsegment  '       '  %  %  
        %   |   %    stringsegment   '     %  %         substring   '     %  %  %    stringreplacecase   '  %       
  '   �    %     	      �    stringsegmentfilter          AdoDb.RecordSet   objectcreate    '    �   '      '  %      fields  '  %     Item    %  %    append     %      open          %   %    stringcountsegment  '       '  %  %  
        %   %  %    stringsegment   ' 	 %     Item    % 	   addnew     %       
  '   �    %      update      Item LIKE ' %  
   '   
  %  !  filter     %   Item          ' 
 %      movefirst      %      eof      % 
     value   ' 	 %  % 	 
  %  
  '  %      movenext               %         stringchopright '  %      close      %  ' 
 %  '  %  '  %     	      X     stringsegmentsort          %   %          stringsegmentsortex    	      �    stringsegmentsortex          AdoDb.RecordSet   objectcreate    '    �   '      '  %      fields  '  %     Item    %  %    append     %      open          %   %    stringcountsegment  '       '  %  %  
        %   %  %    stringsegment   ' 	 %     Item    % 	   addnew     %       
  '   �    %      update     %      Item DESC   %  !  sort        Item    %  !  sort          %   Item          ' 
 %      movefirst      %      eof      % 
     value   ' 	 %  % 	 
  %  
  '  %      movenext              %         stringchopright '  %      close      %  ' 
 %  '  %  '  %     	      h     stringsegmentstartcase          %   %  %         stringsegmentstartex       	      h     stringsegmentstartequiv         %   %  %          stringsegmentstartex       	      �     stringsegmentstartex          %        %  %   
  %  
  %  %  
  %  
    stringcontains  '        %  %   
  %  
  %  %  
  %  
    stringcontainsequiv '     %     	      8    stringsettokens    %   '   \r|\n|\t|\f '   
	    '     %   |     stringcountsegment  '       '  %  %  
        %   |   %    stringsegment   '     %  %         substring   '     %  %  %    stringreplacecase   '  %       
  '   |    %     	      �    stringsortjkm         %     fileexists                	         %   %    .bak    
    filecopy            pathgettempfolder    \temp.ini   
  '     %    iniflush            pause         %    inideletefile           pause         %     inireadsectionnames '     %   |     stringsegmentsort   '     %   |     stringcountsegment  '       '  %  %  
        %   |   %    stringsegment   '     %  %     inireadsectionkeys  '     %   |     stringcountsegment  '       '  %  %  
        %   |   %    stringsegment   ' 	       %  % 	      %     inireadstring     stringtrim  ' 
    % 
        2     stringpadright     % 	   stringtrim  
  '  %  %  
   |   
  '  %       
  '   8      %         stringchopright '     %   |     stringsegmentsort   '     %   |     stringcountsegment  '       '  %  %  
        %   |   %    stringsegment   '        %    2     stringleft    stringtrim  ' 
    %    2     stringchopleft  ' 	    %  % 	 % 
 %         iniwritestring        %    iniflush       %       
  '   �   %       
  '   �      %    iniflush            pause         %    filetostring    '     %   (\r|\n)\s*(\r|\n+)   
    regexpreplacecase   '  %     	      D    stringsortjss         %     fileexists                	         %   %    .bak    
    filecopy          %     filetostring    '   \r\n(( |\t|\w)* +)*((Function)|(Script)) +\w+   '     %  %    regexpcontainsequiv '     %    stringisblank               	            %             stringsegment     stringtoint '     %  %       
    stringleft  '     %  %     %    stringlength      substring   '   (   %  
   )   
  '   

$1 '     %  %  %    regexpreplaceequiv   

   
  '  %   [^]*   
  '     %  %    regexpextractequiv  '       '     %        stringcountsegment  ' 	      ' 
 % 
 % 	 
        %      % 
   stringsegment   '           %   (     (    stringreplacecase     stringtrim    stringlower '     %              stringreplaceallcase    '     %              stringsegment   '     %              stringsegment   '     %              stringsegment   '     %   script    stringequal     s   %  
  '      f   %  
  '        %         2     stringpadright  '  %  %  
  %  
      
  '  % 
      
  ' 
  �      %         stringchopright '     %        stringsegmentsort   '       '     %        stringcountsegment  ' 	      ' 
 % 
 % 	 
        %      % 
   stringsegment   '     %    2     stringchopleft  '  %  %  
   

    
  '  % 
      
  ' 
  �      %         stringchopright '      '       '     %  %  %    regexpreplacecase   '   (
){3,}    '   

    '     %    stringtrim   

    
        %  %  %    regexpreplacecase     stringtrim  
   
  
  '  %     	      h     stringswaplistcase          %   %  %             stringswaplistex       	      h     stringswaplistequiv         %   %  %              stringswaplistex       	      4    stringswaplistex           %   '     %  %    stringcountsegment  '       '  %     %  %  
        %  %  %    stringsegment   '     %   %    stringcontains        %  %  %    stringsegment   ' 	    %  %  % 	   stringreplaceallcase    '     %       
  '   p       %  %  
        %  %  %    stringsegment   '     %   %    stringcontainsequiv       %  %  %    stringsegment   ' 	    %  %  % 	   stringreplaceallequiv   '     %       
  '   H      %     	          stringtofile           %    filedelete         Scripting.FilesystemObject    objectcreate    '       '  %    %  %    createtextfile  '  %    %     write      %      close         %    fileexists  '  %  '  %  '     %     	      P     stringtrailcase        %   %         stringtrailex      	      T     stringtrailequiv           %   %          stringtrailex      	      �     stringtrailex        %           %      %    stringlength      stringright %    stringequal    	            %      %    stringlength      stringright %    stringequiv    	         h     stringtrim           %     stringtrimtrailingblanks      stringtrimleadingblanks    	      �     stringunquote      %   '     %         stringleft   "   
        %         stringchopleft  '           %         stringright  "   
        %         stringchopright '   �    %     	      �     tabpageget     %                getfocus      getrealwindow   '          '     %   %                sendmessage    	      �     tabpageset      %                getfocus      getrealwindow   '          '     %   %  %          sendmessage    	      @    uielevateversion           Elevate Scripts   saystring         %     pathgetname '       pathgetinternetcachefolder   \   
  %  
  '   Download from
  '  %  %   
   
   
  '  %  

 To Internet cache, and run installer?   
  '          %   Y     dialogconfirm    Y   
     	          Please wait   saystring         %   %    shellurltofile           Error downloading file!   saystring      	          "   %  
   "   
              shellrun          %    filedelete            uigotoeditorcol         saveeditorposition     $  ihomereditorline    '  %   $  ihomereditorcol 
  '  %        
 	    %  '  %          nextcharacter           saveeditorposition     $  ihomereditorline    %  
       ����'     $  ihomereditorcol %   
 
      ����'     $  ihomereditorcol %   
           '     %       
  '      �         ����%  
  '  %          priorcharacter          saveeditorposition     $  ihomereditorline    %  
       ����'     $  ihomereditorcol %   
       ����'     $  ihomereditorcol %   
           '     %       
  '      �      $  ihomereditorline    %  
  # �$  ihomereditorcol %   
  
          '     %     	      �    uigotoeditorline            saveeditorposition     $  ihomereditorline    %   
             	          Control+G    Edit      
     keyboardtypeandwait               	            %     inttostring   typestring         Enter    VsTextEditPane    
     keyboardtypeandwait               	           saveeditorposition     $  ihomereditorline    %   
          '     %     	      �     uigotoeditorposition           %     uigotoeditorline                  	         %    uigotoeditorcol               	              	      �    uihandlehomerwindows                         	         %     getrealwindow   '        %    getwindowname    JAWS      stringequal            	              getappfilename   IniForm.exe   stringequiv    %   %  
             	               	              inhjdialog     %   %  
              	            %     getwindowclass   Static    stringequiv             	            %     getwindowclass   Edit      stringequiv        Edit      saystring              	              	             Internal     SuppressTitle          Homer.ini     inireadinteger     	         �    uiselecttoposition           saveeditorposition     $  ihomereditorline    %   
 
 " � $  ihomereditorline    %   
  # � $  ihomereditorcol %  
 
 
  
              	      %   $  ihomereditorline    
  '  $  ihomereditorline    %   
  #  %        
 
 
          selectnextline          saveeditorposition     %       
  '   �    $  ihomereditorline    %   
              	      %  $  ihomereditorcol 
  '  $  ihomereditorline    %   
  # $  ihomereditorcol %  
  
  # 0%        
 
 
          selectnextcharacter         saveeditorposition     %       
  '   �   $  ihomereditorline    %   
  # �$  ihomereditorcol %  
  
          '     %     	      �    uiwaitforcontrol         %        
 
 # X      iskeywaiting      
          '  %        
 
       %    |   %    stringsegment   '     %   |   %    stringsegment   '     %    stringisblank   #    %    stringisblank   
               delay      %       
  '    ����'        %    stringisblank         %    |          stringsegment   '        %    stringisblank         %   |          stringsegment   '        %    stringisblank   " h           getfocus      getwindowclass  %    stringequiv 
  # �   %    stringisblank   " �              getfocus      getcontrolid      inttostring %    stringequiv 
  
          '    ����'    ����'       pause                delay         %       
  '         h     $    %     	      �    uiwaitfortitleandactivate       %  # < %    
             %     findtoplevelwindow  '            delay      %       
  '   (    %  # �    %          getfocus      getrealwindow     intequal      
                  pause         %     appactivate         pause         %   Edit           findwindow  '  %        %    setfocus                  speechoff          JAWS      appactivatetitle            pause           speechon          %    setfocus            pause                          getfocus      gettoplevelwindow     getwindowname   %     stringequal    	          variantgetsubtype       Unknown '  %         %     stringlength          Object  '        %     stringlength          %     inttostring   stringlength    
      Int '      String  '         Null    '     %     	      (     varianttohandle    %      	      (     varianttoint       %      	      (     varianttoobject    %      	      (     varianttostring    %      	      �     vbseval             getjawssettingsdirectory     \Homer\Homer.wsc    
  '      script: %  
    getobject   '  %    %   %  %  %  %    vbseval '  %  '  %     	      �     vbsevalfile           %    \     stringcontains            getjawssettingsdirectory     \Homer\ 
  %   
  '         %     filetostring    '     %  %  %  %  %    vbseval '  %     	      �    voicesavesetting             getactiveconfiguration   .jcf    
  '   Global|Error|Keyboard|Screen|PCCursor|JAWSCursor|Message    '       '  %        %   |   %    stringsegment   '     %    stringisblank            '      eloq-   %  
   Context 
  '     %  %   %  %    iniwriteinteger    %       
  '      �       �     voicesetscreenecho          getscreenecho   '       getscreenecho   %   
          screenecho      8    %     	      �     voicesetspeech          isspeechoff   '  %           speechon               speechoff         %     	      �     voicesetverbosity           getverbosity    '       getverbosity    %   
          verbositylevel      8    %     	           webgeturltofile          getjawssettingsdirectory     \Homer\WebGet.exe   
  '     %    fileexists       	         %    pathgetshort    '       pathgettempfile '     %    pathgetfolder   '     %    folderexists            %    foldercreate             %    folderexists         	      %   
  
  %   
  '     %  %    stringtofile       %       
  %  
  '     %               shellrun          %    filedelete        %    fileexists     	      �     windowactivate      %                getfocus      getrealwindow   '           '     %   %  %          sendmessage    	      �     windowclose    %                getfocus      getrealwindow   '          '    `�  '     %   %  %          sendmessage    	      �     windowhide     %             getfocus    '          '        '     %   %  %          sendmessage    	      �     windowmaximize     %             getfocus    '          '    0�  '     %   %  %          sendmessage    	      �     windowminimize     %             getfocus    '          '     �  '     %   %  %          sendmessage    	      �     windownext     %             getfocus    '          '    @�  '     %   %  %          sendmessage    	      �     windowprevious     %             getfocus    '          '    P�  '     %   %  %          sendmessage    	      �     windowrestore      %             getfocus    '          '     �  '     %   %  %          sendmessage    	      �     windowsetstate      %             getfocus    '           '     %   %  %          sendmessage    	      |     windowshow     %             getfocus    '           '     %   %  %          sendmessage    	      �    $uideletedown        caretvisible    # H      ispccursor  
  # p      isvirtualpccursor     
         Delete down   saystring           speechoff           selecttobottom          pause          Delete    typekey         pause              getfocus      refreshwindow           pause           speechon            sayline            typecurrentscriptkey             �    $uideleteleft        caretvisible    # H      ispccursor  
  # p      isvirtualpccursor     
          speechoff           selectfromstartofline           pause          Delete    typekey         pause              getfocus      refreshwindow           pause           speechon            sayline            typecurrentscriptkey             �    $uideleteline        caretvisible    # H      ispccursor  
  # p      isvirtualpccursor     
          speechoff           pause           jawsend         jawshome            selecttoendofline           pause          Delete    typekey            getfocus      refreshwindow           pause           speechon            sayline            typecurrentscriptkey             �    $uideleteright       caretvisible    # H      ispccursor  
  # p      isvirtualpccursor     
          speechoff           selecttoendofline           pause          Delete    typekey         pause              getfocus      refreshwindow           pause           speechon            sayline            typecurrentscriptkey             �    $uideleteup      caretvisible    # D      ispccursor  
  # l      isvirtualpccursor     
         Delete up     saystring           speechoff           selectfromtop           pause          Delete    typekey         pause              getfocus      refreshwindow           pause           speechon            sayline            typecurrentscriptkey             t     $uisayandtypecurrentscriptkey        saycurrentscriptkeylabel            typecurrentscriptkey          D     $uitypecurrentscriptkey      typecurrentscriptkey          