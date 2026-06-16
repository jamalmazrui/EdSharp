#This file defines a com object wihch can be used to interface between 
#screenreaders and the latex access scripts


import preprocessor
import nemeth
import speech

class latex_access_com:
    def __init__(self):
        self.nemeth_translator=nemeth.nemeth()
        self.speech_translator=speech.speech()
        self.preprocessor=preprocessor.preprocessor()
        self.newcommands=preprocessor.newcommands(self.preprocessor)
    _reg_clsid_ = "{436BC4EC-405C-49ED-A0E7-84945B0BAC03}"
    _reg_progid_ = "latex_access"
    _public_methods_ =["nemeth","speech","preprocessor_add","preprocessor_from_string","preprocessor_write","preprocessor_read","toggle_dollars_nemeth","toggle_dollars_speech"]
    def nemeth(self, input):
        '''Translates the input into Nemeth Braille.'''
        input=self.preprocessor.translate(input)
        return self.nemeth_translator.translate(input)
    
    def speech(self,input):
        '''Translates the input into english speech.'''
        input=self.preprocessor.translate(input)
        return self.speech_translator.translate(input)

    def toggle_dollars_nemeth(self):
        '''Toggles whether dollars are shown in braille.

        Returns a boolian of whether dollars are being removed.'''
        if(self.nemeth_translator.remove_dollars):
            self.nemeth_translator.remove_dollars=False
            return False
        else:
            self.nemeth_translator.remove_dollars=True
            return True

    def toggle_dollars_speech(self):
        '''Toggles whether dollars are spken 

        Returns a boolian of whether dollars are being removed.'''
        if(self.speech_translator.remove_dollars):
            self.speech_translator.remove_dollars=False
            return False
        else:
            self.speech_translator.remove_dollars=True
            return True

    def preprocessor_add(self,command,args,translation_string):
        '''A function to add entries to the preprocessor'''

        self.preprocessor.add_from_string(str(command),args,str(translation_string))

    def preprocessor_from_string(self,input):
        '''Adds preprocessor entries from a LaTeX string containing \newcommand.'''
        self.newcommands.translate(input)

    def preprocessor_write(self, filename):
        self.preprocessor.write(filename)

    def preprocessor_read(self, filename):
        self.preprocessor.read(filename)


#Register the object
if __name__=='__main__':
    import pythoncom,win32com.server.register
    latex_access_com._reg_clsid=pythoncom.CreateGuid()
    win32com.server.register.UseCommandLine(latex_access_com)
