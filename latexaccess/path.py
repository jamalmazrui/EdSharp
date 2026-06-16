'''Provide a function to get the program's directory, needed for tables.'''
import os.path

def get_path():
    return os.path.normpath(os.path.dirname(__file__))
    
