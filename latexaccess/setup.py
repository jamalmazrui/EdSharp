from distutils.core import setup
import py2exe
import sys

class Target:
    def __init__(self, **kw):
        self.__dict__.update(kw)

                    
latex_access_com= Target(description="latex_access com server",
name="latex_access_com",
                    modules=["latex_access_com"],
                    create_exe = False,
                    create_dll = True)
                    
                    
latex_access_matrix= Target(description="latex_access matrix com server",
name="latex_access_matrix",
                    modules=["matrix_processor"],
                    create_exe = False,
                    create_dll = True)
                    
                    
setup (
                    com_server=[latex_access_com,latex_access_matrix])





