#This is part of the latex-access project.  Its purpose is to process matrices in LaTeX into a python array so that they can be manipulated.


class matrix:
    '''A class to store a matrix.'''
    def __init__(self):
        #This list contains the elements of the matrix.  It is a list of lists of strings.
        self.elements=[]
        self.rows=0;
        self.columns=0;
    def tex_init(self,tex):
        '''This function initialises the matrix from standard LaTeX source.

        That is, the parapemeter should be a string of matrix entries delimited by & and \\'''
        #Clear all members, to stop things going wrong
        self.rows=0
        self.columns=0
        self.elements=[]
        tex=str(tex)
        #1st split the matrix into rows.
        rows=tex.split(r"\\")
        self.rows=len(rows)
        #Remove a final blank row, if it exists.  This results from a trailing \\.
        if rows[self.rows-1]=="":
            rows.pop()
            self.rows-=1
        #Now split into columns
        for row in rows:
            entries=row.split("&")
            self.elements.append(entries)
            if len(entries) > self.columns:
                self.columns=len(entries)
        #Now we must expand each row up to to the length of the longest row.

    def get_cell(self,i,j):
        '''Returns the element in the ith row, jth column.'''
        return self.elements[i-1][j-1]

    def get_row(self,i,delim):
        '''Returns the ith row as a string, delimited by delim'''
        return str(delim).join(self.elements[i-1])
    
    def get_col(self,i,delim):
        '''Returns the ith column as a string delimitted by delim.'''
        j=0
        s=""
        while (j < self.rows):
            s+=self.elements[j][i-1]
            s+=str(delim)
            j+=1
        return s
    #Variables required to make this a com object
    _reg_clsid_ ="{723E7D16-7052-403F-976F-DF68B94BD936}"
    _reg_progid_ = "latex_access_matrix"
    _public_methods_ =["tex_init","get_cell","get_row","get_col"]
    _public_attrs_=["rows","columns"]

#Register the object
if __name__=='__main__':
    import pythoncom,win32com.server.register
    matrix._reg_clsid=pythoncom.CreateGuid()
    win32com.server.register.UseCommandLine(matrix)

    
