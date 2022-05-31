Option Explicit
Sub mergefile()           'function name
Dim namefile As String           'declaration variable
Dim allfile As Workbook
Dim lastrow As Range            'to find the last row
Dim findrownumb As Long             'last row variable

namefile = Dir("C:\analisidati")  'taking out all the file in this directory [**PLEASE INSERT YOUR DIRECTORY**}
                                    'Dir is a command to view the first file inside a directory
                                    
    Do Until namefile = ""              'do cycle until he find the last element
'    Debug.Print namefile                'debud print
    
    Set allfile = Workbooks.Open("C:\analisidati" & namefile) 'concat the link to the directory and the file

    allfile.Worksheets(1).Select                  'select the first worksheet 
        Do                                          'star cycle inside the worksheet
'        Debug.Print ActiveSheet.Name                'debug print 
        
         
         Set lastrow = ActiveSheet.Range("A:A").Find(What:="note", LookIn:=xlValues)  'in this case i need to find the last row, on this model it have "note" inside all the file
         'range that select all the elements inside the column A
         findrownumb = lastrow.Row                                      ' associate row number with the variables
         ActiveSheet.Range("B11", "B" & findrownumb - 1).Copy    'select all the element in the range
        '[] or you can use this other function to select element from a cell to other element in the all column
          'ActiveSheet.Range("B9").CurrentRegion.Select [** insert the cell coordinate**]
       ' ActiveSheet.Range(ActiveCell, ActiveCell.End(xlDown)).Select 
         '[]



         ' Debug.Print findrownumb                            
         ThisWorkbook.Worksheets(1).Range("A999999").End(xlUp).Offset(1).PasteSpecial 'this will paste all the elements in the first column 
        
          
        If ActiveSheet.Next Is Nothing Then Exit Do         'if the next elements is null it exit
        ActiveSheet.Next.Select                             'if no it will open the next worksheet
        Loop                                            
    allfile.Close                                     ' close the file
    namefile = Dir                                       ' associate the next file name to the variable
    
    Loop
    
End Sub
