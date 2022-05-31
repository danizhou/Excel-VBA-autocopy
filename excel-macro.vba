Option Explicit
Sub mergefile()           'function name
Dim nomifile As String           'dichiarazione delle variabili
Dim tuttifile As Workbook
Dim lastrow As Range            'per trovare l'ultima riga
Dim findrownumb As Long             'variabile ultima riga

nomifile = Dir("C:\analisidati\")  'estrapoliamo e inseriamo nella variabile nomifile il risultato di dir
                                    'Dir serve per restituire il primo elemento della cartella
                                    
    Do Until nomifile = ""              'ciclo do finche non trova Dir = nulla
'    Debug.Print nomifile                'funzione per printare a debug il nome
    
    Set tuttifile = Workbooks.Open("C:\analisidati\" & nomifile) 'concatena il valore della cartella con il nome del file

    tuttifile.Worksheets(1).Select                  'seleziona il primo foglio
        Do                                          'inizio ciclo per selezione del file
'        Debug.Print ActiveSheet.Name                'printa il nome del file
        
         
         Set lastrow = ActiveSheet.Range("A:A").Find(What:="note", LookIn:=xlValues)  'per trovare la ultima riga
         findrownumb = lastrow.Row                                      ' associo al valore ultima riga
         ActiveSheet.Range("B11", "B" & findrownumb - 1).Copy    'seleziona tutti i valori
          'ActiveSheet.Range("B9").CurrentRegion.Select
       ' ActiveSheet.Range(ActiveCell, ActiveCell.End(xlDown)).Select
         'ActiveSheet.Range("B11").End(xlDown).Row
         Debug.Print findrownumb
         'Range(Range("A5"), Range("A5").End(xlDown)).Select
         ThisWorkbook.Worksheets(1).Range("A999999").End(xlUp).Offset(1).PasteSpecial
        
          
        If ActiveSheet.Next Is Nothing Then Exit Do         ' se il valore prossimo foglio è null allora esce
        ActiveSheet.Next.Select                             'se no seleziona il prossimo foglio
        Loop                                            'ciclo loop
    tuttifile.Close                                     ' chiude il file
    nomifile = Dir                                       ' restituisce il prossimo file
    
    Loop
    
End Sub
