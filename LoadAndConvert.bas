Attribute VB_Name = "LoadnConvert"
Sub LoadnConvert()

Dim pathPDF As String, textPDF As String
Dim openPDF As Object
Dim pathCell As Range, fileCell As Range
Dim objPDF As MsForms.DataObject
Dim textArray() As String
Dim i As Integer, j As Integer



'PATH IS RECEIVED FROM 1, 19 AND FILENAME FROM 2, 19
Set objPDF = New MsForms.DataObject
Set pathCell = ActiveSheet.Cells(1, 19)
Set fileCell = ActiveSheet.Cells(2, 19)
pathPDF = pathCell & fileCell

'IF USER PROVIDED THE FILENAME WITHOUT EXTENSION, ADD IT
i = InStrRev(pathPDF, ".pdf")
If i = 0 Then
    pathPDF = pathPDF & ".pdf"
End If

'ENSURE THAT FILE EXISTS
If Dir(pathPDF) = "" Then
    MsgBox "The filename you provided could not be found!"
Else
    Set openPDF = CreateObject("Shell.Application")
    openPDF.Open (pathPDF)
    'TIME TO WAIT BEFORE/AFTER COPY AND PASTE SENDKEYS
    Application.Wait Now + TimeValue("00:00:2")
    SendKeys "^a"
    Application.Wait Now + TimeValue("00:00:2")
    SendKeys "^c"
    Application.Wait Now + TimeValue("00:00:1")



    AppActivate ActiveWorkbook.Windows(1).Caption
    objPDF.GetFromClipboard
    textPDF = objPDF.GetText(1)
    textArray = Split(textPDF, vbNewLine)


    j = 1
    For Each Row In textArray
        Dim Col() As String
        'IMPORTANT: DELIMITER THAT SPLITS THE DATA INTO CELLS
        Col = Split(Row, " ")
        For i = LBound(Col) To UBound(Col)
            ActiveSheet.Cells(j, i + 1) = Col(i)
        Next i
    j = j + 1
    Next

End If


Rem MsgBox textArray(2)


Rem AppActivate ActiveWorkbook.Windows(1).Caption
Rem ActiveSheet.Range("E6").Select
Rem SendKeys "^v"
Rem ActiveSheet.Paste Destination:=Worksheets("Sheet1").Range("A1:J10")


End Sub
