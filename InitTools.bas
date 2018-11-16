Attribute VB_Name = "InitTools"
Sub clearSheet()
    ActiveSheet.Range("R1:S2").Clear
    ActiveSheet.CommandButton1.Visible = False
    ActiveSheet.CommandButton2.Visible = False
    ActiveSheet.CommandButton3.Visible = False
    ActiveSheet.CommandButton4.Visible = False
End Sub


Sub resetSheet()

 delimiter = " "
 ' DELIMITER BUTTON CAPTION CHANGE
 ActiveSheet.CommandButton4.Caption = "Default Delimiter: (space)"

With Worksheets("Sheet1").Range("R1:R2")
 .Font.Name = "Cambria"
 .Font.Size = 14
 .Font.Bold = True
End With

Worksheets("Sheet1").Range("R1").Value = "Path:"
Worksheets("Sheet1").Range("R2").Value = "Filename:"
Worksheets("Sheet1").Range("S1").Value = ActiveWorkbook.Path & "\"
Worksheets("Sheet1").Range("S2").Value = Dir(ActiveWorkbook.Path & "\*.pdf")

ActiveSheet.CommandButton1.Visible = True
ActiveSheet.CommandButton2.Visible = True
ActiveSheet.CommandButton3.Visible = True
ActiveSheet.CommandButton4.Visible = True

End Sub

