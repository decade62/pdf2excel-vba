VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub commandbutton1_Click()

MsgBox "Please wait until the transfer is complete."

    LoadnConvert

End Sub

Private Sub CommandButton2_Click()
    clearSheet
End Sub


Private Sub CommandButton3_Click()
    ActiveSheet.UsedRange.ClearContents
    resetSheet
End Sub

Private Sub CommandButton4_Click()
Dim delimInput As Variant
delimInput = InputBox("Please provide the new delimiter: ")
delimiter = CStr(delimInput)
ActiveSheet.CommandButton4.Caption = "Current delimiter: " & delimiter
End Sub
