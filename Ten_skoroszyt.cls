VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ten_skoroszyt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

Sheets(1).CommandButton2.Enabled = False
Sheets(1).CommandButton3.Enabled = False
Sheets(1).CommandButton4.Enabled = False
Sheets(1).CommandButton5.Enabled = False
Sheets(1).Range("E1:E13").Clear
Sheets(1).Range("A1:A13").Clear

End Sub

Sub x()

Debug.Print Left(Application.UserLibraryPath, InStr(Application.UserLibraryPath, "\App"))

End Sub

