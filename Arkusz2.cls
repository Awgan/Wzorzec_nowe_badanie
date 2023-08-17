VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

Application.ScreenUpdating = False

    Dim str As String
    
    str = "SELECT TRIM(r.SYMBOL) as SYM, TRIM(r.NAZWA) as NAZ FROM BADANIA r WHERE r.PAKIET = '0' ORDER BY SYM ASC;"
    
    Call czysc_arkusz("Badania")
    
    If GetDataFromWZORZEC(str, "Badania", "A") Then
        ThisWorkbook.Worksheets("Arkusz1").Range("B16").Value = "Badanie:"
        ThisWorkbook.Worksheets("Arkusz1").Range("C16").Value = CStr(Now)
    Else
    
    End If

Application.ScreenUpdating = True

End Sub
