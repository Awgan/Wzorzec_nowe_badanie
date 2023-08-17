VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

Application.ScreenUpdating = False

    Dim str As String
    
    str = "SELECT DISTINCT TRIM(m.SYMBOL) as SYM, TRIM(m.NAZWA) as NAZ, TRIM(a.SYMBOL) as APA FROM METODY m INNER JOIN APARATY a ON m.APARAT = a.ID WHERE m.SYMBOL LIKE 'X-%' AND m.DC > '2020-01-01' ORDER BY SYM, APA ASC;"
    
    Call czysc_arkusz("PracownieWysylkowe")
    
    If GetDataFromWZORZEC(str, "PracownieWysylkowe", "A") Then
        ThisWorkbook.Worksheets("Arkusz1").Range("B18").Value = "PracownieWysylkowe:"
        ThisWorkbook.Worksheets("Arkusz1").Range("C18").Value = CStr(Now)
    Else
    
    End If

Application.ScreenUpdating = True

Call usun_dup("PracownieWysylkowe")

End Sub

Private Sub CommandButton2_Click()

Call usun_dup("PracownieWysylkowe")

End Sub
