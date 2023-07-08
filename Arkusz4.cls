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

Dim str As String

str = "SELECT DISTINCT TRIM(m.SYMBOL) as SYM, TRIM(m.NAZWA) as NAZ, TRIM(a.SYMBOL) as APA FROM METODY m INNER JOIN APARATY a ON m.APARAT = a.ID WHERE m.SYMBOL LIKE 'X-%' AND m.DC > '2020-01-01' ORDER BY SYM, APA ASC;"
Call czysc_arkusz("PracownieWysylkowe")
Call GetDataFromWZORZEC(str, "PracownieWysylkowe", "A")

End Sub

Private Sub CommandButton2_Click()

Call usun_dup("PracownieWysylkowe")

End Sub
