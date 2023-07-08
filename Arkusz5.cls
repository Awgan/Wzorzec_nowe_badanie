VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

Dim str As String

str = "SELECT TRIM(s.SYMBOL) as SYM, TRIM(s.NAZWA) as NAZ FROM SYSTEMY s ORDER BY SYM ASC;"

Call czysc_arkusz("Systemy")

Call GetDataFromWZORZEC(str, "Systemy", "A")

End Sub
