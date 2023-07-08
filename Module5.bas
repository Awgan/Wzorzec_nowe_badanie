Attribute VB_Name = "Module5"
Option Explicit

Public ConnDB As ADODB.Connection

Sub ConnectDB()

'Open the connection to the DB if not already open

    Dim Password As String
    Dim Server_Name As String
    Dim port As String
    Dim User_ID As String
    Dim Database_Name As String
    Dim connOptions As String
    Dim ODBCDriver As String

    On Error GoTo FailedConnection
        
    ODBCDriver = "Firebird/InterBase(r) driver"
    Server_Name = ""
    Database_Name = "XXXX"
    User_ID = "XXXX"
    Password = "XXXX"
    port = "xxxx"

'   connOptions = ";OPTION=" 'Convert LongLong to Int


    If (ConnDB Is Nothing) Then

        Set ConnDB = New ADODB.Connection

    End If

    If Not (ConnDB.State = 1) Then 'if not opened; 1->Connection is opened

        ConnDB.Open "Driver={" & ODBCDriver & "};Server=" & _
        Server_Name & ";Port=" & port & ";Database=" & Database_Name & _
        ";Uid=" & User_ID & ";Pwd=" & Password & connOptions & ";"
    Else
    
        MsgBox "Nie mo¿na utworzyæ po³aczenia. Connection state: " & ConnDB.State
    
    End If

    Exit Sub

FailedConnection:

    MsgBox "Failed connecting to the DB. Please check DB settings.", vbOKOnly + vbCritical, "Database Error"

    Set ConnDB = Nothing

End Sub

Public Function GetDataFromWZORZEC(Optional strSQL As String = "", Optional strArku As String = "", Optional strKolu As String = "A")

    'Dim strSQL As String
    Dim rs As New ADODB.Recordset
    

    On Error GoTo FailedSub

    'strSQL = "Select * From Gwiazdy.Planety;"
    'Badania tylko, z BADANIA
    'strSQL = "SELECT r.ID, r.SYMBOL as SYM FROM BADANIA r WHERE r.PAKIET = '0' ORDER BY SYM ASC;"
    'Pakiety tylko, z BADANIA
    'strSQL = "SELECT r.ID, r.SYMBOL as SYM, r.PAKIET FROM BADANIA r WHERE (r.PAKIET = '1' AND r.DEL = '0') ORDER BY SYM ASC;"
    'Metody wys³kowe, z METODY
    'strSQL = "SELECT DISTINCT m.SYMBOL as SYM, a.SYMBOL as APA FROM METODY m INNER JOIN APARATY a ON m.APARAT = a.ID WHERE m.SYMBOL LIKE 'X-%' ORDER BY SYM, APA ASC;"
    'Systemy, z SYSTEMY
    'strSQL = "SELECT s.SYMBOL as SYM, s.NAZWA as NAZ FROM SYSTEMY s ORDER BY SYM ASC;"
    
    'Note we need not qualify the schema name, as we specified our schema in the connection string already!

    ConnectDB
    
    rs.Open strSQL, ConnDB
 

    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim sh As Worksheet
    If strArku = "" Then
        Set sh = wb.Worksheets("Arkusz2")
    Else
        Set sh = wb.Worksheets(strArku)
    End If
    
    Dim kolumnaNaWyniki As String
    kolumnaNaWyniki = CStr(strKolu & "1")
    Dim rng As Range
    Set rng = sh.Range(kolumnaNaWyniki)
    
'    'Kopiowanie do array
'    rs.MoveFirst
'    Dim arrPaczka() As Variant
'    arrPaczka = rs.GetRows
'    arrPaczka = Application.WorksheetFunction.Transpose(arrPaczka)
    
     'Wersja I, do arkusza
'    rs.MoveFirst
'
'    Dim i As Integer
'    i = 0
'
'    Do While Not rs.EOF
'
'        'Debug.Print rs!Symbol
'        rng.Offset(i, 0).Value = rs!SYM
'        rng.Offset(i, 1).Value = rs!APA
'        i = i + 1
'
'        rs.MoveNext
'
'    Loop
    
    'Wersja II, do arkusza
    rs.MoveFirst
    Call rng.CopyFromRecordset(rs)
    
    
CloseSub:

    If Not (rs Is Nothing) Then

       If (rs.State = 1) Then   '1->Recordset is opened

            rs.Close

        End If

    End If

    DisconnectDB

    Exit Function

FailedSub:

    MsgBox "Failed reading from the Database.", vbOKOnly + vbCritical, "Database Error"

    GoTo CloseSub

End Function

Sub DisconnectDB()

'Close the connection to the DB if open

 

    On Error GoTo FailedConnection

   

    If Not (ConnDB Is Nothing) Then

        If Not (ConnDB.State = 0) Then  '0-> Connection is closed

            ConnDB.Close

        End If

    End If

 

    Exit Sub

FailedConnection:

    MsgBox "Failed closing the DB connection.", vbOKOnly + vbCritical, "Database Error"

    Set ConnDB = Nothing

End Sub



