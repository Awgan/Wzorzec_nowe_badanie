Attribute VB_Name = "Module4"
'funkcje ma³o u¿ywane lub wcale



Sub initAll()

    arrBadanieWiersze() = Array("Symbol Badania", "Nazwa Badania", "Nazwa alternatywna", "Kod ICD9", "Symbol Materia³u", "Nazwa Materia³u", "Grupa badañ", "Grupa do rejestracji", "Grupa do wydruku", "Czas oczekiwania")

End Sub

'Check existance of the NAME in RANGE
Function checkNameInRange(ByVal rng As Range, ByVal name As String) As Boolean

    If rng.Value = name Then
    
        checkNameInRange = True
    
    Else
    
        checkNameInRange = False
    
    End If

End Function

'Copy information from workbook with new badanie from sheet1 to array
'Parametry: arrWierszeNaglowka - predefiniowane nazwy wierszy z arkusza 1
'           wbName - nazwa skoroszytu
'Zwracana:tabela z informacjami o nag³ówkach wierszy i wartoœciach
'
Function NoweBadanie(ByRef arrWierszeNaglowki() As Variant, ByVal wbName As String) As Variant()

    Const cellAdd As String = "A1"
    
    Dim shNoweBadanie As Worksheet
    Set shNoweBadanie = Workbooks(wbName).Worksheets(strNoweBadanie)
    
    Dim rng As Range
    Set rng = shNoweBadanie.Range(cellAdd)
    
    Dim arrDimension As Integer
    arrDimension = shNoweBadanie.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim arrNoweBadanie()
    ReDim arrNoweBadanie(0 To arrDimension, 0 To 1)
        Dim i As Integer
        For i = LBound(arrNoweBadanie, 1) To UBound(arrNoweBadanie, 1)
            
            arrNoweBadanie(i, 0) = rng.Offset(i, 0).Value   'Nag³ówek z wiersza
            arrNoweBadanie(i, 1) = rng.Offset(i, 1).Value   'Wartoœæ z wiersza
         
        Next i
    
    NoweBadanie = arrNoweBadanie

End Function

'Wybierz symbole badañ z tabeli z Arkusza '1'("Nowe Badanie")
'Parametry funkcji: arrNoweBadania() - tabela z danymi z arkusza 1
'                   szukFraza - wskazuje, które informacje z tabeli arrNoweBadania funkcja ma przetarzaæ
'Zwraca: funkcja zwraca tabelê z symbolami nowych badañ
'
Function wybierzSymboleBadan(ByRef arrNoweBadania() As Variant, Optional ByVal szukFraza As String = "Symbol Badania") As Variant()
    
    Dim symbole()
    Dim j, k, u As Integer
    j = 0
    k = 0
    
    Dim i As Integer
    For i = 0 To UBound(arrNoweBadania)
    
        If arrNoweBadania(i, 0) = szukFraza Then
                        
            Dim str As String
            str = arrNoweBadania(i, 1)
            
            Dim paczka() As String
            paczka = Split(str, " ")
            
            'Debug.Print "lbound(paczka): " & LBound(paczka)
            'Debug.Print "ubound(paczka): " & UBound(paczka)
            
            j = k + UBound(paczka)
            'Debug.Print "j: " & j
            
            ReDim Preserve symbole(j)
            
            Dim h As Integer
            h = 0
            
            For u = k To j
                'Debug.Print "u: " & u
                symbole(u) = paczka(h)
                h = h + 1
            Next u
            
            k = j + 1
            'Debug.Print "k: " & k
            'Debug.Print "lbound(symbole): " & LBound(symbole)
            'Debug.Print "ubound(symbole): " & UBound(symbole)
            'Debug.Print

        End If
    Next i

wybierzSymboleBadan = symbole

End Function

Sub buttonsDisable()

    If "wysy³kowe" = "wysylkowe" Then
    
        Debug.Print "Tak"
    
    End If

End Sub
