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

Public Function czy_jest_w_zakresie(s As String, sh_arkusz As Worksheet, Optional kolu As String = "A") As Boolean

    'Element 's' oraz arkusz musz¹ istnieæ
    If sh_arkusz Is Nothing Or s = "" Then
    
        Exit Function
    
    End If
    
    'Wyznaczanie ostaniego wiersza
    Dim last As Integer
    last = sh_arkusz.Cells(Rows.Count, sh_arkusz.Columns(kolu).Column).End(xlUp).Row
    
    'Wyznaczanie zakresu do przeszukiwania
    Dim findRng As Range
    Dim tmp As String
    tmp = kolu + CStr(1) + ":" + kolu + CStr(last)
    
    Set findRng = sh_arkusz.Range(tmp)
    
    'Wyszukiwanie elementu 's' w okreœlonym zakresie
    Dim adres As Range
    Set adres = findRng.Find(s, , xlValues, xlWhole, xlByRows, xlNext, False)
    
    'Return funkcji
    If adres Is Nothing Then
    
        czy_jest_w_zakresie = False
    Else
        czy_jest_w_zakresie = True
        
    End If


End Function

Public Function czy_pakiet(s As String, sh_pakiety As Worksheet, Optional kolu As String = "A") As Boolean

    czy_pakiet = czy_jest_w_zakresie(s, sh_pakiety, kolu)

End Function

Public Function czy_system(s As String, sh_system As Worksheet, Optional kolu As String = "A") As Boolean

    czy_system = czy_jest_w_zakresie(s, sh_system, kolu)

End Function

Public Function czy_pracownia_wysylkowa(s As String, sh_pracownia As Worksheet, Optional kolu As String = "A") As Boolean

    czy_pracownia_wysylkowa = czy_jest_w_zakresie(s, sh_pracownia, kolu)

End Function

'Sortowanie wskazanej kolumny 'kolu' z arkusza 'sh_arkusz'
'
Public Function sortuj(sh_arkusz As Worksheet, Optional kolu As String = "A")

'Wyznacza ostatni wiersz oraz ostani¹ kolumnê
Dim lastRow, lastColumn As Integer
lastRow = sh_arkusz.Cells(Rows.Count, sh_arkusz.Columns(kolu).Column).End(xlUp).Row
lastColumn = sh_arkusz.Cells(1, sh_arkusz.Columns.Count).End(xlToLeft).Column

'Wyci¹ganie symbolu literowego(np. 'C') ostatniej kolumny danych do sortowania
Dim tmp, nextKolu As String
nextKolu = Split(Cells(1, lastColumn).Address, "$")(1)

'Ustalanie zakresu do sortowania; wskazuje kilka kolumn, najczêœciej bêd¹ to dwie kolumny
tmp = kolu + CStr(1) + ":" + nextKolu + CStr(lastRow)

Dim rng As Range
Set rng = sh_arkusz.Range(tmp)

'Sortowanie
rng.Sort Key1:=Range(CStr(kolu + CStr(1))), order1:=xlAscending



End Function

Sub test()

'Debug.Print czy_pakiet("PZa", ThisWorkbook.Worksheets(strPakiety))
'Debug.Print czy_pakiet("ZAWODZI", ThisWorkbook.Worksheets(strSystemy))
Dim sdsd As String
sdsd = sortuj(ActiveWorkbook.Worksheets("PracownieWysylkowe"), "L")


    

End Sub


Sub czysc_arkusz(ByVal shName As String)

Dim wb As Workbook
Set wb = ThisWorkbook
Dim sh As Worksheet
Set sh = wb.Worksheets(shName)

sh.Cells.Clear

End Sub

Sub buttonsDisable()

    If "wysy³kowe" = "wysylkowe" Then
    
        Debug.Print "Tak"
    
    End If

End Sub
