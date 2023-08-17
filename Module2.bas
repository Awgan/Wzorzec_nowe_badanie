Attribute VB_Name = "Module2"
Option Explicit

'Global Info:
'Funkcje do wyci�gni�cia danych o badaniach i pracowniach oraz funkcje do przygotowania WzorzecAK2


'Copy information from sheet2("Pracownie") to array
'Systemy i Pracownie w pliku konfiguracyjnym dostarczane s� w formie tabeli poziomej
'Funkcja zamienia tabel� poziom� na tabel� pionow�
'Funkcja zwraca tabel� zawieraj�c� nazw� badania oraz pracownie w systemach
'Argument funkcji 'wbName' to nazwa skoroszytu, w kt�rym znajduje si� arkusz o nazwie 'Pracownie'
'       nazwa arkusza 'Pracownie' sprawdzana jest przy otwieraniu skoroszytu
'
Function Pracownie(ByVal wbName As String) As Variant()

Debug.Print "[ ] pracownie"
    
    Const cellAdd As String = "A1"
    
    Dim shPracownie As Worksheet
    Set shPracownie = Workbooks(wbName).Worksheets(strPracownie)
    
    Dim arrPracownie() As Variant
        
        
        
    
    
    'Kiedy jeden wiersz zawiera mniej pracowni ni� drugi, to generuje si� b��dny plik WzorzecAK2
    'Znale�� wiersz o najwi�kszej liczbie pracowni
    
    Dim lastColumn, lastRow As Integer
    
    lastColumn = 0
    lastRow = shPracownie.Cells(Rows.Count, 1).End(xlUp).Row    'Pierwsza kolumna to kolumna z badaniami
    
'Debug.Print "lastRow: " & lastRow
    
    Dim tempRow, tempColumn As Integer
    tempColumn = 0
    
    For tempRow = 2 To lastRow
'Debug.Print "tempRow: " & tempRow
        tempColumn = shPracownie.Cells(tempRow, Columns.Count).End(xlToLeft).Column
        
        If tempColumn > lastColumn Then
            lastColumn = tempColumn
        End If
'Debug.Print "lastColumn: " & lastColumn
    Next tempRow
    
    If lastColumn = 0 Then
        MsgBox "Error: function Pracownie(): B��d obliczenia liczby kolumn"
        Exit Function
    End If
    
    
    
    
    'Wiersze maj� d�ugo�� liczby kolumn, a kolumny wierszy
    ReDim arrPracownie(0 To lastColumn - 1, 0 To lastRow - 1) As Variant
    
    Dim rng As Range
    Set rng = shPracownie.Range(cellAdd)
        
        '   Zamiana Wierszy na Kolumny i,j = j,i           i1      i2     i3
        '   *-------*-------*-------*------*----      *-------*-------*------*------*---
        '   |Badanie|system1|system2|...   |...    j1 |Badanie|Bad1   |Bad2  |      |
        '   *-------*-------*-------*------*----      *-------*-------*------*------*---
        '   |Bad1   |X-WYS1 |X-WYS2 |X-WYS3|...  ---> |system1|X-WYS1 |...   |      |
        '   *-------*-------*-------*------*----      *-------*-------*------*------*---
        '   |Bad2   |...    |...    |...   |...    j3 |system2|X-WYS2 |...   |      |
        '   *-------*-------*-------*------*----      *-------*-------*------*------*---
        Dim i, j As Integer
        For i = 0 To lastColumn - 1
            For j = 0 To lastRow - 1
                arrPracownie(i, j) = rng.Offset(j, i).Value
                'arrPracownie(i_row, j_column) = rng.Offset(j_row, i_column).Value
            Next j
        Next i
    
    Pracownie = arrPracownie
    
Debug.Print "[x] pracownie"
End Function



'Pick unique values 'Metody Wysy�kowe' from array 'arrPracownie()' and return
Function pickDistinctValue(ByRef arrPracownie() As Variant) As Variant
    
Debug.Print "[ ] pickDistinctValue"
'Debug.Print "arrPracownie L dim1: " & LBound(arrPracownie, 1)
'Debug.Print "arrPracownie U dim1: " & UBound(arrPracownie, 1)
'Debug.Print "arrPracownie L dim2: " & LBound(arrPracownie, 2)
'Debug.Print "arrPracownie U dim2: " & UBound(arrPracownie, 2)
        
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim paczka() As Variant
    'ReDim paczka(UBound(arrPracownie, 1) * UBound(arrPracownie, 2))
    
    '###ToDo:
    'Usprawni� mechanizm wyznaczania wielko�ci tabeli, teraz jest teraz po sparata�sku
    'Teraz tabela jest nadmiarowa i istnieje du�o pustych miejsc,
    'kt�re trzeba korygowa� parametrem 'k' przy wpisywaniu warto�ci do arkuszy Metody, ParametryWMetodach, Powi�zanieMetod
    'parametr 'k' zwi�ksza swoj� warto�� przy ka�dym wyst�pieniu pustego pola
    'a nast�pnie warto�� 'k' odejmowana jest od inkrementowanego 'i'
    'Mo�liwe, �e nie ma innej metody, poniewa� mo�e zdarzy� si�, �e
    'jedno badanie b�dzie mia�o r�ne pracownie dla r�znych system�w
    
    'Tabela nadmiarowa, poniewa� trudno przewidzie� liczb� r�nych pracowni
    ReDim paczka(UBound(arrPracownie, 1), UBound(arrPracownie, 2) - 1) 'Odejmujemy, poniewa� nie interesuje nas pierwsza kolumna z Systemami
    
    Dim h As Integer
    h = 0
        
    'Przeszukujemy tabel� 'arrPracownie' w poszukiwaniu pojedynczych symboli pracowni oraz badania
    'Uwzgl�dniamy r�wnie� symbol badania, aby rozr�ni� w nowej tabeli zakresy pracowni
    Dim i, j As Integer
    For j = LBound(arrPracownie, 2) + 1 To UBound(arrPracownie, 2)
        
'SPRAWDZENIE
        'Je�eli PAKIET lub nieBADANIE, to id� do nast�pnej kolumny

        If czy_pakiet(CStr(arrPracownie(0, j)), WB_menu.Worksheets(strPakiety)) _
            Or _
            czy_badanie(CStr(arrPracownie(0, j)), WB_menu.Worksheets(strBadania)) = False Then

            GoTo next_j
        End If
        
        dict.RemoveAll
        
        For i = LBound(arrPracownie, 1) To UBound(arrPracownie, 1)
   

            'Odrzu�, je�eli pusta pracownia lub je�eli pierwszy wiersz lub r�ne od "X-" lub X-GENCZvX-GENET(pracownie wy��czone)
'SPRAWDZENIE
            If arrPracownie(i, j) = "" _
            Or _
                i <> LBound(arrPracownie, 1) _
                And _
                czy_pracownia_wysylkowa(CStr(arrPracownie(i, j)), WB_menu.Worksheets(strPracownieWysylkowe)) = False Then
'
'
'
'                (i <> LBound(arrPracownie, 1) And Left(CStr(arrPracownie(i, j)), 2) <> "X-") _
'            Or _
'                (i <> LBound(arrPracownie, 1) And Left(CStr(arrPracownie(i, j)), 7) = "X-GENCZ") _
'            Or _
'                (i <> LBound(arrPracownie, 1) And Left(CStr(arrPracownie(i, j)), 7) = "X-GENET") Then
                        
                'Do nothing, as it is not 'metoda wysy�kowa'
                        
            Else
                'Umie�� metody wysy�kowe w s�owniku
                dict(arrPracownie(i, j)) = 1
                
            End If
            
            
        Next i
        
        
        'Umie�� elementy ze s�ownika w tabeli
        h = 0
        Dim fred As Variant
        
        'Warto�ci umieszczane w tabeli kolumnami
        'Ka�d� seri� pracowni poprzedza symbol badania, do kt�rego pracownie nale��
        For Each fred In dict.keys
            'Debug.Print "h= " & h & " dist: " & fred
            paczka(h, j - 1) = fred
            h = h + 1
        Next
        
next_j:
    Next j
    
    'return
    pickDistinctValue = paczka
    
'For Each fred In paczka
''    If fred = "" Then GoTo continueLoop
''    End If
'    Debug.Print "paczk: " & fred
'continueLoop:
'Next

Debug.Print "[x] pickDistinctValue"

End Function

'DODAJ METODY
'
'Uzupe�nij Metody Wysy�kowe w arkuszu "Metody"
Sub dodajMetody(ByVal wbName As String, ByRef arrTypyPrac() As Variant)
            
Debug.Print "[ ] dodajMetody"
'Debug.Print Workbooks(wbName).Worksheets("pracownie wysy�kowe").range("B5").Value

    
    
    Dim lastRow As Integer
    lastRow = Workbooks(wbName).Worksheets("Metody").Cells(Rows.Count, 1).End(xlUp).Row + 1
    If lastRow < 4 Then 'Ograniczenie, aby nie skasowa� dw�ch pierwszych wierszy
        lastRow = 4
    End If
    
    Dim cellAdd As String
    cellAdd = "C" & lastRow 'Zacznij w tej kom�rce
    
    
    
    errChecking = ""
    errChecking = errChecking & "Kom�rki zawieraj�ce b��d w ""Metody"":" & vbCrLf
    
    'Okre�lenie kom�rki pocz�tkowej
    Dim rng As Range
    Set rng = Workbooks(wbName).Worksheets("Metody").Range(cellAdd)
    
    Dim b As Integer
    b = 0

    Dim plikBadanie As String
    Dim i, j As Integer
'Debug.Print "j Ubound,2 : " & UBound(arrTypyPrac, 2)
'Debug.Print "i Ubound,1 : " & UBound(arrTypyPrac, 1)

    For j = 0 To UBound(arrTypyPrac, 2)
    'Dim k As Integer
    'k = 0   'Wsp�czynnik korekty przy pustych
        For i = 1 To UBound(arrTypyPrac, 1) 'Start from '1' as the first row '0' contains info about 'Badanie'

'SPRAWDZENIE
            If _
                arrTypyPrac(i, j) = "" Then
'            Or _
'                czy_pakiet(CStr(arrTypyPrac(b, j)), WB_menu.Worksheets("Pakiety")) Then 'ten warunek sprawdzany jest ju� w funkcji 'pickDistinctValue'
            
                'Do nothing
                'Debug.Print "i in If= " & i
                'GoTo continueLoop
                'k = k + 1 'lub Exit For
                Exit For
            Else
            
                'Akcja synchronizatora, '+' dodaj
                rng.Offset(i - 1, -2) = "+"
                'Musi by�, ale dlaczego?
                rng.Offset(i - 1, -1) = "1"
                'Symbol metody wysy�kowej
                rng.Offset(i - 1, 0) = arrTypyPrac(i, j)
                'Symbol Badania
                rng.Offset(i - 1, 1) = arrTypyPrac(b, j)
                'Nazwa metody wysy�kowej
'                rng.Offset(i - 1, 2) = Application.VLookup(rng.Offset(i - 1, 0).Value, _
'                Workbooks(wbName).Worksheets("pracownie wysy�kowe").Range("A:B"), 2, False)
                rng.Offset(i - 1, 2) = Application.VLookup(rng.Offset(i - 1, 0).Value, _
                ThisWorkbook.Worksheets("PracownieWysylkowe").Range("A:B"), 2, False)
                
                'Sprawdzanie b��d�w VLookUp
                If Application.WorksheetFunction.IsNA(rng.Offset(i - 1, 2)) Then
                
                    errChecking = errChecking & rng.Offset(i - 1, 2).Address & vbCrLf
                    rng.Offset(i - 1, 2).Interior.Color = RGB(255, 100, 100)
                
                End If
                
                'Kod, puste
                rng.Offset(i - 1, 3) = ""
                'Pracownia
                rng.Offset(i - 1, 4) = arrTypyPrac(i, j)
                'Aparat
'                rng.Offset(i - 1, 5) = Application.VLookup(rng.Offset(i - 1, 0).Value, _
'                Workbooks(wbName).Worksheets("pracownie wysy�kowe").Range("E:F"), 2, False)
                rng.Offset(i - 1, 5) = Application.VLookup(rng.Offset(i - 1, 0).Value, _
                ThisWorkbook.Worksheets("PracownieWysylkowe").Range("A:C"), 3, False)
                
                'Sprawdzanie b��d�w VLookUp
                If Application.WorksheetFunction.IsNA(rng.Offset(i - 1, 5)) Then
                
                    errChecking = errChecking & rng.Offset(i - 1, 5).Address & vbCrLf
                    rng.Offset(i - 1, 5).Interior.Color = RGB(255, 100, 100)
                
                End If
                
                'Koszt, puste
                rng.Offset(i - 1, 6) = ""
                'Punkty, puste
                rng.Offset(i - 1, 7) = ""
                'Badanie �rod�owe
                rng.Offset(i - 1, 8) = "WYSYLKA"
                'Metoda �r�d�owa
                rng.Offset(i - 1, 9) = "WYSYLKA"
                'Badanie
                rng.Offset(i - 1, 10) = "WYSYLKA"
                'Serwis, puste by� powinno, brak w pliku na podstawie, kt�rego stworzono to makro
                rng.Offset(i - 1, 11) = ""
                'Nieprze��cza�, zero
                rng.Offset(i - 1, 12) = "0"
                'Grupa, powinno by� puste, brak w pliku na podstawie, kt�rego stworzono to makro
                'rng.Offset(i - 1, 13) = ""
            
            End If
                        
        Next i
'continueLoop:
        
        'Zamiana punktu odniesienia kom�rki pocz�tkowej uwzgl�dniajac parametr korekty 'k'
        'i = i - 1 - k  'nie jest potrzebne bo uzywamy 'Exit For'
        i = i - 1
        Set rng = rng.Offset(i, 0)
        'Debug.Print "rng addr= " & rng.Address
    Next j
    
Debug.Print "[x] dodajMetody"

End Sub

'DODAJ PARAMETRY
'
'Uzupenij Metody Wysy�kowe w arkuszu "Parametr w metodach"
Sub dodajParametryWMetodach(ByVal wbName As String, ByRef arrTypyPrac() As Variant)

Debug.Print "[ ] dodajParametryWMetodach"

    
    
    Dim lastRow As Integer
    lastRow = Workbooks(wbName).Worksheets("ParametryWMetodach").Cells(Rows.Count, 1).End(xlUp).Row + 1
    If lastRow < 4 Then 'Ograniczenie, aby nie skasowa� dw�ch pierwszych wierszy
        lastRow = 4
    End If
    
    Dim cellAdd As String
    cellAdd = "B" & lastRow 'Zacznij w tej kom�rce



    errChecking = errChecking & "Kom�rki zawieraj�ce b��d w ""ParametryWMetodach"":" & vbCrLf
    
    
    
    'Okre�lenie kom�rki pocz�tkowej
    Dim rng As Range
    Set rng = Workbooks(wbName).Worksheets("ParametrywMetodach").Range(cellAdd)
    
    
    'To samo co w 'dodajMetody'
    Dim b As Integer
    b = 0
    
    Dim i, j As Integer
    For j = 0 To UBound(arrTypyPrac, 2)
    'Dim k As Integer
    'k = 0
        For i = 1 To UBound(arrTypyPrac, 1) 'Start from '1' as the first row contains inf about 'Badanie'
            
'SPRAWDZENIE
            If _
                arrTypyPrac(i, j) = "" Then
'            Or _
'                czy_pakiet(CStr(arrTypyPrac(b, j)), WB_menu.Worksheets("Pakiety")) Then 'ten warunek sprawdzany jest ju� w funkcji 'pickDistinctValue'
                'Debug.Print "i in If= " & i
                'GoTo continueLoop
                'k = k + 1
                Exit For
            Else
            
                'Akcja synchronizatora
                rng.Offset(i - 1, -1) = "+"
                'Metoda wysy�kowa
                rng.Offset(i - 1, 0) = arrTypyPrac(i, j)
                'Badanie
                rng.Offset(i - 1, 1) = arrTypyPrac(b, j)
                'Parametr
                rng.Offset(i - 1, 2) = "WYSYLKA"
                'Metoda
                rng.Offset(i - 1, 3) = "WYSYLKA"
                'Badanie
                rng.Offset(i - 1, 4) = "WYSYLKA"
                'Dorejestrowywany
                rng.Offset(i - 1, 5) = "0"
                'Kolejno��
                rng.Offset(i - 1, 6) = "0"
                'Format, powinno by� puste, brak w pliku na podstawie, kt�rego stworzono to makro
                'rng.Offset(i - 1, 7) = ""
            
            End If
           
        Next i
'continueLoop:
        
        'Zamiana punktu odniesienia kom�rki pocz�tkowej uwzgl�dniajac parametr korekty 'k'
        'i = i - 1 - k  'nie jest potrzebne bo uzywamy 'Exit For'
        i = i - 1
        Set rng = rng.Offset(i, 0)
        'Debug.Print "rng addr= " & rng.Address
    Next j
    
Debug.Print "[x] dodajParametryWMetodach"

End Sub

'DODAJ POWIAZANIA METOD
'
'Uzupenij Metody Wysy�kowe w arkuszu "Powiazania w metodach"
Sub dodajPowiazaniaMetod(ByVal wbName As String, ByRef arrTypyPrac() As Variant)
    
Debug.Print "[ ] dodajPowiazaniaMetod"
'Debug.Print "arr L dim1: " & LBound(arrTypyPrac, 1)
'Debug.Print "arr U dim1: " & UBound(arrTypyPrac, 1)
'Debug.Print "arr L dim2: " & LBound(arrTypyPrac, 2)
'Debug.Print "arr U dim2: " & UBound(arrTypyPrac, 2)
    
    
    
    Dim lastRow As Integer
    lastRow = Workbooks(wbName).Worksheets("PowiazaniaMetod").Cells(Rows.Count, 1).End(xlUp).Row + 1
    If lastRow < 4 Then 'Ograniczenie, aby nie skasowa� dw�ch pierwszych wierszy
        lastRow = 4
    End If
    
    Dim cellAdd As String
    cellAdd = "B" & lastRow 'Zacznij w tej kom�rce
    
    
        
    errChecking = errChecking & "Kom�rki zawieraj�ce b��d w ""PowiazaniaMetod"":" & vbCrLf
    
    
    'Okre�lenie kom�rki pocz�tkowej
    Dim rng As Range
    Set rng = Workbooks(wbName).Worksheets("PowiazaniaMetod").Range(cellAdd)

    Dim b As Integer
    b = 0
    
    Dim i, j As Integer
    
        '   |   j
        '---*-------*-------*------*------*---
        ' i |Badanie|Bad1   |Bad2  |      |
        '   *-------*-------*------*------*---
        '   |system1|X-WYS1 |...   |      |
        '   *-------*-------*------*------*---
        '   |system2|X-WYS2 |...   |      |
        '   *-------*-------*------*------*---
    
    ' 'j' to kolumny dla kolejnych bada�
    For j = 1 To UBound(arrTypyPrac, 2)
    Dim k As Integer
    k = 0
    i = 1
    
        If czy_pakiet(CStr(arrTypyPrac(b, j)), WB_menu.Worksheets("Pakiety")) = False Then
            
            ' 'i' to wiersze dla kolejnych system�w
            For i = 1 To UBound(arrTypyPrac, 1) 'Start from '1' as the first row contains inf about 'Badanie'
                
                'Je�eli System nie ma zdefiniowanej pracowni lub nie jest pracowni� Wysy�kow� "X-" lub jest "X-GENCZ lub "X-GENET"
                'Mo�e si� zdarzy�, �e niekt�re systemy nie maj� zdefiniowanej pracowni lub jest to pracownia lokalna dla Systemu
                'Wtedy omijamy i szukamy dalej
                'Po to jest parametr 'k', aby usuwa� przerwy w arkuszu WzorzecAK2
                'Warunek 'czy_pakiet' musi zosta�; nie tak jak w powy�szych. Ta funkcja nie u�ywa tabeli po przej�ciu przez 'pickDistinctValue'
                If _
                    arrTypyPrac(i, j) = "" _
                Or _
                    czy_pracownia_wysylkowa(CStr(arrTypyPrac(i, j)), WB_menu.Worksheets("PracownieWysylkowe")) = False _
                Or _
                    czy_system(CStr(arrTypyPrac(i, 0)), WB_menu.Worksheets("Systemy")) = False Then
                    'Debug.Print "i in If= " & i
                    'GoTo continueLoop
                    k = k + 1
                Else
                    
                    'Akcja synchronizatora
                    rng.Offset(i - 1 - k, -1) = "+"
                    'Badanie
                    rng.Offset(i - 1 - k, 0) = arrTypyPrac(b, j)
                    'Dowolny typ zlecenia, jedynka
                    rng.Offset(i - 1 - k, 1) = "1"
                    'Typ zlecenia, puste
                    rng.Offset(i - 1 - k, 2) = ""
                    'Dowolna rejestracja, jedynka
                    rng.Offset(i - 1 - k, 3) = "1"
                    'Rejestracja, puste
                    rng.Offset(i - 1 - k, 4) = ""
                    'Dowolny system, zero
                    rng.Offset(i - 1 - k, 5) = "0"
                    'System
                    rng.Offset(i - 1 - k, 6) = arrTypyPrac(i, 0)
                    'Metoda wysy�kowa
                    rng.Offset(i - 1 - k, 7) = arrTypyPrac(i, j)
                    'Badanie
                    rng.Offset(i - 1 - k, 8) = arrTypyPrac(b, j)
                    'Inna pracownia, zero
                    rng.Offset(i - 1 - k, 9) = "0"
                    'Pracownia, puste
                    rng.Offset(i - 1 - k, 10) = ""
                    'Do rozlicze�, zero
                    rng.Offset(i - 1 - k, 11) = "0"
                    'Dowolny materia�, jedynka
                    rng.Offset(i - 1 - k, 12) = "1"
                    'Materia�, puste
                    rng.Offset(i - 1 - k, 13) = ""
                    'Dowolny oddzia�, jeden
                    rng.Offset(i - 1 - k, 14) = "1"
                    'Oddzia�, puste
                    rng.Offset(i - 1 - k, 15) = ""
                    'Dowolny p�atnik, jeden
                    rng.Offset(i - 1 - k, 16) = "1"
                    'P�atnik, powinno by� puste, brak w pliku na podstawie, kt�rego stworzono to makro
                    'rng.Offset(i - 1 - k, 17) = ""
                
                    'Sprawdzenie, czy istnieje Badanie
    '                rng.Offset(i - 1 - k, 19) = Application.VLookup(rng.Offset(i - 1 - k, 8).Value, _
    '                Workbooks(wbName).Worksheets("pracownie dom 24.04.23").Range("A:B"), 2, False)
                    rng.Offset(i - 1 - k, 19) = Application.VLookup(rng.Offset(i - 1 - k, 8).Value, _
                    ThisWorkbook.Worksheets("Badania").Range("A:B"), 2, False)
                    
                    'Sprawdzenie b��d�w VLookUp
                    If Application.WorksheetFunction.IsNA(rng.Offset(i - 1 - k, 19).Value) Then
                    
                        errChecking = errChecking & CStr(rng.Offset(i - 1 - k, 19).Address) & vbCrLf
                        rng.Offset(i - 1 - k, 19).Interior.Color = RGB(255, 100, 100)
                    End If
                    
                    'Sprawdzenie, czy istnieje System
    '                rng.Offset(i - 1 - k, 21) = Application.VLookup(rng.Offset(i - 1 - k, 6).Value, _
    '                Workbooks(wbName).Worksheets("pracownie dom 24.04.23").Range("P:P"), 1, False)
                    rng.Offset(i - 1 - k, 21) = Application.VLookup(rng.Offset(i - 1 - k, 6).Value, _
                    ThisWorkbook.Worksheets("Systemy").Range("A:A"), 1, False)
                           
                    'Sprawdzenie b��d�w VLookUp
                    If Application.WorksheetFunction.IsNA(rng.Offset(i - 1 - k, 21).Value) Then
                    
                        errChecking = errChecking & CStr(rng.Offset(i - 1 - k, 21).Address) & vbCrLf
                        rng.Offset(i - 1 - k, 6).Interior.ColorIndex = 3
                        rng.Offset(i - 1 - k, 21).Interior.ColorIndex = 3
                    End If
                
                End If
               
            Next i
    'continueLoop:
        End If
        
        'Zamiana punktu odniesienia kom�rki pocz�tkowej uwzgl�dniajac parametr korekty 'k'
        i = i - 1 - k
        Set rng = rng.Offset(i, 0)
        'Debug.Print "rng addr= " & rng.Address
    Next j


Debug.Print "[x] dodajPowiazaniaMetod"

End Sub

