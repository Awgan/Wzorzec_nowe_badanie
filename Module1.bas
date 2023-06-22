Attribute VB_Name = "Module1"
Option Explicit

Private flagKonfWczyt As Boolean
Private flagKonfPrzygot As Boolean
Private flagWzorPrzygot As Boolean

Private flagKolejneBadanie As Boolean

Public WB_menu As Workbook
Public WB_konf As Workbook

Public pathSlashType As String          'dla rozró¿nienia czy plik uruchamiany jest z OneDrivea
Public pathActual As String
Public pathCommon As String             'stworzenie jednej œcie¿ki dostêpu dla komputera i chmury

Public numerPliku As Integer

Public arrBadanieWiersze() As Variant   'nazwy wiersz z tabelki z pierwszego arkusza pliku konfiguracyjnego
Public plikKonfig As String             'plik z danymi o nowym badaniu
Public plikKonfigNazwa As String        'plik z danymi o nowym badaniu, nazwa pliku

Public Const wzorzecAK As String = "WzorzecAK2.xlsx"
Public Const plikDat As String = "Wzor_02042020.dat"

Public errChecking As String


Sub ttt()

Dim wbpath, cpath As String
'wbpath = Workbooks("construct_02.xlsm").Path
'MsgBox wbpath

MsgBox "flagKolejnebadanie: " & flagKolejneBadanie

'FileCopy wbpath & "\construct_01.xlsm", wbpath & "\c_copy.txt"

'cpath = "C:\Users\" & Environ$("Username") & "\OneDrive - Alab Laboratoria Sp. z o.o\Scripts\Wzorzec_nowe_badanie\"

'MsgBox CurDir
End Sub


'CommandButton 1
Sub wskaz_plik_konfiguracyjny()

flagKonfWczyt = False
flagKonfPrzygot = False
flagWzorPrzygot = False

    
    Set WB_menu = ActiveWorkbook
    
    
    
    'Okreœlenie pochodzenia pliku: komputer lub chmura
    If InStr(WB_menu.Path, "\") <> 0 Then
    
        pathSlashType = "\" 'dla pliku bezpoœrednio z komputera
    Else
    
        pathSlashType = "/" 'dla pliku z chmury, np. OneDrive, œcie¿ka zaczyna siê od http://....
    End If
    
    
    
    
    Dim konfFilePath As String
    
    Dim fd As Office.FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogOpen)
        fd.Filters.Clear
        fd.Filters.Add "Excel Files", "*.xlsx?", 1
        fd.AllowMultiSelect = False
        fd.InitialFileName = WB_menu.Path & pathSlashType
        
    If fd.Show = True Then
        konfFilePath = fd.SelectedItems(1)
        WB_menu.Sheets(1).Range("A4").Interior.ColorIndex = 4
    Else
        
        WB_menu.Sheets(1).Range("A1:A13").Clear
        WB_menu.Sheets(1).Range("A4").Interior.ColorIndex = 3
        WB_menu.Sheets(1).CommandButton2.Enabled = False
        WB_menu.Sheets(1).CommandButton3.Enabled = False
        WB_menu.Sheets(1).CommandButton4.Enabled = False
        Exit Sub
    End If
    
    
    
    
    'Przygotuj nazwê pliku
    'Przygotuj wspóln¹ œcie¿kê dostêpu do pliku

    If pathSlashType = "\" Then
        'Normalna œcie¿ka dostêpu dla Windowsów
        
        Dim lastBSlash As Integer
        lastBSlash = InStrRev(konfFilePath, "\")
        
        pathActual = WB_menu.Path
        
        'Redukcja do nazwy pliku z rozszerzeniem
        konfFilePath = Right(konfFilePath, Len(konfFilePath) - lastBSlash)
    
    Else
        'Przypadek, kiedy plik wskazywany jest jako pochodz¹cy z 'OneDrivea'
        
        Dim lastSlash As Integer
        lastSlash = InStrRev(konfFilePath, "/")
        
        'Sk³adanie œcie¿ki z dwóch: CurDir daje œcie¿kê z liter¹ dysku na komputerze
        'do momentu nazwy udzia³u w chmurze OneDrive(koñczy siê na '/Documents');
        'Zmienna 'konfFilePath' daje dalsz¹ œcie¿kê ju¿ z chmury
        'Odcinamy odpowednie fragmenty i sklejamy
        Dim tempName As String
        tempName = Left(konfFilePath, lastSlash - 1)
        pathActual = Left(CurDir, InStr(CurDir, "\Documents") - 1) & _
        Replace(Right(tempName, Len(tempName) - InStrRev(tempName, "/Documents/") - 9), "/", "\")
        
        'Redukcja do nazwy pliku z rozszerzeniem
        konfFilePath = Right(konfFilePath, Len(konfFilePath) - lastSlash)
    
    End If
    'Debug.Print "pathActual: " & pathActual

        



    plikKonfig = konfFilePath
    plikKonfigNazwa = Left(plikKonfig, InStrRev(plikKonfig, ".") - 1)
    
    
    
    WB_menu.Sheets(1).Range("E1").Value = pathActual        'dla debugowania
    WB_menu.Sheets(1).Range("E2").Value = WB_menu.Path
    WB_menu.Sheets(1).Range("E3").Value = plikKonfig
    WB_menu.Sheets(1).Range("E4").Value = plikKonfigNazwa   'dla debugowania
    WB_menu.Sheets(1).CommandButton2.Enabled = False
    WB_menu.Sheets(1).CommandButton3.Enabled = False
    WB_menu.Sheets(1).CommandButton4.Enabled = False
    WB_menu.Sheets(1).Range("A7:A13").Clear
    
    
    
    flagKonfWczyt = True
    WB_menu.Sheets(1).CommandButton2.Enabled = True


End Sub


'CommandButton 2
Sub otworz_plik_konfiguracyjny()


Dim flagWB As Boolean
flagWB = False

Dim wName As Workbook
        
For Each wName In Workbooks

    If wName.name = plikKonfig Then
        Debug.Print "wName: JEST"
        flagWB = True
    End If
Next wName



    
    If flagKonfWczyt Then
    
        '#KOPIOWANIE PLIKU
        'Je¿eli pracujesz z OneDrive, to skrypt nie zadzia³a w tej wersji.
        'Problemy z manipulacj¹ plikami za pomoc¹ funkcji VBA
        '
        If Left(pathActual, 4) = "http" Then
        
            MsgBox "Kopia pliku konfiguracyjnego badnia nie zosta³a utworzona!" & vbCrLf & vbCrLf _
            & "Skrypt nie jest przygotowany do pracy z OneDrive oraz innymi rozwi¹zaniami chmurowymi." & vbCrLf _
            & "Je¿eli potrzebujesz kopi zapasowej, to u¿ywaj plików bezpoœrednio na dysku komputera.", _
            vbExclamation, "Kopia Zapasowa"
            
        Else
        
            '###ToDo:
            'Stwórz kopiê pliku, aby zosta³ orygina³
    
            Dim pKonf, eKonf As String
            pKonf = Left(plikKonfig, InStrRev(plikKonfig, ".") - 1) '
            eKonf = Right(plikKonfig, 4)
                        
                        
            'Je¿eli workbook jest otwarty, to kopiowanie nie uda siê
            'Kopiowanie orygina³u, plik musi byæ zamkniêty
            If Not flagWB Then
                'Kopiowanie chmura
                'FileCopy nie radzi sobie z OneDrive,problem z protoko³em https://
                
                'Kopiowanie komputer
                
                FileCopy CStr(pathActual & "\" & plikKonfig), CStr(pathActual & "\" & pKonf & "-oryginal." & eKonf)
            Else
                MsgBox "Plik jest otwarty, nie mo¿na zrobiæ kopii przed edycj¹.", , "Kopia Zapasowa"
            End If
        End If
        '#KOPIOWANIE PLIKU: KONIEC
        
        
        
        If flagWB = False Then
            Set WB_konf = Workbooks.Open(pathActual & pathSlashType & plikKonfig)
            'flagK = True
            Debug.Print "WB_konf.name #1: " & WB_konf.name
            
        Else
                
            Set WB_konf = Workbooks(plikKonfig)
            WB_konf.Activate
            Debug.Print "WB_konf.name #2: " & WB_konf.name
        
        End If
        
        
        
        'Je¿eli brak pliku, to zakoñcz funkcjê
        If WB_konf Is Nothing Then
            flagKonfPrzygot = False
            WB_menu.Sheets(1).CommandButton2.Enabled = False
            WB_menu.Sheets(1).CommandButton3.Enabled = False
            WB_menu.Sheets(1).CommandButton4.Enabled = False
            WB_menu.Sheets(1).Range("A7").Interior.ColorIndex = 3
            MsgBox "Brakuje pliku konfiguracyjnego", , "Brak Pliku Konfiguracyjnego"
            Exit Sub
        End If
        
        
        flagKonfPrzygot = True
        WB_menu.Sheets(1).CommandButton3.Enabled = True
        WB_menu.Sheets(1).Range("A7").Interior.ColorIndex = 4
    Else
        flagKonfPrzygot = False
        WB_menu.Sheets(1).CommandButton2.Enabled = False
        WB_menu.Sheets(1).CommandButton3.Enabled = False
        WB_menu.Sheets(1).CommandButton4.Enabled = False
        WB_menu.Sheets(1).Range("A7").Interior.ColorIndex = 3
        MsgBox "Brakuje pliku konfiguracyjnego", , "Brak Pliku Konfiguracyjnego"
        Exit Sub
    End If
    

    
    Dim flagShPrac As Boolean
    flagShPrac = False

    'SprawdŸ, czy plik zawiera arkusz o nazwie 'Pracownie'
    Dim sh As Worksheet
    For Each sh In WB_konf.Worksheets
    
        If sh.name = "Pracownie" Then
            flagShPrac = True
            Exit For
        End If
        
    Next sh

  
    
    
    If flagShPrac Then
                        
        Dim shPracownie As Worksheet
        Set shPracownie = WB_konf.Worksheets("Pracownie")
                        
    'Usuñ formatowanie wszystkich komórek w arkuszu Pracownie
        shPracownie.Cells.ClearFormats
        shPracownie.Columns.AutoFit
        WB_konf.Activate
        shPracownie.Activate
        ActiveWindow.FreezePanes = False
                
        
        Dim lastColumn, lastrow As Integer
        lastColumn = shPracownie.Cells(2, Columns.Count).End(xlToLeft).Column - 1 'zaczynamy od 2 wiersza, poniewa¿ w 1 wierszu zdarzaj¹ siê artefakty, np. w postaci zer(0)
        lastrow = shPracownie.Cells(Rows.Count, 1).End(xlUp).Row - 1
        
        Dim rng As Range
        Set rng = shPracownie.Range("A1")
        
        Dim i, j As Integer
        For i = 0 To lastColumn
            For j = 1 To lastrow
        
            'SprawdŸ b³êdne komórki i zaznacz na czerwono
                If rng.Offset(j, i).Value = "" Or Left(CStr(rng.Offset(j, i).Value), 2) <> "X-" Or Left(CStr(rng.Offset(j, i).Value), 7) = "X-LIMBA" Then
                    
                    'Debug.Print "Jest b³êdna komórka"
                    rng.Offset(j, i).Interior.Color = RGB(255, 100, 100)
        
                End If
        
            Next j
        Next i
        
        WB_konf.Activate
        shPracownie.Range("A1").Select
    
'Wyœwietl komunikat, ¿e brakuje arkusza o nazwie 'Pracownie' i zakoñcz sub-a
    Else
        
        flagKonfPrzygot = False
        WB_menu.Activate
        WB_menu.Sheets(1).CommandButton3.Enabled = False
        WB_menu.Sheets(1).Range("A7").Interior.ColorIndex = 3
        MsgBox "Brak arkusza: ""Pracownie""" & vbCrLf & _
        "Zmieñ nazwê odpowiedniego arkusza na ""Pracownie"" ", , "Brak Arkusza PRACOWNIE"
        Exit Sub
    
    End If

End Sub


'CommandButton 3
Sub stworzAK2()
    
    
    'initAll
    
    Dim WB_AK2 As Workbook
    Dim flagWAK2 As Boolean
    flagWAK2 = False
    
'SprawdŸ czy istnieje WzorzeAK2
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.name = "WzorzecAK2.xlsx" Then
            Set WB_AK2 = Workbooks("WzorzecAK2.xlsx")
            flagWAK2 = True
            Exit For
        End If
    Next

    If flagWAK2 = False Then
    
        'Dim wbpath As String
        'wbpath = WB_konf.path
        'On Error Resume Next
        Set WB_AK2 = Workbooks.Open(pathActual & pathSlashType & "WzorzecAK2.xlsx")
        flagWAK2 = True
    End If
    
    If WB_AK2 Is Nothing Then
        MsgBox "Brakuje pliku ""WzorzecAK2"" ", , "Brak Pliku WZORZECAK2"
        Exit Sub
    End If
  

'Czy dopisaæ kolejne badania do pliku WzorzecAK2
Dim odp As Integer
odp = 0
If Not flagKolejneBadanie Then

    odp = 7

Else
    
    odp = MsgBox("Czy dodaæ kolejne badania do pliku WzorzecAK2?", vbYesNo)
    
End If
  

If odp = 7 Then     '7 = No; formatowanie pliku WzorzecAK2

'Czyszczenie pliku WzorzecAK2
    Dim lastrow
    lastrow = WB_AK2.Worksheets("Metody").Cells(Rows.Count, 1).End(xlUp).Row
    If lastrow < 4 Then 'Ograniczenie, aby nie skasowaæ dwóch pierwszych wierszy
        lastrow = 4
    End If
    WB_AK2.Worksheets("Metody").Range("A" & Cells(4, 1).Row & ":" & "P" & Cells(lastrow, 1).Row).Clear
    
    lastrow = WB_AK2.Worksheets("ParametryWMetodach").Cells(Rows.Count, 1).End(xlUp).Row
    If lastrow < 4 Then 'Ograniczenie, aby nie skasowaæ dwóch pierwszych wierszy
        lastrow = 4
    End If
    WB_AK2.Worksheets("ParametryWMetodach").Range("A" & Cells(4, 1).Row & ":" & "I" & Cells(lastrow, 1).Row).Clear
    
    lastrow = WB_AK2.Worksheets("PowiazaniaMetod").Cells(Rows.Count, 1).End(xlUp).Row
    If lastrow < 4 Then 'Ograniczenie, aby nie skasowaæ dwóch pierwszych wierszy
        lastrow = 4
    End If
    WB_AK2.Worksheets("PowiazaniaMetod").Range("A" & Cells(4, 1).Row & ":" & "W" & Cells(lastrow, 1).Row).Clear
  
Else
    flagKolejneBadanie = False
    '6 = Yes

End If
  
   
'Wyci¹gnij pojedyncze nazwy pracowni
    Dim disLabor()
    disLabor = pickDistinctValue(Pracownie(WB_konf.name))
    
   
    
'G³owne dzia³anie
    Call dodajMetody(WB_AK2.name, disLabor)

    Call dodajParametryWMetodach(WB_AK2.name, disLabor)

    Call dodajPowiazaniaMetod(WB_AK2.name, Pracownie(WB_konf.name))
    
    
'Aktywacja kolejnego Buttona 4
    flagWzorPrzygot = True
    WB_menu.Sheets(1).CommandButton4.Enabled = True
    WB_menu.Sheets(1).Range("A10").Interior.ColorIndex = 4
    
    WB_AK2.Activate
    WB_AK2.Worksheets("Metody").Activate
    WB_AK2.Worksheets("Metody").Range("A1").Select
    WB_AK2.Worksheets("ParametryWMetodach").Activate
    WB_AK2.Worksheets("ParametryWMetodach").Range("A1").Select
    WB_AK2.Worksheets("PowiazaniaMetod").Activate
    WB_AK2.Worksheets("PowiazaniaMetod").Range("A1").Select
    
    flagKolejneBadanie = True

    
    MsgBox errChecking

End Sub

'CommandButton 4
Sub stworzDat()

Dim WB_wz As Workbook
Set WB_wz = Workbooks(wzorzecAK)

Dim SH_Metody, SH_Parametry, SH_Powiazania As Worksheet
Set SH_Metody = WB_wz.Sheets("Metody")
Set SH_Parametry = WB_wz.Sheets("ParametryWMetodach")
Set SH_Powiazania = WB_wz.Sheets("PowiazaniaMetod")



Dim Wzorzec_dat As String
Wzorzec_dat = pathActual & "\" & plikDat



'If pathSlashType = "\" Then
'
'    Wzorzec_dat = pathActual & "\" & plikDat
'
'Else
'
'    '###ToDo:
'    '#Stworzyæ œcie¿ke dostêpu do pracy z chmur¹ np.: OneDrive
'
''    Wzorzec_dat = "C:\Users\" & Application.UserName & "\Desktop\" & plikDat
''    Debug.Print Wzorzec_dat
''    Wzorzec_dat = "C:\Users\" & "rszlachetka" & "\Desktop\" & plikDat
''    Debug.Print Wzorzec_dat
''    Wzorzec_dat = pathActual & "\" & plikDat
''    Debug.Print Wzorzec_dat
''    Wzorzec_dat = Left(Application.UserLibraryPath, InStr(Application.UserLibraryPath, "\App")) & "Desktop\" & plikDat
''    Debug.Print Wzorzec_dat
''    Wzorzec_dat = Left(Application.UserLibraryPath, InStr(Application.UserLibraryPath, "\App") - 1) & "\" & pathCommon & "\" & plikDat
''    Debug.Print Wzorzec_dat
'    Wzorzec_dat = Left(CurDir, InStr(CurDir, "\Documents") - 1) & Replace(Right(pathActual, Len(pathActual) - InStrRev(pathActual, "/Documents/") - 9), "/", "\") & "\" & plikDat
'    Debug.Print Wzorzec_dat
''    Exit Sub
'End If



Debug.Print "wzorzec_dat: " & Wzorzec_dat



numerPliku = FreeFile


If Wzorzec_dat <> "" Then
    Open Wzorzec_dat For Output As numerPliku
Else
    MsgBox "Brak œcie¿ki do pliku DAT", , "Plik DAT"
    Exit Sub
End If




'Metody
Print #numerPliku, kolumna_1

Dim lastrow, lastColumn As Integer
lastrow = SH_Metody.Cells(Rows.Count, 1).End(xlUp).Row
lastColumn = SH_Metody.Range("O4").Column

    Dim rng As Range
    Set rng = SH_Metody.Range("A4")

    Dim strLinia As String
    Dim i, j As Integer
    
    For i = 4 To lastrow
    strLinia = ""
        For j = 1 To lastColumn
        
            strLinia = strLinia + CStr(rng.Offset(i - 4, j - 1).Value)
            If j < lastColumn Then
                strLinia = strLinia + vbTab
            End If
        
        Next j
        Print #numerPliku, strLinia
    Next i
strLinia = strLinia + vbTab



'Parametry
Print #numerPliku, kolumna_2
    
lastrow = SH_Parametry.Cells(Rows.Count, 1).End(xlUp).Row
lastColumn = SH_Parametry.Range("H4").Column
Set rng = SH_Parametry.Range("A4")
    For i = 4 To lastrow
    strLinia = ""
        For j = 1 To lastColumn
            
            strLinia = strLinia + CStr(rng.Offset(i - 4, j - 1).Value)
            If j < lastColumn Then
                strLinia = strLinia + vbTab
            End If
        
        Next j
        Print #numerPliku, strLinia
    Next i
strLinia = strLinia + vbTab



'Powiazania
Print #numerPliku, kolumna_3

lastrow = SH_Powiazania.Cells(Rows.Count, 1).End(xlUp).Row
lastColumn = SH_Powiazania.Range("R4").Column
Set rng = SH_Powiazania.Range("A4")
    For i = 4 To lastrow
    strLinia = ""
        For j = 1 To lastColumn
        
            strLinia = strLinia + CStr(rng.Offset(i - 4, j - 1).Value)
            If j < lastColumn Then
                strLinia = strLinia + vbTab
            End If
        
        Next j
        Print #numerPliku, strLinia
    Next i
strLinia = strLinia + vbTab

Close numerPliku



Dim returnvalue As Integer
returnvalue = Shell("C:\Program Files\Notepad++\notepad++.exe " & """" & Wzorzec_dat & """", vbNormalFocus)
'Debug.Print "returnvalue: " & returnvalue
WB_menu.Sheets(1).Range("A13").Interior.ColorIndex = 4
End Sub



