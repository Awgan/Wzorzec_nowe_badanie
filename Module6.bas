Attribute VB_Name = "Module6"
'Wyszukaj duplikaty i wybierz ten, który ma X-WYSYL, inne usuñ
'Je¿eli nie ma X-WYSYL, to zostaw X-WY-WA

'Usuñ duplikaty w kolumnie, zaczynaj¹c od pierwszego wiersza
Function usun_dup(ByRef shName As String, Optional ByVal kol As String = "A")

    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim sh As Worksheet
    Set sh = wb.Worksheets(shName)
    
    Dim lastRow As Integer
    lastRow = sh.Cells(Rows.Count, 2).End(xlUp).Row
    
    Dim rng1, rng2 As Range
    Set rng1 = sh.Range(CStr(kol & "1:" & kol & lastRow))
    
    Dim strDups As String
    strDups = ""
    Dim arrTmp() As String
    
    
    
    For Each rng2 In rng1
        
        'Je¿eli obecne rng2 jest równe poprzedniemu, to pomiñ
'        If CStr(rng2.Value) = strDups Then
'
'            GoTo endfor
'
'        End If
        
        
        'Policz ile istnieje duplikatów dla obecnego rng2
        Dim dups As Integer
        dups = ile_dup(sh, rng2)
    
        'Debug.Print rng2.Value & " " & dups
        
        If dups > 1 Then
        
            strDups = rng2.Value
            ReDim arrTmp(dups, 2)
            Dim i As Integer
            
            'Zapisz duplikaty do array
            For i = 0 To dups - 1
            
                arrTmp(i, 0) = CStr(rng2.Offset(i, 1).Value) 'Nazwa Metody
                arrTmp(i, 1) = CStr(rng2.Offset(i, 2).Value) 'Aparat
                'Debug.Print arrTmp(i)
            
            Next i
            
            
            Dim strNazw, strXwy As String
            strXwy = ""
            
            
            
'            Dim a As Variant
'            a = ""
'            For Each a In arrTmp
'
'                'Je¿eli jest X-WYSYL, to u¿yj i pomiñ resztê.
'                If a = "X-WYSYL" Then
'                    strXwy = a
'
'                    'Debug.Print "Exit"
'                    Exit For
'                End If
'
'            Next a


            For i = 0 To dups - 1
            
                'Je¿eli jest X-WYSYL, to u¿yj i pomiñ resztê.
                If arrTmp(i, 1) = "X-WYSYL" Then
                    strXwy = arrTmp(i, 1)
                    strNazw = arrTmp(i, 0)

                    'Debug.Print "Exit"
                    Exit For
                End If
            
            Next i
            
            
            'Je¿eli brakuje X-WYSYL, to przypisz pierwszy aparat w kolejce
            If strXwy <> "X-WYSYL" Then
                strXwy = arrTmp(0, 1)
                strNazw = arrTmp(0, 0)
            End If
            
            
            'Wpisz poprawny aparat wysy³kowy
            rng2.Offset(0, 1).Value = strNazw
            rng2.Offset(0, 2).Value = strXwy
            
            
            'Usuñ resztê duplikatów
            Dim rngTemp As Range
            Set rngTemp = sh.Range(rng2.Offset(1, 0).Address & ":" & rng2.Offset(dups - 1, 1).Address)
            rngTemp.EntireRow.Delete
            
            
            'Debug.Print "m= " & strDups & " s= " & strXwy
                    
        End If
endfor:
    Next rng2


End Function

'Policz duplikaty w kolumnie, zaczynaj¹c od 'rn'
Function ile_dup(sh As Worksheet, rn As Range) As Integer

    ile_dup = 0
    
    Dim rntmp As Range
    Dim i As Integer
    i = 0
    
    Do
        i = i + 1
        Set rntmp = rn.Offset(i, 0)
    
    Loop Until rn.Value <> rntmp.Value
    
    ile_dup = i


End Function

