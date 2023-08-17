Attribute VB_Name = "Module7"
'Dodaj przedrostek "lab:" do SYMBOL_LABU
'

Option Explicit


Function DodajAfiks(ByVal strSlowo As String, ByVal strAfiks As String, Optional ByVal boolPrefiks As Boolean = True) As String
'Afiks to grupa fragment�w wyraz�w dodwanaych do wyrazu przed, po, w �rdoku itd
    
'Sprawd� argumenty
    If strSlowo = "" Or strAfiks = "" Then
        DodajAfiks = "Error:Brak s�owa lub afiksu"
        Exit Function
    End If
    
    
    
'Oczy�� z bia�ych znak�w
    strSlowo = Trim(strSlowo)
    strAfiks = Trim(strAfiks)
    
    
    
'Tw�rz nwe s�owo
    If boolPrefiks = True Then
        'Dodaj przedrostek
        strSlowo = strAfiks + strSlowo
    Else
        'Dodaj przyrostek
        strSlowo = strSlowo + strAfiks
    End If
    
    DodajAfiks = strSlowo

End Function
