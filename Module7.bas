Attribute VB_Name = "Module7"
'Dodaj przedrostek "lab:" do SYMBOL_LABU
'

Option Explicit


Function DodajAfiks(ByVal strSlowo As String, ByVal strAfiks As String, Optional ByVal boolPrefiks As Boolean = True) As String
'Afiks to grupa fragmentów wyrazów dodwanaych do wyrazu przed, po, w œrdoku itd
    
'SprawdŸ argumenty
    If strSlowo = "" Or strAfiks = "" Then
        DodajAfiks = "Error:Brak s³owa lub afiksu"
        Exit Function
    End If
    
    
    
'Oczyœæ z bia³ych znaków
    strSlowo = Trim(strSlowo)
    strAfiks = Trim(strAfiks)
    
    
    
'Twórz nwe s³owo
    If boolPrefiks = True Then
        'Dodaj przedrostek
        strSlowo = strAfiks + strSlowo
    Else
        'Dodaj przyrostek
        strSlowo = strSlowo + strAfiks
    End If
    
    DodajAfiks = strSlowo

End Function
