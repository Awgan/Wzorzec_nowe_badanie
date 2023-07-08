Attribute VB_Name = "Module3"
'Nazwy arkuszy pliku z nowym badaniem, teoretycznie powinny byæ takie
'
Public Const strNoweBadanie As String = "Nowe Badanie"
Public Const strPracownie As String = "Pracownie"
Public Const strKonfiguracja As String = "Konfiguracja"
Public Const strBadania As String = "Badania"
Public Const strPakiety As String = "Pakiety"
Public Const strPracownieWysylkowe As String = "PracownieWysylkowe"
Public Const strSystemy As String = "Systemy"


'Nazwy dla ka¿dego z trzech modu³ów pliku DAT
'
Public Const kolumna_1 As String = _
            "[Metody]" & vbCrLf & _
            "*Symbol" & vbCrLf & _
            "*Badanie@Badania" & vbCrLf & _
            "Nazwa" & vbCrLf & _
            "Kod" & vbCrLf & _
            "Pracownia@Pracownie" & vbCrLf & _
            "Aparat@Aparaty" & vbCrLf & _
            "Koszt" & vbCrLf & _
            "Punkty" & vbCrLf & _
            "*BadanieZrodlowe@Badania-" & vbCrLf & _
            "MetodaZrodlowa@Metody&Metody.Badanie@Badania" & vbCrLf & _
            "Serwis" & vbCrLf & _
            "NiePrzelaczac" & vbCrLf & _
            "Grupa" & vbCrLf
            
Public Const kolumna_2 As String = _
            vbCrLf & vbCrLf & _
            "[ParametryWMetodach]" & vbCrLf & _
            "*Metoda@Metody&Metody.Badanie@Badania" & vbCrLf & _
            "*Parametr@Parametry&Parametry.Metoda@Metody&Metody.Badanie@Badania" & vbCrLf & _
            "Dorejestrowywany" & vbCrLf & _
            "Kolejnosc" & vbCrLf & _
            "Format" & vbCrLf

Public Const kolumna_3 As String = _
            vbCrLf & vbCrLf & _
            "[PowiazaniaMetod]" & vbCrLf & _
            "*Badanie@Badania" & vbCrLf & _
            "*DowolnyTypZlecenia" & vbCrLf & _
            "*TypZlecenia@TypyZlecen" & vbCrLf & _
            "*DowolnaRejestracja" & vbCrLf & _
            "*Rejestracja@Rejestracje" & vbCrLf & _
            "*DowolnySystem" & vbCrLf & _
            "*System@Systemy" & vbCrLf & _
            "Metoda@Metody&Metody.Badanie@Badania" & vbCrLf & _
            "InnaPracownia" & vbCrLf & _
            "Pracownia@Pracownie" & vbCrLf & _
            "*DoRozliczen" & vbCrLf & _
            "*DowolnyMaterial" & vbCrLf & _
            "*Material@Materialy" & vbCrLf & _
            "*DowolnyOddzial" & vbCrLf & _
            "*Oddzial@Oddzialy" & vbCrLf & _
            "*DowolnyPlatnik" & vbCrLf & _
            "*Platnik@Platnicy" & vbCrLf


