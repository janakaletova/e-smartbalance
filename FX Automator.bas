Attribute VB_Name = "FX Automator"
Option Compare Database
Option Explicit

' =====================================================================================
' MODUL: FX Automator (Súèas inovatívneho modulu Smart-Pairing)
' PROJEKT: e-smartbalance s.r.o.
' POPIS: Zabezpeèuje automatické sahovanie aktuálnych aj historickıch kurzovıch
'        lístkov z API Národnej banky Slovenska (NBS) priamo do relaènej databázy.
' =====================================================================================


' ==============================================================================
' PROMPT PRE UMELÚ INTELIGENCIU (AI) NA GENEROVANIE TOHTO KÓDU:
' Rola: Si expert na MS Access, VBA a integráciu REST API.
'
' Úloha: Vytvor VBA modul pre MS Access, ktorı automaticky sahuje kurzové
' lístky z API Národnej banky Slovenska (NBS) vo formáte XML.
' Vytvor funkciu 'NacitajKurzyNBS'.
' Funkcia musí naèíta údaje z nbs 'https://nbs.sk/export/sk/exchange-rate/{YYYY-MM-DD}/xml'.
' Následne musí vyparsova XML uzly (Cube) a uloi meny (currency) a kurzy (rate)
' do tabu¾ky 'Tbl_kurzy_nbs'.
' Ošetri slovenské desatinné èiarky pri prevode na Double a skontroluj, èi
' kurzy pre danı deò u v databáze neexistujú.
' Na záver pridaj testovaciu procedúru 'Test_NacitajHistorickeKurzy'.
'
' KONTEXT DÁT (Struktúra databázy):
' 1. Tbl_kurzy_nbs: ID_kurzu (AutoNumber/PK), Time (Date/Time),
'    currency (Short Text - napr. USD, CZK), Rate (Number/Double).
' ==============================================================================

' -------------------------------------------------------------------------------------
' FUNKCIA: NacitajKurzyNBS
' ÚÈEL:    Dynamicky stiahne XML kurzovı lístok z NBS pre zadanı dátum a uloí ho.
' NÁVRATOVÁ HODNOTA: Boolean (True = úspech, False = chyba pripojenia alebo spracovania)
' PARAMETRE:
'   - datumUhrady (Date): Dátum, pre ktorı potrebujeme získa kurz.
'   - tichyRezim (Boolean): Ak je True, nehlási chyby ani úspech oknami (MsgBox).
' -------------------------------------------------------------------------------------
Function NacitajKurzyNBS(datumUhrady As Date, Optional tichyRezim As Boolean = False) As Boolean
    Dim http As Object
    Dim xmlDoc As Object
    Dim xmlNodes As Object
    Dim node As Object
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim url As String
    Dim datumZ_XML As Date
    Dim pocetExistujucich As Long
    Dim menaZ_XML As String
    Dim kurzRaw As String
    Dim kurzZ_XML As Double
    Dim sqlDatum As String
    Dim formatovanyDatumPreURL As String
    
    On Error GoTo ErrorHandler

    ' PREDVOLENİ STAV: Funkcia zaèína s predpokladom neúspechu
    NacitajKurzyNBS = False

    ' 1. PRÍPRAVA SPOJENIA S API NBS
    formatovanyDatumPreURL = Format(datumUhrady, "yyyy-mm-dd")
    url = "https://nbs.sk/export/sk/exchange-rate/" & formatovanyDatumPreURL & "/xml"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.send

    ' Kontrola dostupnosti servera
    If http.Status <> 200 Then
        If Not tichyRezim Then MsgBox "Chyba pripojenia k NBS pre dátum " & datumUhrady & "! (Status: " & http.Status & ")", vbCritical
        GoTo Cistka
    End If

    ' 2. SPRACOVANIE XML SÚBORU
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:ns='http://www.ecb.int/vocabulary/2002-08-01/eurofxref'"
    xmlDoc.loadXML http.responseText

    ' Extrakcia dátumu z XML
    On Error Resume Next
    Set node = xmlDoc.selectSingleNode("//ns:Cube[@time]")
    If Not node Is Nothing Then
        datumZ_XML = CDate(node.Attributes.getNamedItem("time").Text)
    End If
    On Error GoTo ErrorHandler

    If datumZ_XML = 0 Then
        If Not tichyRezim Then MsgBox "V XML sa nepodarilo nájs dátum pre " & datumUhrady & ".", vbCritical
        GoTo Cistka
    End If

    ' 3. KONTROLA EXISTUJÚCICH DÁT
    sqlDatum = Format(datumZ_XML, "mm\/dd\/yyyy")
    pocetExistujucich = DCount("*", "Tbl_kurzy_nbs", "[time] = #" & sqlDatum & "#")
    
    If pocetExistujucich > 0 Then
        If Not tichyRezim Then
            If MsgBox("Kurzy pre dátum " & datumZ_XML & " u v databáze sú. Prepísa?", vbQuestion + vbYesNo) = vbNo Then
                GoTo Cistka
            End If
        End If
    End If

    ' 4. ZÁPIS DO TABU¼KY
    Set xmlNodes = xmlDoc.selectNodes("//ns:Cube[@currency]")
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM Tbl_kurzy_nbs WHERE [time] = #" & sqlDatum & "#", dbOpenDynaset)

    For Each node In xmlNodes
        menaZ_XML = node.Attributes.getNamedItem("currency").Text
        kurzRaw = node.Attributes.getNamedItem("rate").Text
        
        ' Prevod s ošetrením slovenskej desatinnej èiarky
        kurzZ_XML = CDbl(Replace(kurzRaw, ".", ","))
        
        rs.FindFirst "[currency] = '" & menaZ_XML & "'"
        
        If rs.NoMatch Then
            rs.AddNew
            rs![Time] = datumZ_XML
            rs![currency] = menaZ_XML
        Else
            rs.Edit
        End If
        
        rs![Rate] = kurzZ_XML
        rs.Update
    Next node

    ' AK SME SA DOSTALI A SEM, VŠETKO PREBEHLO ÚSPEŠNE
    NacitajKurzyNBS = True
    If Not tichyRezim Then MsgBox "Import úspešne dokonèenı pre: " & datumZ_XML, vbInformation, "Hotovo"

Cistka:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    Set db = Nothing: Set xmlDoc = Nothing: Set http = Nothing
    Exit Function

ErrorHandler:
    If Not tichyRezim Then MsgBox "Chyba: " & Err.Description, vbCritical
    Resume Cistka
End Function

' -------------------------------------------------------------------------------------
' PROCEDÚRA: Test_NacitajHistorickeKurzy
' ÚÈEL:      Overenie funkènosti volania funkcie s rôznymi dátumami.
' -------------------------------------------------------------------------------------
Sub Test_NacitajHistorickeKurzy()
    Dim testDátum As Date
    Dim vysledok As Boolean
    
    testDátum = DateSerial(2023, 2, 15)
    
    Debug.Print "Testujem sahovanie pre: " & testDátum
    
    ' Volanie funkcie a spracovanie jej návratovej hodnoty
    vysledok = NacitajKurzyNBS(testDátum, True)
    
    If vysledok = True Then
        Debug.Print "TEST ÚSPEŠNİ: Dáta boli stiahnuté a uloené."
        MsgBox "Test prebehol úspešne!", vbInformation
    Else
        Debug.Print "TEST ZLYHAL: Skontrolujte pripojenie alebo logy."
        MsgBox "Test zlyhal!", vbExclamation
    End If
End Sub


' ==============================================================================
' AI PROMPT (ZADANIE):
' "Vytvor procedúru, ktorá prejde tabu¾ku 'Tbl_faktura' a pre všetky faktúry
' v cudzej mene, ktoré nemajú vyplnenı kurz, ho automaticky doplní.
' Ak kurz v databáze chıba, zavolaj funkciu 'NacitajKurzyNBS'.
' Ošetri víkendy pomocou vyh¾adania posledného dostupného kurzu."
' ==============================================================================

Public Sub AktualizujKurzyVFakturach()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim menaTxt As String
    Dim datumFaktury As Date
    Dim kurzNBS As Double
    Dim maxDatum As Variant
    Dim sqlDatum As String
    Dim upravenePocet As Long
    
    Set db = CurrentDb
    ' Vyberieme len faktúry v cudzej mene (FK_mena <> 1), kde kurz chıba [cite: 61]
    Set rs = db.OpenRecordset("SELECT * FROM Tbl_faktura WHERE FK_mena <> 1 AND (kurz_vystavenia Is Null OR kurz_vystavenia = 0)")
    
    If rs.EOF Then
        MsgBox "Všetky faktúry majú kurzy doplnené.", vbInformation, "Hotovo"
        Exit Sub
    End If
    
    Do While Not rs.EOF
        datumFaktury = rs!Datum_vystavenia
        
        ' 1. Preklad ID meny na textovı kód pre NBS [cite: 11, 12, 83]
        Select Case rs!FK_mena
            Case 2: menaTxt = "USD"
            Case 4: menaTxt = "CZK"
            Case 6: menaTxt = "GBP"
            Case 7: menaTxt = "HUF"
            Case Else: menaTxt = ""
        End Select
        
        If menaTxt <> "" Then
            sqlDatum = Format(datumFaktury, "mm\/dd\/yyyy")
            
            ' 2. Vyh¾adanie najnovšieho kurzu v databáze k danému dòu (rieši víkendy) [cite: 38, 46]
            maxDatum = DMax("[time]", "Tbl_kurzy_nbs", "[currency]='" & menaTxt & "' AND [time]<=#" & sqlDatum & "#")
            
            ' 3. Ak kurz v DB nie je, skúsime ho stiahnu z API NBS [cite: 35, 42]
            If IsNull(maxDatum) Then
                Call NacitajKurzyNBS(datumFaktury, True) ' Tichı reim
                maxDatum = DMax("[time]", "Tbl_kurzy_nbs", "[currency]='" & menaTxt & "' AND [time]<=#" & sqlDatum & "#")
            End If
            
            ' 4. Zápis kurzu do faktúry
            If Not IsNull(maxDatum) Then
                kurzNBS = Nz(DLookup("rate", "Tbl_kurzy_nbs", "[currency]='" & menaTxt & "' AND [time]=#" & Format(maxDatum, "mm\/dd\/yyyy") & "#"), 1)
                
                rs.Edit
                rs!kurz_vystavenia = kurzNBS
                rs.Update
                upravenePocet = upravenePocet + 1
            End If
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    MsgBox "Aktualizácia kurzov dokonèená!" & vbCrLf & "Upravenıch faktúr: " & upravenePocet, vbInformation
End Sub
