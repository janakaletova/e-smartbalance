Attribute VB_Name = "FX Automator"
Option Compare Database
Option Explicit

' ==============================================================================
' PROMPT PRE UMELÚ INTELIGENCIU (AI) NA GENEROVANIE TOHTO KÓDU:
' "Uprav existujúci VBA skript na sahovanie kurzov z NBS. Zmeò procedúru na
' funkciu vracajúcu Boolean (True=úspech, False=chyba) pre lepšie ošetrenie
' vıpadkov. Pridaj parameter 'datumUhrady' pre sahovanie historickıch lístkov
' z dynamickej URL (RRRR-MM-DD) a parameter 'tichyRezim' (Boolean), ktorı pri
' automatickom importe skryje všetky vyskakovacie okná a pri zhode dátumov
' automaticky aktualizuje dáta v tabu¾ke.
'
' PRIDANIE NOVEJ FUNKCIE: Vytvor novú subrutinu 'DoplnKurzyDoFaktur', ktorá prejde
' tabu¾ku 'Tbl_faktura'. Pre faktúry, ktoré nie sú v EUR a nemajú vyplnené pole
' 'kurz_vystavenia', doh¾adá najbliší historickı kurz v 'Tbl_kurzy_nbs'.
' Ak kurz chıba, funkcia si ho automaticky stiahne cez 'NacitajKurzyNBS'."
' ==============================================================================

' =====================================================================================
' MODUL: FX Automator (Súèas inovatívneho modulu Smart-Pairing)
' PROJEKT: e-smartbalance s.r.o.
' POPIS: Zabezpeèuje automatické sahovanie aktuálnych aj historickıch kurzovıch
'        lístkov z API Národnej banky Slovenska (NBS) priamo do relaènej databázy.
' =====================================================================================

' -------------------------------------------------------------------------------------
' FUNKCIA: NacitajKurzyNBS
' ÚÈEL:    Dynamicky stiahne XML kurzovı lístok z NBS pre zadanı dátum a uloí ho.
' NÁVRATOVÁ HODNOTA: Boolean (True = úspech, False = chyba pripojenia alebo spracovania)
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

    NacitajKurzyNBS = False

    ' 1. PRÍPRAVA SPOJENIA S API NBS
    formatovanyDatumPreURL = Format(datumUhrady, "yyyy-mm-dd")
    url = "https://nbs.sk/export/sk/exchange-rate/" & formatovanyDatumPreURL & "/xml"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.send

    If http.Status <> 200 Then
        If Not tichyRezim Then MsgBox "Chyba pripojenia k NBS pre dátum " & datumUhrady & "! (Status: " & http.Status & ")", vbCritical
        GoTo Cistka
    End If

    ' 2. SPRACOVANIE XML SÚBORU
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:ns='http://www.ecb.int/vocabulary/2002-08-01/eurofxref'"
    xmlDoc.loadXML http.responseText

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

' ==============================================================================
' AI PROMPT (ZMENOVÁ POIADAVKA PRE UMELÚ INTELIGENCIU):
' "Vytvor novú procedúru 'DoplnKurzyDoFaktur', ktorá prejde tabu¾ku 'Tbl_faktura'.
' Úloha: Pre všetky faktúry v cudzej mene, ktoré nemajú vyplnenı 'kurz_vystavenia',
' doh¾adaj správny historickı kurz z tabu¾ky 'Tbl_kurzy_nbs' pod¾a dátumu vystavenia.
' Ak kurz v databáze chıba, zavolaj funkciu 'NacitajKurzyNBS' v tichom reime (True),
' stiahni ho z API Národnej banky Slovenska a ulo priamo k hlavièke faktúry.
' Zabezpeè aj riešenie pre víkendy (kedy NBS nevydáva kurzy) a na konci zobraz
' štatistiku úspešnosti aktualizovanıch záznamov."
' ==============================================================================

' -------------------------------------------------------------------------------------
' PROCEDÚRA: DoplnKurzyDoFaktur
' ÚÈEL:      Retrospektívne preh¾adá Tbl_faktura a doplní chıbajúce kurzy vystavenia
'            pre cudzie meny. Automaticky dopytuje chıbajúce dni cez API NBS.
' -------------------------------------------------------------------------------------
Sub DoplnKurzyDoFaktur()
    Dim db As DAO.Database
    Dim rsFaktury As DAO.Recordset
    Dim datumFaktury As Date
    Dim menaID As Integer
    Dim menaStr As String
    Dim maxDatum As Variant
    Dim zistenyKurz As Double
    Dim pocetAktualizovanych As Integer
    Dim sqlDatum As String
    
    Dim trebaStiahnut As Boolean
    Dim lastTriedDate As Date
    
    Set db = CurrentDb
    
    ' OPRAVA: Zoradíme faktúry pod¾a dátumu (ORDER BY), aby sme postupovali chronologicky
    Set rsFaktury = db.OpenRecordset("SELECT * FROM Tbl_faktura WHERE FK_mena <> 1 AND kurz_vystavenia Is Null ORDER BY Datum_vystavenia")
    
    pocetAktualizovanych = 0
    lastTriedDate = 0 ' Pomocná premenná pre ochranu pred spamovaním API poèas sviatkov
    
    If rsFaktury.EOF Then
        MsgBox "Všetky faktúry v cudzej mene u majú kurz úspešne vyplnenı!", vbInformation, "Kontrola kurzov"
        rsFaktury.Close
        Set rsFaktury = Nothing
        Set db = Nothing
        Exit Sub
    End If
    
    ' Preh¾adávame chıbajúce kurzy
    Do While Not rsFaktury.EOF
        datumFaktury = rsFaktury!Datum_vystavenia
        menaID = rsFaktury!FK_mena
        
        ' Mapovanie na textovı kód meny
        Select Case menaID
            Case 4: menaStr = "CZK"
            Case 6: menaStr = "GBP"
            Case 2: menaStr = "USD"
            Case 7: menaStr = "HUF"
            Case Else: menaStr = ""
        End Select
        
        If menaStr <> "" Then
            sqlDatum = Format(datumFaktury, "mm\/dd\/yyyy")
            
            ' 1. KROK: H¾adáme akıko¾vek najnovší dostupnı kurz v našej DB
            maxDatum = DMax("[Time]", "Tbl_kurzy_nbs", "[currency]='" & menaStr & "' AND [Time] <= #" & sqlDatum & "#")
            
            trebaStiahnut = False
            
            ' INTELIGENTNÁ LOGIKA SAHOVANIA (Oprava Lazy Fetching chyby)
            If IsNull(maxDatum) Then
                ' Tabu¾ka je úplne prázdna
                trebaStiahnut = True
            Else
                ' Akı je to deò v tıdni? (1 = Nede¾a, 2 = Pondelok ... 7 = Sobota vo vbSunday, my pouijeme vbMonday pre európsky štandard)
                If Weekday(datumFaktury, vbMonday) <= 5 Then
                    ' Je to pracovnı deò: Ak kurz v DB je starší ako dátum faktúry, systém sa ho pokúsi stiahnu.
                    ' Vınimka (lastTriedDate): Ak sme u tento dátum na API dopytovali a NBS nám dalo starší kurz (štátny sviatok), nebudeme API spamova znova.
                    If maxDatum < datumFaktury And datumFaktury <> lastTriedDate Then
                        trebaStiahnut = True
                    End If
                Else
                    ' Je to víkend: Staèí nám piatkovı kurz. Ak je ale "piatkovı" kurz starší viac ako 4 dni (napríklad pre Ve¾kú noc), API zavoláme.
                    If DateDiff("d", maxDatum, datumFaktury) > 4 And datumFaktury <> lastTriedDate Then
                        trebaStiahnut = True
                    End If
                End If
            End If
            
            ' 2. KROK: Ak nám kurz reálne chıba, stiahneme ho z API NBS!
            If trebaStiahnut Then
                Call NacitajKurzyNBS(datumFaktury, True)
                
                ' Zapamätáme si, e sme tento konkrétny dátum u vyskúšali, aby sme sa nezasekli na sviatkoch
                lastTriedDate = datumFaktury
                
                ' Po stiahnutí znova vyh¾adáme aktuálny najnovší kurz
                maxDatum = DMax("[Time]", "Tbl_kurzy_nbs", "[currency]='" & menaStr & "' AND [Time] <= #" & sqlDatum & "#")
            End If
            
            ' 3. KROK: Zapíšeme nájdenı kurz do faktúry
            If Not IsNull(maxDatum) Then
                zistenyKurz = DLookup("Rate", "Tbl_kurzy_nbs", "[currency]='" & menaStr & "' AND [Time] = #" & Format(maxDatum, "mm\/dd\/yyyy") & "#")
                
                rsFaktury.Edit
                rsFaktury!kurz_vystavenia = zistenyKurz
                rsFaktury.Update
                
                pocetAktualizovanych = pocetAktualizovanych + 1
            End If
        End If
        
        rsFaktury.MoveNext
    Loop
    
    rsFaktury.Close
    Set rsFaktury = Nothing
    Set db = Nothing
    
    MsgBox "Kontrola a sahovanie kurzov úspešne dokonèené!" & vbCrLf & vbCrLf & _
           "Poèet aktualizovanıch faktúr: " & pocetAktualizovanych, vbInformation, "FX Automator: Faktúry"
End Sub

