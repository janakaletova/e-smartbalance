Attribute VB_Name = "ETL Import"
Option Compare Database
Option Explicit

' ==============================================================================
' AI PROMPT (ZADANIE PRE UMELÚ INTELIGENCIU):
' "Vytvor VBA program pre MS Access, ktorı dokáe preèíta bankovı vıpis
' uloenı v súbore CSV a automaticky ho nahra do tabu¾ky 'tbl_platba'.
'
' Hlavné poiadavky:
' 1. Bezpeènos: Zabezpeè, aby sa iadna platba nenahrala dvakrát, aj keï
'    pouívate¾ spustí import opakovane (vyui ochranu cez chybové hlásenie
'    o duplicite).
' 2. Inteligencia: Program musí vedie preloi textové skratky mien
'    (napr. EUR, CZK) na èíselné ID, ktoré pouíva naša databáza.
' 3. Preh¾adnos: Na konci importu uká pouívate¾ovi správu o tom, ko¾ko
'    novıch platieb sa úspešne pridalo a ko¾ko sa ich preskoèilo,
'    lebo u v systéme boli."
' ==============================================================================
' ==============================================================================
' AI PROMPT (ZMENOVÁ POIADAVKA PRE UMELÚ INTELIGENCIU):
' "Uprav náš existujúci importnı program tak, aby nebol napevno zviazanı
' len s jednım konkrétnym súborom. Namiesto toho z neho urob univerzálnu
' funkciu, ktorá dokáe prija cestu k súboru (parameter), ktorú jej pošle
' pouívate¾ po kliknutí na tlaèidlo vo formulári.
'
' Ïalšie poiadavky:
' 1. Auditná stopa: Zabezpeè, aby sa celá táto prijatá cesta k súboru
'    uloila do databázy ku kadej jednej nahranej platbe. V budúcnosti
'    tak budeme presne vedie doh¾ada zdrojovı súbor.
' 2. Zachovaj všetky doterajšie ochrany: Inteligentné èítanie dátumov
'    (aby nepadal na rôznych formátoch), automatické priradenie 'Bankového
'    prevodu' a ochranu pred duplicitami (tiché preskoèenie existujúcich platieb)."
' ==============================================================================



' ==============================================================================
' AI PROMPT (ZMENOVÁ POIADAVKA PRE UMELÚ INTELIGENCIU):
' "Uprav existujúcu procedúru 'ImportujBankovyVypis' pre import CSV súboru.
' Zabezpeè, aby systém pri platbe v cudzej mene automaticky vyh¾adal správny
' kurz v tabu¾ke 'Tbl_kurzy_nbs' pod¾a dátumu platby.
' Ak kurz v databáze chıba, systém musí potichu (bez vyskakovacích okien)
' zavola funkciu 'NacitajKurzyNBS', stiahnu historické dáta z API Národnej
' banky Slovenska a následne tento stiahnutı kurz priradi k nahrávanej platbe.
' Vyrieš aj problém s víkendmi, kedy NBS kurzy nevydáva."
' ==============================================================================

' =====================================================================================
' MODUL: Automatické nahrávanie bankového vıpisu (V4 - Integrácia s FX Automatorom)
' POPIS AKTUÁLNEHO SPRÁVANIA FUNKCIE:
' 1. CSV Import: Prijíma presnú cestu k súboru, èíta dáta a parsuje ich.
' 2. Preklad dát: Inteligentne prekladá menové textové skratky na èíselné ID meny.
' 3. Ochrana pred duplicitami: Existujúce platby bezpeène preskoèí bez pádu aplikácie.
' 4. SMART-FX LOGIKA (Kurzové rozdiely):
'    - Pri platbách v inej mene ako EUR systém h¾adá najnovší kurz k danému dòu.
'    - Pouíva funkciu DMax na prekonanie víkendov (v nede¾u zoberie piatkovı kurz).
'    - Ak kurz pre danı deò v systéme neexistuje, funkcia prevezme kontrolu a cez
'      REST API z NBS stiahne chıbajúci kurzovı lístok (XML).
'    - Zistenı kurz zapíše priamo do ståpca 'pouzity_kurz_nbs' ku kadej jednej platbe.
' =====================================================================================
' =====================================================================================
' MODUL: Automatické nahrávanie bankového vıpisu (Parametrizovaná verzia s FX Automator)
' POPIS: Prijíma cestu k súboru, naèíta dáta, stiahne chıbajúce kurzy NBS a vloí do tabu¾ky.
' =====================================================================================
Sub ImportujBankovyVypis(ByVal filePath As String)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fso As Object
    Dim ts As Object
    Dim lineText As String
    Dim dataArray() As String
    
    Dim successCount As Integer
    Dim duplicateCount As Integer
    Dim menaID As Integer
    
    ' NOVÉ PREMENNÉ PRE PRÁCU S KURZAMI:
    Dim datumPlatby As Date
    Dim menaTxt As String
    Dim kurzNBS As Double
    Dim maxDatum As Variant
    Dim sqlDatum As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Kontrola, èi súbor na odovzdanej ceste naozaj existuje
    If Not fso.FileExists(filePath) Then
        MsgBox "Súbor na ceste (" & filePath & ") sa nenašiel!", vbCritical, "Súbor chıba"
        Exit Sub
    End If
    
    ' Otvorenie súboru (1 = Iba na èítanie, False = Nevytvára novı, -1 = Unicode formát)
    Set ts = fso.OpenTextFile(filePath, 1, False, -1)
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbl_platba")
    
    ' Preskoèenie hlavièky (názvov ståpcov v CSV)
    If Not ts.AtEndOfStream Then ts.ReadLine
    
    ' Zapnutie ochrany pred chybami (kvôli duplicitám)
    On Error GoTo ErrorHandler
    
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        
        If Trim(lineText) <> "" Then
            dataArray = Split(lineText, ";")
            
            ' Extrakcia dát a prevod na správne typy
            menaTxt = UCase(Trim(dataArray(2)))
            datumPlatby = InteligentnyParserDatumu(dataArray(3))
            
            ' Priradenie správnej meny (ID) pod¾a textu
            Select Case menaTxt
                Case "EUR": menaID = 1
                Case "CZK": menaID = 4
                Case "USD": menaID = 2
                Case "GBP": menaID = 6
                Case "HUF": menaID = 7
                Case Else: menaID = 1
            End Select
            
            ' =========================================================
            ' SMART-LOGIKA: ZISOVANIE A SAHOVANIE KURZU NBS
            ' =========================================================
            If menaTxt = "EUR" Then
                kurzNBS = 1 ' Pre eurá je kurz vdy 1
            Else
                ' Formát pre SQL dopyt
                sqlDatum = Format(datumPlatby, "mm\/dd\/yyyy")
                
                ' 1. Pokus: Nájs najnovší platnı kurz k dátumu platby (rieši aj víkendy)
                maxDatum = DMax("[time]", "Tbl_kurzy_nbs", "[currency]='" & menaTxt & "' AND [time]<=#" & sqlDatum & "#")
                
                ' Ak kurz neexistuje, zavoláme FX Automator
                If IsNull(maxDatum) Then
                    Call NacitajKurzyNBS(datumPlatby, True) ' True = Tichı reim bez vyskakovacích okien
                    
                    ' 2. Pokus: Znova preèítame najnovší dátum kurzu po stiahnutí
                    maxDatum = DMax("[time]", "Tbl_kurzy_nbs", "[currency]='" & menaTxt & "' AND [time]<=#" & sqlDatum & "#")
                End If
                
                ' Ak sa kurz našiel (alebo stiahol), vytiahneme jeho hodnotu (Rate)
                If Not IsNull(maxDatum) Then
                    kurzNBS = Nz(DLookup("rate", "Tbl_kurzy_nbs", "[currency]='" & menaTxt & "' AND [time]=#" & Format(maxDatum, "mm\/dd\/yyyy") & "#"), 1)
                Else
                    kurzNBS = 1 ' Fallback, ak NBS neodpovedá
                End If
            End If
            ' =========================================================
            
            ' Pridanie záznamu do databázy
            rs.AddNew
            rs!var_symbol_banka = dataArray(0)
            
            ' Ošetrenie desatinnej èiarky pri sume
            rs!suma = Val(Replace(dataArray(1), ",", "."))
            
            rs!FK_mena = menaID
            rs!dátum = datumPlatby
            
            ' Uloenie nášho automaticky zisteného kurzu z NBS
            rs!pouzity_kurz_nbs = kurzNBS
            
            ' Ošetrenie prázdnych hodnôt pre IBAN a Názov protistrany
            If UBound(dataArray) >= 4 Then rs!iban_protistrany = dataArray(4)
            If UBound(dataArray) >= 5 Then rs!nazov_protistrany = dataArray(5)
            
            ' Informácie pre kontrolu (audit)
            rs!nazov_zdrojoveho_suboru = filePath
            rs!sparovane_automaticky = False
            rs!FK_sposob_platby = 2 ' Bankovı prevod
            
            rs.Update
            successCount = successCount + 1
            
ContinueLoop:
        End If
    Loop
    
    ' Upratovanie pamäte
    rs.Close
    ts.Close
    Set rs = Nothing
    Set ts = Nothing
    Set db = Nothing
    Set fso = Nothing
    
    On Error GoTo 0
    
    ' Závereèné hlásenie pre pouívate¾a
    MsgBox "Nahrávanie vıpisu bolo dokonèené!" & vbCrLf & vbCrLf & _
           "Úspešne pridané nové platby: " & successCount & vbCrLf & _
           "Preskoèené (u existujúce) platby: " & duplicateCount, vbInformation, "Vısledok importu"
    Exit Sub

ErrorHandler:
    ' Ak Access narazí na identickú platbu (z rovnakého súboru/rovnaké ID), preskoèí ju
    If Err.Number = 3022 Then
        rs.CancelUpdate
        duplicateCount = duplicateCount + 1
        Resume ContinueLoop
    Else
        ' Neèakaná chyba vypíše detail
        MsgBox "Chyba na riadku: " & lineText & vbCrLf & _
               "Èíslo chyby: " & Err.Number & vbCrLf & _
               "Popis: " & Err.Description, vbCritical, "Detail chyby"
        rs.CancelUpdate
        Resume ContinueLoop
    End If
End Sub
' ------------------------------------------------------------------------------
' POMOCNÁ PROCEDÚRA: Test_Import
' POPIS: Slúi na rıchle otestovanie importu priamo z VBA editora bez nutnosti
'        vybera súbor cez formulár.
' ------------------------------------------------------------------------------
Sub Test_Import()
    Dim testovaciaCesta As String
    
    ' Vytvorenie cesty k súboru v rovnakej zloke, kde beí táto databáza
    testovaciaCesta = CurrentProject.Path & "\bankovy_vypis_komplexny.csv"
    
    ' Volanie hlavnej procedúry a odovzdanie parametra s cestou
    Call ImportujBankovyVypis(testovaciaCesta)
End Sub




