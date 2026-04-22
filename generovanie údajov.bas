Attribute VB_Name = "generovanie údajov"
Option Compare Database

' ==============================================================================
' PROMPT PRE UMELÚ INTELIGENCIU (AI) NA GENEROVANIE TOHTO KÓDU:
' Rola: Si expert na MS Access a programovanie vo VBA.
'
' Úloha: Vygeneruj VBA kód pre MS Access, ktorı obsahuje dve procedúry.
' Prvá procedúra (DoplnFakturyDo350_IbaPracovneDni) doplní chıbajúce faktúry do
' tabu¾ky 'Tbl_faktura' tak, aby ich bolo celkovo presne 350. Dátumy musia by len
' pracovné dni a nesmú presiahnu 17.4.2026.
' Druhá procedúra (GenerujKomplexnyBankovyVypisCSV) preèíta dáta z 'Tbl_faktura'
' a vygeneruje bankovı vıpis vo formáte CSV. Nasimuluj 4 scenáre úhrad: presná zhoda (60%),
' preklep vo variabilnom symbole (15%), èiastoèná úhrada (15%) a nezmysel vo VS (10%).
'
' KONTEXT DÁT (Striktne dodriavaj tieto názvy polí a typy):
' 1. Tbl_faktura: ID_faktura (AutoNumber/PK), FK_partner_ID (Number/FK),
'    Typ_faktury (Yes/No), Datum_vystavenia (Date/Time), Suma (Currency),
'    Variabilny_symbol (Short Text), FK_mena (Number/FK), pouzity_kurz_nbs (Number).
' 2. Tbl_partner: PK_partner (AutoNumber/PK), typ_partnera (Yes/No), názov, ièo.
' 3. Tbl_mena: PK_mena (PK), Skratka (1=EUR, 2=USD, 4=CZK, 6=GBP, 7=HUF).
' ==============================================================================

' ------------------------------------------------------------------------------
' PROCEDÚRA 1: DoplnFakturyDo350_IbaPracovneDni
' POPIS: Táto procedúra slúi na hromadné vytvorenie testovacích dát.
'        Najprv zistí, ko¾ko faktúr u v tabu¾ke Tbl_faktura je, a následne
'        dogeneruje chıbajúci poèet do 400. Zabezpeèuje, aby dátumy vystavenia
'        pripadli vıluène na pracovné dni (preskakuje víkendy) a aby nepresiahli
'        dátum 17.4.2026. Taktie náhodne simuluje vydané a prijaté faktúry v rôznych menách.
' ------------------------------------------------------------------------------
Sub DoplnFakturyDo400_IbaPracovneDni()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim datum As Date
    Dim pocetExistujucich As Long
    Dim pocetNaVygenerovanie As Long
    Dim i As Long
    Dim partnerID As Integer
    Dim menaID As Integer
    Dim typ As Boolean
    Dim akt_datum As Date

    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Tbl_faktura")
    
    ' 1. Zistíme, ko¾ko faktúr v tabu¾ke u reálne máš
    pocetExistujucich = DCount("*", "Tbl_faktura")
    pocetNaVygenerovanie = 400 - pocetExistujucich
    
    If pocetNaVygenerovanie <= 0 Then
        MsgBox "U máš " & pocetExistujucich & " faktúr! Nie je potrebné generova ïalšie.", vbInformation
        Exit Sub
    End If
    
    Randomize
    ' Zaèíname generova od polovice januára
    datum = DateSerial(2026, 1, 20)
    
    akt_datum = Now()
    
    For i = 1 To pocetNaVygenerovanie
        ' Kadú 3. faktúru posunieme o deò dopredu, aby boli nasekané tesne za sebou
        If i Mod 3 = 0 Then datum = datum + 1
        
        ' K¾úèová funkcia: Preskoèíme víkendy (Sobota = 7, Nede¾a = 1 vo vbSunday)
        While Weekday(datum, vbMonday) > 5
            datum = datum + 1
        Wend
        
        ' Zastavíme generovanie na akt. datume, aby sme nepresiahli dnešnı deò
        If datum > akt_datum Then
            datum = DateSerial(2026, 1, 20) ' Ak sme na konci, zaèneme opä od januára
        End If
        
        ' Vıber partnera a logiky (Vınos/Náklad)
        If (i Mod 4 = 0) Then
            typ = False ' Prijatá (Náklad)
            partnerID = Choose(Int(Rnd() * 2) + 1, 6, 19)
        Else
            typ = True ' Vydaná (Vınos)
            partnerID = Choose(Int(Rnd() * 5) + 1, 10, 11, 20, 21, 22)
        End If
        
        ' Priradenie správnej meny pod¾a partnera
        If partnerID = 19 Then
            menaID = 4 ' CZK
        ElseIf partnerID = 20 Then
            menaID = 6 ' GBP
        Else
            menaID = 1 ' EUR
        End If
        
        ' Zápis nového riadku
        rs.AddNew
        rs!FK_partner_ID = partnerID
        'rs!Typ_faktury = typ
        rs!Datum_vystavenia = datum
        rs!suma = Round((Rnd() * 1500) + 100, 2)
        ' VS vo formáte YYYYMMDD + poradové èíslo pre unikátnos
        rs!Variabilny_symbol = Format(datum, "yyyymmdd") & Format(i, "000")
        rs!FK_mena = menaID
        rs.Update
    Next i
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    MsgBox "Úspech! Zvyšnıch " & pocetNaVygenerovanie & " faktúr bolo dogenerovanıch." & vbCrLf & _
           "Teraz máš v tabu¾ke presne 400 záznamov, bez víkendov a nasekané do dnes.", vbInformation
End Sub

' ------------------------------------------------------------------------------
' PROCEDÚRA: GenerujLogickeVS_PreOdberatelov
' ÚÈEL: Prejde všetky vystavené faktúry a pridelí im novı, unikátny variabilnı
'       symbol vo formáte RRRRMMXXXX.
' OPRAVA: Typ faktúry sa urèuje dynamicky z prepojenej tabu¾ky Tbl_partner
'         (odstránená redundancia po¾a Typ_faktury).
' LOGIKA: Poradové èíslo (XXXX) sa automaticky resetuje pri zmene mesiaca.
' ------------------------------------------------------------------------------
Sub GenerujLogickeVS_PreOdberatelov()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim aktualnyRokMesiac As String
    Dim pocitadlo As Integer
    Dim novyVS As String
    Dim pocetUpravenych As Long
    Dim strSQL As String
    
    Set db = CurrentDb
    
    ' SQL dotaz prepojí faktúry s partnermi a vyfiltruje iba odberate¾ov (typ_partnera = True)
    strSQL = "SELECT Tbl_faktura.* FROM Tbl_faktura " & _
             "INNER JOIN Tbl_partner ON Tbl_faktura.FK_partner_ID = Tbl_partner.PK_partner " & _
             "WHERE Tbl_partner.typ_partnera = True " & _
             "ORDER BY Tbl_faktura.Datum_vystavenia ASC, Tbl_faktura.ID_faktura ASC"
    
    Set rs = db.OpenRecordset(strSQL)
    
    If rs.EOF Then
        MsgBox "Nenašli sa iadne vydané faktúry pre odberate¾ov.", vbInformation, "Chyba"
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        Exit Sub
    End If
    
    aktualnyRokMesiac = ""
    pocetUpravenych = 0
    
    ' Prechádzame záznamy v sluèke
    Do While Not rs.EOF
        ' Skontrolujeme, èi sa zmenil mesiac alebo rok. Ak áno, resetujeme poèítadlo na 1.
        If Format(rs!Datum_vystavenia, "yyyymm") <> aktualnyRokMesiac Then
            aktualnyRokMesiac = Format(rs!Datum_vystavenia, "yyyymm")
            pocitadlo = 1
        End If
        
        ' Zloenie nového VS: RokMesiac + 4-miestne poradové èíslo
        ' Vısledok napr.: 2026010001
        novyVS = aktualnyRokMesiac & Format(pocitadlo, "0000")
        
        ' Aktualizácia záznamu v tabu¾ke
        rs.Edit
        rs!Variabilny_symbol = novyVS
        rs.Update
        
        ' Zvıšenie poèítadiel pre ïalší krok
        pocitadlo = pocitadlo + 1
        pocetUpravenych = pocetUpravenych + 1
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    MsgBox "Úspech! Bolo vygenerovanıch " & pocetUpravenych & " novıch variabilnıch symbolov pre odberate¾ov." & vbCrLf & _
           "Nová logika: ROK + MESIAC + 4-miestne poradové èíslo (napr. 2026010001).", vbInformation, "Generátor VS"
End Sub
' ==============================================================================
' AI PROMPT (ZMENOVÁ POIADAVKA PRE UMELÚ INTELIGENCIU):
' "Uprav generátor bankového vıpisu tak, aby namiesto jedného ve¾kého súboru
' vygeneroval samostatné CSV súbory za kadı kalendárny mesiac (napr. vıpis_01.csv).
'
' Logika rozdelenia a nové scenáre:
' 1. Program prejde všetky faktúry a pre kadú nasimuluje platbu (scenáre).
' 2. Rozde¾ platby do scenárov: ideálna (40%), oneskorená pre cudzie meny (10%),
'    splátky - 2 platby (8%) splatené, 2 - platby z piatich ( nesplatené) (2%)  , preklep vo VS (15%), zmena IBANu (10%) a úplnı nezmysel (5%).
' 3. Ak je partner Dodávate¾ (True), suma v banke musí by záporná (vıdavok).
' 4. Zachovaj ochranu pred budúcimi dátumami (dnes je 19.4.2026).
' 5. Bankové poplatky generuj mesaène a vlo ich vdy do správneho mesaèného súboru."
' ==============================================================================

' ------------------------------------------------------------------------------
' PROCEDÚRA: GenerujMesenéBankovéVıpisyCSV
' POPIS: Vytvorí sadu CSV súborov rozdelenıch pod¾a mesiacov so splátkovou logikou
'        a oneskorenımi platbami. Zisuje celkovı poèet faktúr: ak je ich viac
'        ako 340, pri prvıch 340 platbách je VS bezchybnı. Pre zvyšné platby
'        (nad 340) môu vznika preklepy alebo nezmysly vo VS. Splátková logika
'        (úplná aj èiastoèná úhrada) je zachovaná plnohodnotne pre správne
'        aj pre chybné variabilné symboly.
' ------------------------------------------------------------------------------
Sub GenerujMesenéBankovéVıpisyCSV()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fso As Object
    Dim tsArray(1 To 12) As Object ' Pole pre súborové streamy
    Dim filePath As String
    Dim varSymbol As String
    Dim pouzityVS As String
    Dim suma As Double
    Dim datumVystavenia As Date
    Dim datumPlatby As Date
    Dim datumPlatby1 As Date, datumPlatby2 As Date
    Dim dnesnyDatum As Date
    Dim menaID As Integer
    Dim menaStr As String
    Dim typPlatby As Integer
    Dim outLine As String
    Dim m As Integer
    
    Dim partnerIBAN As String
    Dim partnerNazov As String
    Dim skutocnyIbanPreCSV As String
    Dim pouzityIBAN As String
    
    Dim pocetFaktur As Long
    Dim cisloFaktury As Long
    
    dnesnyDatum = Date ' Berie aktuálny dátum (napr. 19. apríl 2026)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set db = CurrentDb
    
    ' 1. PRÍPRAVA SÚBOROV
    For m = 1 To Month(dnesnyDatum)
        filePath = CurrentProject.Path & "\bankovy_vypis_" & Year(dnesnyDatum) & "_" & Format(m, "00") & ".csv"
        Set tsArray(m) = fso.CreateTextFile(filePath, True, True)
        tsArray(m).WriteLine "Var_Symbol_Banka;Suma_Prijata;Mena_Pohybu;Datum_Prijmu;IBAN_Protistrany;Nazov_Protistrany"
    Next m
    
    ' 2. NAÈÍTANIE FAKTÚR A ZISTENIE POÈTU
    Dim sqlQuery As String
    sqlQuery = "SELECT F.*, P.iban, P.názov AS NazovPartnera, P.typ_partnera " & _
               "FROM Tbl_faktura AS F INNER JOIN Tbl_partner AS P " & _
               "ON F.FK_partner_ID = P.PK_partner " & _
               "ORDER BY F.Datum_vystavenia"
               
    Set rs = db.OpenRecordset(sqlQuery)
    
    If rs.EOF Then
        MsgBox "iadne faktúry na spracovanie.", vbExclamation
        rs.Close
        Set rs = Nothing: Set db = Nothing: Set fso = Nothing
        Exit Sub
    End If
    
    ' Zistenie celkového poètu faktúr presunom na koniec a spä
    rs.MoveLast
    pocetFaktur = rs.RecordCount
    rs.MoveFirst
    
    Randomize
    cisloFaktury = 0
    
    ' 3. GENERUJEME PLATBY A ROZDE¼UJEME ICH DO SÚBOROV
    Do While Not rs.EOF
        cisloFaktury = cisloFaktury + 1
        
        varSymbol = Nz(rs!Variabilny_symbol, "")
        suma = Nz(rs!suma, 0)
        
        ' Mínusové sumy pre dodávate¾ov
        If rs!typ_partnera = True Then
            suma = suma * -1
        End If
        
        datumVystavenia = rs!Datum_vystavenia
        menaID = Nz(rs!FK_mena, 1)
        partnerIBAN = Nz(rs!iban, "")
        partnerNazov = Nz(rs!NazovPartnera, "Neznámy partner")
        
        ' Dynamickı IBAN
        skutocnyIbanPreCSV = partnerIBAN
        If skutocnyIbanPreCSV = "" Then
            skutocnyIbanPreCSV = "SK" & Int(Rnd() * 90 + 10) & "0900" & Format(Int(Rnd() * 999999999), "000000000000")
        End If
        
        Select Case menaID
            Case 1: menaStr = "EUR": Case 4: menaStr = "CZK": Case 6: menaStr = "GBP"
            Case 2: menaStr = "USD": Case 7: menaStr = "HUF": Case Else: menaStr = "EUR"
        End Select
        
        ' ---------------------------------------------------------
        ' A) ODDELENÁ LOGIKA PRE VARIABILNİ SYMBOL
        ' ---------------------------------------------------------
        pouzityVS = varSymbol
        If pocetFaktur > 340 And cisloFaktury > 340 Then
            ' Zvyšné platby MÔU ma problém vo VS (napr. 50% šanca na chybu)
            Dim vsChyba As Integer
            vsChyba = Int(Rnd() * 100) + 1
            
            If vsChyba <= 30 Then
                ' 30% šanca na preklep (napríklad zámena 0 za O)
                If InStr(pouzityVS, "0") > 0 Then pouzityVS = Replace(pouzityVS, "0", "O", 1, 1)
            ElseIf vsChyba <= 50 Then
                ' 20% šanca na úplnı nezmysel
                pouzityVS = "UHRADA" & Int(Rnd() * 99)
            End If
            ' Ak padne 51-100, VS zostáva správny aj nad limit 340
        End If
        
        ' ---------------------------------------------------------
        ' B) ODDELENÁ LOGIKA PRE IBAN (Zmena banky)
        ' ---------------------------------------------------------
        pouzityIBAN = skutocnyIbanPreCSV
        If Int(Rnd() * 100) + 1 <= 10 Then ' 10% šanca na zmenu banky pre akúko¾vek platbu
            pouzityIBAN = "SK" & Int(Rnd() * 90 + 10) & "1100" & Format(Int(Rnd() * 999999999), "000000000000")
        End If
        
        ' ---------------------------------------------------------
        ' C) TYP PLATBY (Urèuje iba termíny a splátky)
        ' ---------------------------------------------------------
        typPlatby = Int(Rnd() * 100) + 1
        
        Select Case typPlatby
            Case 1 To 60 ' 1. Ideálna platba v celku (60%)
                datumPlatby = datumVystavenia + Int(Rnd() * 14) + 1
                If datumPlatby <= dnesnyDatum Then
                    outLine = pouzityVS & ";" & Replace(Format(suma, "0.00"), ",", ".") & ";" & menaStr & ";" & Format(datumPlatby, "dd.mm.yyyy") & ";" & pouzityIBAN & ";" & partnerNazov
                    tsArray(Month(datumPlatby)).WriteLine outLine
                End If
                
            Case 61 To 75 ' 2. Oneskorená platba v celku pre kurzovı rozdiel (15%)
                If menaID <> 1 Then
                    datumPlatby = datumVystavenia + Int(Rnd() * 20) + 30
                Else
                    datumPlatby = datumVystavenia + Int(Rnd() * 14) + 1
                End If
                
                If datumPlatby <= dnesnyDatum Then
                    outLine = pouzityVS & ";" & Replace(Format(suma, "0.00"), ",", ".") & ";" & menaStr & ";" & Format(datumPlatby, "dd.mm.yyyy") & ";" & pouzityIBAN & ";" & partnerNazov
                    tsArray(Month(datumPlatby)).WriteLine outLine
                End If
                
            Case 76 To 90 ' 3. Platba na SPLÁTKY plne splatená (15%)
                Dim s1 As Double: s1 = Round(suma / 2, 2)
                Dim s2 As Double: s2 = suma - s1
                
                If menaID <> 1 Then
                    datumPlatby1 = datumVystavenia + Int(Rnd() * 15) + 20
                    datumPlatby2 = datumPlatby1 + Int(Rnd() * 20) + 15
                Else
                    datumPlatby1 = datumVystavenia + Int(Rnd() * 5) + 2
                    datumPlatby2 = datumPlatby1 + Int(Rnd() * 10) + 5
                End If
                
                If datumPlatby1 <= dnesnyDatum Then
                    tsArray(Month(datumPlatby1)).WriteLine pouzityVS & ";" & Replace(Format(s1, "0.00"), ",", ".") & ";" & menaStr & ";" & Format(datumPlatby1, "dd.mm.yyyy") & ";" & pouzityIBAN & ";" & partnerNazov
                    If datumPlatby2 <= dnesnyDatum Then
                        tsArray(Month(datumPlatby2)).WriteLine pouzityVS & ";" & Replace(Format(s2, "0.00"), ",", ".") & ";" & menaStr & ";" & Format(datumPlatby2, "dd.mm.yyyy") & ";" & pouzityIBAN & ";" & partnerNazov
                    End If
                End If
                
            Case 91 To 100 ' 4. TRVALİ NEDOPLATOK: 2 splátky z 5 (10%)
                s1 = Round(suma / 5, 2)
                s2 = Round(suma / 5, 2)
                
                If menaID <> 1 Then
                    datumPlatby1 = datumVystavenia + Int(Rnd() * 15) + 20
                    datumPlatby2 = datumPlatby1 + Int(Rnd() * 20) + 15
                Else
                    datumPlatby1 = datumVystavenia + Int(Rnd() * 5) + 2
                    datumPlatby2 = datumPlatby1 + Int(Rnd() * 10) + 5
                End If
                
                If datumPlatby1 <= dnesnyDatum Then
                    tsArray(Month(datumPlatby1)).WriteLine pouzityVS & ";" & Replace(Format(s1, "0.00"), ",", ".") & ";" & menaStr & ";" & Format(datumPlatby1, "dd.mm.yyyy") & ";" & pouzityIBAN & ";" & partnerNazov
                    If datumPlatby2 <= dnesnyDatum Then
                        tsArray(Month(datumPlatby2)).WriteLine pouzityVS & ";" & Replace(Format(s2, "0.00"), ",", ".") & ";" & menaStr & ";" & Format(datumPlatby2, "dd.mm.yyyy") & ";" & pouzityIBAN & ";" & partnerNazov
                    End If
                End If
        End Select
        
        rs.MoveNext
    Loop
    
    ' 4. ZÁPIS MESAÈNİCH POPLATKOV
    For m = 1 To Month(dnesnyDatum)
        If m = Month(dnesnyDatum) And 28 > Day(dnesnyDatum) Then
            datumPlatby = dnesnyDatum
        Else
            datumPlatby = DateSerial(Year(dnesnyDatum), m, 28)
        End If
        outLine = "POPLATOK " & Format(m, "00") & "/" & Year(dnesnyDatum) & ";-7.50;EUR;" & Format(datumPlatby, "dd.mm.yyyy") & ";;Mesaènı poplatok za úèet"
        tsArray(m).WriteLine outLine
    Next m
    
    ' 5. UPRATOVANIE
    On Error Resume Next
    For m = 1 To 12
        If Not tsArray(m) Is Nothing Then tsArray(m).Close
    Next m
    
    rs.Close
    Set rs = Nothing: Set db = Nothing: Set fso = Nothing
    
    MsgBox "Generovanie mesaènıch vıpisov úspešne dokonèené!", vbInformation
End Sub

' ==============================================================================
' PROMPT PRE UMELÚ INTELIGENCIU (AI) NA GENEROVANIE TOHTO KÓDU:
' Rola: Si expert na MS Access a VBA programovanie.
'
' Úloha: Vytvor VBA procedúru, ktorá prejde tabu¾ku 'Tbl_partner' a vygeneruje
' fiktívny, ale štrukturálne správny IBAN pod¾a krajiny.
'
' NOVINKA (Biznis logika): Pridaj overenie typu partnera zo ståpca 'typ_partnera'
' (Boolean). Ak ide o Dodávate¾a (True), MUSÍ sa mu vygenerova a zapísa IBAN.
' Ak ide o Odberate¾a (False), poui náhodnı faktor 50 na 50 (šanca 0.5), èi sa
' mu IBAN zapíše alebo sa do databázy vloí hodnota Null (simulácia chıbajúcich dát).
' ==============================================================================

' =====================================================================================
' MODUL: Generovanie údajov (V3 - Podmienka pod¾a typu partnera)
' POPIS: Generuje Dummy IBANy s rešpektovaním biznis logiky:
'        - Dodávatelia (ktorım platíme my) musia ma IBAN vdy na 100 %.
'        - Odberatelia (ktorí platia nám) majú 50 % šancu, e ich IBAN u v systéme máme.
' =====================================================================================

Sub DoplnDummyIBANyPrePartnerov()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim kodKrajiny As String
    Dim idPartnera As Integer
    Dim novyIban As String
    Dim jeDodavatel As Boolean
    Dim sancaNaIBAN As Double
    Dim pocetVyplnenych As Integer
    Dim pocetPrazdnych As Integer
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Tbl_partner")
    
    pocetVyplnenych = 0
    pocetPrazdnych = 0
    
    ' Inicializácia náhodnıch èísel
    Randomize
    
    Do While Not rs.EOF
        ' 1. Zistíme, èi ide o Dodávate¾a (True) alebo Odberate¾a (False)
        jeDodavatel = Nz(rs!typ_partnera, False)
        
        ' 2. Bezpeèné naèítanie a ošetrenie kódu krajiny
        kodKrajiny = UCase(Trim(Nz(rs!krajina, "SK")))
        If Len(kodKrajiny) < 2 Then
            kodKrajiny = "SK"
        Else
            kodKrajiny = Left(kodKrajiny, 2)
        End If
        
        idPartnera = rs!PK_partner
        
        ' 3. Generovanie správneho formátu pre danú krajinu
        Select Case kodKrajiny
            Case "SK" ' 24 znakov
                novyIban = "SK" & Int(Rnd() * 90 + 10) & "0900000000" & Format(idPartnera, "0000000000")
            Case "CZ" ' 24 znakov
                novyIban = "CZ" & Int(Rnd() * 90 + 10) & "0100000000" & Format(idPartnera, "0000000000")
            Case "GB" ' 22 znakov
                novyIban = "GB" & Int(Rnd() * 90 + 10) & "BARC" & Format(idPartnera, "00000000000000")
            Case "IE" ' 22 znakov
                novyIban = "IE" & Int(Rnd() * 90 + 10) & "BOFI" & Format(idPartnera, "00000000000000")
            Case "DE" ' 22 znakov
                novyIban = "DE" & Int(Rnd() * 90 + 10) & "10040000" & Format(idPartnera, "0000000000")
            Case "US" ' Simulácia US
                novyIban = "US" & Int(Rnd() * 90 + 10) & "BOFA0000" & Format(idPartnera, "0000000000")
            Case Else ' Univerzálny fallback
                novyIban = kodKrajiny & Int(Rnd() * 90 + 10) & "0000000000" & Format(idPartnera, "0000000000")
        End Select
        
        ' --- ROZHODOVACIA LOGIKA (Biznis podmienka) ---
        rs.Edit
        
        If jeDodavatel = True Then
            ' A: Dodávate¾ MUSÍ ma IBAN
            rs!iban = novyIban
            pocetVyplnenych = pocetVyplnenych + 1
        Else
            ' B: Odberate¾ - náhodná 50/50 šanca
            sancaNaIBAN = Rnd() ' Vygeneruje èíslo od 0 do 1
            If sancaNaIBAN <= 0.5 Then
                rs!iban = novyIban
                pocetVyplnenych = pocetVyplnenych + 1
            Else
                rs!iban = Null
                pocetPrazdnych = pocetPrazdnych + 1
            End If
        End If
        
        rs.Update
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    MsgBox "Logika IBANov úspešne aplikovaná!" & vbCrLf & vbCrLf & _
           "Vyplnené IBANy: " & pocetVyplnenych & vbCrLf & _
           "Prázdne IBANy (Null): " & pocetPrazdnych, vbInformation, "Dáta ošetrené"
End Sub

