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
'        dogeneruje chıbajúci poèet do 350. Zabezpeèuje, aby dátumy vystavenia
'        pripadli vıluène na pracovné dni (preskakuje víkendy) a aby nepresiahli
'        dátum 17.4.2026. Taktie náhodne simuluje vydané a prijaté faktúry v rôznych menách.
' ------------------------------------------------------------------------------
Sub DoplnFakturyDo350_IbaPracovneDni()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim datum As Date
    Dim pocetExistujucich As Long
    Dim pocetNaVygenerovanie As Long
    Dim i As Long
    Dim partnerID As Integer
    Dim menaID As Integer
    Dim typ As Boolean
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Tbl_faktura")
    
    ' 1. Zistíme, ko¾ko faktúr v tabu¾ke u reálne máš
    pocetExistujucich = DCount("*", "Tbl_faktura")
    pocetNaVygenerovanie = 350 - pocetExistujucich
    
    If pocetNaVygenerovanie <= 0 Then
        MsgBox "U máš " & pocetExistujucich & " faktúr! Nie je potrebné generova ïalšie.", vbInformation
        Exit Sub
    End If
    
    Randomize
    ' Zaèíname generova od polovice januára
    datum = DateSerial(2026, 1, 20)
    
    For i = 1 To pocetNaVygenerovanie
        ' Kadú 3. faktúru posunieme o deò dopredu, aby boli nasekané tesne za sebou
        If i Mod 3 = 0 Then datum = datum + 1
        
        ' K¾úèová funkcia: Preskoèíme víkendy (Sobota = 7, Nede¾a = 1 vo vbSunday)
        While Weekday(datum, vbMonday) > 5
            datum = datum + 1
        Wend
        
        ' Zastavíme generovanie na 17.4.2026 (Piatok), aby sme nepresiahli dnešnı deò
        If datum > DateSerial(2026, 4, 17) Then
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
        rs!Typ_faktury = typ
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
           "Teraz máš v tabu¾ke presne 350 záznamov, bez víkendov a nasekané do 17. 4. 2026.", vbInformation
End Sub


' ==============================================================================
' AI PROMPT (ZMENOVÁ POIADAVKA):
' "Uprav generátor bankového vıpisu tak, aby namiesto jedného ve¾kého súboru
' vygeneroval samostatné CSV súbory za kadı kalendárny mesiac (napr. vıpis_01.csv,
' vıpis_02.csv atï.).
'
' Logika rozdelenia:
' 1. Program prejde všetky faktúry a pre kadú nasimuluje platbu (scenáre).
' 2. Platba sa automaticky zapíše do súboru prislúchajúcemu danému mesiacu.
' 3. Zachovaj ochranu pred budúcimi dátumami (dnes je 19.4.2026).
' 4. Bankové poplatky generuj mesaène a vlo ich vdy do správneho mesaèného súboru."
' ==============================================================================

' ------------------------------------------------------------------------------
' PROCEDÚRA: GenerujMesenéBankovéVıpisyCSV
' POPIS: Vytvorí sadu CSV súborov rozdelenıch pod¾a mesiacov.
' ------------------------------------------------------------------------------
Sub GenerujMesenéBankovéVıpisyCSV()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fso As Object
    Dim tsArray(1 To 12) As Object ' Pole pre súborové streamy (pre kadı mesiac jeden)
    Dim filePath As String
    Dim varSymbol As String
    Dim suma As Double
    Dim datumVystavenia As Date
    Dim datumPlatby As Date
    Dim dnesnyDatum As Date
    Dim menaID As Integer
    Dim menaStr As String
    Dim scenario As Integer
    Dim outLine As String
    Dim m As Integer
    
    Dim partnerIBAN As String
    Dim partnerNazov As String
    Dim skutocnyIbanPreCSV As String
    
    dnesnyDatum = Date ' 19. apríl 2026
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set db = CurrentDb
    
    ' 1. PRÍPRAVA SÚBOROV (Otvoríme súbory pre všetky relevantné mesiace)
    For m = 1 To Month(dnesnyDatum)
        filePath = CurrentProject.Path & "\bankovy_vypis_2026_" & Format(m, "00") & ".csv"
        Set tsArray(m) = fso.CreateTextFile(filePath, True, True)
        ' Zápis hlavièky do kadého mesaèného súboru
        tsArray(m).WriteLine "Var_Symbol_Banka;Suma_Prijata;Mena_Pohybu;Datum_Prijmu;IBAN_Protistrany;Nazov_Protistrany"
    Next m
    
    ' 2. NAÈÍTANIE FAKTÚR
    Dim sqlQuery As String

    sqlQuery = "SELECT F.*, P.iban, P.názov AS NazovPartnera, P.typ_partnera " & _
               "FROM Tbl_faktura AS F INNER JOIN Tbl_partner AS P " & _
               "ON F.FK_partner_ID = P.PK_partner " & _
               "ORDER BY F.Datum_vystavenia"
               
    Set rs = db.OpenRecordset(sqlQuery)
    Randomize
    
    ' 3. GENERUJEME PLATBY A ROZDE¼UJEME ICH DO SÚBOROV
    Do While Not rs.EOF
        varSymbol = Nz(rs!Variabilny_symbol, "")
        suma = Nz(rs!suma, 0)
' --- OPRAVA: Mínusové sumy pre dodávate¾ov ---
        ' Ak je partner Dodávate¾ (True), my platíme jemu -> peniaze z nášho úètu odchádzajú
        If rs!typ_partnera = True Then
            suma = suma * -1
        End If
        ' ---------------------------------------------
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
        
        scenario = Int(Rnd() * 100) + 1
        
        ' Logika scenárov (vıpoèet dátumu platby)
        Select Case scenario
            Case 1 To 50 ' Ideálna platba
                datumPlatby = datumVystavenia + Int(Rnd() * 14) + 1
                If datumPlatby <= dnesnyDatum Then
                    outLine = varSymbol & ";" & Replace(Format(suma, "0.00"), ",", ".") & ";" & menaStr & ";" & Format(datumPlatby, "dd.mm.yyyy") & ";" & skutocnyIbanPreCSV & ";" & partnerNazov
                    tsArray(Month(datumPlatby)).WriteLine outLine ' Zápis do správneho mesiaca
                End If
                
            Case 51 To 65 ' Preklep vo VS
                datumPlatby = datumVystavenia + Int(Rnd() * 10) + 1
                If datumPlatby <= dnesnyDatum Then
                    Dim chybnyVS As String: chybnyVS = varSymbol
                    If InStr(chybnyVS, "0") > 0 Then chybnyVS = Replace(chybnyVS, "0", "O", 1, 1)
                    outLine = chybnyVS & ";" & Replace(Format(suma, "0.00"), ",", ".") & ";" & menaStr & ";" & Format(datumPlatby, "dd.mm.yyyy") & ";" & skutocnyIbanPreCSV & ";" & partnerNazov
                    tsArray(Month(datumPlatby)).WriteLine outLine
                End If
                
            Case 66 To 75 ' Zmena banky
                datumPlatby = datumVystavenia + Int(Rnd() * 10) + 1
                If datumPlatby <= dnesnyDatum Then
                    Dim zmenenyIBAN As String: zmenenyIBAN = "SK" & Int(Rnd() * 90 + 10) & "1100" & Format(Int(Rnd() * 999999999), "000000000000")
                    outLine = varSymbol & ";" & Replace(Format(suma, "0.00"), ",", ".") & ";" & menaStr & ";" & Format(datumPlatby, "dd.mm.yyyy") & ";" & zmenenyIBAN & ";" & partnerNazov
                    tsArray(Month(datumPlatby)).WriteLine outLine
                End If
                
            Case 76 To 90 ' Èiastoèná platba (2 splátky môu by v rôznych mesiacoch!)
                Dim s1 As Double: s1 = Round(suma / 2, 2)
                Dim s2 As Double: s2 = suma - s1
                ' 1. splátka
                datumPlatby = datumVystavenia + Int(Rnd() * 3) + 1
                If datumPlatby <= dnesnyDatum Then
                    tsArray(Month(datumPlatby)).WriteLine varSymbol & ";" & Replace(Format(s1, "0.00"), ",", ".") & ";" & menaStr & ";" & Format(datumPlatby, "dd.mm.yyyy") & ";" & skutocnyIbanPreCSV & ";" & partnerNazov
                    ' 2. splátka
                    datumPlatby = datumPlatby + Int(Rnd() * 15) + 5
                    If datumPlatby <= dnesnyDatum Then
                        tsArray(Month(datumPlatby)).WriteLine varSymbol & ";" & Replace(Format(s2, "0.00"), ",", ".") & ";" & menaStr & ";" & Format(datumPlatby, "dd.mm.yyyy") & ";" & skutocnyIbanPreCSV & ";" & partnerNazov
                    End If
                End If
                
            Case Else ' Úplnı nezmysel
                datumPlatby = datumVystavenia + Int(Rnd() * 7) + 1
                If datumPlatby <= dnesnyDatum Then
                    outLine = "UHRADA" & Int(Rnd() * 99) & ";" & Replace(Format(suma, "0.00"), ",", ".") & ";" & menaStr & ";" & Format(datumPlatby, "dd.mm.yyyy") & ";" & skutocnyIbanPreCSV & ";" & partnerNazov
                    tsArray(Month(datumPlatby)).WriteLine outLine
                End If
        End Select
        
        rs.MoveNext
    Loop
    
    ' 4. ZÁPIS MESAÈNİCH POPLATKOV
    For m = 1 To Month(dnesnyDatum)
        ' Fixnı dátum poplatku
        If m = Month(dnesnyDatum) And 28 > Day(dnesnyDatum) Then
            datumPlatby = dnesnyDatum
        Else
            datumPlatby = DateSerial(2026, m, 28)
        End If
        
        outLine = "POPLATOK " & Format(m, "00") & "/2026;-7.50;EUR;" & Format(datumPlatby, "dd.mm.yyyy") & ";;Mesaènı poplatok za úèet"
        tsArray(m).WriteLine outLine
    Next m
    
    ' 5. UPRATOVANIE (Zatvorenie všetkıch otvorenıch súborov)
    On Error Resume Next
    For m = 1 To 12
        tsArray(m).Close
    Next m
    
    rs.Close
    Set rs = Nothing: Set db = Nothing: Set fso = Nothing
    
    MsgBox "Generovanie mesaènıch vıpisov úspešne dokonèené v prieèinku databázy!", vbInformation
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

