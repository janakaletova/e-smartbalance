Attribute VB_Name = "fuzzy logika p·rovania"
Option Compare Database
Option Explicit

' ==============================================================================
' POMOCN¡ FUNKCIA 1: V˝poËet Levenshteinovej vzdialenosti (PoËet preklepov)
' ==============================================================================
Private Function LevenshteinDistance(ByVal s1 As String, ByVal s2 As String) As Integer
    Dim i As Integer, j As Integer
    Dim l1 As Integer, l2 As Integer
    Dim d() As Integer
    Dim min1 As Integer, min2 As Integer, min3 As Integer
    
    l1 = Len(s1)
    l2 = Len(s2)
    ReDim d(l1, l2)
    
    For i = 0 To l1
        d(i, 0) = i
    Next i
    For j = 0 To l2
        d(0, j) = j
    Next j
    
    For i = 1 To l1
        For j = 1 To l2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                d(i, j) = d(i - 1, j - 1)
            Else
                min1 = d(i - 1, j) + 1
                min2 = d(i, j - 1) + 1
                min3 = d(i - 1, j - 1) + 1
                If min2 < min1 Then min1 = min2
                If min3 < min1 Then min1 = min3
                d(i, j) = min1
            End If
        Next j
    Next i
    LevenshteinDistance = d(l1, l2)
End Function

' ==============================================================================
' POMOCN¡ FUNKCIA 2: V˝poËet percentu·lnej zhody (0 - 100%)
' ==============================================================================
Private Function Similarity(ByVal s1 As String, ByVal s2 As String) As Double
    Dim maxLen As Integer
    maxLen = IIf(Len(s1) > Len(s2), Len(s1), Len(s2))
    If maxLen = 0 Then
        Similarity = 100
    Else
        Similarity = (maxLen - LevenshteinDistance(s1, s2)) / maxLen * 100
    End If
End Function

' ==============================================================================
' HLAVN¡ PROCED⁄RA: Fuzzy p·rovanie s transakËn˝m potvrdzovanÌm
' ==============================================================================
Public Sub Parovanie_FuzzyLogic()
    Dim db As DAO.Database
    Dim wrk As DAO.Workspace
    Dim rsPlatby As DAO.Recordset
    Dim rsFaktury As DAO.Recordset
    
    Dim p_id As Long, f_id As Long
    Dim p_vs As String, f_vs As String
    Dim p_suma As Double, f_zostatok As Double
    
    Dim sim As Double
    Dim threshold As Double
    Dim countPaired As Integer
    Dim promptMsg As String
    Dim ans As VbMsgBoxResult
    
    ' Nastavenie citlivosti na preklepy (napr. 75% zhoda znakov)
    threshold = 75
    countPaired = 0
    
    ' Inicializ·cia transakËnÈho prostredia
    Set wrk = DBEngine.Workspaces(0)
    Set db = CurrentDb
    
    ' SPUSTENIE TRANSAKCIE (Vöetky zmeny sa drûia len v pam‰ti)
    wrk.BeginTrans
    On Error GoTo ErrorHandler
    
    ' 1. NaËÌtanie len NESP¡ROVAN›CH platieb z banky
    Set rsPlatby = db.OpenRecordset("SELECT ID_platby, suma, var_symbol_banka FROM tbl_platba WHERE FK_faktura Is Null AND var_symbol_banka Is Not Null")
    
    If Not rsPlatby.EOF Then
        rsPlatby.MoveFirst
        Do Until rsPlatby.EOF
            p_id = rsPlatby!ID_platby
            ' Odstr·nime prÌpadnÈ medzery pre lepöie porovnanie
            p_vs = Replace(CStr(rsPlatby!var_symbol_banka), " ", "")
            ' Pouûijeme absol˙tnu hodnotu sumy (rieöi problÈm s mÌnusov˝mi ˙hradami dod·vateæom)
            p_suma = Abs(Nz(rsPlatby!suma, 0))
            
            ' 2. Pre kaûd˙ platbu otvorÌme zoznam NEUHRADEN›CH fakt˙r a zostatkov
            Set rsFaktury = db.OpenRecordset("SELECT ID_faktura, Variabilny_symbol, Chyba_Doplatit FROM qry_Faktury_Na_Vyber WHERE Variabilny_symbol Is Not Null")
            
            If Not rsFaktury.EOF Then
                rsFaktury.MoveFirst
                Do Until rsFaktury.EOF
                    f_id = rsFaktury!ID_faktura
                    f_vs = Replace(CStr(rsFaktury!Variabilny_symbol), " ", "")
                    f_zostatok = Abs(Nz(rsFaktury!Chyba_Doplatit, 0))
                    
                    ' V˝poËet pravdepodobnosti, ûe ide o preklep
                    sim = Similarity(p_vs, f_vs)
                    
                    ' Podmienka 1: Variabiln˝ symbol sa musÌ podobaù na aspoÚ 75%
                    If sim >= threshold Then
                        ' Podmienka 2: Suma na platbe je rovnak· ako aktu·lny zostatok na doplatenie fakt˙ry
                        If Round(p_suma, 2) = Round(f_zostatok, 2) Then
                            
                            ' Naöli sme zhodu! ZapÌöeme ju cez aktualizaËn˝ SQL prÌkaz
                            ' Tento prÌkaz sa vÔaka wrk.BeginTrans zatiaæ neuloûÌ natrvalo
                            db.Execute "UPDATE tbl_platba SET " & _
                                       "FK_faktura = " & f_id & ", " & _
                                       "autoparovaci_dotaz = 'fuzzy logic', " & _
                                       "sparovane_automaticky = True " & _
                                       "WHERE ID_platby = " & p_id
                                       
                            countPaired = countPaired + 1
                            Exit Do ' Platba je vybaven·, preskoËÌme na Ôalöiu platbu
                            
                        End If
                    End If
                    rsFaktury.MoveNext
                Loop
            End If
            
            If Not rsFaktury Is Nothing Then rsFaktury.Close
            rsPlatby.MoveNext
        Loop
    End If
    
    ' 3. Vyhodnotenie transakcie a zobrazenie okna na potvrdenie
    If countPaired > 0 Then
        promptMsg = "Algoritmus (Fuzzy Logic) identifikoval " & countPaired & " platieb s preklepom vo VS." & vbCrLf & _
                    "Suma t˝chto platieb sedÌ so zostatkom na fakt˙rach a VS vykazuje vysok˙ podobnosù." & vbCrLf & vbCrLf & _
                    "Chcete POTVRDIç t˙to transakciu a z·v‰zne ich sp·rovaù?"
                    
        ans = MsgBox(promptMsg, vbYesNo + vbQuestion + vbDefaultButton2, "Potvrdenie inteligentnÈho p·rovania")
        
        If ans = vbYes Then
            wrk.CommitTrans ' UloûÌ vöetky zmeny do datab·zy
            MsgBox "V˝borne! " & countPaired & " platieb bolo ˙speöne sp·rovan˝ch.", vbInformation, "Hotovo"
        Else
            wrk.Rollback ' Vr·ti datab·zu do pÙvodnÈho stavu pred spustenÌm kÛdu
            MsgBox "Oper·cia bola zruöen·. Z·znamy zostali nesp·rovanÈ.", vbExclamation, "ZruöenÈ"
        End If
    Else
        wrk.Rollback ' Upratovanie
        MsgBox "Nenaöli sa ûiadne platby, ktorÈ by spÂÚali podmienky (zhoda sumy a preklep vo VS).", vbInformation, "éiadne v˝sledky"
    End If
    
    ' 4. Uvoænenie pam‰te
    On Error Resume Next
    rsPlatby.Close
    Set rsPlatby = Nothing
    Set rsFaktury = Nothing
    Set db = Nothing
    Set wrk = Nothing
    Exit Sub
    
ErrorHandler:
    ' Ak niekde nastane IT chyba (naprÌklad zamknut· tabuæka), vr·time zmeny sp‰ù
    wrk.Rollback
    MsgBox "Nastala neoËak·van· chyba: " & Err.Description, vbCritical, "Kritick· chyba"
End Sub

