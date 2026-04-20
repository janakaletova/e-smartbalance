Attribute VB_Name = "Pomocné funkcie"
Option Compare Database

' ==============================================================================
' AI PROMPT (ZADANIE PRE UMELÚ INTELIGENCIU):
' "Vytvor inteligentného asistenta pre spracovanie dátumov v MS Access.
' Táto funkcia musí vzia akıko¾vek text, ktorı vyzerá ako dátum,
' a správne z neho vyèíta deò, mesiac a rok.
'
' Musí by pripravená na to, e:
' 1. Odde¾ovaèe môu by rôzne (bodky, lomenice alebo pomlèky).
' 2. Poradie môe by európske (deò na zaèiatku) alebo technické (rok na zaèiatku).
' 3. Program nesmie skolabova, ak má pouívate¾ v poèítaèi nastavenı inı
'    jazyk alebo formát èasu, ne je v tom súbore."
' ==============================================================================

' ------------------------------------------------------------------------------
' POMOCNÁ FUNKCIA: Inteligentnı prekladaè dátumov
' POPIS: Táto funkcia slúi ako 'ochrannı štít'. Zoberie text z bankového vıpisu
'        a premení ho na skutoènı dátum, ktorému databáza rozumie za kadıch
'        okolností. Poradí si s formátmi ako 06.01.2026, 2026-01-06 aj 06/01/26.
' ------------------------------------------------------------------------------
Public Function InteligentnyParserDatumu(ByVal strDatum As String) As Date
    Dim dParts() As String
    Dim sClean As String
    Dim r, m, d As Integer
    
    ' 1. Vyèistíme text - zjednotíme rôzne odde¾ovaèe na bodky
    sClean = Replace(strDatum, "/", ".")
    sClean = Replace(sClean, "-", ".")
    sClean = Trim(sClean)
    
    ' 2. Rozdelíme text na jednotlivé kúsky (deò, mesiac, rok)
    dParts = Split(sClean, ".")
    
    If UBound(dParts) = 2 Then
        ' Zisujeme, kde sa nachádza rok (h¾adáme 4-miestne èíslo)
        If Len(dParts(2)) = 4 Then
            ' Benı formát: 06.01.2026 (Deò.Mesiac.Rok)
            r = CInt(dParts(2))
            m = CInt(dParts(1))
            d = CInt(dParts(0))
        ElseIf Len(dParts(0)) = 4 Then
            ' Technickı formát: 2026.01.06 (Rok.Mesiac.Deò)
            r = CInt(dParts(0))
            m = CInt(dParts(1))
            d = CInt(dParts(2))
        Else
            ' Skrátenı rok: 06.01.26 (Pridáme 2000)
            r = 2000 + CInt(dParts(2))
            m = CInt(dParts(1))
            d = CInt(dParts(0))
        End If
        
        ' 3. Zloíme bezpeènı dátum, ktorı je imúnny voèi nastaveniam Windows
        InteligentnyParserDatumu = DateSerial(r, m, d)
    Else
        ' Ak je formát úplne neštandardnı, skúsime poslednú záchranu
        InteligentnyParserDatumu = CDate(strDatum)
    End If
End Function

' ------------------------------------------------------------------------------
' PROCEDÚRA PRE FORMULÁR: SpustiImportZGui
' POPIS: Otvorí Windows okno pre vıber CSV súboru, zapíše cestu do formulára,
'        spustí import dát a automaticky vyfiltruje podformulár tak, aby
'        zobrazil len platby z tohto konkrétneho importu.
' ------------------------------------------------------------------------------
Public Sub SpustiImportZGui(ByRef aktualnyFormular As Form)
    Dim fd As Object
    Dim vybranySubor As String
    
    ' Otvorenie štandardného Windows okna pre vıber súboru
    Set fd = Application.FileDialog(3) ' 3 = msoFileDialogFilePicker
    
    With fd
        .Title = "Vyberte CSV súbor s bankovım vıpisom"
        .Filters.Clear
        .Filters.Add "CSV Súbory", "*.csv"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            vybranySubor = .SelectedItems(1)
            
            ' 1. Zapíšeme cestu do textového po¾a na formulári (ak existuje)
            On Error Resume Next
            aktualnyFormular.Controls("txtCestaKSuboru").Value = vybranySubor
            On Error GoTo 0
            
            ' 2. ZAVOLÁME NÁŠ HLAVNİ IMPORTNİ SKRIPT
            Call ImportujBankovyVypis(vybranySubor)
            
            ' 3. Zobrazenie a vyfiltrovanie dát v podformulári
            ' (Predpokladáme, e podformulár sa volá "subfrm_Platby")
            On Error Resume Next
            With aktualnyFormular.Controls("subfrm_Platby").Form
                ' Vyfiltrujeme záznamy, kde sa zdrojovı súbor zhoduje s vybranou cestou
                .Filter = "nazov_zdrojoveho_suboru = '" & vybranySubor & "'"
                .FilterOn = True
                .Requery
            End With
            On Error GoTo 0
            
        Else
            MsgBox "Import bol zrušenı.", vbExclamation, "Zrušené"
        End If
    End With
    
    Set fd = Nothing
End Sub

' ==============================================================================
' AI PROMPT (ZMENOVÁ POIADAVKA):
' "Uprav zdroj dát pre podformulár tak, aby namiesto èíselnıch ID zobrazoval
' reálne názvy z èíselníkov (Mena, Spôsob platby). Urob to pomocou SQL dotazu,
' ktorı tieto tabu¾ky prepojí. Filter na zdrojovı súbor musí zosta zachovanı
' a funkènı aj nad tımto novım dotazom."
' ==============================================================================

Public Sub AplikujFilterImportu(ByRef frm As Form)
    Dim filterPath As String
    filterPath = Replace(Nz(frm.txtCestaKSuboru, ""), "'", "''")
    
    ' Filter teraz beí nad dotazom qry_Platby_Prehlad
    With frm.subfrm_Platby.Form
        If filterPath = "" Then
            .Filter = "[nazov_zdrojoveho_suboru] = ''"
            .FilterOn = True
        Else
            .Filter = "[nazov_zdrojoveho_suboru] = '" & filterPath & "'"
            .FilterOn = True
        End If
        .Requery
    End With
End Sub

Public Sub AplikujFilterNBS(ByRef frm As Form)
    Dim sqlDatum As String
    
    ' Kontrola, èi je dátum zadanı
    If IsNull(frm.txtDatumKurzu) Then
        ' Ak je políèko prázdne, podformulár ostane prázdny
        frm.subfrm_Kurzy.Form.Filter = "[time] IS NULL"
    Else
        ' POZOR: Access SQL vyaduje dátum v US formáte #mm/dd/yyyy#
        ' Spätné lomky \/ zabezpeèia, e Access nepouije slovenské bodky
        sqlDatum = Format(frm.txtDatumKurzu, "mm\/dd\/yyyy")
        
        frm.subfrm_Kurzy.Form.Filter = "[time] = #" & sqlDatum & "#"
    End If
    
    ' Zapnutie filtra a obnovenie dát
    frm.subfrm_Kurzy.Form.FilterOn = True
    frm.subfrm_Kurzy.Requery
End Sub

