Attribute VB_Name = "Main"
Option Explicit
' Workbooks
Public wb As Workbook
' Worksheets
Public sht_code As Worksheet, sht_ryhmät As Worksheet, sht_lapset As Worksheet, sht_rsuodatus As Worksheet, _
sht_lsuodatus As Worksheet, sht_lasna As Worksheet, sht_lasna2 As Worksheet, sht_päiväkoti As Worksheet
' ListObjects
Public tbl_ryhmät As ListObject, tbl_lapset As ListObject
' Lapset
Public ls_järjestys As Range, ls_kutsumanimi As Range, _
ls_kokonimi As Range, ls_ryhmä As Range, ls_dieetti As Range, ls_työntekijä As Range, ls_matulo As Range, _
ls_malähtö As Range, ls_titulo As Range, ls_tilähtö As Range, ls_ketulo As Range, ls_kelähtö As Range, _
ls_totulo As Range, ls_tolähtö As Range, ls_petulo As Range, ls_pelähtö As Range, ls_matulo2 As Range, _
ls_malähtö2 As Range, ls_titulo2 As Range, ls_tilähtö2 As Range, ls_ketulo2 As Range, ls_kelähtö2 As Range, _
ls_totulo2 As Range, ls_tolähtö2 As Range, ls_petulo2 As Range, ls_pelähtö2 As Range, ls_sukunimi As Range, _
ls_päivystys As Range, ls_latulo As Range, ls_lalähtö As Range, ls_sutulo As Range, ls_sulähtö As Range, _
ls_pienryhmä As Range, ls_arkiPoissa As Range, ls_vklPoissa As Range
' Ryhmät
Public rs_järjestys As Range, rs_käytössä As Range, rs_ryhmännimi As Range, rs_aakkosjärjestys As Range, _
rs_boldaus1 As Range, rs_boldaus2 As Range, rs_pboldaus1 As Range, rs_pboldaus2 As Range, _
rs_ruokatulostus As Range, rs_ruokaAsetukset As Range, rs_ruokayhdistys As Range, rs_aamupala1 As Range, _
rs_aamupala2 As Range, rs_lounas1 As Range, rs_lounas2 As Range, rs_välipala1 As Range, rs_välipala2 As Range, _
rs_päivällinen1 As Range, rs_päivällinen2 As Range, rs_iltapala1 As Range, rs_iltapala2 As Range, _
rs_listatulostus As Range, rs_listayhdistys As Range, rs_yhdistettynimi As Range, rs_yhdistettytyyli As Range, _
rs_päivälaput As Range, rs_plabc As Range, rs_plpohja As Range, rs_pltyhjät As Range, rs_päivystys As Range
' Misc Ranges
Public rng_ryhmät As Range, rng_lapset As Range, laps As Range
' Variables
Public Päivät As Long, Ruokailut As Long, Sarakkeet As Long, ekaDieettiRivi As Long, _
tokaDieettiRivi As Long, kolDieettiRivi As Long, nelDieettiRivi As Long, _
VL_Vkl As Long, PL_Poissa As Long, LapsetSarake As Long, ListaLisäys As Long, _
Rivi As Long, pienennys As Long, Ruokakoonti As Boolean, koontiruoat() As Integer

Sub Pääohjelma()
    '1       2         3         4         5         6         7         8         9        10        11        12        13        14
    '2345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678902345
    Dim af_lapset As ListObject
    Dim rng_lapsisuodatus As Range

    ' Testiympäristö
    Dim testiymparisto As Boolean
    testiymparisto = 0
    If testiymparisto = 0 Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    End If

    Dim sort_lapset_ryhmä As Range: Set sort_lapset_ryhmä = Range("tbl_lapset[Ryhmä]")
    Dim sort_lapset_aakkoset As Range: Set sort_lapset_aakkoset = Range("tbl_lapset[Koko nimi]")

    ' Näytetään kaikki välilehdet
    Dim WSname As Variant
    For Each WSname In Array("lasna", "Code", "Pohja", "Aamuilta_pohja", "R_arki", "R_ilta", "R_arkikaikki", "R_kaikki", "pl_pysty", "pl_pystywc", "pl_vaaka", "pl_vaakawc", "Ruokakoonti_pohja")
        Worksheets(WSname).Visible = True
    Next WSname

    Set wb = ThisWorkbook

    Set sht_code = wb.Worksheets("Code")
    Set sht_ryhmät = wb.Worksheets("Ryhmät")
    Set sht_lapset = wb.Worksheets("Lapset")
    Set sht_päiväkoti = wb.Worksheets("Päiväkoti")

    Set tbl_ryhmät = sht_ryhmät.ListObjects("tbl_ryhmät")
    Set tbl_lapset = sht_lapset.ListObjects("tbl_lapset")

    ' lapset-taulukon sarakkeiden nimet
    Set ls_järjestys = tbl_lapset.ListColumns(1).Range
    Set ls_kutsumanimi = tbl_lapset.ListColumns(2).Range
    Set ls_kokonimi = tbl_lapset.ListColumns(3).Range
    Set ls_ryhmä = tbl_lapset.ListColumns(4).Range
    Set ls_dieetti = tbl_lapset.ListColumns(5).Range
    Set ls_työntekijä = tbl_lapset.ListColumns(6).Range
    Set ls_päivystys = tbl_lapset.ListColumns(7).Range
    Set ls_matulo = tbl_lapset.ListColumns(8).Range
    Set ls_malähtö = tbl_lapset.ListColumns(9).Range
    Set ls_titulo = tbl_lapset.ListColumns(10).Range
    Set ls_tilähtö = tbl_lapset.ListColumns(11).Range
    Set ls_ketulo = tbl_lapset.ListColumns(12).Range
    Set ls_kelähtö = tbl_lapset.ListColumns(13).Range
    Set ls_totulo = tbl_lapset.ListColumns(14).Range
    Set ls_tolähtö = tbl_lapset.ListColumns(15).Range
    Set ls_petulo = tbl_lapset.ListColumns(16).Range
    Set ls_pelähtö = tbl_lapset.ListColumns(17).Range
    Set ls_latulo = tbl_lapset.ListColumns(18).Range
    Set ls_lalähtö = tbl_lapset.ListColumns(19).Range
    Set ls_sutulo = tbl_lapset.ListColumns(20).Range
    Set ls_sulähtö = tbl_lapset.ListColumns(21).Range
    Set ls_sukunimi = tbl_lapset.ListColumns(22).Range
    Set ls_arkiPoissa = tbl_lapset.ListColumns(23).Range
    Set ls_vklPoissa = tbl_lapset.ListColumns(24).Range

    ' ryhmät-taulukon sarakkeiden nimet
    Set rs_järjestys = tbl_ryhmät.ListColumns(1).Range
    Set rs_käytössä = tbl_ryhmät.ListColumns(2).Range
    Set rs_ryhmännimi = tbl_ryhmät.ListColumns(3).Range
    Set rs_aakkosjärjestys = tbl_ryhmät.ListColumns(4).Range
    Set rs_boldaus1 = tbl_ryhmät.ListColumns(5).Range
    Set rs_boldaus2 = tbl_ryhmät.ListColumns(6).Range
    Set rs_pboldaus1 = tbl_ryhmät.ListColumns(7).Range
    Set rs_pboldaus2 = tbl_ryhmät.ListColumns(8).Range
    Set rs_ruokatulostus = tbl_ryhmät.ListColumns(9).Range
    Set rs_ruokaAsetukset = tbl_ryhmät.ListColumns(10).Range
    Set rs_ruokayhdistys = tbl_ryhmät.ListColumns(11).Range
    Set rs_aamupala1 = tbl_ryhmät.ListColumns(12).Range
    Set rs_aamupala2 = tbl_ryhmät.ListColumns(13).Range
    Set rs_lounas1 = tbl_ryhmät.ListColumns(14).Range
    Set rs_lounas2 = tbl_ryhmät.ListColumns(15).Range
    Set rs_välipala1 = tbl_ryhmät.ListColumns(16).Range
    Set rs_välipala2 = tbl_ryhmät.ListColumns(17).Range
    Set rs_päivällinen1 = tbl_ryhmät.ListColumns(18).Range
    Set rs_päivällinen2 = tbl_ryhmät.ListColumns(19).Range
    Set rs_iltapala1 = tbl_ryhmät.ListColumns(20).Range
    Set rs_iltapala2 = tbl_ryhmät.ListColumns(21).Range
    Set rs_listatulostus = tbl_ryhmät.ListColumns(22).Range
    Set rs_listayhdistys = tbl_ryhmät.ListColumns(23).Range
    Set rs_yhdistettynimi = tbl_ryhmät.ListColumns(24).Range
    Set rs_yhdistettytyyli = tbl_ryhmät.ListColumns(25).Range
    Set rs_päivystys = tbl_ryhmät.ListColumns(26).Range
    Set rs_päivälaput = tbl_ryhmät.ListColumns(27).Range
    Set rs_plabc = tbl_ryhmät.ListColumns(28).Range
    Set rs_plpohja = tbl_ryhmät.ListColumns(29).Range
    Set rs_pltyhjät = tbl_ryhmät.ListColumns(30).Range

    Set rng_ryhmät = tbl_ryhmät.ListColumns(3).DataBodyRange
    Set rng_lapset = tbl_lapset.ListColumns(2).DataBodyRange
    
    If sht_päiväkoti.Range("G8").Value = "Kyllä" Then
        Ruokakoonti = True
        ReDim koontiruoat(6, 4)
        Dim nollarivit As Integer, nollasarakkeet As Integer
        For nollarivit = 0 To 6
            For nollasarakkeet = 0 To 4
                koontiruoat(nollarivit, nollasarakkeet) = 0
            Next nollasarakkeet
        Next nollarivit
    Else
        Ruokakoonti = False
    End If
    
    Dim merkki As String
    Dim ei_merkki As String

    merkki = "X"
    ei_merkki = ""

    ' Nollataan suodatukset
    Call PoistaSuodatukset

    ' Poistetaan vanhat ryhmät
    Dim lRow As Long
    Dim vanharyh As Range
    lRow = sht_code.Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    If lRow <> 1 Then
        sht_code.Select
        For Each vanharyh In sht_code.Range(Cells(2, 2), Cells(lRow, 2))
            If Not Sheets(vanharyh.Value) Is Nothing Then
                Sheets(vanharyh.Value).Delete
                Sheets(vanharyh.Value & "_ruoka").Delete
                Sheets(vanharyh.Value & "_ma").Delete
                Sheets(vanharyh.Value & "_ti").Delete
                Sheets(vanharyh.Value & "_ke").Delete
                Sheets(vanharyh.Value & "_to").Delete
                Sheets(vanharyh.Value & "_pe").Delete
            End If
        Next vanharyh
        sht_code.Range(Cells(2, 2), Cells(lRow, 2)).Value = vbNullString
    End If
    Sheets("Aamulista").Delete
    Sheets("Iltalista").Delete
    On Error GoTo 0

    ' Kun kutsumanimen kohdalla ei lue mitään --> koko rivi tuhotaan
    Call Rivienpoisto(rng_ryhmät)
    Call Rivienpoisto(rng_lapset)

    ' Nimet
    Dim sort_ryhmät_järjestys As Range
    Set sort_ryhmät_järjestys = rs_järjestys

    Call AlkuTarkistus

    ' Suodatus: Käytössä olevat
    tbl_ryhmät.ShowAutoFilter = True
    'tbl_ryhmät.Range.AutoFilter Field:=rs_käytössä.Column, Criteria1:="Kyllä"

    ' Sorttaus: Järjestyksen mukaan, käänteisesti
    With tbl_ryhmät.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rs_käytössä, SortOn:=xlSortOnValues, Order:=xlDescending
        .SortFields.Add Key:=rs_järjestys, SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .Apply
    End With

    ' Aloitetaan numerointi ykkösestä
    Dim Ryhmänumero As Long: Ryhmänumero = 1
    Dim Code_ryhmät() As Variant

    Dim ryh As Range
    ' Viikkolista & päivälappu -yhdistelytapa
    Dim VL_Yhd As Long, PL_Yhd As Long

    'Errorit pois jotta ei valittaisi kun yhtään ryhmää ei käytössä
    On Error Resume Next
    Dim Ryhmäsolut As Long
    Ryhmäsolut = rng_ryhmät.SpecialCells(xlCellTypeVisible).Cells.Count
    If Ryhmäsolut > 0 Then
        On Error GoTo 0

        ' Luupataan ryhmät yksi kerrallaan läpi (näkyvät solut suodatuksen ja filtterin jälkeen)
        For Each ryh In rng_ryhmät.SpecialCells(xlCellTypeVisible)
            ' Onko ryhmä käytössä
            If sht_ryhmät.Cells(ryh.Row, rs_käytössä.Column) = "Kyllä" Then
        
                ' Nollaa suodatukset
                tbl_lapset.AutoFilter.ShowAllData
                ' Ruokalappu
                If sht_ryhmät.Cells(ryh.Row, rs_ruokatulostus.Column).Value <> "Ei" Then
                    Call Ryhmä_Ruoka(sht_ryhmät.Cells(ryh.Row, rs_ruokatulostus.Column).Value, sht_ryhmät.Cells(ryh.Row, rs_ryhmännimi.Column).Value, _
                                     sht_ryhmät.Cells(ryh.Row, rs_aakkosjärjestys.Column).Value, Ryhmänumero, _
                                     sht_ryhmät.Cells(ryh.Row, rs_ruokatulostus.Column).Value, sht_ryhmät.Cells(ryh.Row, rs_ruokaAsetukset.Column).Value, _
                                     spaceremove(sht_ryhmät.Cells(ryh.Row, rs_ruokayhdistys.Column).Value), _
                                     Code_ryhmät, sht_ryhmät.Cells(ryh.Row, rs_päivystys.Column).Value, ryh.Row)
                End If
                
        
                ' VL & PL Yhdistelytapa
                Select Case sht_ryhmät.Cells(ryh.Row, rs_yhdistettytyyli.Column).Value
                Case "Viikkolista & päivälaput"
                    VL_Yhd = 0
                    PL_Yhd = 0
                Case "Viikkolista & yhdistetyt päivälaput"
                    VL_Yhd = 0
                    PL_Yhd = 1
                Case "Viikkolista & 2-puoleiset päivälaput"
                    VL_Yhd = 0
                    PL_Yhd = 2
                Case "Yhdistetty viikkolista & päivälaput"
                    VL_Yhd = 1
                    PL_Yhd = 0
                Case "Yhdistetty viikkolista & yhdistetyt päivälaput"
                    VL_Yhd = 1
                    PL_Yhd = 1
                Case "Yhdistetty viikkolista & 2-puoleiset päivälaput"
                    VL_Yhd = 1
                    PL_Yhd = 2
                End Select
        
                ' Viikkolistan tulostus
                Select Case sht_ryhmät.Cells(ryh.Row, rs_listatulostus.Column).Value
                Case "Ma-pe"
                    VL_Vkl = 0
                Case "Ma-su"
                    VL_Vkl = 1
                End Select
    
                ' Nollaa suodatukset
                tbl_lapset.AutoFilter.ShowAllData
                ' Hoitoaikalista
                If sht_ryhmät.Cells(ryh.Row, rs_listatulostus.Column).Value <> "Ei" Then
                    Call Ryhmä_Lista(sht_ryhmät.Cells(ryh.Row, rs_ryhmännimi.Column).Value, _
                                     sht_ryhmät.Cells(ryh.Row, rs_aakkosjärjestys.Column).Value, Ryhmänumero, _
                                     sht_ryhmät.Cells(ryh.Row, rs_listatulostus.Column).Value, _
                                     spaceremove(sht_ryhmät.Cells(ryh.Row, rs_listayhdistys.Column).Value), _
                                     sht_ryhmät.Cells(ryh.Row, rs_yhdistettynimi.Column).Value, _
                                     Code_ryhmät, sht_ryhmät.Cells(ryh.Row, rs_päivystys.Column).Value, _
                                     ryh.Row, VL_Yhd, VL_Vkl)
                End If
    
                ' Nollaa suodatukset
                tbl_lapset.AutoFilter.ShowAllData
    
                Select Case sht_ryhmät.Cells(ryh.Row, rs_päivälaput.Column).Value
                Case "Kyllä"
                    PL_Poissa = 0
                    '    PL_Pr = 0
                Case "Kyllä - poissaolevat"
                    PL_Poissa = 1
                    '    PL_Pr = 0
                    'Case "Kyllä + pienryhmät"
                    '    PL_Poissa = 0
                    '    PL_Pr = 1
                    'Case "Kyllä - poissa + pienryhmät"
                    '    PL_Poissa = 1
                    '    PL_Pr = 1
                End Select
    
                ' Päivälaput
                If sht_ryhmät.Cells(ryh.Row, rs_päivälaput.Column).Value <> "Ei" Then
                    Call Ryhmä_Päivälaput(sht_ryhmät.Cells(ryh.Row, rs_ryhmännimi.Column).Value, _
                                          sht_ryhmät.Cells(ryh.Row, rs_plabc.Column).Value, Ryhmänumero, _
                                          spaceremove(sht_ryhmät.Cells(ryh.Row, rs_listayhdistys.Column).Value), _
                                          sht_ryhmät.Cells(ryh.Row, rs_yhdistettynimi.Column).Value, Code_ryhmät, _
                                          sht_ryhmät.Cells(ryh.Row, rs_päivystys.Column).Value, _
                                          sht_ryhmät.Cells(ryh.Row, rs_päivälaput.Column).Value, ryh.Row, PL_Yhd, _
                                          sht_ryhmät.Cells(ryh.Row, rs_plpohja.Column).Value, VL_Yhd, PL_Poissa, sht_ryhmät.Cells(ryh.Row, rs_pltyhjät.Column).Value)
                End If
                Ryhmänumero = Ryhmänumero + 1
                ' Onko käytössä
            End If
        Next ryh
    
        ' Tehdään koontilappu
        If Ruokakoonti Then Call KoontiTulostus(Code_ryhmät)
    
    End If


    ' Aamu- ja iltalistat
    Call PoistaSuodatukset
    Dim aamuiltalistat As String: aamuiltalistat = sht_päiväkoti.Range("I5").Value
    If aamuiltalistat = sht_code.Range("A42").Value Then Call Aamu_ilta_listat

    ' Tapiolan viikonloppulaput
    Dim tapiolavkl As String: tapiolavkl = sht_päiväkoti.Range("I8").Value
    If tapiolavkl = sht_code.Range("A42").Value Then Call Tapiolan_Viikonloppulaput

    If ThisWorkbook.Worksheets("Code").Range("F2").Value = 1 Then Call PoistaTiedot

    ' Lisätään codeen tehdyt ryhmät (jos ryhmiä edes yksi)
    If (Not Not Code_ryhmät) <> 0 Then
        sht_code.Range("B2").Resize(UBound(Code_ryhmät) + 1).Value = Application.Transpose(Code_ryhmät)
    End If
    
    With tbl_lapset.Sort
        ' Sorttaus: Järjestyksen mukaan, käänteisesti
        .SortFields.Clear
        .SortFields.Add Key:=sort_lapset_ryhmä, SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=sort_lapset_aakkoset, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    ' Loppusäädöt
    Call Lopetus("Hoitolistat on tehty.", vbInformation, "Valmista")

End Sub

Sub AdvancedFilter()

    sht_lapset.Range("tbl_lapset").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=sht_code.Range("A100:R105")

End Sub

Sub KoontiTulostus(ByRef Code_ryhmät() As Variant)

    Sheets("Ruokakoonti_pohja").Copy After:=Sheets("Lapset")
    With ActiveSheet
        .Name = "Ruokakoonti"
        .Range("C4:G10") = koontiruoat()
    End With

    ' Napataan ryhmien nimet talteen, jotta voidaan poistaa ne ensi kerralla
    ' Jos Code_Ryhmät Array on tyhjä, muuta kooksi 0 ja lisää ryhmän nimi
    ' Muutoin lisää yksi merkintä + ryhmän nimi
    Dim arrayIsNothing As Boolean
    On Error Resume Next
    arrayIsNothing = IsNumeric(UBound(Code_ryhmät)) And False
    If Err.Number <> 0 Then arrayIsNothing = True
    On Error GoTo 0

    ' Arrayyn ryhmän nimi
    If arrayIsNothing Then
        ReDim Code_ryhmät(0)
        Code_ryhmät(0) = "Ruokakoonti"
    Else
        ReDim Preserve Code_ryhmät(UBound(Code_ryhmät) + 1)
        Code_ryhmät(UBound(Code_ryhmät)) = "Ruokakoonti"
    End If

    ' Ryhmän lisäys Codeen
    sht_code.Range("B2").Resize(UBound(Code_ryhmät) + 1).Value = Application.Transpose(Code_ryhmät)


End Sub

Sub Ryhmä_Ruoka(Pohja As String, Ryhmänimi As String, Aakkosjärjestys As String, Ryhmänumero As Long, R_Tulostus As String, R_Asetukset As String, R_Yhdistely As String, _
                ByRef Code_ryhmät() As Variant, R_päivystys As String, ryhrow As Long)

    ' Napataan ryhmien nimet talteen, jotta voidaan poistaa ne ensi kerralla
    ' Jos Code_Ryhmät Array on tyhjä, muuta kooksi 0 ja lisää ryhmän nimi
    ' Muutoin lisää yksi merkintä + ryhmän nimi
    Dim arrayIsNothing As Boolean
    On Error Resume Next
    arrayIsNothing = IsNumeric(UBound(Code_ryhmät)) And False
    If Err.Number <> 0 Then arrayIsNothing = True
    On Error GoTo 0

    If arrayIsNothing Then
        ReDim Code_ryhmät(0)
        Code_ryhmät(0) = Ryhmänimi & "_ruoka"
    Else
        ReDim Preserve Code_ryhmät(UBound(Code_ryhmät) + 1)
        Code_ryhmät(UBound(Code_ryhmät)) = Ryhmänimi & "_ruoka"
    End If

    ' päivämäärien lisääminen
    Dim pvm_vuosi As Long
    Dim pvm_kk As Long
    Dim pvm_pv As Long
    Dim pvm As Date

    ' Haetaan pvm Codesta ja muunnetaan sopivaan muotoon
    pvm_vuosi = Year(Now)
    pvm_pv = sht_code.[C2].Value2
    pvm_kk = sht_code.[C3].Value2

    ' Päivämäärän koonti
    pvm = DateSerial(pvm_vuosi, pvm_kk, pvm_pv)

    ' Suodata pois tyhjät, näytä vain erityisruokavaliot
    ' jotta voidaan laskea dieetit
    Dim sort_lapset_dieetti As Range
    Set sort_lapset_dieetti = ls_dieetti


    ' ryhmien yhdistelemisen tsekkaus
    With tbl_lapset
        If R_päivystys = "Kyllä" Then
            tbl_lapset.Range.AutoFilter Field:=ls_päivystys.Column, Criteria1:=Ryhmänimi, Operator:=xlFilterValues
        End If
        ' Tsekataan yhdistetyt ruokalaput
        If R_Yhdistely = vbNullString Then
            ' Jos ei yhdistetä mitään, suodatetaan ryhmän nimen mukaan
            If R_päivystys = "Ei" Then
                .Range.AutoFilter Field:=ls_ryhmä.Column, Criteria1:=Ryhmänimi, Operator:=xlFilterValues
            End If
        Else
            ' Lisätään Ryhmänimi yhdistelyyn
            R_Yhdistely = Ryhmänimi + ", " + R_Yhdistely
            ' Tehdään Array ruokalistan yhdistelyryhmistä
            Dim r_yhdistely_array() As String
            r_yhdistely_array = Split(R_Yhdistely, ",", , vbTextCompare)
            ' Suodatetaan arrayn ryhmien mukaan
            tbl_lapset.Range.AutoFilter Field:=ls_ryhmä.Column, Criteria1:=r_yhdistely_array, Operator:=xlFilterValues
        End If
     
        ' Jos kyseessä päivystysryhmä, pidetään vain päivystyslapset
        ' Vain erikoisruokavaliot
        .Range.AutoFilter Field:=ls_dieetti.Column, Criteria1:="<>", Operator:=xlFilterValues
        ' Ei työntekijöitä
        '.Range.AutoFilter Field:=ls_työntekijä.Column, Criteria1:="=", Operator:=xlFilterValues
     
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=sort_lapset_dieetti, SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With

    End With

    ' Ryhmän lisäys Codeen
    sht_code.Range("B2").Resize(UBound(Code_ryhmät) + 1).Value = Application.Transpose(Code_ryhmät)

    Dim pohjaNimi As String
    ' Lasketaan pohjan parametrit
    Select Case Pohja
    Case "Ma-pe"
        Päivät = 5
        Ruokailut = 3
        pohjaNimi = "R_arki"
    Case "Ma-pe, vain ilta"
        Päivät = 5
        Ruokailut = 2
        pohjaNimi = "R_ilta"
    Case "Ma-pe + ilta"
        Päivät = 5
        Ruokailut = 5
        pohjaNimi = "R_arkikaikki"
    Case "Ma-su + ilta"
        Päivät = 7
        Ruokailut = 5
        pohjaNimi = "R_kaikki"
    End Select

    Dim isoFontti As Boolean: isoFontti = False

    Dim päiväPoissa As Boolean
    Dim iltaPoissa As Boolean
    Dim vklPoissa As Boolean

    Dim matulorng As Range

    Dim pv As Long
    Dim lnimi As Range

    Dim a1 As String: a1 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_aamupala1.Column).Value2
    Dim a2 As String: a2 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_aamupala2.Column).Value2
    Dim l1 As String: l1 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_lounas1.Column).Value2
    Dim l2 As String: l2 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_lounas2.Column).Value2
    Dim v1 As String: v1 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_välipala1.Column).Value2
    Dim v2 As String: v2 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_välipala2.Column).Value2
    Dim p1 As String: p1 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_päivällinen1.Column).Value2
    Dim p2 As String: p2 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_päivällinen1.Column).Value2
    Dim i1 As String: i1 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_iltapala1.Column).Value2
    Dim i2 As String: i2 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_iltapala2.Column).Value2

    Dim maTulo As Range
    Dim maLähtö As Range
    Dim ruokaPäivä As Long
    Dim tulo As String
    Dim lähtö As String
    Dim ruokaRivi As Long: ruokaRivi = 4
    Dim tuloArr() As String
    Dim lähtöArr() As String
    Dim yhdArr() As String
    Dim kerta As Long
    Dim tyhjienLasku As Long

    ' Poistetaan dieettilistalta lapset, jotka eivät ole ruokailuissa mukana
    If tyhjasuodatin(rng_lapset) = False Then
        For Each lnimi In rng_lapset.SpecialCells(xlCellTypeVisible)
            tyhjienLasku = 0
            'Jos kyseessä ei ole työntekijä
            If sht_lapset.Cells(lnimi.Row, ls_työntekijä.Column).Value2 = "" Then
                ' Oletuksena lapset ovat pois, kunnes toisin hoitoajoista todetaan
                vklPoissa = True
                iltaPoissa = True
                päiväPoissa = True
                ' Käydään läpi arkipäivät
                For pv = 0 To 4
                    Set matulorng = sht_lapset.Cells(lnimi.Row, ls_matulo.Column + 2 * pv)
                    ' Yksittäinen hoitoaika
                    If IsDate(Format(sht_lapset.Cells(lnimi.Row, ls_matulo.Column + 2 * pv), "h:mm")) = True Then
                        ' Aamupala
                        If matulorng <= a2 And matulorng.Offset(, 1) >= a1 Then päiväPoissa = False
                        ' Lounas
                        If matulorng <= l2 And matulorng.Offset(, 1) >= l1 Then päiväPoissa = False
                        ' Välipala
                        If matulorng <= v2 And matulorng.Offset(, 1) >= v1 Then päiväPoissa = False
                        ' Päivällinen
                        If matulorng <= p2 And matulorng.Offset(, 1) >= p1 Then iltaPoissa = False
                        ' Iltapala
                        If matulorng <= i2 And matulorng.Offset(, 1) >= i1 Then iltaPoissa = False
                
                        ' Useampi hoitoaika
                    ElseIf InStr(matulorng, ",") > 0 Then
                        ' Tehdään hoitoajoista arrayt
                        tuloArr = Split(matulorng, ",")
                        lähtöArr = Split(matulorng.Offset(, 1), ",")
                        ' Käydään array läpi
                        For kerta = 0 To ArrayLen(tuloArr) - 1
                            ' Aamupala
                            If tuloArr(kerta) <= CDate(a2) And lähtöArr(kerta) >= CDate(a1) Then päiväPoissa = False
                            ' Lounas
                            If tuloArr(kerta) <= CDate(l2) And lähtöArr(kerta) >= CDate(l1) Then päiväPoissa = False
                            ' Välipala
                            If tuloArr(kerta) <= CDate(v2) And lähtöArr(kerta) >= CDate(v1) Then päiväPoissa = False
                            ' Päivällinen
                            If tuloArr(kerta) <= CDate(p2) And lähtöArr(kerta) >= CDate(p1) Then iltaPoissa = False
                            ' Iltapala
                            If tuloArr(kerta) <= CDate(i2) And lähtöArr(kerta) >= CDate(i1) Then iltaPoissa = False
                        Next kerta
                        ' Ei hoitoaikaa tai kirjain
                    Else
                        If matulorng = "" Then tyhjienLasku = tyhjienLasku + 1
               
                    End If
                Next pv
                For pv = 5 To 6
                    Set matulorng = sht_lapset.Cells(lnimi.Row, ls_matulo.Column + 2 * pv)
                    ' Yksittäinen hoitoaika
                    If IsDate(Format(sht_lapset.Cells(lnimi.Row, ls_matulo.Column + 2 * pv), "h:mm")) = True Then
                        ' Aamupala
                        If matulorng <= a2 And matulorng.Offset(, 1) >= a1 Then vklPoissa = False
                        ' Lounas
                        If matulorng <= l2 And matulorng.Offset(, 1) >= l1 Then vklPoissa = False
                        ' Välipala
                        If matulorng <= v2 And matulorng.Offset(, 1) >= v1 Then vklPoissa = False
                        ' Päivällinen
                        If matulorng <= p2 And matulorng.Offset(, 1) >= p1 Then vklPoissa = False
                        ' Iltapala
                        If matulorng <= i2 And matulorng.Offset(, 1) >= i1 Then vklPoissa = False
                
                        ' Useampi hoitoaika
                    ElseIf InStr(matulorng, ",") > 0 Then
                        ' Tehdään hoitoajoista arrayt
                        tuloArr = Split(matulorng, ",")
                        lähtöArr = Split(matulorng.Offset(, 1), ",")
                        ' Käydään array läpi
                        For kerta = 0 To ArrayLen(tuloArr) - 1
                            ' Aamupala
                            If tuloArr(kerta) <= CDate(a2) And lähtöArr(kerta) >= CDate(a1) Then vklPoissa = False
                            ' Lounas
                            If tuloArr(kerta) <= CDate(l2) And lähtöArr(kerta) >= CDate(l1) Then vklPoissa = False
                            ' Välipala
                            If tuloArr(kerta) <= CDate(v2) And lähtöArr(kerta) >= CDate(v1) Then vklPoissa = False
                            ' Päivällinen
                            If tuloArr(kerta) <= CDate(p2) And lähtöArr(kerta) >= CDate(p1) Then vklPoissa = False
                            ' Iltapala
                            If tuloArr(kerta) <= CDate(i2) And lähtöArr(kerta) >= CDate(i1) Then vklPoissa = False
                        Next kerta
                    End If
                Next pv
                sht_lapset.Cells(lnimi.Row, ls_arkiPoissa.Column) = ""
                sht_lapset.Cells(lnimi.Row, ls_vklPoissa.Column) = ""
                If päiväPoissa And iltaPoissa Then sht_lapset.Cells(lnimi.Row, ls_arkiPoissa.Column) = "Ei syö"
                If päiväPoissa And iltaPoissa = False Then sht_lapset.Cells(lnimi.Row, ls_arkiPoissa.Column) = "Päivä"
                If päiväPoissa = False And iltaPoissa Then sht_lapset.Cells(lnimi.Row, ls_arkiPoissa.Column) = "Ilta"
                If vklPoissa Then sht_lapset.Cells(lnimi.Row, ls_vklPoissa.Column) = "Ei syö"
                ' Jos ma-pe = "", näytä listalla jos ei erikseen estetty
                If tyhjienLasku = 5 And (R_Asetukset = "Pieni fontti - poissa" Or R_Asetukset = "Iso fontti - poissa") Then
                    sht_lapset.Cells(lnimi.Row, ls_arkiPoissa.Column) = ""
                End If
            End If
        Next lnimi
    End If

    ' Suodatetaan ne dieettilapset pois, jotka eivät ole ruokailuissa mukana
    If R_Asetukset = "Pieni fontti - poissa" Or R_Asetukset = "Pieni fontti - poissa - tyhjät" Or _
       R_Asetukset = "Iso fontti - poissa" Or R_Asetukset = "Iso fontti - poissa - tyhjät" Then
        With tbl_lapset
            ' Päivä
            If Ruokailut = 3 Then .Range.AutoFilter Field:=ls_arkiPoissa.Column, Criteria1:="<>Päivä", Criteria2:="<>Ei syö", Operator:=xlFilterValues
            ' Ilta
            If Ruokailut = 2 Then .Range.AutoFilter Field:=ls_arkiPoissa.Column, Criteria1:="<>Ilta", Criteria2:="<>Ei syö", Operator:=xlFilterValues
            ' Päivä & ilta (ei syö)
            If Ruokailut = 5 Then .Range.AutoFilter Field:=ls_arkiPoissa.Column, Criteria1:="<>Ei syö", Operator:=xlFilterValues
            ' Viikonloput (ei syö)
            If Päivät = 7 Then .Range.AutoFilter Field:=ls_vklPoissa.Column, Criteria1:="<>Ei syö", Operator:=xlFilterValues
        End With
    End If

    If R_Asetukset = "Iso fontti" Or R_Asetukset = "Iso fontti - poissa" Or _
       R_Asetukset = "Iso fontti - poissa - tyhjät" Then isoFontti = True
   
    'sht_lapset.Range.AutoFilter Field:=ls_arkiPoissa.Column, Criteria1:="", Operator:=xlFilterValues
    'sht_lapset.Range.AutoFilter Field:=ls_vklPoissa.Column, Criteria1:="", Operator:=xlFilterValues



    Sarakkeet = 30 / Ruokailut
    If pohjaNimi = "R_ilta" Then Sarakkeet = 10

    ' Montako lasta mahtuu pohjalle. 1. lapulle, 2. lapulle ja molemmille lapuille yhteensä
    ' Ekan lapun ekalle riville: Sarakkeet -1
    ' Muille riveille mahtuu:    Sarakkeet
    Dim maxDieetit As Long
    maxDieetit = (4 * Sarakkeet) - 1

    Dim Dieettilapset As Long: Dieettilapset = 0
    Dim a As Range
    If tyhjasuodatin(rng_lapset) = False Then
        For Each a In rng_lapset.SpecialCells(xlCellTypeVisible)
            Dieettilapset = Dieettilapset + 1
        Next a
    Else

    End If

    ' Selvitetään montako lasta on milläkin dieettirivillä
    If Dieettilapset <> 0 Then
        If Dieettilapset >= Sarakkeet - 1 Then
            ekaDieettiRivi = Sarakkeet - 1
            If Dieettilapset >= ekaDieettiRivi + 2 * Sarakkeet Then
                tokaDieettiRivi = Sarakkeet
                If Dieettilapset >= ekaDieettiRivi + tokaDieettiRivi + 2 * Sarakkeet Then
                    kolDieettiRivi = Sarakkeet
                    If Dieettilapset > ekaDieettiRivi + tokaDieettiRivi + kolDieettiRivi + 2 * Sarakkeet Then
                        Call Lopetus("O-ou! Ryhmässä " & Ryhmänimi & " on liikaa erityisruokavalioita. Ohjelma tukee max " & maxDieetit & " erityisruokavaliota tällä " & Pohja & " -pohjalla." _
                                   & vbCrLf & "Laita minulle viestiä jos tarpeesi on suurempi, niin katsotaan mitä voin tehdä asialle :)" & vbCrLf & "jaakko.haavisto@jyvaskyla.fi", _
                                     vbExclamation, "Virhe")
                    Else
                        nelDieettiRivi = Dieettilapset - ekaDieettiRivi - tokaDieettiRivi - kolDieettiRivi
                    End If
                Else
                    kolDieettiRivi = Dieettilapset - ekaDieettiRivi - tokaDieettiRivi
                End If
            Else
                tokaDieettiRivi = Dieettilapset - ekaDieettiRivi
            End If
        Else
            ekaDieettiRivi = Dieettilapset
        End If
    End If

    Sheets(pohjaNimi).Copy After:=Sheets("Lapset")

    With ActiveSheet
        .Name = Ryhmänimi & "_ruoka"
        If R_Yhdistely = vbNullString Then
            .Range("F1").Value = UCase(Ryhmänimi)
        Else
            .Range("F1").Value = UCase(Join(r_yhdistely_array, ", "))
        End If
        .Range("A1").Value = pvm
        .Range("D1").Value = DateAdd("d", 4, CDate(pvm))
    End With

    Dim Ruokaryhmä As Worksheet: Set Ruokaryhmä = wb.Worksheets(Ryhmänimi & "_ruoka")

    Dim Lapsinumero As Long
    Lapsinumero = 1
    Dim rng_dieetti As Range

    Dim rivi1 As Long: rivi1 = Sarakkeet - 1
    Dim rivi2 As Long: rivi2 = rivi1 + Sarakkeet
    Dim rivi3 As Long: rivi3 = rivi2 + Sarakkeet
    Dim rivi4 As Long: rivi4 = rivi3 + Sarakkeet
    Dim sivu As Long: sivu = 1
    Dim DieettiRivi As Long
    Dim DieettiSarake As Long

    If Dieettilapset > rivi2 Then sivu = 2

    ' Dieettilapset, nimet ja ruksit
    With Ruokaryhmä
        If Dieettilapset <> 0 Then
            For Each rng_dieetti In rng_lapset.SpecialCells(xlCellTypeVisible)
                ' rivi 1
                If Lapsinumero <= rivi1 Then
                    DieettiRivi = 2
                    DieettiSarake = 3 + Ruokailut + Ruokailut * (Lapsinumero - 1)
                
                    ' rivi 2
                ElseIf rivi1 < Lapsinumero And Lapsinumero <= rivi2 Then
                    DieettiRivi = 2 + 2 + Päivät
                    DieettiSarake = 3 + Ruokailut * Lapsinumero - Ruokailut * Sarakkeet
           
                    ' rivi 3
                ElseIf rivi2 < Lapsinumero And Lapsinumero <= rivi3 Then
                    DieettiRivi = 23
                    DieettiSarake = 3 + Ruokailut * Lapsinumero - 2 * (Ruokailut * Sarakkeet)
           
                    ' rivi 4
                ElseIf rivi3 < Lapsinumero And Lapsinumero <= rivi4 Then
                    DieettiRivi = 23 + 2 + Päivät
                    DieettiSarake = 3 + Ruokailut * Lapsinumero - 3 * (Ruokailut * Sarakkeet)
                End If
            
                ' Nimi
                .Cells(DieettiRivi, DieettiSarake) = UCase(rng_dieetti.Value)
                ' Ruksit
                Call Dieetti(rng_dieetti.Offset(0, 1).Value, Ruokaryhmä, Ryhmänimi, Päivät, Ruokailut, 2 + DieettiRivi, DieettiSarake)
            
                Lapsinumero = Lapsinumero + 1
            Next rng_dieetti
        End If

    End With
    ' Nollaa suodatukset
    tbl_lapset.AutoFilter.ShowAllData

    With tbl_lapset
    
        If R_päivystys = "Kyllä" Then
            tbl_lapset.Range.AutoFilter Field:=ls_päivystys.Column, Criteria1:=Ryhmänimi, Operator:=xlFilterValues
        End If
    
        ' Nollaus
        ' Sorttaus valinnan mukaan
        If R_Yhdistely = vbNullString Then
            ' Jos ei yhdistetä mitään, suodatetaan ryhmän nimen mukaan
            If R_päivystys = "Ei" Then
                tbl_lapset.Range.AutoFilter Field:=ls_ryhmä.Column, Criteria1:=Ryhmänimi, Operator:=xlFilterValues
            End If
       
        Else
            tbl_lapset.Range.AutoFilter Field:=ls_ryhmä.Column, Criteria1:=r_yhdistely_array, Operator:=xlFilterValues
        
        End If
    End With

    Dim kaikkiHoitoajat As Boolean: kaikkiHoitoajat = True
    ' Ruokien laskeminen
    With Ruokaryhmä
        For Each lnimi In rng_lapset.SpecialCells(xlCellTypeVisible)
            ' Nollataan tyhjät
            tyhjienLasku = 0
            Set maTulo = sht_lapset.Cells(lnimi.Row, ls_matulo.Column)
            Set maLähtö = sht_lapset.Cells(lnimi.Row, ls_matulo.Column).Offset(, 1)
            ' arkipäivät:  5 pv -> 8
            ' koko viikko: 7 pv -> 12
            For ruokaPäivä = 0 To (2 * Päivät) - 2 Step 2
                ' Yksi hoitoaika
                If IsDate(Format(maTulo.Offset(, ruokaPäivä).Value2, "h:mm")) = True Then
                    tulo = maTulo.Offset(, ruokaPäivä).Value2
                    lähtö = maLähtö.Offset(, ruokaPäivä).Value2
                    If Ruokailut = 2 Then        ' ilta
                        If tulo <= p2 And lähtö >= p1 Then .Cells(ruokaRivi, 3) = .Cells(ruokaRivi, 3) + 1
                        If tulo <= i2 And lähtö >= i1 Then .Cells(ruokaRivi, 4) = .Cells(ruokaRivi, 4) + 1
                    End If
                    If Ruokailut = 3 Then        ' päivä
                        If tulo <= a2 And lähtö >= a1 Then .Cells(ruokaRivi, 3) = .Cells(ruokaRivi, 3) + 1
                        If tulo <= l2 And lähtö >= l1 Then .Cells(ruokaRivi, 4) = .Cells(ruokaRivi, 4) + 1
                        If tulo <= v2 And lähtö >= v1 Then .Cells(ruokaRivi, 5) = .Cells(ruokaRivi, 5) + 1
                    End If
                    If Ruokailut = 5 Then        ' kaikki ruokailut
                        If tulo <= a2 And lähtö >= a1 Then .Cells(ruokaRivi, 3) = .Cells(ruokaRivi, 3) + 1
                        If tulo <= l2 And lähtö >= l1 Then .Cells(ruokaRivi, 4) = .Cells(ruokaRivi, 4) + 1
                        If tulo <= v2 And lähtö >= v1 Then .Cells(ruokaRivi, 5) = .Cells(ruokaRivi, 5) + 1
                        If tulo <= p2 And lähtö >= p1 Then .Cells(ruokaRivi, 6) = .Cells(ruokaRivi, 6) + 1
                        If tulo <= i2 And lähtö >= i1 Then .Cells(ruokaRivi, 7) = .Cells(ruokaRivi, 7) + 1
                    End If
                
                    ' Useampi hoitoaika
                ElseIf InStr(maTulo.Offset(, ruokaPäivä), ",") > 0 Then
                    ' Tehdään hoitoajoista arrayt
                    tuloArr = Split(maTulo.Offset(, ruokaPäivä), ",")
                    lähtöArr = Split(maLähtö.Offset(, ruokaPäivä), ",")
                    ' Käydään array läpi
                    For kerta = 0 To ArrayLen(tuloArr) - 1
                        If Ruokailut = 2 Then    ' ilta
                            If tuloArr(kerta) <= CDate(p2) And lähtöArr(kerta) >= CDate(p1) Then .Cells(ruokaRivi, 3) = .Cells(ruokaRivi, 3) + 1
                            If tuloArr(kerta) <= CDate(i2) And lähtöArr(kerta) >= CDate(i1) Then .Cells(ruokaRivi, 4) = .Cells(ruokaRivi, 4) + 1
                        End If
                        If Ruokailut = 3 Then    ' päivä
                            If tuloArr(kerta) <= CDate(a2) And lähtöArr(kerta) >= CDate(a1) Then .Cells(ruokaRivi, 3) = .Cells(ruokaRivi, 3) + 1
                            If tuloArr(kerta) <= CDate(l2) And lähtöArr(kerta) >= CDate(l1) Then .Cells(ruokaRivi, 4) = .Cells(ruokaRivi, 4) + 1
                            If tuloArr(kerta) <= CDate(v2) And lähtöArr(kerta) >= CDate(v1) Then .Cells(ruokaRivi, 5) = .Cells(ruokaRivi, 5) + 1
                        End If
                        If Ruokailut = 5 Then    ' kaikki ruokailut
                            If tuloArr(kerta) <= CDate(a2) And lähtöArr(kerta) >= CDate(a1) Then .Cells(ruokaRivi, 3) = .Cells(ruokaRivi, 3) + 1
                            If tuloArr(kerta) <= CDate(l2) And lähtöArr(kerta) >= CDate(l1) Then .Cells(ruokaRivi, 4) = .Cells(ruokaRivi, 4) + 1
                            If tuloArr(kerta) <= CDate(v2) And lähtöArr(kerta) >= CDate(v1) Then .Cells(ruokaRivi, 5) = .Cells(ruokaRivi, 5) + 1
                            If tuloArr(kerta) <= CDate(p2) And lähtöArr(kerta) >= CDate(p1) Then .Cells(ruokaRivi, 6) = .Cells(ruokaRivi, 6) + 1
                            If tuloArr(kerta) <= CDate(i2) And lähtöArr(kerta) >= CDate(i1) Then .Cells(ruokaRivi, 7) = .Cells(ruokaRivi, 7) + 1
                        End If
                    
                    Next kerta
                Else
                    ' Lasketaan tyhjiä. Jos ei yhtäkään tyhjiä,
                    If maTulo.Offset(, ruokaPäivä).Value2 = "" Then
                        tyhjienLasku = tyhjienLasku + 1
                        If tyhjienLasku = Päivät Or (tyhjienLasku = 5 And ruokaPäivä = 8) Then kaikkiHoitoajat = False
                    End If
                End If
                ruokaRivi = ruokaRivi + 1
            Next ruokaPäivä
            ruokaRivi = 4
        Next lnimi
        ' Käytetään isoa fonttia jos asetuksissa valittu
        If isoFontti Or kaikkiHoitoajat Then .Range(.Cells(4, 3), .Cells(3 + Päivät, 3 + Ruokailut - 1)).Style = "IsoFontti"
    
        ' Jos ei mene 2 sivulle, poistetaan 2. sivun sisältö ja asetetaan tulostumaan vain 1. sivu
        If sivu = 1 Then
            .Range(Cells(2 * Päivät + 6, 1), Cells(40, 1)).EntireRow.Delete
            If Dieettilapset <= rivi1 Then
                .PageSetup.PrintArea = .Range(Cells(1, 1), Cells(Päivät + 3, 2 + Sarakkeet * Ruokailut)).Address
                .Range(Cells(Päivät + 4, 1), Cells(40, 1)).EntireRow.Delete
            Else
                .PageSetup.PrintArea = .Range(Cells(1, 1), Cells(2 * Päivät + 5, 2 + Sarakkeet * Ruokailut)).Address
            End If
        Else
            If Dieettilapset <= rivi3 Then
                .Range(Cells(23 + 2 + Päivät, 1), Cells(40, 1)).EntireRow.Delete
            End If
        End If
        ' Koonti
        If Ruokakoonti Then
            
            Dim k_päivä As Integer, k_ruokailu As Integer, iltalisä As Integer: iltalisä = 0
            ' Iltaruokien kohdalla siirretään focusta iltalisän verran
            If Ruokailut = 2 Then iltalisä = 3
            
            For k_päivä = 0 To Päivät - 1
                For k_ruokailu = 0 To Ruokailut - 1
                    koontiruoat(k_päivä, k_ruokailu + iltalisä) = koontiruoat(k_päivä, k_ruokailu + iltalisä) + .Cells(4 + k_päivä, 3 + k_ruokailu)
                Next k_ruokailu
            Next k_päivä
        End If
    End With
End Sub

Sub Ryhmä_Lista(Ryhmänimi As String, Aakkosjärjestys As String, Ryhmänumero As Long, VL_Tulostus As String, _
                VL_Yhdistely As String, VL_Nimi As String, ByRef Code_ryhmät() As Variant, VL_päivystys As String, _
                ryhrow As Long, VL_Yhd As Long, VL_Vkl As Long)

    ' Kloonataan pohja
    Sheets("Pohja").Copy After:=Sheets("Lapset")

    Dim Ryhmäsivu As Worksheet
    Dim Nimi As Worksheet

    ' Napataan ryhmien nimet talteen, jotta voidaan poistaa ne ensi kerralla
    ' Jos Code_Ryhmät Array on tyhjä, muuta kooksi 0 ja lisää ryhmän nimi
    ' Muutoin lisää yksi merkintä + ryhmän nimi
    Dim arrayIsNothing As Boolean
    On Error Resume Next
    arrayIsNothing = IsNumeric(UBound(Code_ryhmät)) And False
    If Err.Number <> 0 Then arrayIsNothing = True
    On Error GoTo 0

    ' Jos Yhdistettyjä ryhmiä, käytä niiden omaa nimeä. Muutoin Ryhmän nimeä
    Dim oikeanimi As String
    If VL_Yhd = 0 Or VL_Yhdistely = vbNullString Then
        ActiveSheet.Name = Ryhmänimi
        Set Nimi = wb.Worksheets(Ryhmänimi)
        Set Ryhmäsivu = wb.Worksheets(Ryhmänimi)
        oikeanimi = Ryhmänimi
        ' Arrayyn ryhmän nimi
        If arrayIsNothing Then
            ReDim Code_ryhmät(0)
            Code_ryhmät(0) = Ryhmänimi
        Else
            ReDim Preserve Code_ryhmät(UBound(Code_ryhmät) + 1)
            Code_ryhmät(UBound(Code_ryhmät)) = Ryhmänimi
        End If
        ' Ryhmän nimi capseilla ylälaitaan
        Ryhmäsivu.Range("A1").Value = UCase(Ryhmänimi)
    ElseIf VL_Yhd = 1 Then
        ActiveSheet.Name = VL_Nimi
        Set Nimi = wb.Worksheets(VL_Nimi)
        Set Ryhmäsivu = wb.Worksheets(VL_Nimi)
        oikeanimi = VL_Nimi
        ' Arrayyn yhdistetyn ryhmän nimi
        If arrayIsNothing Then
            ReDim Code_ryhmät(0)
            Code_ryhmät(0) = VL_Nimi
        Else
            ReDim Preserve Code_ryhmät(UBound(Code_ryhmät) + 1)
            Code_ryhmät(UBound(Code_ryhmät)) = VL_Nimi
        End If
        ' Ryhmän nimi capseilla ylälaitaan
        Ryhmäsivu.Range("A1").Value = UCase(VL_Nimi)
    End If

    ' Ryhmän lisäys Codeen
    sht_code.Range("B2").Resize(UBound(Code_ryhmät) + 1).Value = Application.Transpose(Code_ryhmät)

    ' päivämäärien lisääminen
    Dim pvm_vuosi As Long
    Dim pvm_kk As Long
    Dim pvm_pv As Long
    Dim pvm As Date

    ' Haetaan pvm Codesta ja muunnetaan sopivaan muotoon
    pvm_vuosi = Year(Now)
    pvm_pv = sht_code.[C2].Value2
    pvm_kk = sht_code.[C3].Value2

    ' Päivämäärän koonti
    pvm = DateSerial(pvm_vuosi, pvm_kk, pvm_pv)

    ' Lisätään pvm:t oikeille päiville
    Nimi.Range("D1").Value = pvm
    Nimi.Range("G1").Value = DateAdd("d", 1, CDate(Nimi.Range("D1")))
    Nimi.Range("J1").Value = DateAdd("d", 2, CDate(Nimi.Range("D1")))
    Nimi.Range("M1").Value = DateAdd("d", 3, CDate(Nimi.Range("D1")))
    Nimi.Range("P1").Value = DateAdd("d", 4, CDate(Nimi.Range("D1")))
    If VL_Vkl = 1 Then
        Nimi.Range("S1").Value = DateAdd("d", 5, CDate(Nimi.Range("D1")))
        Nimi.Range("V1").Value = DateAdd("d", 6, CDate(Nimi.Range("D1")))
    End If

    Dim i As Long

    Dim pituus As Long
    Dim listaus As Long: listaus = 1

    With tbl_lapset

        ' Jos kyseessä päivystysryhmä, pidetään vain päivystyslapset
        If VL_päivystys = "Kyllä" Then
            tbl_lapset.Range.AutoFilter Field:=ls_päivystys.Column, Criteria1:=Ryhmänimi, Operator:=xlFilterValues
        End If
        ' Ei työntekijöitä listalle
        tbl_lapset.Range.AutoFilter Field:=ls_työntekijä.Column, Criteria1:="=", Operator:=xlFilterValues

        ' Poistetaan koko viikon poissaolevat
        'If VL_Tulostus = "Kyllä - poissaolevat" Then
        '    tbl_lapset.Range.AutoFilter Field:=ls_poissaoleva.Column, Criteria1:="=", Operator:=xlFilterValues
        'End If

        ' Jos ei yhdistettyjä ryhmiä
        If VL_Yhdistely = vbNullString Then
            pituus = 1
            ' Jos ei yhdistetä mitään, suodatetaan ryhmän nimen mukaan
            If VL_päivystys = "Ei" Then
                .Range.AutoFilter Field:=ls_ryhmä.Column, Criteria1:=Ryhmänimi, Operator:=xlFilterValues
            End If
            ' Yhdistettyjä ryhmiä löytyy
        Else
            ' Lisätään Ryhmänimi yhdistelyyn
            VL_Yhdistely = Ryhmänimi + ", " + VL_Yhdistely
            ' Tehdään Array ruokalistan yhdistelyryhmistä
            Dim vl_yhdistely_array() As String
            vl_yhdistely_array = Split(VL_Yhdistely, ",", , vbTextCompare)
            pituus = arrayitems(vl_yhdistely_array)
            .Range.AutoFilter Field:=ls_ryhmä.Column, Criteria1:=vl_yhdistely_array, Operator:=xlFilterValues
        End If

    End With

    ' Lapsiryhmän sorttaus
    Dim sort_lapset_abc As Range
    If (Aakkosjärjestys = "Oma järjestys") Then
        Set sort_lapset_abc = Range("A1")
    ElseIf (Aakkosjärjestys = "Kutsumanimi") Then
        Set sort_lapset_abc = Range("B1")
    ElseIf (Aakkosjärjestys = "Sukunimi") Then
        Set sort_lapset_abc = Range("tbl_lapset[Sukunimi]")
    End If

    With tbl_lapset.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sort_lapset_abc, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    ' Aloitetaan numerointi ykkösestä
    Dim Lapsinumero As Long: Lapsinumero = 1
    Dim Tj_numero As Long: Tj_numero = 1

    ' Luupataan lapset yksi kerrallaan läpi (näkyvät solut suodatuksen ja filtterin jälkeen)
    For Each laps In rng_lapset.SpecialCells(xlCellTypeVisible)

        'Aliohjelma Lapsi
        Call Lapsi(sht_lapset.Cells(laps.Row, ls_kutsumanimi.Column).Value, _
                   Lapsinumero, Ryhmäsivu, sht_ryhmät.Cells(ryhrow, _
                                                            rs_boldaus1.Column).Value, _
                   sht_ryhmät.Cells(ryhrow, rs_boldaus2.Column).Value, _
                   sht_ryhmät.Cells(ryhrow, rs_pboldaus1.Column).Value, _
                   sht_ryhmät.Cells(ryhrow, rs_pboldaus2.Column).Value, _
                   sht_lapset.Cells(laps.Row, ls_matulo.Column), VL_Vkl)

        ' Numerointiin lisätään aina yksi (Code-välilehteä varten, jotta ryhmät voidaan myöhemmin poistaa)
        Lapsinumero = Lapsinumero + 1
    Next laps

End Sub

Sub Ryhmä_Päivälaput(Ryhmänimi As String, Aakkosjärjestys As String, Ryhmänumero As Long, VL_Yhdistely As String, _
                     VL_Nimi As String, ByRef Code_ryhmät() As Variant, VL_päivystys As String, VL_päivälaput As String, _
                     ryhrow As Long, PL_Yhd As Long, PL_Pohja As String, VL_Yhd As Long, PL_Poissa As Long, _
                     pl_tyhjät As Integer)

    ' TODO
    ' - poissaolevat
    ' + pienryhmät
    
    
    Dim Ryhmäsivu As Worksheet
    Dim Nimi As Worksheet

    ' Napataan ryhmien nimet talteen, jotta voidaan poistaa ne ensi kerralla
    ' Jos Code_Ryhmät Array on tyhjä, muuta kooksi 0 ja lisää ryhmän nimi
    ' Muutoin lisää yksi merkintä + ryhmän nimi
    Dim arrayIsNothing As Boolean
    On Error Resume Next
    arrayIsNothing = IsNumeric(UBound(Code_ryhmät)) And False
    If Err.Number <> 0 Then arrayIsNothing = True
    On Error GoTo 0

    ' päivämäärien lisääminen
    Dim pvm_vuosi As Long
    Dim pvm_kk As Long
    Dim pvm_pv As Long
    Dim pvm As Date

    ' Haetaan pvm Codesta ja muunnetaan sopivaan muotoon
    pvm_vuosi = Year(Now)
    pvm_pv = sht_code.[C2].Value2
    pvm_kk = sht_code.[C3].Value2

    ' Päivämäärän koonti
    pvm = DateSerial(pvm_vuosi, pvm_kk, pvm_pv)

    Dim i As Long

    Dim pituus As Long
    Dim listaus As Long: listaus = 1

    With tbl_lapset
        ' Jos kyseessä päivystysryhmä, pidetään vain päivystyslapset
        If VL_päivystys = "Kyllä" Then
            .Range.AutoFilter Field:=ls_päivystys.Column, Criteria1:=Ryhmänimi, Operator:=xlFilterValues
        End If
        
      
        ' Ei työntekijöitä listalle
        .Range.AutoFilter Field:=ls_työntekijä.Column, Criteria1:="=", Operator:=xlFilterValues

        ' Yhdistellyt ryhmät
        If Not VL_Yhdistely = vbNullString And PL_Yhd > 0 Then
            ' Lisätään Ryhmänimi yhdistelyyn
            VL_Yhdistely = Ryhmänimi + ", " + VL_Yhdistely
            ' Tehdään Array ruokalistan yhdistelyryhmistä
            Dim vl_yhdistely_array() As String
            vl_yhdistely_array = Split(VL_Yhdistely, ",", , vbTextCompare)
    
            ' Välilyönnit pois arraysta
            Dim väli As Long
            For väli = LBound(vl_yhdistely_array) To UBound(vl_yhdistely_array)
                vl_yhdistely_array(väli) = Trim(vl_yhdistely_array(väli))
            Next
    
            pituus = arrayitems(vl_yhdistely_array)
            
            ' Sama lista
            If PL_Yhd = 1 Then .Range.AutoFilter Field:=ls_ryhmä.Column, Criteria1:=vl_yhdistely_array, Operator:=xlFilterValues
            ' 2-puoliset
            If PL_Yhd = 2 Then .Range.AutoFilter Field:=ls_ryhmä.Column, Criteria1:=vl_yhdistely_array(0), Operator:=xlFilterValues
            
            ' Vain 1 ryhmä
        Else
            pituus = 1
            ' Jos ei yhdistetä mitään, suodatetaan ryhmän nimen mukaan
            If VL_päivystys = "Ei" Then
                .Range.AutoFilter Field:=ls_ryhmä.Column, Criteria1:=Ryhmänimi, Operator:=xlFilterValues
            End If
            
            ' Yhdistettyjä ryhmiä löytyy
        End If

    End With

    ' Lapsiryhmän sorttaus
    Dim sort_lapset_abc As Range
    If (Aakkosjärjestys = "Oma järjestys") Then
        Set sort_lapset_abc = Range("A1")
    ElseIf (Aakkosjärjestys = "Kutsumanimi") Then
        Set sort_lapset_abc = Range("B1")
    ElseIf (Aakkosjärjestys = "Sukunimi") Then
        Set sort_lapset_abc = Range("tbl_lapset[Sukunimi]")
    End If

    With tbl_lapset
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=sort_lapset_abc, SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With
    End With


    Dim oikeanimi As String
    ' Jos ei yhdistettyjä ryhmiä, käytetään ryhmän nimeä
    ' Jos Yhdistettyjä ryhmiä, käytä niiden omaa nimeä
    If PL_Yhd = 0 Then
        oikeanimi = Ryhmänimi
        ' Lisätään arrayhyn (codeen)
        If arrayIsNothing Then
            ReDim Code_ryhmät(0)
            Code_ryhmät(0) = Ryhmänimi
        Else
            ' Jos viimeisin array on ryhmän nimi -> ei laiteta listalle (ei duplikaatteja)
            ' TODO: tsekkaus, että toimii myös yhdistetyissä ryhmissä
            Dim viimeisinRyhmä As Integer
            viimeisinRyhmä = ArrayLen(Code_ryhmät) - 1
            If Code_ryhmät(viimeisinRyhmä) = oikeanimi Then
            Else
                ReDim Preserve Code_ryhmät(UBound(Code_ryhmät) + 1)
                Code_ryhmät(UBound(Code_ryhmät)) = oikeanimi
            End If
        End If
    Else
        oikeanimi = VL_Nimi
        ' Lisätään arrayhyn (codeen)
        If arrayIsNothing Then
            If PL_Yhd = 1 Then
                ReDim Code_ryhmät(0)
                Code_ryhmät(0) = oikeanimi
            End If
            If PL_Yhd = 2 Then
                ReDim Code_ryhmät(1)
                Code_ryhmät(0) = vl_yhdistely_array(0)
                Code_ryhmät(1) = vl_yhdistely_array(1)
            End If
        Else
            If PL_Yhd = 1 Then
                ReDim Preserve Code_ryhmät(UBound(Code_ryhmät) + 1)
                Code_ryhmät(UBound(Code_ryhmät)) = oikeanimi
            End If
            If PL_Yhd = 2 Then
                ReDim Preserve Code_ryhmät(UBound(Code_ryhmät) + 1)
                Code_ryhmät(UBound(Code_ryhmät)) = vl_yhdistely_array(0)
                ReDim Preserve Code_ryhmät(UBound(Code_ryhmät) + 1)
                Code_ryhmät(UBound(Code_ryhmät)) = vl_yhdistely_array(1)
            End If
        End If
    End If
        
    ' Aloitetaan numerointi ykkösestä
    Dim Lapsinumero As Long: Lapsinumero = 1
    Dim Tj_numero As Long: Tj_numero = 1
    Dim PL_Pohja2 As String, pl_tyhjät2 As Integer
    
    Dim lappupv(1 To 5) As String
    lappupv(1) = "pe"
    lappupv(2) = "to"
    lappupv(3) = "ke"
    lappupv(4) = "ti"
    lappupv(5) = "ma"
    
    Dim vko As Variant

    
    
    Lapsinumero = 0

    ' Jos 2-puoleinen -- Haetaan asetukset
    If PL_Yhd = 2 Then
        PL_Pohja2 = sht_ryhmät.Cells(rs_ryhmännimi.Find(vl_yhdistely_array(1)).Row, rs_plpohja.Column).Value2
        pl_tyhjät2 = sht_ryhmät.Cells(rs_ryhmännimi.Find(vl_yhdistely_array(1)).Row, rs_pltyhjät.Column).Value2
    End If
    
    For Each vko In lappupv
        Call Päivälappu(PL_Pohja, PL_Yhd, vko, oikeanimi, Ryhmänimi, pvm, pl_tyhjät, PL_Poissa)
        
        ' 2-puoleiset
        If PL_Yhd = 2 Then
            ' Haetaan 2-puolen lapset
            tbl_lapset.Range.AutoFilter Field:=ls_ryhmä.Column, Criteria1:=vl_yhdistely_array(1), Operator:=xlFilterValues
            Call Päivälappu(PL_Pohja2, PL_Yhd, vko, "", vl_yhdistely_array(1), pvm, pl_tyhjät2, PL_Poissa)
            ' Haetaan taas 1-puolen lapset
            If VL_päivystys = "Ei" Then
                tbl_lapset.Range.AutoFilter Field:=ls_ryhmä.Column, Criteria1:=Ryhmänimi, Operator:=xlFilterValues
            Else
                tbl_lapset.Range.AutoFilter Field:=ls_päivystys.Column, Criteria1:=Ryhmänimi, Operator:=xlFilterValues
            End If
            
        End If
    
    Next vko
    
End Sub

Sub Päivälappupohja_kopsuri()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Näytetään välilehdet
    Dim WSname As Variant
    For Each WSname In Array("pl_pysty", "pl_pystywc", "pl_vaaka", "pl_vaakawc")
        Worksheets(WSname).Visible = True
    Next WSname

    Dim PL_Pohja As String, Kohde As String

    PL_Pohja = ThisWorkbook.Worksheets("Päiväkoti").Range("D8").Value2
    Kohde = ThisWorkbook.Worksheets("Päiväkoti").Range("E8").Value2
    
    If Kohde = "" Then
        MsgBox "Et ole kirjoittanut Ryhmän nimeksi mitään."
        'Sheets(Array("pl_pysty", "pl_pystywc", "pl_vaaka", "pl_vaakawc")).Select
        Call Lopetus_Simple
    End If
    
    ' Onko välilehti jo olemassa
    Dim i As Integer
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "pl_" & Kohde Then
            MsgBox "pl_" & Kohde & " -niminen välilehti löytyy jo."
            Call Lopetus_Simple
        End If
    Next i
    
    Select Case PL_Pohja
    Case "Pysty"
        Sheets("pl_pysty").Copy After:=Sheets("pl_pysty")
    Case "Pysty + wc"
        Sheets("pl_pystywc").Copy After:=Sheets("pl_pystywc")
    Case "Vaaka"
        Sheets("pl_vaaka").Copy After:=Sheets("pl_vaaka")
    Case "Vaaka + wc"
        Sheets("pl_vaakawc").Copy After:=Sheets("pl_vaakawc")
    Case Else
        MsgBox "Et ole valinnut oikeaa pohjaa."
        Call Lopetus_Simple
    End Select
   
    ActiveSheet.Name = "pl_" & Kohde
    
    ' Piilotetaan välilehdet
    Sheets(Array("pl_pysty", "pl_pystywc", "pl_vaaka", "pl_vaakawc")).Select
    ActiveWindow.SelectedSheets.Visible = False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Uusi pohja lisätty." & vbCrLf & "Välilehden nimi on pl_" & Kohde


End Sub

Sub Päivälappu(PL_Pohja As String, PL_Yhd As Long, vko As Variant, oikeanimi As String, Ryhmänimi As String, pvm As Date, pl_tyhjät As Integer, PL_Poissa As Long)
    
    ' - poissa -asetusta varten suodatusten poisto
    If PL_Poissa = 1 Then
        With tbl_lapset
            .Range.AutoFilter Field:=ls_matulo.Column
            .Range.AutoFilter Field:=ls_titulo.Column
            .Range.AutoFilter Field:=ls_ketulo.Column
            .Range.AutoFilter Field:=ls_totulo.Column
            .Range.AutoFilter Field:=ls_petulo.Column
        End With
    End If
    ' Kloonataan pohja
    Select Case PL_Pohja
    Case "Pysty"
        Sheets("pl_pysty").Copy After:=Sheets("pl_pysty")
    Case "Pysty + wc"
        Sheets("pl_pystywc").Copy After:=Sheets("pl_pystywc")
    Case "Vaaka"
        Sheets("pl_vaaka").Copy After:=Sheets("pl_vaaka")
    Case "Vaaka + wc"
        Sheets("pl_vaakawc").Copy After:=Sheets("pl_vaakawc")
    Case "Kustomoitu"
        Sheets("pl_" & Ryhmänimi).Copy After:=Sheets("pl_pysty")
    End Select
    
    Dim Ryhmäsivu As Worksheet
    If PL_Yhd = 1 Then
        ActiveSheet.Name = oikeanimi & "_" & vko
        Set Ryhmäsivu = wb.Worksheets(oikeanimi & "_" & vko)
    Else
        ActiveSheet.Name = Ryhmänimi & "_" & vko
        Set Ryhmäsivu = wb.Worksheets(Ryhmänimi & "_" & vko)
    End If
    
    Dim RyhmäPaikka As Range, PvmPaikka As Range, NimiPaikka As Range, VikaSarakePaikka As Range, LoppuRivit As Range
    
    On Error Resume Next
        
    ' Napataan koodit pohjasta
    Dim pl_Ryhmänimi As Range, pl_Pvm As Range, pl_Nimi As Range, pl_Hoitoaika As Range, pl_Alateksti As Range, pl_Vikasarake As Range, pl_Vikarivi As Range
    With Ryhmäsivu.Range("A1:O200")
        Set pl_Ryhmänimi = .Find("pl-ryhmänimi", MatchCase:=False)
        Set pl_Pvm = .Find("pl-pvm", MatchCase:=False)
        Set pl_Nimi = .Find("pl-nimi", MatchCase:=False)
        Set pl_Hoitoaika = .Find("pl-hoitoaika", MatchCase:=False)
        Set pl_Alateksti = .Find("pl-alateksti", MatchCase:=False)
        Set pl_Vikasarake = .Find("pl-vikasarake", MatchCase:=False)
        Set pl_Vikarivi = .Find("pl-vikarivi", MatchCase:=False)
    End With
    On Error GoTo 0
    ' Ryhmän nimen lisäys
    If PL_Yhd = 1 Then
        pl_Ryhmänimi = UCase(oikeanimi)
    Else
        pl_Ryhmänimi = UCase(Ryhmänimi)
    End If
    ' Päivämäärän lisäys
    If vko = "ma" Then pl_Pvm = Format(pvm, "ddd d.m.")
    If vko = "ti" Then pl_Pvm = Format(DateAdd("d", 1, CDate(pvm)), "ddd d.m.")
    If vko = "ke" Then pl_Pvm = Format(DateAdd("d", 2, CDate(pvm)), "ddd d.m.")
    If vko = "to" Then pl_Pvm = Format(DateAdd("d", 3, CDate(pvm)), "ddd d.m.")
    If vko = "pe" Then pl_Pvm = Format(DateAdd("d", 4, CDate(pvm)), "ddd d.m.")
    
    ' Käydään ryhmän lapset läpi ja lisätään nimet ja hoitoajat
    Dim Lapsinumero As Long
    Lapsinumero = 0
    Dim päiväapuri As Long, päiväapuri2 As Long
    
    Rivi = 0
    ' Ryhmäsivu.Cells(Lapsinumero + 2, 1)
    pienennys = 0
    
   
    ' Jos poissaolevia:
    If PL_Poissa = 1 Then
        With tbl_lapset
            If vko = "ma" Then .Range.AutoFilter Field:=ls_matulo.Column, Criteria1:=">0", Operator:=xlFilterValues
            If vko = "ti" Then .Range.AutoFilter Field:=ls_titulo.Column, Criteria1:=">0", Operator:=xlFilterValues
            If vko = "ke" Then .Range.AutoFilter Field:=ls_ketulo.Column, Criteria1:=">0", Operator:=xlFilterValues
            If vko = "to" Then .Range.AutoFilter Field:=ls_totulo.Column, Criteria1:=">0", Operator:=xlFilterValues
            If vko = "pe" Then .Range.AutoFilter Field:=ls_petulo.Column, Criteria1:=">0", Operator:=xlFilterValues
        End With
    End If
    With Ryhmäsivu
        Dim lapsimäärä As Integer: lapsimäärä = 0
        If tyhjasuodatin(rng_lapset) = False Then
            For Each laps In rng_lapset.SpecialCells(xlCellTypeVisible)
                tbl_lapset.Range.AutoFilter Field:=ls_petulo.Column
        
                ' Jos lapsirivit meinaavat mennä alatekstin päälle, lisätään rivi ja otetaan ylös sen koko.
                ' Myöhemmin pienennetään nimilistaa tämän koon perusteella
        
                If pl_Nimi.Row + Rivi = pl_Alateksti.Row Then
                    .Range("A" & pl_Nimi.Row + Rivi).EntireRow.Insert
                    pienennys = pienennys + .Range("A" & pl_Nimi.Row + Rivi).RowHeight
                End If

                ' Kopioidaan tyyli seuraavalle riville (ei ensimmäistä riviä)
                If Not Rivi <= 1 Then
                    .Range(Cells(pl_Nimi.Offset(Rivi, 0).Row, 1), Cells(pl_Nimi.Offset(Rivi, 0).Row, pl_Vikasarake.Column)).Offset(-2, 0).Copy
                    If WorksheetFunction.IsEven(Rivi) Then
                        ' Pariliset rivit
                        .Range(Cells(pl_Nimi.Offset(Rivi, 0).Row, 1), Cells(pl_Nimi.Offset(Rivi, 0).Row, pl_Vikasarake.Column)).PasteSpecial Paste:=xlPasteFormats
                    Else
                        ' Parittomat rivit
                        .Range(Cells(pl_Nimi.Offset(Rivi, 0).Row, 1), Cells(pl_Nimi.Offset(Rivi, 0).Row, pl_Vikasarake.Column)).PasteSpecial Paste:=xlPasteFormats
                    End If
                End If
        
                ' Lapsen nimen lisäys
                'If Dieetti Then
                '    pl_Nimi.Offset(Rivi) = UCase(sht_lapset.Cells(laps.Row, ls_kutsumanimi.Column).Value) + " " + ChrW(&HD83C) & ChrW(&HDF74)
                'Else
                pl_Nimi.Offset(Rivi) = UCase(sht_lapset.Cells(laps.Row, ls_kutsumanimi.Column).Value)
        
                ' Hoitoajan lisäys
                If vko = "ma" Then
                    päiväapuri = ls_matulo.Column
                    päiväapuri2 = ls_malähtö.Column
                ElseIf vko = "ti" Then
                    päiväapuri = ls_titulo.Column
                    päiväapuri2 = ls_tilähtö.Column
                ElseIf vko = "ke" Then
                    päiväapuri = ls_ketulo.Column
                    päiväapuri2 = ls_kelähtö.Column
                ElseIf vko = "to" Then
                    päiväapuri = ls_totulo.Column
                    päiväapuri2 = ls_tolähtö.Column
                ElseIf vko = "pe" Then
                    päiväapuri = ls_petulo.Column
                    päiväapuri2 = ls_pelähtö.Column
                End If

                ' Call pl_Lapsi(Lapsinumero, Ryhmäsivu, sht_lapset.Cells(laps.Row, ls_matulo.Column))
                Call pl_Lapsi(Lapsinumero, pl_Hoitoaika, sht_lapset.Cells(laps.Row, päiväapuri), sht_lapset.Cells(laps.Row, päiväapuri2))
                ' Numerointiin lisätään aina yksi (Code-välilehteä varten, jotta ryhmät voidaan myöhemmin poistaa)
                lapsimäärä = lapsimäärä + 1
            Next laps
    
        
            ' 1. sivun viimeinen rivi
            ' Ryhmäsivu.HPageBreaks.Item(1).Location.Row - 1
    
            ' 2. sivun ensimmäinen rivi
            ' Ryhmäsivu.HPageBreaks.Item(1).Location.Row
    
            ' 2. sivun viimeinen rivi:
            ' Ryhmäsivu.HPageBreaks.Item(2).Location.Row - 1
        
            ' TYHJIEN LISÄYS
            ' Lisätään tyhjiä, kunnes kaikki tyhjät lisätty TAI alateksti tulee vastaan
            Dim T_Lisäys As Integer
            T_Lisäys = 0
            If lapsimäärä <> 1 Then
                Do Until pl_Nimi.Offset(Rivi).Row = pl_Alateksti.Row Or T_Lisäys = pl_tyhjät
                    If Not Rivi <= 1 Then
                        .Range(Cells(pl_Nimi.Offset(Rivi, 0).Row, 1), Cells(pl_Nimi.Offset(Rivi, 0).Row, pl_Vikasarake.Column)).Offset(-2, 0).Copy
                        If WorksheetFunction.IsEven(pl_Nimi.Row + Rivi) Then
                            ' Pariliset rivit
                            .Range(Cells(pl_Nimi.Offset(Rivi, 0).Row, 1), Cells(pl_Nimi.Offset(Rivi, 0).Row, pl_Vikasarake.Column)).PasteSpecial Paste:=xlPasteFormats
                        Else
                            ' Parittomat rivit
                            .Range(Cells(pl_Nimi.Offset(Rivi, 0).Row, 1), Cells(pl_Nimi.Offset(Rivi, 0).Row, pl_Vikasarake.Column)).PasteSpecial Paste:=xlPasteFormats
                        End If
                        Rivi = Rivi + 1
                        T_Lisäys = T_Lisäys + 1
                    End If
                Loop
        
                Dim Nimirivit As Range
                ' Jos lapsirivejä on liikaa, pienennetään niitä
                If pienennys > 0 Then
                    For Each Nimirivit In Ryhmäsivu.Range(Cells(pl_Nimi.Row, 1), Cells(pl_Nimi.Offset(Rivi).Row - 1, 1))
                        Nimirivit.RowHeight = Application.WorksheetFunction.RoundDown(Nimirivit.RowHeight - (pienennys / Rivi), 1)
                    Next Nimirivit
                    ' Muussa tapauksessa
                End If
     
                ' Poistetaan rivien ja alatekstin välinen tila JOS ei alateksti ole ihan kiinni riveissä
                If pl_Nimi.Offset(Rivi).Row = pl_Alateksti.Row Then
                Else
                    .Range(Cells(pl_Nimi.Offset(Rivi).Row, 1), Cells(pl_Alateksti.Row - 1, 1)).Rows.EntireRow.Delete
            
                    ' Silloin on myös tilaa suurentaa rivejä
                    ' RIVIEN KOON SUURENTAMINEN
                    ' Lasketaan eri osien korkeudet
                    Dim rivikorkeus As Long, LoppuOsa As Range, D_Korkeus As Long
                    D_Korkeus = 0
                    For Each LoppuOsa In Ryhmäsivu.Range(Cells(pl_Vikarivi.Row + 1, 1), Cells(Ryhmäsivu.HPageBreaks.Item(1).Location.Row - 1, 1))
                        rivikorkeus = rivikorkeus + LoppuOsa.Rows.Height
                    Next LoppuOsa
                
                    ' Jaetaan loppuosan tyhjän tilan korkeus jokaisen lapsirivin kesken (myös tyhjät rivit) ja suurennetaan rivien kokoa
                    For Each Nimirivit In Ryhmäsivu.Range(Cells(pl_Nimi.Row, 1), Cells(pl_Nimi.Offset(Rivi).Row - 1, 1))
                        Nimirivit.RowHeight = Nimirivit.RowHeight + (rivikorkeus / Rivi)
                    Next Nimirivit
            
                

                
                End If
                
                On Error Resume Next
                ' TODO
                ' Set Printarea
                ' Entä 2-puolisena?
                .PageSetup.PrintArea = Ryhmäsivu.Range(Cells(1, 1), Cells(Ryhmäsivu.HPageBreaks.Item(1).Location.Row - 1, pl_Vikasarake.Column)).Address
                With .Range("A1:O200")
                    Set pl_Hoitoaika = .Find("pl-hoitoaika", MatchCase:=False)
                    Set pl_Alateksti = .Find("pl-alateksti", MatchCase:=False)
                    Set pl_Vikasarake = .Find("pl-vikasarake", MatchCase:=False)
                    Set pl_Vikarivi = .Find("pl-vikarivi", MatchCase:=False)
                    pl_Hoitoaika.Value2 = ""
                    pl_Alateksti.Value2 = ""
                    pl_Vikasarake.Value2 = ""
                    pl_Vikarivi.Value2 = ""
                End With
                On Error GoTo 0
End If

            ' Jos ei lapsia listalla, poista lappu
        Else
            .Delete
        End If
    End With


End Sub

Sub Lapsi(Kutsumanimi As String, Lapsinumero As Long, Ryhmäsivu As Worksheet, tulobold As String, _
          menobold As String, puokkariyks As String, puokkarikaks As String, maTulo As Range, Optional VL_Vkl)

    Dim viiva As String: viiva = "-"
    Dim tuloArr() As String
    Dim lähtöArr() As String
    Dim kerta As Long
    Dim arrRivi As Long
    Dim maxArr As Long

    If Lapsinumero = 1 Then Rivi = 2

    With Ryhmäsivu
        ' Lapsen numero ja nimi
        .Cells(Rivi, 1) = Lapsinumero
        .Cells(Rivi, 3) = UCase(Kutsumanimi)
        LapsetSarake = 0
        ListaLisäys = 0
        ' ma-pe
        Päivät = 8
        ' myös viikonloput
        If VL_Vkl = 1 Then Päivät = 12
    
        For LapsetSarake = 0 To Päivät Step 2
    
            ' Yksittäinen hoitoaika
            If IsNumeric(maTulo.Offset(, LapsetSarake).Value2) = True Then
                If Not IsEmpty(maTulo.Offset(, LapsetSarake).Value2) Then
                    ' Muodostetaan hoitoaika
                    ' Puoliyö tulo (muutetaan format stringiksi)
                    If maTulo.Offset(, LapsetSarake).Value2 = 0 Then
                        .Cells(Rivi, 4 + ListaLisäys).NumberFormat = "@"
                        .Cells(Rivi, 4 + ListaLisäys) = "0:00"
                    Else
                        .Cells(Rivi, 4 + ListaLisäys) = maTulo.Offset(, LapsetSarake).Value2
                    End If
                    .Cells(Rivi, 5 + ListaLisäys) = viiva
                    .Cells(Rivi, 6 + ListaLisäys) = maTulo.Offset(, LapsetSarake + 1).Value2
                    ' Puokkariboldaukset
                    If maTulo.Offset(, LapsetSarake).Value2 >= puokkariyks And maTulo.Offset(, LapsetSarake).Value2 <= puokkarikaks Then .Cells(Rivi, 4 + ListaLisäys).Style = "Boldaus3"
                    If maTulo.Offset(, LapsetSarake + 1).Value2 >= puokkariyks And maTulo.Offset(, LapsetSarake + 1).Value2 <= puokkarikaks Then .Cells(Rivi, 6 + ListaLisäys).Style = "Boldaus3"
                    ' Boldaukset
                    If maTulo.Offset(, LapsetSarake).Value2 <= tulobold Or maTulo.Offset(, LapsetSarake).Value2 >= menobold Then .Cells(Rivi, 4 + ListaLisäys).Style = "Boldaus2"
                    If maTulo.Offset(, LapsetSarake + 1).Value2 <= tulobold Or maTulo.Offset(, LapsetSarake + 1).Value2 >= menobold Then .Cells(Rivi, 6 + ListaLisäys).Style = "Boldaus2"
                    ' Puoliyö lähtö (poistetaan hoitoaika, jos 23:55)
                    If CDate(maTulo.Offset(, LapsetSarake + 1).Value2) = "23:55" Then .Cells(Rivi, 6 + ListaLisäys) = "23:59"

                 
                End If
                ' Useampia hoitoaikoja
            ElseIf InStr(maTulo.Offset(, LapsetSarake).Value2, ",") > 0 Then
                ' Tehdään hoitoajoista arrayt
                tuloArr = Split(maTulo.Offset(, LapsetSarake).Value2, ",")
                lähtöArr = Split(maTulo.Offset(, LapsetSarake + 1).Value2, ",")
                
                ' Hoitoaikojen määrä arrayssa
                ' Käydään array läpi
                For kerta = 0 To ArrayLen(tuloArr) - 1
                    ' Eka kerta
                    If kerta = 0 Then
                        arrRivi = Rivi
                    Else
                        arrRivi = arrRivi + 1
                    End If
                    ' Puoliyö tulo (muutetaan format stringiksi)
                    If CDate(tuloArr(kerta)) = "00:00" Then
                        .Cells(Rivi, 4 + ListaLisäys).NumberFormat = "@"
                        .Cells(Rivi, 4 + ListaLisäys) = "0:00"
                    Else
                        .Cells(arrRivi, 4 + ListaLisäys) = tuloArr(kerta)
                    End If
                
                                
                    .Cells(arrRivi, 5 + ListaLisäys) = viiva
                    .Cells(arrRivi, 6 + ListaLisäys) = lähtöArr(kerta)
                
                    ' Puokkariboldaukset
                    If tuloArr(kerta) >= CDate(puokkariyks) And tuloArr(kerta) <= CDate(puokkarikaks) Then .Cells(arrRivi, 4 + ListaLisäys).Style = "Boldaus3"
                    If lähtöArr(kerta) >= CDate(puokkariyks) And lähtöArr(kerta) <= CDate(puokkarikaks) Then .Cells(arrRivi, 6 + ListaLisäys).Style = "Boldaus3"
                    ' Boldaukset
                    If tuloArr(kerta) <= CDate(tulobold) Or tuloArr(kerta) >= CDate(menobold) Then .Cells(arrRivi, 4 + ListaLisäys).Style = "Boldaus2"
                    If lähtöArr(kerta) <= CDate(tulobold) Or lähtöArr(kerta) >= CDate(menobold) Then .Cells(arrRivi, 6 + ListaLisäys).Style = "Boldaus2"
                
                    ' Puoliyö lähtö (poistetaan hoitoaika, jos 23:55)
                    If CDate(lähtöArr(kerta)) = "23:55" Then .Cells(arrRivi, 6 + ListaLisäys) = "23:59"
                
                Next kerta
                ' Max arrayn pituus
                If ArrayLen(tuloArr) > maxArr Then maxArr = ArrayLen(tuloArr)
                ' Yksi hoitoaika
            Else
                .Cells(Rivi, 5 + ListaLisäys) = maTulo.Offset(, LapsetSarake).Value2
            End If
            ListaLisäys = ListaLisäys + 3
        Next LapsetSarake

        ' Yläviiva
        If Lapsinumero <> 1 Then
            If VL_Vkl = 1 Then
                .Range(.Cells(Rivi, 1), .Cells(Rivi, 24)).Style = "Yläviiva"
            Else
                .Range(.Cells(Rivi, 1), .Cells(Rivi, 18)).Style = "Yläviiva"
            End If
        End If
    
        ' Harmaatausta
        .Range(.Cells(1, 4), .Cells(Rivi, 6)).Style = "Harmaatausta"
        .Range(.Cells(1, 10), .Cells(Rivi, 12)).Style = "Harmaatausta"
        .Range(.Cells(1, 16), .Cells(Rivi, 18)).Style = "Harmaatausta"
        If VL_Vkl = 1 Then
            .Range(.Cells(1, 22), .Cells(Rivi, 24)).Style = "Harmaatausta"
            .PageSetup.PrintArea = Ryhmäsivu.Range(Cells(1, 1), Cells(Rivi, 24)).Address
        End If
        .Range("C1").EntireColumn.AutoFit
    End With

    ' Lisätään rivejä (jos on ollut useampia hoitoaikoja, lisätään maksimiarrayn pituuden verran rivejä)
    If maxArr > 0 Then
        Rivi = Rivi + maxArr
    Else
        Rivi = Rivi + 1
    End If

End Sub

Sub pl_Lapsi(Lapsinumero As Long, Ryhmäsivu As Range, maTulo As Range, maLähtö As Range)

    Dim viiva As String: viiva = "-"
    Dim tuloArr() As String
    Dim lähtöArr() As String
    Dim yhdArr() As String
    Dim kerta As Long
    Dim AlkuKorkeus As Long: AlkuKorkeus = 0

    With Ryhmäsivu
        '        Debug.Print (ls_dieetti.Column)
        '       Debug.Print (maTulo.Row)
        ' Lisätään kuvake, jos dieettilapsi
        ' TODO: jotain häikkää, ei toimi.
'        If Trim(sht_lapset.Cells(maTulo.Row, ls_dieetti.Column).Offset(Rivi).Value2) <> vbNullString Then
        If Trim(sht_lapset.Cells(maTulo.Row, ls_dieetti.Column).Value2) <> vbNullString Then
            .Offset(Rivi, -1) = .Offset(Rivi, -1) + " " + ChrW(&HD83C) & ChrW(&HDF74)
        End If
    
        ' Yksittäinen hoitoaika
        If IsNumeric(maTulo.Value2) = True Then
            If Not IsEmpty(maTulo.Value2) Then
                ' Muodostetaan hoitoaika
                ' Puoliyö tulo (muutetaan format stringiksi)
                If maTulo.Value2 = 0 Then
                    .Offset(Rivi).NumberFormat = "@"
                    .Offset(Rivi) = "0:00"
                End If
                
                If CDate(maLähtö.Value2) > "23:54" Then
                    .Offset(Rivi) = Format(maTulo.Value2, "h:mm") & " " & ChrW(8594)
                ElseIf CDate(maTulo.Value2) = "0:00" Then
                    .Offset(Rivi) = " " & ChrW(8594) & " " & Format(maLähtö.Value2, "h:mm")
                Else
                    .Offset(Rivi) = Format(maTulo.Value2, "h:mm") & " - " & Format(maLähtö.Value2, "h:mm")
                End If
                
            End If
            ' Useampia hoitoaikoja
        ElseIf InStr(maTulo.Value2, ",") > 0 Then
            AlkuKorkeus = .Offset(1).RowHeight
            ' Tehdään hoitoajoista arraytnj
            tuloArr = Split(maTulo.Value2, ",")
            lähtöArr = Split(maLähtö.Value2, ",")
            
            ' Hoitoaikojen määrä arrayssa
            ' Käydään array läpi
            ReDim yhdArr(1)
            For kerta = 0 To ArrayLen(tuloArr) - 1
               
                ' TODO
                ' Puoliyö tulo (muutetaan format stringiksi)
                'If CDate(tuloArr(kerta)) = "00:00" Then
                '    .Offset(Rivi).NumberFormat = "@"
                '     .Offset(Rivi) = "0:00"
                ' End If
               
                ' Array hoitoajoista esim. '11:00 - 14:00'
                'ReDim Preserve yhdArr(kerta)
                
                '                ReDim Preserve Code_ryhmät(UBound(Code_ryhmät) + 1)
                '                Code_ryhmät(UBound(Code_ryhmät)) = Ryhmänimi
                ' Yöhoito
                If CDate(lähtöArr(kerta)) > "23:54" Then
                    yhdArr(kerta) = tuloArr(kerta) & " " & ChrW(8594)
                ElseIf CDate(tuloArr(kerta)) = "0:00" Then
                    yhdArr(kerta) = " " & ChrW(8594) & " " & lähtöArr(kerta)
                Else
                    ' Muutoin vain hoitoaika
                    yhdArr(kerta) = tuloArr(kerta) & " - " & lähtöArr(kerta)
                End If
            Next kerta
           
            
            ' Lisätään hoitoajat soluun. Hoitoaikojen välille rivinvaihto, jotta menevät samalle solulle
            .Offset(Rivi) = Join(yhdArr, Chr(10))
            ' Koon säätö
            
            pienennys = pienennys + (.Offset(1).RowHeight - AlkuKorkeus)
            .Offset(Rivi).EntireRow.AutoFit
            
        Else
            .Offset(Rivi) = maTulo.Value2
        End If
    
    End With

    Rivi = Rivi + 1

End Sub

Sub Dieetti(kokonimi As String, Ruokaryhmä As Worksheet, Ryhmänimi As String, Päivät As Long, Ruokailut As Long, Rivi As Long, Sarake As Long)

    ' Ryhmän ruokailuajat
    Dim a1 As String: a1 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_aamupala1.Column).Value2
    Dim a2 As String: a2 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_aamupala2.Column).Value2
    Dim l1 As String: l1 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_lounas1.Column).Value2
    Dim l2 As String: l2 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_lounas2.Column).Value2
    Dim v1 As String: v1 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_välipala1.Column).Value2
    Dim v2 As String: v2 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_välipala2.Column).Value2
    Dim p1 As String: p1 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_päivällinen1.Column).Value2
    Dim p2 As String: p2 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_päivällinen1.Column).Value2
    Dim i1 As String: i1 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_iltapala1.Column).Value2
    Dim i2 As String: i2 = sht_ryhmät.Cells(rs_ryhmännimi.Find(Ryhmänimi).Row, rs_iltapala2.Column).Value2

    tbl_lapset.Range.AutoFilter Field:=ls_kokonimi.Column, Criteria1:=kokonimi, Operator:=xlFilterValues

    'Dim Listanumero As Long: Listanumero = 1
    Dim lnimi As Range
    Dim matulorng As Range
    Dim pv As Long: pv = 0
    Dim a As Long: a = 30
    Dim L As Long: L = 31
    Dim V As Long: V = 32

    Dim P As Long: P = 33
    Dim i As Long: i = 34

    Dim tuloArr() As String
    Dim lähtöArr() As String
    Dim yhdArr() As String
    Dim kerta As Long

    With Ruokaryhmä
        For Each lnimi In rng_lapset.SpecialCells(xlCellTypeVisible)
            ' Lopeta jos kyseessä työntekijä
            If sht_lapset.Cells(lnimi.Row, ls_työntekijä.Column).Value2 <> "" Then Exit Sub
            ' Käydään läpi kaikki päivät
            For pv = 0 To Päivät - 1
                Set matulorng = sht_lapset.Cells(lnimi.Row, ls_matulo.Column + 2 * pv)
                ' Yksittäinen hoitoaika
                If IsDate(Format(sht_lapset.Cells(lnimi.Row, ls_matulo.Column + 2 * pv), "h:mm")) = True Then
                    If Ruokailut = 2 Then
                        ' Päivällinen
                        If matulorng <= p2 And matulorng.Offset(, 1) >= p1 Then .Cells(Rivi + pv, Sarake).Value = "X"
                        ' Iltapala
                        If matulorng <= i2 And matulorng.Offset(, 1) >= i1 Then .Cells(Rivi + pv, Sarake + 1).Value = "X"
                    
                    ElseIf Ruokailut >= 3 Then
                        ' Aamupala
                        If matulorng <= a2 And matulorng.Offset(, 1) >= a1 Then .Cells(Rivi + pv, Sarake).Value = "X"
                        ' Lounas
                        If matulorng <= l2 And matulorng.Offset(, 1) >= l1 Then .Cells(Rivi + pv, Sarake + 1).Value = "X"
                        ' Välipala
                        If matulorng <= v2 And matulorng.Offset(, 1) >= v1 Then .Cells(Rivi + pv, Sarake + 2).Value = "X"
                    
                        If Ruokailut = 5 Then
                            ' Päivällinen
                            If matulorng <= p2 And matulorng.Offset(, 1) >= p1 Then .Cells(Rivi + pv, Sarake + 3).Value = "X"
                            ' Iltapala
                            If matulorng <= i2 And matulorng.Offset(, 1) >= i1 Then .Cells(Rivi + pv, Sarake + 4).Value = "X"
                        End If
                    End If

                    ' Useampi hoitoaika
                ElseIf InStr(matulorng, ",") > 0 Then
                    ' Tehdään hoitoajoista arrayt
                    tuloArr = Split(matulorng, ",")
                    lähtöArr = Split(matulorng.Offset(, 1), ",")
                    ' Käydään array läpi
                    For kerta = 0 To ArrayLen(tuloArr) - 1
                        If Ruokailut = 2 Then
                            ' Päivällinen
                            If tuloArr(kerta) <= CDate(p2) And lähtöArr(kerta) >= CDate(p1) Then .Cells(Rivi + pv, Sarake).Value = "X"
                            ' Iltapala
                            If tuloArr(kerta) <= CDate(i2) And lähtöArr(kerta) >= CDate(i1) Then .Cells(Rivi + pv, Sarake + 1).Value = "X"
                    
                        ElseIf Ruokailut >= 3 Then
                            ' Aamupala
                            If tuloArr(kerta) <= CDate(a2) And lähtöArr(kerta) >= CDate(a1) Then .Cells(Rivi + pv, Sarake).Value = "X"
                            ' Lounas
                            If tuloArr(kerta) <= CDate(l2) And lähtöArr(kerta) >= CDate(l1) Then .Cells(Rivi + pv, Sarake + 1).Value = "X"
                            ' Välipala
                            If tuloArr(kerta) <= CDate(v2) And lähtöArr(kerta) >= CDate(v1) Then .Cells(Rivi + pv, Sarake + 2).Value = "X"
                        
                            If Ruokailut = 5 Then
                                ' Päivällinen
                                If tuloArr(kerta) <= CDate(p2) And lähtöArr(kerta) >= CDate(p1) Then .Cells(Rivi + pv, Sarake + 3).Value = "X"
                                ' Iltapala
                                If tuloArr(kerta) <= CDate(i2) And lähtöArr(kerta) >= CDate(i1) Then .Cells(Rivi + pv, Sarake + 4).Value = "X"
                            End If
                        End If
                    Next kerta
                    ' Kirjain tai tyhjä
                Else
                    ' Jos koko viikko tyhjänä, merkkaa kysymysmerkki ja siirry seuraavaan lapseen
                    If pv = 0 And IsEmpty(matulorng.Offset(, 2).Value2) And IsEmpty(matulorng.Offset(, 4).Value2) And IsEmpty(matulorng.Offset(, 6).Value2) And IsEmpty(matulorng.Offset(, 8).Value2) Then
                        .Cells(Rivi + pv, Sarake).Value = "?"
                        .Cells(Rivi + pv, Sarake).Style = "Pieni"
                        Exit Sub
                    End If
                End If
            Next pv
        Next lnimi
    End With

End Sub

Sub Aamu_ilta_listat()

    ' Sorttausten nollaus. Tarpeellistako muka?
    tbl_lapset.Sort.SortFields.Clear
    tbl_lapset.Sort.Apply
    tbl_ryhmät.Sort.SortFields.Clear
    tbl_ryhmät.Sort.Apply

    ' Haetaan aamuilta-listojen kellonajat
    Dim aamu1 As String: aamu1 = Replace(sht_päiväkoti.Range("J5").Value, ",", ".")
    Dim aamu2 As String: aamu2 = Replace(sht_päiväkoti.Range("K5").Value, ",", ".")
    Dim ilta1 As String: ilta1 = Replace(sht_päiväkoti.Range("L5").Value, ",", ".")
    Dim ilta2 As String: ilta2 = Replace(sht_päiväkoti.Range("M5").Value, ",", ".")

    ' Aamulista
    Sheets("Aamuilta_pohja").Copy After:=Sheets("lapset")
    With ActiveSheet
        .Name = "Aamulista"
    End With

    Dim sht_aamulista As Worksheet: Set sht_aamulista = wb.Worksheets("Aamulista")

    ' päivämäärien lisääminen
    Dim pvm_vuosi As Long
    Dim pvm_kk As Long
    Dim pvm_pv As Long
    Dim pvm As Date

    ' Haetaan pvm Codesta ja muunnetaan sopivaan muotoon
    ' Päivämäärän koonti
    pvm = DateSerial(Year(Now), sht_code.[C3].Value, sht_code.[C2].Value)
    Dim viikonpäivät(1 To 5) As Range

    Set viikonpäivät(1) = Range("tbl_lapset[Ma tulo]")
    Set viikonpäivät(2) = Range("tbl_lapset[Ti tulo]")
    Set viikonpäivät(3) = Range("tbl_lapset[Ke tulo]")
    Set viikonpäivät(4) = Range("tbl_lapset[To tulo]")
    Set viikonpäivät(5) = Range("tbl_lapset[Pe tulo]")

    Dim päivänimi(1 To 5) As String
    sht_aamulista.Range("A2") = pvm
    sht_aamulista.Range("C2") = DateAdd("d", 1, CDate(pvm))
    sht_aamulista.Range("E2") = DateAdd("d", 2, CDate(pvm))
    sht_aamulista.Range("G2") = DateAdd("d", 3, CDate(pvm))
    sht_aamulista.Range("I2") = DateAdd("d", 4, CDate(pvm))

    Dim Rivi As Long: Rivi = 2
    Dim vk As Variant
    Dim u As Range
    Dim paivano As Double
    Dim kerta As Double
    kerta = 1
    paivano = 0

    sht_aamulista.Range("A1").Value2 = "AAMULISTA"

    For Each vk In viikonpäivät
        tbl_lapset.AutoFilter.ShowAllData
        'sht_aamulista.Range("A1").Offset(rivi - 1).Value = päivänimi(kerta)
        'sht_aamulista.Cells(2, 1 + paivano).value = päivänimi(kerta)
        If paivano = 0 Then
            ' ma
            tbl_lapset.Range.AutoFilter Field:=ls_matulo.Column, Criteria1:=">=" & aamu1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & aamu2, Operator:=xlFilterValues
        ElseIf paivano = 2 Then
            ' ti
            tbl_lapset.Range.AutoFilter Field:=ls_titulo.Column, Criteria1:=">=" & aamu1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & aamu2, Operator:=xlFilterValues
        ElseIf paivano = 4 Then
            ' ke
            tbl_lapset.Range.AutoFilter Field:=ls_ketulo.Column, Criteria1:=">=" & aamu1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & aamu2, Operator:=xlFilterValues
        ElseIf paivano = 6 Then
            ' to
            tbl_lapset.Range.AutoFilter Field:=ls_totulo.Column, Criteria1:=">=" & aamu1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & aamu2, Operator:=xlFilterValues
        
        ElseIf paivano = 8 Then
            ' pe
            tbl_lapset.Range.AutoFilter Field:=ls_petulo.Column, Criteria1:=">=" & aamu1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & aamu2, Operator:=xlFilterValues
        End If

        With tbl_lapset.Sort
            ' Sorttaus: Järjestyksen mukaan, käänteisesti
            .SortFields.Clear
            .SortFields.Add Key:=vk, SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With

        'On Error Resume Next
        On Error GoTo 0

        If tyhjasuodatin(rng_lapset.Offset(, 5)) = True Then
        Else
            For Each u In rng_lapset.SpecialCells(xlCellTypeVisible)
                ' Jos tyhjä
                If sht_aamulista.Cells(1 + Rivi, 1 + paivano).Value2 = vbNullString Then
                    sht_aamulista.Cells(1 + Rivi, 1 + paivano).Value2 = u.Offset(, 6 + paivano).Value2
                    sht_aamulista.Cells(1 + Rivi, 2 + paivano).Value2 = u.Value2
                    Rivi = Rivi + 1

                Else
                    ' Jos sama
                    If sht_aamulista.Cells(1 + Rivi, 1 + paivano).Value2 = u.Offset(, 6 + paivano).Value2 Then
                        sht_aamulista.Cells(1 + Rivi, 2 + paivano).Value2 = u.Value
                        Rivi = Rivi + 1
                    Else
                        ' Jos eri
                        ' sht_aamulista.Range(Cells(rivi + 2, 2), Cells(rivi + 2, 3)).Style = "Yläviiva"
                        Rivi = Rivi + 1
                        sht_aamulista.Cells(1 + Rivi, 1 + paivano).Value2 = u.Offset(, 6 + paivano).Value2
                        sht_aamulista.Cells(1 + Rivi, 2 + paivano).Value2 = u.Value2
                    End If
                End If
            Next u

            Dim päivä As Range
            Set päivä = sht_aamulista.Range(Cells(3, 1 + paivano), Cells(3, 1 + paivano).End(xlDown).End(xlDown).End(xlUp))
    
            Dim n As Long: n = 0
            For n = 1 To päivä.Rows.Count
                If päivä.Cells(n, 1) = päivä.Cells(n + 1, 1) Then päivä.Cells(n + 1, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 2, 1) Then päivä.Cells(n + 2, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 3, 1) Then päivä.Cells(n + 3, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 4, 1) Then päivä.Cells(n + 4, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 5, 1) Then päivä.Cells(n + 5, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 6, 1) Then päivä.Cells(n + 6, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 7, 1) Then päivä.Cells(n + 7, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 8, 1) Then päivä.Cells(n + 8, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 9, 1) Then päivä.Cells(n + 9, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 10, 1) Then päivä.Cells(n + 10, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 11, 1) Then päivä.Cells(n + 11, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 12, 1) Then päivä.Cells(n + 12, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 13, 1) Then päivä.Cells(n + 13, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 14, 1) Then päivä.Cells(n + 14, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 15, 1) Then päivä.Cells(n + 15, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 16, 1) Then päivä.Cells(n + 16, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 17, 1) Then päivä.Cells(n + 17, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 18, 1) Then päivä.Cells(n + 18, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 19, 1) Then päivä.Cells(n + 19, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 20, 1) Then päivä.Cells(n + 20, 1).Value2 = ""
            Next n
        End If
        On Error GoTo 0
        Rivi = 2
        paivano = paivano + 2
        kerta = kerta + 1
    Next vk

    'Sorttaukset pois
    tbl_lapset.Sort.SortFields.Clear
    tbl_lapset.Sort.Apply
    tbl_ryhmät.Sort.SortFields.Clear
    tbl_ryhmät.Sort.Apply

    Sheets("Aamuilta_pohja").Copy After:=Sheets("Aamulista")

    With ActiveSheet
        .Name = "Iltalista"
    End With

    Dim sht_iltalista As Worksheet: Set sht_iltalista = wb.Worksheets("Iltalista")

    Set viikonpäivät(1) = Range("tbl_lapset[Ma lähtö]")
    Set viikonpäivät(2) = Range("tbl_lapset[Ti lähtö]")
    Set viikonpäivät(3) = Range("tbl_lapset[Ke lähtö]")
    Set viikonpäivät(4) = Range("tbl_lapset[To lähtö]")
    Set viikonpäivät(5) = Range("tbl_lapset[Pe lähtö]")

    Rivi = 2
    kerta = 1
    paivano = 0

    sht_iltalista.Range("A2") = pvm
    sht_iltalista.Range("C2") = DateAdd("d", 1, CDate(pvm))
    sht_iltalista.Range("E2") = DateAdd("d", 2, CDate(pvm))
    sht_iltalista.Range("G2") = DateAdd("d", 3, CDate(pvm))
    sht_iltalista.Range("I2") = DateAdd("d", 4, CDate(pvm))

    sht_iltalista.Range("A1").Value2 = "ILTALISTA"

    For Each vk In viikonpäivät

        ' Nollataan suodatukset
        tbl_lapset.AutoFilter.ShowAllData

        If paivano = 0 Then
            ' ma
            tbl_lapset.Range.AutoFilter Field:=ls_malähtö.Column, Criteria1:=">=" & ilta1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & ilta2, Operator:=xlFilterValues

        ElseIf paivano = 2 Then
            ' ti
            tbl_lapset.Range.AutoFilter Field:=ls_tilähtö.Column, Criteria1:=">=" & ilta1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & ilta2, Operator:=xlFilterValues
        ElseIf paivano = 4 Then
            ' ke
            tbl_lapset.Range.AutoFilter Field:=ls_kelähtö.Column, Criteria1:=">=" & ilta1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & ilta2, Operator:=xlFilterValues
        ElseIf paivano = 6 Then
            ' to
            tbl_lapset.Range.AutoFilter Field:=ls_tolähtö.Column, Criteria1:=">=" & ilta1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & ilta2, Operator:=xlFilterValues
        ElseIf paivano = 8 Then
            ' pe
            tbl_lapset.Range.AutoFilter Field:=ls_pelähtö.Column, Criteria1:=">=" & ilta1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & ilta2, Operator:=xlFilterValues
        End If

        With tbl_lapset.Sort
            ' Sorttaus: Järjestyksen mukaan, käänteisesti
            .SortFields.Clear
            .SortFields.Add Key:=vk, SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With

        On Error GoTo 0
        If tyhjasuodatin(rng_lapset.Offset(, 4)) = True Then
        Else
            For Each u In rng_lapset.SpecialCells(xlCellTypeVisible)
                ' Jos tyhjä
                If sht_iltalista.Cells(1 + Rivi, 1 + paivano).Value2 = vbNullString Then
                    sht_iltalista.Cells(1 + Rivi, 1 + paivano).Value2 = u.Offset(, 7 + paivano).Value2
                    sht_iltalista.Cells(1 + Rivi, 2 + paivano).Value2 = u.Value2
                    Rivi = Rivi + 1
                Else
                    ' Jos sama
                    If sht_iltalista.Cells(1 + Rivi, 1 + paivano).Value2 = u.Offset(, 7 + paivano).Value2 Then
                        sht_iltalista.Cells(1 + Rivi, 2 + paivano).Value2 = u.Value
                        Rivi = Rivi + 1
                    Else
                        ' Jos eri
                        ' sht_aamulista.Range(Cells(rivi + 2, 2), Cells(rivi + 2, 3)).Style = "Yläviiva"
                        Rivi = Rivi + 1
                        sht_iltalista.Cells(1 + Rivi, 1 + paivano).Value2 = u.Offset(, 7 + paivano).Value2
                        sht_iltalista.Cells(1 + Rivi, 2 + paivano).Value2 = u.Value2
                    End If
                End If
            Next u
    
            Set päivä = sht_iltalista.Range(Cells(3, 1 + paivano), Cells(3, 1 + paivano).End(xlDown).End(xlDown).End(xlUp))
    
    
            n = 0
            For n = 1 To päivä.Rows.Count
                If päivä.Cells(n, 1) = päivä.Cells(n + 1, 1) Then päivä.Cells(n + 1, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 2, 1) Then päivä.Cells(n + 2, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 3, 1) Then päivä.Cells(n + 3, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 4, 1) Then päivä.Cells(n + 4, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 5, 1) Then päivä.Cells(n + 5, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 6, 1) Then päivä.Cells(n + 6, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 7, 1) Then päivä.Cells(n + 7, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 8, 1) Then päivä.Cells(n + 8, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 9, 1) Then päivä.Cells(n + 9, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 10, 1) Then päivä.Cells(n + 10, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 11, 1) Then päivä.Cells(n + 11, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 12, 1) Then päivä.Cells(n + 12, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 13, 1) Then päivä.Cells(n + 13, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 14, 1) Then päivä.Cells(n + 14, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 15, 1) Then päivä.Cells(n + 15, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 16, 1) Then päivä.Cells(n + 16, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 17, 1) Then päivä.Cells(n + 17, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 18, 1) Then päivä.Cells(n + 18, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 19, 1) Then päivä.Cells(n + 19, 1).Value2 = ""
                If päivä.Cells(n, 1) = päivä.Cells(n + 20, 1) Then päivä.Cells(n + 20, 1).Value2 = ""
    
            Next n
    
        End If
        On Error GoTo 0
        Rivi = 2
        paivano = paivano + 2
        kerta = kerta + 1
    Next vk

    'If rivi > rivi2 Then
    '    sht_aamulista.Range(Cells(1, 5), Cells(rivi - 1, 5)).Style = "Oikeaviiva"
    'Else
    '    sht_aamulista.Range(Cells(1, 5), Cells(rivi2 - 1, 5)).Style = "Oikeaviiva"
    'End If

End Sub

Sub Tapiolan_Viikonloppulaput()

End Sub

Sub Rivienpoisto(Table_nimi As Range)
    Dim rng As Range
    On Error Resume Next
    Set rng = Table_nimi.SpecialCells(xlCellTypeBlanks)
    
    If Not rng Is Nothing Then
        rng.Delete Shift:=xlUp
    End If
    On Error GoTo 0
End Sub

Function tyhjasuodatin(ByRef rngstart As Range) As Boolean

    tyhjasuodatin = False
    Dim rngFiltered As Range
    'here I get an error if there are no cells
    On Error GoTo hell
    Set rngFiltered = rngstart.SpecialCells(xlCellTypeVisible)

    Exit Function

hell:
    tyhjasuodatin = True

End Function

Sub AlkuTarkistus()

    '************************
    ' PÄIVÄKOTI
    '************************

    'If sht_code.Range("D2").Value = 0 Then
    '    Lopetus
    '    MsgBox "Hoitoaikoja ei ole lisätty oikein." & vbCrLf & "Käy lisäämässä hoitoajat Päiväkoti-välilehdellä. _
    Lue ohjeet. Jos ongelmat jatkuvat, ota yhteyttä: jaakko.haavisto@jyvaskyla.fi", vbExclamation, "Virhe"
    '    End
    'End If

    '************************
    ' AAMU JA ILTALISTAT
    '************************
    Dim o As Range
    ' Käytössä: tyhjä --> "Ei"
    If sht_päiväkoti.Range("I5") = vbNullString Then sht_päiväkoti.Range("I5") = "Ei"
    ' Jos Käytössä, kellonajat eivät saa olla tyhjät
    If sht_päiväkoti.Range("I5") = "Kyllä" Then
        For Each o In sht_päiväkoti.Range("I5:M5")
            If o = vbNullString Then Call Lopetus("Olet valinnut aamu-ja iltalistat, mutta et ole merkannut kellonaikoja." & vbCrLf & "Lisää ne Päiväkoti -välilehdellä.", vbExclamation, "Virhe")
        Next o
    End If
    ' Alku pitää olla aikaisempi kuin loppu
    If sht_päiväkoti.Range("J5") > sht_päiväkoti.Range("K5") Then Call Lopetus("Aamulistan alku-klo ei voi olla isompi kuin loppu-klo." & vbCrLf & "Tarkasta kellonajat Päiväkoti -välilehdeltä.", vbExclamation, "Virhe")
    If sht_päiväkoti.Range("L5") > sht_päiväkoti.Range("M5") Then Call Lopetus("Iltalistan alku-klo ei voi olla isompi kuin loppu-klo." & vbCrLf & "Tarkasta kellonajat Päiväkoti -välilehdeltä.", vbExclamation, "Virhe")

    '************************
    ' RYHMÄT
    '************************

    For Each o In rng_ryhmät.SpecialCells(xlCellTypeVisible)
        'Jos Käytössä = tyhjä --> "Kyllä"
        If sht_ryhmät.Cells(o.Row, rs_käytössä.Column).Value = vbNullString Then sht_ryhmät.Cells(o.Row, rs_käytössä.Column).Value = "Kyllä"
    Next o

    ' Suodatetaan Käytössä -mukaan
    'tbl_ryhmät.Range.AutoFilter 2, "Kyllä"
    With tbl_ryhmät.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rs_käytössä, SortOn:=xlSortOnValues, Order:=xlDescending
        .SortFields.Add Key:=rs_järjestys, SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .Apply
    End With

    On Error Resume Next
    If tyhjasuodatin(rng_ryhmät.SpecialCells(xlCellTypeVisible)) = True Then Call Lopetus("Et ole merkinnyt yhtäkään ryhmää käytössä olevaksi." & vbCrLf & "Ole hyvä ja merkitse joku ryhmistä käyttöön, vaikka tarkoituksenasi olisi tulostaa pelkästään aamu- ja iltalistat.", vbExclamation, "Virhe")
    
    On Error GoTo 0
    With sht_ryhmät
        For Each o In rng_ryhmät.SpecialCells(xlCellTypeVisible)
            ' Onko käytössä
            If .Cells(o.Row, rs_käytössä.Column) = "Kyllä" Then
            
                ' Jos asetukset tyhjinä, merkitään oletusasetukset
                If .Cells(o.Row, rs_järjestys.Column).Value = vbNullString Then .Cells(o.Row, rs_järjestys.Column).Value = "1"
                If .Cells(o.Row, rs_aakkosjärjestys.Column).Value = vbNullString Then .Cells(o.Row, rs_aakkosjärjestys.Column).Value = "Kutsumanimi"
            
                If .Cells(o.Row, rs_boldaus1.Column).Value = vbNullString Then .Cells(o.Row, rs_boldaus1.Column).Value = "6:59"
                If .Cells(o.Row, rs_boldaus2.Column).Value = vbNullString Then .Cells(o.Row, rs_boldaus2.Column).Value = "17:01"
                If .Cells(o.Row, rs_pboldaus1.Column).Value = vbNullString Then .Cells(o.Row, rs_pboldaus1.Column).Value = "10:00"
                If .Cells(o.Row, rs_pboldaus2.Column).Value = vbNullString Then .Cells(o.Row, rs_pboldaus2.Column).Value = "14:30"
            
                If .Cells(o.Row, rs_ruokatulostus.Column).Value = vbNullString Then .Cells(o.Row, rs_ruokatulostus.Column).Value = "Ma-pe"
                If .Cells(o.Row, rs_ruokaAsetukset.Column).Value = vbNullString Then .Cells(o.Row, rs_listatulostus.Column).Value = "Pieni fontti"
            
                If .Cells(o.Row, rs_aamupala1.Column).Value = vbNullString Then .Cells(o.Row, rs_aamupala1.Column).Value = "7:50"
                If .Cells(o.Row, rs_aamupala2.Column).Value = vbNullString Then .Cells(o.Row, rs_aamupala2.Column).Value = "8:15"
                If .Cells(o.Row, rs_lounas1.Column).Value = vbNullString Then .Cells(o.Row, rs_lounas1.Column).Value = "11:50"
                If .Cells(o.Row, rs_lounas2.Column).Value = vbNullString Then .Cells(o.Row, rs_lounas2.Column).Value = "12:15"
                If .Cells(o.Row, rs_välipala1.Column).Value = vbNullString Then .Cells(o.Row, rs_välipala1.Column).Value = "13:55"
                If .Cells(o.Row, rs_välipala2.Column).Value = vbNullString Then .Cells(o.Row, rs_välipala2.Column).Value = "14:15"
                If .Cells(o.Row, rs_iltapala1.Column).Value = vbNullString Then .Cells(o.Row, rs_iltapala1.Column).Value = "18:55"
                If .Cells(o.Row, rs_iltapala2.Column).Value = vbNullString Then .Cells(o.Row, rs_iltapala2.Column).Value = "19:15"
            
                If .Cells(o.Row, rs_listatulostus.Column).Value = vbNullString Then .Cells(o.Row, rs_listatulostus.Column).Value = "Ma-pe"
                If .Cells(o.Row, rs_listatulostus.Column).Value = "Kyllä" Then .Cells(o.Row, rs_listatulostus.Column).Value = "Ma-pe"
                
                If .Cells(o.Row, rs_päivystys.Column).Value = vbNullString Then .Cells(o.Row, rs_päivystys.Column).Value = "Ei"
                If .Cells(o.Row, rs_päivälaput.Column).Value = vbNullString Then .Cells(o.Row, rs_päivälaput.Column).Value = "Ei"
                If .Cells(o.Row, rs_plabc.Column).Value = vbNullString Then .Cells(o.Row, rs_plabc.Column).Value = "Kutsumanimi"
                If .Cells(o.Row, rs_plpohja.Column).Value = vbNullString Then .Cells(o.Row, rs_plpohja.Column).Value = "Pysty"
                If .Cells(o.Row, rs_pltyhjät.Column).Value = vbNullString Then .Cells(o.Row, rs_pltyhjät.Column).Value = "0"
                
                
                If .Cells(o.Row, rs_päivystys.Column).Value <> "Ei" And .Cells(o.Row, rs_ruokayhdistys.Column).Value <> vbNullString And .Cells(o.Row, rs_ruokatulostus.Column).Value <> "Ei" Then
                    Call Lopetus(o.Value & " -ryhmä on merkattu päivystysryhmäksi. Siinä on myös yhdistettyjä ruokatilaus-ryhmiä. " _
                               & vbCrLf & vbCrLf & "Se ei ole tuettu ominaisuus tällä hetkellä. Poista yhdistetyt ruokaryhmät.", vbExclamation, "Virhe")
                End If
                
                If .Cells(o.Row, rs_päivystys.Column).Value <> "Ei" And .Cells(o.Row, rs_listayhdistys.Column).Value <> vbNullString And .Cells(o.Row, rs_listatulostus.Column).Value <> "Ei" Then
                    Call Lopetus(o.Value & " -ryhmä on merkattu päivystysryhmäksi. Siinä on myös yhdistettyjä ryhmiä. " _
                               & vbCrLf & vbCrLf & "Se ei ole tuettu ominaisuus tällä hetkellä. Poista yhdistetyt ryhmät.", vbExclamation, "Virhe")
                End If
                
                If .Cells(o.Row, rs_päivystys.Column).Value <> "Ei" And .Cells(o.Row, rs_listayhdistys.Column).Value <> vbNullString And .Cells(o.Row, rs_päivälaput.Column).Value <> "Ei" Then
                    Call Lopetus(o.Value & " -ryhmä on merkattu päivystysryhmäksi. Siinä on myös yhdistettyjä ryhmiä. " _
                               & vbCrLf & vbCrLf & "Se ei ole tuettu ominaisuus tällä hetkellä. Poista yhdistetyt ryhmät.", vbExclamation, "Virhe")
                End If
                
                ' TODO
                ' Yhdistettyjä ruokalistoja, mutta nimi puuttuu --> Ryhmien nimet
   
                ' Yhdistettyjen ryhmien tsekkaus
                Dim validi_array() As Variant
                validi_array = Application.Transpose(sht_code.Range("G2:G" & sht_code.Range("G2").End(xlDown).Row).Value2)
                Dim vl_yhdistely_array() As String
                Dim element As Variant
            
                ' Ruokalapun yhdistetyt ryhmät pitää olla valideja
                If .Cells(o.Row, rs_ruokayhdistys.Column).Value <> "" Then
                    vl_yhdistely_array = Split(spaceremove(.Cells(o.Row, rs_listayhdistys.Column).Value), ",", , vbTextCompare)
                    For Each element In vl_yhdistely_array
                        If IsInArray(CStr(element), validi_array) = False Then
                            Call Lopetus(o.Value & " -ryhmän kanssa yhdistetyn ruokalapun kanssa on ongelma. " & element & _
                                         " ei ole oikea ryhmä." & vbCrLf & vbCrLf & "Mahdollisia ryhmiä ovat: " & _
                                         Join(validi_array, ", "), vbExclamation, "Virhe")
                        End If
                    Next element
                End If
            
                ' Viikkolistan yhdistetyt ryhmät pitää olla valideja
                If .Cells(o.Row, rs_listayhdistys.Column).Value <> "" Then
                    vl_yhdistely_array = Split(spaceremove(.Cells(o.Row, rs_listayhdistys.Column).Value), ",", , vbTextCompare)
                    For Each element In vl_yhdistely_array
                        If IsInArray(CStr(element), validi_array) = False Then Call Lopetus(o.Value & " -ryhmän kanssa yhdistetyn viikkolistan kanssa on ongelma. " & element & _
                                                                                            " ei ole oikea ryhmä." & vbCrLf & vbCrLf & "Mahdollisia ryhmiä ovat: " & _
                                                                                            Join(validi_array, ", "), vbExclamation, "Virhe")
                    Next element
                End If
            
                ' Yhdistettyjä viikkolistoja, mutta listan nimi puuttuu (Viikkolistan tulostus, Ryhmiä valittuna, asetuksissa yhdistetty viikkolista, mutta listan nimi tyhjä)
                If .Cells(o.Row, rs_listatulostus.Column).Value <> "Ei" And _
                                                                .Cells(o.Row, rs_listayhdistys.Column).Value <> "" And _
                                                                Left(.Cells(o.Row, rs_yhdistettytyyli.Column).Value, 10) = "Yhdistetty" And _
                                                                .Cells(o.Row, rs_yhdistettynimi.Column).Value = "" Then
                    Call Lopetus(o.Value & " -ryhmällä on yhdistettyjä viikkolistoja, mutta listan nimi puuttuu." & vbCrLf & "Käy lisäämässä yhdistetyn listan nimi ja luo listat uudelleen.", vbExclamation, "Virhe")
                End If
            
                ' Jos päivälaput ja asetuksista 2-puoliset päivälaput, mutta ryhmien yhdistäminen on tyhjä --> Valitus
                If .Cells(o.Row, rs_päivälaput.Column).Value <> "Ei" And _
                                                             Right(.Cells(o.Row, rs_yhdistettytyyli.Column).Value, 22) = "2-puoleiset päivälaput" And _
                                                             .Cells(o.Row, rs_listayhdistys.Column).Value = vbNullString Then
                    Call Lopetus("Olet valinnut ryhmälle " & o.Value & " 2-puoleisten päivälappujen tulostamisen," & vbCrLf & "mutta et ole valinnut 2. sivulle tulevaa ryhmää." & vbCrLf & _
                                 "Ole hyvä ja merkitse ryhmän nimi Ryhmien yhdistäminen-sarakkeeseen Ryhmät-välilehdellä.", vbExclamation, "Virhe")
                End If
            
                Dim arrayIsNothing As Boolean
                On Error Resume Next
                arrayIsNothing = IsNumeric(UBound(vl_yhdistely_array)) And False
                If Err.Number <> 0 Then arrayIsNothing = True
                On Error GoTo 0

                ' Jos yhdistettyjen ryhmien array on tyhjä, ei tehdä alkutarkistusta.
                If arrayIsNothing = False Then
                    ' Jos 2-puoleiset päivälaput, saa olla ainoastaan 1 yhdistettävä ryhmä, jos enemmän ryhmiä --> Valitus
                    If .Cells(o.Row, rs_päivälaput.Column).Value <> "Ei" And _
                                                                 Right(.Cells(o.Row, rs_yhdistettytyyli.Column).Value, 22) = "2-puoleiset päivälaput" And _
                                                                 ArrayLen(vl_yhdistely_array) > 1 Then
                        Call Lopetus("Olet valinnut ryhmälle " & o.Value & " 2-puoleisten päivälappujen tulostamisen," & vbCrLf & "Tämä asetus rajoittaa yhdistettävien ryhmien määrän yhteen." & vbCrLf & _
                                     "Käy poistamassa ylimääräiset yhdistetyt ryhmät.", vbExclamation, "Virhe")
                    End If
                End If
           
                ' Jos päivälaput + 2-puolinen, tarkistetaan 2.lapun asetukset
                If .Cells(o.Row, rs_päivälaput.Column).Value <> "Ei" And _
                                                             Right(.Cells(o.Row, rs_yhdistettytyyli.Column).Value, 22) = "2-puoleiset päivälaput" Then
                    Dim tokaryhmä As Integer: tokaryhmä = rs_ryhmännimi.Find(vl_yhdistely_array(0)).Row
                    ' Oletusasetukset pohjalle (pysty) ja tyhjien määrälle (0
                    If .Cells(tokaryhmä, rs_plpohja.Column).Value2 = vbNullString Then .Cells(tokaryhmä, rs_plpohja.Column).Value2 = "Pysty"
                    If .Cells(tokaryhmä, rs_pltyhjät.Column).Value2 = vbNullString Then .Cells(tokaryhmä, rs_pltyhjät.Column).Value2 = 0
                    ' Jos 2. päivälapun pohja on kustomoitu, eikä oikeaa pohjaa löydy --> Valitus
                    If .Cells(tokaryhmä, rs_plpohja.Column).Value2 = "Kustomoitu" And WorksheetExists("pl_" & .Cells(tokaryhmä, rs_ryhmännimi.Column).Value) = False Then
                        Call Lopetus(o.Value & " -ryhmällä on 2-puoleinen päivälappu-asetus. Kuitenkaan 2. ryhmän kustomoitua päivälappua ei ole olemassa." & vbCrLf & "Käy generoimassa kustomoitu pohja " & sht_ryhmät.Cells(tokaryhmä, rs_ryhmännimi.Column).Value & " -ryhmälle Päiväkoti-välilehdellä.", vbExclamation, "Virhe")
                    End If
                End If
                ' Jos päivälaput + kustomoitu pohja, mutta pohjaa ei löydy --> Valitus
                If .Cells(o.Row, rs_päivälaput.Column).Value <> "Ei" And _
                                                             .Cells(o.Row, rs_plpohja.Column).Value = "Kustomoitu" And _
                                                             WorksheetExists("pl_" & .Cells(o.Row, rs_ryhmännimi.Column).Value) = False Then
                    Call Lopetus(o.Value & " -ryhmällä on kustomoitu päivälappu, mutta sitä ei ole olemassa." & vbCrLf & "Käy generoimassa kustomoitu pohja Päiväkoti-välilehdellä.", vbExclamation, "Virhe")
                End If
                    

            End If
        Next o
    End With

End Sub

Sub Lopetus_Simple()
    ThisWorkbook.Worksheets("Päiväkoti").Select

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    End
End Sub

Sub Lopetus(Optional VirheViesti As String, Optional VirheTyyppi As String, Optional VirheOtsikko As String)

    Dim sort_ryhmät_järjestys As Range
    Set sort_ryhmät_järjestys = Range("tbl_ryhmät[Järjestys]")

    ' Sorttaus: järjestys alusta loppuun
    With sht_ryhmät.ListObjects("tbl_ryhmät").Sort
        .SortFields.Clear
        .SortFields.Add Key:=sort_ryhmät_järjestys, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    ' Lisätään suojaukset
    'For Each WSprotect In Array("R_pohja1", "R_pohja2", "Ri_pohja1", "Ri_pohja2", "Pohja")
    '    Worksheets(WSprotect).Protect AllowFormattingCells:=True, AllowFormattingColumns:=True, _
    '        AllowFormattingRows:=True
    'Next WSprotect

    ' Nollataan suodatukset
    Dim sort_lapset_abc As Range
    Set sort_lapset_abc = Range("tbl_lapset[Ryhmä]")

    'tbl_lapset.AutoFilter.ShowAllData
    With tbl_lapset.Sort
        ' Sorttaus: Järjestyksen mukaan, käänteisesti
        .SortFields.Clear
        .SortFields.Add Key:=sort_lapset_abc, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    ' Piilotetaan välilehdet
    Sheets(Array("lasna", "Code", "Pohja", "Aamuilta_pohja", "R_arki", "R_ilta", "R_arkikaikki", "R_kaikki", "pl_pysty", "pl_pystywc", "pl_vaaka", "pl_vaakawc", "Ruokakoonti_pohja")).Select
    ActiveWindow.SelectedSheets.Visible = False

    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Worksheets
        If sht.Visible Then
            sht.Activate
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End If
    Next sht

    wb.Worksheets("Päiväkoti").Select
    Call PoistaSuodatukset

    If VirheViesti <> vbNullString Then MsgBox VirheViesti, VirheTyyppi, VirheOtsikko

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    End

End Sub

Sub PoistaSuodatukset()

    Set wb = ThisWorkbook
    Set sht_ryhmät = wb.Worksheets("Ryhmät")
    Set sht_lapset = wb.Worksheets("Lapset")
    Set tbl_ryhmät = sht_ryhmät.ListObjects("tbl_ryhmät")
    Set tbl_lapset = sht_lapset.ListObjects("tbl_lapset")
    tbl_ryhmät.ShowAutoFilter = True
    tbl_ryhmät.AutoFilter.ShowAllData
    tbl_lapset.ShowAutoFilter = True
    tbl_lapset.AutoFilter.ShowAllData

End Sub

Sub PoistaTiedot()

    ' Poistetaan tiedot, jos codessa valittuna
    If tyhjasuodatin(sht_lista.ListObjects("tbl_lista").DataBodyRange) = False Then
        sht_lista.ListObjects("tbl_lista").DataBodyRange.Delete
    End If
    sht_code.Range("D2").Value = 0

    If tyhjasuodatin(sht_TLista.ListObjects("tbl_tlista").DataBodyRange) = False Then
        sht_TLista.ListObjects("tbl_tlista").DataBodyRange.Delete
    End If
    sht_code.Range("E2").Value = 0

End Sub

Private Function RangeToString(ByRef rngDisplay As Range, ByVal strSeparator As String) As String

    'The string to separate elements on the message box,
    'if the range size is more than one cell
    
    Dim strMessage      As String
    Dim astrMessage()   As String
    
    Dim avarRange()     As Variant
    Dim varElement      As Variant
    
    Dim i               As Long
    
    'If the range is only one cell, we will return that that
    If rngDisplay.Cells.Count = 1 Then
        strMessage = rngDisplay.Value
        'Else the range is multiple cells, so we need to concatenate their values
    Else
        'Assign range to a variant array
        avarRange = rngDisplay
        
        'Loop through each element to build a one-dimensional array of the range
        For Each varElement In avarRange
        
            ReDim Preserve astrMessage(i)
            astrMessage(i) = CStr(varElement)
            i = 1 + i
            
        Next varElement
        
        'Build the string to return
        strMessage = Join(astrMessage, strSeparator)
        
    End If
    
    RangeToString = strMessage
    
End Function

Function ConcatenateRow(rowRange As Range, joinString As String) As String
    Dim x As Variant, temp As String

    temp = " "
    For Each x In rowRange
        temp = temp & x & joinString
    Next

    ConcatenateRow = Left(temp, Len(temp) - Len(joinString))
End Function

Sub Poistaryhmat()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Näytetään kaikki välilehdet
    Dim WSname As Variant
    For Each WSname In Array("lasna", "Code", "Pohja", "Aamuilta_pohja")
        Worksheets(WSname).Visible = True
    Next WSname

    Set wb = ThisWorkbook
    Set sht_code = wb.Worksheets("Code")

    ' Poistetaan vanhat ryhmät
    Dim lRow As Long
    Dim vanharyh As Range
    lRow = sht_code.Cells(Rows.Count, 2).End(xlUp).Row
    On Error Resume Next
    If lRow <> 1 Then
        sht_code.Select
        For Each vanharyh In sht_code.Range(Cells(2, 2), Cells(lRow, 2))
            If Not Sheets(vanharyh.Value) Is Nothing Then
                Sheets(vanharyh.Value).Delete
                Sheets(vanharyh.Value & "_ruoka").Delete
                Sheets(vanharyh.Value & "_ma").Delete
                Sheets(vanharyh.Value & "_ti").Delete
                Sheets(vanharyh.Value & "_ke").Delete
                Sheets(vanharyh.Value & "_to").Delete
                Sheets(vanharyh.Value & "_pe").Delete
            End If
        Next vanharyh
        sht_code.Range(Cells(2, 2), Cells(lRow, 2)).Value = vbNullString
        Sheets("Aamulista").Delete
        Sheets("Iltalista").Delete
    End If

    On Error GoTo 0

    ' Piilotetaan välilehdet
    Sheets(Array("lasna", "Code", "Pohja", "Aamuilta_pohja")).Select
    ActiveWindow.SelectedSheets.Visible = False

    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Worksheets
        If sht.Visible Then
            sht.Activate
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End If
    Next sht

    wb.Worksheets("Päiväkoti").Select

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

Sub Hoitoaikadata()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Näytetään kaikki välilehdet
    Dim WSname As Variant
    For Each WSname In Array("lasna", "Code", "Pohja", "Aamuilta_pohja")
        Worksheets(WSname).Visible = True
    Next WSname

    Set wb = ThisWorkbook

    Set sht_code = wb.Worksheets("Code")
    Set sht_lasna = wb.Worksheets("lasna")

    sht_code.Select
    sht_code.Range("G2:G2000").Value2 = ""
    sht_lasna.Select
    sht_lasna.Range("A1:AF3000").Value2 = ""

    ' Piilotetaan välilehdet
    Sheets(Array("lasna", "Code", "Pohja", "Aamuilta_pohja")).Select
    ActiveWindow.SelectedSheets.Visible = False

    Dim sht As Worksheet
    For Each sht In ActiveWorkbook.Worksheets
        If sht.Visible Then
            sht.Activate
            Range("A1").Select
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End If
    Next sht

    wb.Worksheets("Päiväkoti").Select

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Function spaceremove(strs) As String
    Dim str As String
    Dim nstr As String
    Dim sstr As String
    Dim x As Long
    str = strs

    For x = 1 To VBA.Len(str)
        sstr = Left(Mid(str, x), 1)
        If sstr = " " Or sstr = " " Then
        Else
            nstr = nstr & "" & sstr
        End If

    Next x
    spaceremove = nstr
End Function

' Tulostaa arrayn lukumäärän
Function arrayitems(arr As Variant) As Long
    Dim lukumäärä As Long: lukumäärä = UBound(arr) - LBound(arr) + 1
    arrayitems = lukumäärä
End Function

Public Function ArrayLen(arr As Variant) As Long
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Sub Ruokatilausten_lähettäminen()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim WSname As Variant
    Worksheets("Code").Visible = True
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sht_code As Worksheet: Set sht_code = wb.Worksheets("Code")
    Dim lRow As Long: lRow = sht_code.Cells(Rows.Count, 2).End(xlUp).Row
    
    Dim pdflista As Collection
    Set pdflista = New Collection
    
    Dim i As Long
    If lRow <> 1 Then
        sht_code.Select
        For i = sht_code.Range(Cells(2, 2), Cells(lRow, 2)).Count To 1 Step -1
            If Right(sht_code.Cells(1 + i, 2).Value2, 5) = "ruoka" Then
                pdflista.Add sht_code.Cells(1 + i, 2).Value
            End If
        Next i
    End If
        
    Set sht_päiväkoti = wb.Worksheets("Päiväkoti")
    
    Dim emailApplication As Object
    Dim emailItem As Object
    Dim strPath As String
    Dim lngPos As Long
    ' Build the PDF file name
    'strPath = ActiveWorkbook.FullName
    strPath = Application.ActiveWorkbook.Path + "\" + sht_päiväkoti.Range("I11").Value2 + ".pdf"
    'MsgBox strPath
    'lngPos = InStrRev(strPath, ".")
    'strPath = Left(strPath, lngPos) & "pdf"
    ' Export workbook as PDF
    If pdflista.Count > 0 Then
        Sheets(collectionToArray(pdflista)).Select
    Else
        MsgBox "Ei löydetty yhtään ruokatilauslappua."
        Call Lopetus_Simple
    End If
    
    ActiveSheet.ExportAsFixedFormat xlTypePDF, strPath, OpenAfterPublish:=False, IgnorePrintAreas:=False
    
    '    ActiveWorkbook.ExportAsFixedFormat xlTypePDF, strPath
    Set emailApplication = CreateObject("Outlook.Application")
    Set emailItem = emailApplication.CreateItem(0)
    ' Now we build the email.
    emailItem.To = sht_päiväkoti.Range("I9").Value2
    emailItem.Subject = sht_päiväkoti.Range("I13").Value2
    emailItem.Body = sht_päiväkoti.Range("I15").Value2
    ' Attach the PDF file
    emailItem.Attachments.Add strPath
    ' Send the Email
    ' Use this OR .Display, but not both together.
    emailItem.Send
    ' Display the Email so the user can change it as desired before sending it
    ' Use this OR .Send, but not both together.
    'emailItem.Display
    Set emailItem = Nothing
    Set emailApplication = Nothing
    ' Delete the PDF file
    Kill strPath
    Call Lopetus_Simple
    
End Sub

Function collectionToArray(c As Collection) As Variant()
    Dim a() As Variant: ReDim a(0 To c.Count - 1)
    Dim i As Integer
    For i = 1 To c.Count
        a(i - 1) = c.Item(i)
    Next
    collectionToArray = a
End Function


