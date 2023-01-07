Attribute VB_Name = "Main"
Option Explicit
' Workbooks
Public wb As Workbook
' Worksheets
Public sht_code As Worksheet, sht_ryhm�t As Worksheet, sht_lapset As Worksheet, sht_rsuodatus As Worksheet, _
sht_lsuodatus As Worksheet, sht_lasna As Worksheet, sht_lasna2 As Worksheet, sht_p�iv�koti As Worksheet
' ListObjects
Public tbl_ryhm�t As ListObject, tbl_lapset As ListObject
' Lapset
Public ls_j�rjestys As Range, ls_kutsumanimi As Range, _
ls_kokonimi As Range, ls_ryhm� As Range, ls_dieetti As Range, ls_ty�ntekij� As Range, ls_matulo As Range, _
ls_mal�ht� As Range, ls_titulo As Range, ls_til�ht� As Range, ls_ketulo As Range, ls_kel�ht� As Range, _
ls_totulo As Range, ls_tol�ht� As Range, ls_petulo As Range, ls_pel�ht� As Range, ls_matulo2 As Range, _
ls_mal�ht�2 As Range, ls_titulo2 As Range, ls_til�ht�2 As Range, ls_ketulo2 As Range, ls_kel�ht�2 As Range, _
ls_totulo2 As Range, ls_tol�ht�2 As Range, ls_petulo2 As Range, ls_pel�ht�2 As Range, ls_sukunimi As Range, _
ls_p�ivystys As Range, ls_latulo As Range, ls_lal�ht� As Range, ls_sutulo As Range, ls_sul�ht� As Range, _
ls_pienryhm� As Range, ls_arkiPoissa As Range, ls_vklPoissa As Range
' Ryhm�t
Public rs_j�rjestys As Range, rs_k�yt�ss� As Range, rs_ryhm�nnimi As Range, rs_aakkosj�rjestys As Range, _
rs_boldaus1 As Range, rs_boldaus2 As Range, rs_pboldaus1 As Range, rs_pboldaus2 As Range, _
rs_ruokatulostus As Range, rs_ruokaAsetukset As Range, rs_ruokayhdistys As Range, rs_aamupala1 As Range, _
rs_aamupala2 As Range, rs_lounas1 As Range, rs_lounas2 As Range, rs_v�lipala1 As Range, rs_v�lipala2 As Range, _
rs_p�iv�llinen1 As Range, rs_p�iv�llinen2 As Range, rs_iltapala1 As Range, rs_iltapala2 As Range, _
rs_listatulostus As Range, rs_listayhdistys As Range, rs_yhdistettynimi As Range, rs_yhdistettytyyli As Range, _
rs_p�iv�laput As Range, rs_plabc As Range, rs_plpohja As Range, rs_pltyhj�t As Range, rs_p�ivystys As Range
' Misc Ranges
Public rng_ryhm�t As Range, rng_lapset As Range, laps As Range
' Variables
Public P�iv�t As Long, Ruokailut As Long, Sarakkeet As Long, ekaDieettiRivi As Long, _
tokaDieettiRivi As Long, kolDieettiRivi As Long, nelDieettiRivi As Long, _
VL_Vkl As Long, PL_Poissa As Long, LapsetSarake As Long, ListaLis�ys As Long, _
Rivi As Long, pienennys As Long, Ruokakoonti As Boolean, koontiruoat() As Integer

Sub P��ohjelma()
    '1       2         3         4         5         6         7         8         9        10        11        12        13        14
    '2345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678902345
    Dim af_lapset As ListObject
    Dim rng_lapsisuodatus As Range

    ' Testiymp�rist�
    Dim testiymparisto As Boolean
    testiymparisto = 0
    If testiymparisto = 0 Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    End If

    Dim sort_lapset_ryhm� As Range: Set sort_lapset_ryhm� = Range("tbl_lapset[Ryhm�]")
    Dim sort_lapset_aakkoset As Range: Set sort_lapset_aakkoset = Range("tbl_lapset[Koko nimi]")

    ' N�ytet��n kaikki v�lilehdet
    Dim WSname As Variant
    For Each WSname In Array("lasna", "Code", "Pohja", "Aamuilta_pohja", "R_arki", "R_ilta", "R_arkikaikki", "R_kaikki", "pl_pysty", "pl_pystywc", "pl_vaaka", "pl_vaakawc", "Ruokakoonti_pohja")
        Worksheets(WSname).Visible = True
    Next WSname

    Set wb = ThisWorkbook

    Set sht_code = wb.Worksheets("Code")
    Set sht_ryhm�t = wb.Worksheets("Ryhm�t")
    Set sht_lapset = wb.Worksheets("Lapset")
    Set sht_p�iv�koti = wb.Worksheets("P�iv�koti")

    Set tbl_ryhm�t = sht_ryhm�t.ListObjects("tbl_ryhm�t")
    Set tbl_lapset = sht_lapset.ListObjects("tbl_lapset")

    ' lapset-taulukon sarakkeiden nimet
    Set ls_j�rjestys = tbl_lapset.ListColumns(1).Range
    Set ls_kutsumanimi = tbl_lapset.ListColumns(2).Range
    Set ls_kokonimi = tbl_lapset.ListColumns(3).Range
    Set ls_ryhm� = tbl_lapset.ListColumns(4).Range
    Set ls_dieetti = tbl_lapset.ListColumns(5).Range
    Set ls_ty�ntekij� = tbl_lapset.ListColumns(6).Range
    Set ls_p�ivystys = tbl_lapset.ListColumns(7).Range
    Set ls_matulo = tbl_lapset.ListColumns(8).Range
    Set ls_mal�ht� = tbl_lapset.ListColumns(9).Range
    Set ls_titulo = tbl_lapset.ListColumns(10).Range
    Set ls_til�ht� = tbl_lapset.ListColumns(11).Range
    Set ls_ketulo = tbl_lapset.ListColumns(12).Range
    Set ls_kel�ht� = tbl_lapset.ListColumns(13).Range
    Set ls_totulo = tbl_lapset.ListColumns(14).Range
    Set ls_tol�ht� = tbl_lapset.ListColumns(15).Range
    Set ls_petulo = tbl_lapset.ListColumns(16).Range
    Set ls_pel�ht� = tbl_lapset.ListColumns(17).Range
    Set ls_latulo = tbl_lapset.ListColumns(18).Range
    Set ls_lal�ht� = tbl_lapset.ListColumns(19).Range
    Set ls_sutulo = tbl_lapset.ListColumns(20).Range
    Set ls_sul�ht� = tbl_lapset.ListColumns(21).Range
    Set ls_sukunimi = tbl_lapset.ListColumns(22).Range
    Set ls_arkiPoissa = tbl_lapset.ListColumns(23).Range
    Set ls_vklPoissa = tbl_lapset.ListColumns(24).Range

    ' ryhm�t-taulukon sarakkeiden nimet
    Set rs_j�rjestys = tbl_ryhm�t.ListColumns(1).Range
    Set rs_k�yt�ss� = tbl_ryhm�t.ListColumns(2).Range
    Set rs_ryhm�nnimi = tbl_ryhm�t.ListColumns(3).Range
    Set rs_aakkosj�rjestys = tbl_ryhm�t.ListColumns(4).Range
    Set rs_boldaus1 = tbl_ryhm�t.ListColumns(5).Range
    Set rs_boldaus2 = tbl_ryhm�t.ListColumns(6).Range
    Set rs_pboldaus1 = tbl_ryhm�t.ListColumns(7).Range
    Set rs_pboldaus2 = tbl_ryhm�t.ListColumns(8).Range
    Set rs_ruokatulostus = tbl_ryhm�t.ListColumns(9).Range
    Set rs_ruokaAsetukset = tbl_ryhm�t.ListColumns(10).Range
    Set rs_ruokayhdistys = tbl_ryhm�t.ListColumns(11).Range
    Set rs_aamupala1 = tbl_ryhm�t.ListColumns(12).Range
    Set rs_aamupala2 = tbl_ryhm�t.ListColumns(13).Range
    Set rs_lounas1 = tbl_ryhm�t.ListColumns(14).Range
    Set rs_lounas2 = tbl_ryhm�t.ListColumns(15).Range
    Set rs_v�lipala1 = tbl_ryhm�t.ListColumns(16).Range
    Set rs_v�lipala2 = tbl_ryhm�t.ListColumns(17).Range
    Set rs_p�iv�llinen1 = tbl_ryhm�t.ListColumns(18).Range
    Set rs_p�iv�llinen2 = tbl_ryhm�t.ListColumns(19).Range
    Set rs_iltapala1 = tbl_ryhm�t.ListColumns(20).Range
    Set rs_iltapala2 = tbl_ryhm�t.ListColumns(21).Range
    Set rs_listatulostus = tbl_ryhm�t.ListColumns(22).Range
    Set rs_listayhdistys = tbl_ryhm�t.ListColumns(23).Range
    Set rs_yhdistettynimi = tbl_ryhm�t.ListColumns(24).Range
    Set rs_yhdistettytyyli = tbl_ryhm�t.ListColumns(25).Range
    Set rs_p�ivystys = tbl_ryhm�t.ListColumns(26).Range
    Set rs_p�iv�laput = tbl_ryhm�t.ListColumns(27).Range
    Set rs_plabc = tbl_ryhm�t.ListColumns(28).Range
    Set rs_plpohja = tbl_ryhm�t.ListColumns(29).Range
    Set rs_pltyhj�t = tbl_ryhm�t.ListColumns(30).Range

    Set rng_ryhm�t = tbl_ryhm�t.ListColumns(3).DataBodyRange
    Set rng_lapset = tbl_lapset.ListColumns(2).DataBodyRange
    
    If sht_p�iv�koti.Range("G8").Value = "Kyll�" Then
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

    ' Poistetaan vanhat ryhm�t
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

    ' Kun kutsumanimen kohdalla ei lue mit��n --> koko rivi tuhotaan
    Call Rivienpoisto(rng_ryhm�t)
    Call Rivienpoisto(rng_lapset)

    ' Nimet
    Dim sort_ryhm�t_j�rjestys As Range
    Set sort_ryhm�t_j�rjestys = rs_j�rjestys

    Call AlkuTarkistus

    ' Suodatus: K�yt�ss� olevat
    tbl_ryhm�t.ShowAutoFilter = True
    'tbl_ryhm�t.Range.AutoFilter Field:=rs_k�yt�ss�.Column, Criteria1:="Kyll�"

    ' Sorttaus: J�rjestyksen mukaan, k��nteisesti
    With tbl_ryhm�t.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rs_k�yt�ss�, SortOn:=xlSortOnValues, Order:=xlDescending
        .SortFields.Add Key:=rs_j�rjestys, SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .Apply
    End With

    ' Aloitetaan numerointi ykk�sest�
    Dim Ryhm�numero As Long: Ryhm�numero = 1
    Dim Code_ryhm�t() As Variant

    Dim ryh As Range
    ' Viikkolista & p�iv�lappu -yhdistelytapa
    Dim VL_Yhd As Long, PL_Yhd As Long

    'Errorit pois jotta ei valittaisi kun yht��n ryhm�� ei k�yt�ss�
    On Error Resume Next
    Dim Ryhm�solut As Long
    Ryhm�solut = rng_ryhm�t.SpecialCells(xlCellTypeVisible).Cells.Count
    If Ryhm�solut > 0 Then
        On Error GoTo 0

        ' Luupataan ryhm�t yksi kerrallaan l�pi (n�kyv�t solut suodatuksen ja filtterin j�lkeen)
        For Each ryh In rng_ryhm�t.SpecialCells(xlCellTypeVisible)
            ' Onko ryhm� k�yt�ss�
            If sht_ryhm�t.Cells(ryh.Row, rs_k�yt�ss�.Column) = "Kyll�" Then
        
                ' Nollaa suodatukset
                tbl_lapset.AutoFilter.ShowAllData
                ' Ruokalappu
                If sht_ryhm�t.Cells(ryh.Row, rs_ruokatulostus.Column).Value <> "Ei" Then
                    Call Ryhm�_Ruoka(sht_ryhm�t.Cells(ryh.Row, rs_ruokatulostus.Column).Value, sht_ryhm�t.Cells(ryh.Row, rs_ryhm�nnimi.Column).Value, _
                                     sht_ryhm�t.Cells(ryh.Row, rs_aakkosj�rjestys.Column).Value, Ryhm�numero, _
                                     sht_ryhm�t.Cells(ryh.Row, rs_ruokatulostus.Column).Value, sht_ryhm�t.Cells(ryh.Row, rs_ruokaAsetukset.Column).Value, _
                                     spaceremove(sht_ryhm�t.Cells(ryh.Row, rs_ruokayhdistys.Column).Value), _
                                     Code_ryhm�t, sht_ryhm�t.Cells(ryh.Row, rs_p�ivystys.Column).Value, ryh.Row)
                End If
                
        
                ' VL & PL Yhdistelytapa
                Select Case sht_ryhm�t.Cells(ryh.Row, rs_yhdistettytyyli.Column).Value
                Case "Viikkolista & p�iv�laput"
                    VL_Yhd = 0
                    PL_Yhd = 0
                Case "Viikkolista & yhdistetyt p�iv�laput"
                    VL_Yhd = 0
                    PL_Yhd = 1
                Case "Viikkolista & 2-puoleiset p�iv�laput"
                    VL_Yhd = 0
                    PL_Yhd = 2
                Case "Yhdistetty viikkolista & p�iv�laput"
                    VL_Yhd = 1
                    PL_Yhd = 0
                Case "Yhdistetty viikkolista & yhdistetyt p�iv�laput"
                    VL_Yhd = 1
                    PL_Yhd = 1
                Case "Yhdistetty viikkolista & 2-puoleiset p�iv�laput"
                    VL_Yhd = 1
                    PL_Yhd = 2
                End Select
        
                ' Viikkolistan tulostus
                Select Case sht_ryhm�t.Cells(ryh.Row, rs_listatulostus.Column).Value
                Case "Ma-pe"
                    VL_Vkl = 0
                Case "Ma-su"
                    VL_Vkl = 1
                End Select
    
                ' Nollaa suodatukset
                tbl_lapset.AutoFilter.ShowAllData
                ' Hoitoaikalista
                If sht_ryhm�t.Cells(ryh.Row, rs_listatulostus.Column).Value <> "Ei" Then
                    Call Ryhm�_Lista(sht_ryhm�t.Cells(ryh.Row, rs_ryhm�nnimi.Column).Value, _
                                     sht_ryhm�t.Cells(ryh.Row, rs_aakkosj�rjestys.Column).Value, Ryhm�numero, _
                                     sht_ryhm�t.Cells(ryh.Row, rs_listatulostus.Column).Value, _
                                     spaceremove(sht_ryhm�t.Cells(ryh.Row, rs_listayhdistys.Column).Value), _
                                     sht_ryhm�t.Cells(ryh.Row, rs_yhdistettynimi.Column).Value, _
                                     Code_ryhm�t, sht_ryhm�t.Cells(ryh.Row, rs_p�ivystys.Column).Value, _
                                     ryh.Row, VL_Yhd, VL_Vkl)
                End If
    
                ' Nollaa suodatukset
                tbl_lapset.AutoFilter.ShowAllData
    
                Select Case sht_ryhm�t.Cells(ryh.Row, rs_p�iv�laput.Column).Value
                Case "Kyll�"
                    PL_Poissa = 0
                    '    PL_Pr = 0
                Case "Kyll� - poissaolevat"
                    PL_Poissa = 1
                    '    PL_Pr = 0
                    'Case "Kyll� + pienryhm�t"
                    '    PL_Poissa = 0
                    '    PL_Pr = 1
                    'Case "Kyll� - poissa + pienryhm�t"
                    '    PL_Poissa = 1
                    '    PL_Pr = 1
                End Select
    
                ' P�iv�laput
                If sht_ryhm�t.Cells(ryh.Row, rs_p�iv�laput.Column).Value <> "Ei" Then
                    Call Ryhm�_P�iv�laput(sht_ryhm�t.Cells(ryh.Row, rs_ryhm�nnimi.Column).Value, _
                                          sht_ryhm�t.Cells(ryh.Row, rs_plabc.Column).Value, Ryhm�numero, _
                                          spaceremove(sht_ryhm�t.Cells(ryh.Row, rs_listayhdistys.Column).Value), _
                                          sht_ryhm�t.Cells(ryh.Row, rs_yhdistettynimi.Column).Value, Code_ryhm�t, _
                                          sht_ryhm�t.Cells(ryh.Row, rs_p�ivystys.Column).Value, _
                                          sht_ryhm�t.Cells(ryh.Row, rs_p�iv�laput.Column).Value, ryh.Row, PL_Yhd, _
                                          sht_ryhm�t.Cells(ryh.Row, rs_plpohja.Column).Value, VL_Yhd, PL_Poissa, sht_ryhm�t.Cells(ryh.Row, rs_pltyhj�t.Column).Value)
                End If
                Ryhm�numero = Ryhm�numero + 1
                ' Onko k�yt�ss�
            End If
        Next ryh
    
        ' Tehd��n koontilappu
        If Ruokakoonti Then Call KoontiTulostus(Code_ryhm�t)
    
    End If


    ' Aamu- ja iltalistat
    Call PoistaSuodatukset
    Dim aamuiltalistat As String: aamuiltalistat = sht_p�iv�koti.Range("I5").Value
    If aamuiltalistat = sht_code.Range("A42").Value Then Call Aamu_ilta_listat

    ' Tapiolan viikonloppulaput
    Dim tapiolavkl As String: tapiolavkl = sht_p�iv�koti.Range("I8").Value
    If tapiolavkl = sht_code.Range("A42").Value Then Call Tapiolan_Viikonloppulaput

    If ThisWorkbook.Worksheets("Code").Range("F2").Value = 1 Then Call PoistaTiedot

    ' Lis�t��n codeen tehdyt ryhm�t (jos ryhmi� edes yksi)
    If (Not Not Code_ryhm�t) <> 0 Then
        sht_code.Range("B2").Resize(UBound(Code_ryhm�t) + 1).Value = Application.Transpose(Code_ryhm�t)
    End If
    
    With tbl_lapset.Sort
        ' Sorttaus: J�rjestyksen mukaan, k��nteisesti
        .SortFields.Clear
        .SortFields.Add Key:=sort_lapset_ryhm�, SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=sort_lapset_aakkoset, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    ' Loppus��d�t
    Call Lopetus("Hoitolistat on tehty.", vbInformation, "Valmista")

End Sub

Sub AdvancedFilter()

    sht_lapset.Range("tbl_lapset").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=sht_code.Range("A100:R105")

End Sub

Sub KoontiTulostus(ByRef Code_ryhm�t() As Variant)

    Sheets("Ruokakoonti_pohja").Copy After:=Sheets("Lapset")
    With ActiveSheet
        .Name = "Ruokakoonti"
        .Range("C4:G10") = koontiruoat()
    End With

    ' Napataan ryhmien nimet talteen, jotta voidaan poistaa ne ensi kerralla
    ' Jos Code_Ryhm�t Array on tyhj�, muuta kooksi 0 ja lis�� ryhm�n nimi
    ' Muutoin lis�� yksi merkint� + ryhm�n nimi
    Dim arrayIsNothing As Boolean
    On Error Resume Next
    arrayIsNothing = IsNumeric(UBound(Code_ryhm�t)) And False
    If Err.Number <> 0 Then arrayIsNothing = True
    On Error GoTo 0

    ' Arrayyn ryhm�n nimi
    If arrayIsNothing Then
        ReDim Code_ryhm�t(0)
        Code_ryhm�t(0) = "Ruokakoonti"
    Else
        ReDim Preserve Code_ryhm�t(UBound(Code_ryhm�t) + 1)
        Code_ryhm�t(UBound(Code_ryhm�t)) = "Ruokakoonti"
    End If

    ' Ryhm�n lis�ys Codeen
    sht_code.Range("B2").Resize(UBound(Code_ryhm�t) + 1).Value = Application.Transpose(Code_ryhm�t)


End Sub

Sub Ryhm�_Ruoka(Pohja As String, Ryhm�nimi As String, Aakkosj�rjestys As String, Ryhm�numero As Long, R_Tulostus As String, R_Asetukset As String, R_Yhdistely As String, _
                ByRef Code_ryhm�t() As Variant, R_p�ivystys As String, ryhrow As Long)

    ' Napataan ryhmien nimet talteen, jotta voidaan poistaa ne ensi kerralla
    ' Jos Code_Ryhm�t Array on tyhj�, muuta kooksi 0 ja lis�� ryhm�n nimi
    ' Muutoin lis�� yksi merkint� + ryhm�n nimi
    Dim arrayIsNothing As Boolean
    On Error Resume Next
    arrayIsNothing = IsNumeric(UBound(Code_ryhm�t)) And False
    If Err.Number <> 0 Then arrayIsNothing = True
    On Error GoTo 0

    If arrayIsNothing Then
        ReDim Code_ryhm�t(0)
        Code_ryhm�t(0) = Ryhm�nimi & "_ruoka"
    Else
        ReDim Preserve Code_ryhm�t(UBound(Code_ryhm�t) + 1)
        Code_ryhm�t(UBound(Code_ryhm�t)) = Ryhm�nimi & "_ruoka"
    End If

    ' p�iv�m��rien lis��minen
    Dim pvm_vuosi As Long
    Dim pvm_kk As Long
    Dim pvm_pv As Long
    Dim pvm As Date

    ' Haetaan pvm Codesta ja muunnetaan sopivaan muotoon
    pvm_vuosi = Year(Now)
    pvm_pv = sht_code.[C2].Value2
    pvm_kk = sht_code.[C3].Value2

    ' P�iv�m��r�n koonti
    pvm = DateSerial(pvm_vuosi, pvm_kk, pvm_pv)

    ' Suodata pois tyhj�t, n�yt� vain erityisruokavaliot
    ' jotta voidaan laskea dieetit
    Dim sort_lapset_dieetti As Range
    Set sort_lapset_dieetti = ls_dieetti


    ' ryhmien yhdistelemisen tsekkaus
    With tbl_lapset
        If R_p�ivystys = "Kyll�" Then
            tbl_lapset.Range.AutoFilter Field:=ls_p�ivystys.Column, Criteria1:=Ryhm�nimi, Operator:=xlFilterValues
        End If
        ' Tsekataan yhdistetyt ruokalaput
        If R_Yhdistely = vbNullString Then
            ' Jos ei yhdistet� mit��n, suodatetaan ryhm�n nimen mukaan
            If R_p�ivystys = "Ei" Then
                .Range.AutoFilter Field:=ls_ryhm�.Column, Criteria1:=Ryhm�nimi, Operator:=xlFilterValues
            End If
        Else
            ' Lis�t��n Ryhm�nimi yhdistelyyn
            R_Yhdistely = Ryhm�nimi + ", " + R_Yhdistely
            ' Tehd��n Array ruokalistan yhdistelyryhmist�
            Dim r_yhdistely_array() As String
            r_yhdistely_array = Split(R_Yhdistely, ",", , vbTextCompare)
            ' Suodatetaan arrayn ryhmien mukaan
            tbl_lapset.Range.AutoFilter Field:=ls_ryhm�.Column, Criteria1:=r_yhdistely_array, Operator:=xlFilterValues
        End If
     
        ' Jos kyseess� p�ivystysryhm�, pidet��n vain p�ivystyslapset
        ' Vain erikoisruokavaliot
        .Range.AutoFilter Field:=ls_dieetti.Column, Criteria1:="<>", Operator:=xlFilterValues
        ' Ei ty�ntekij�it�
        '.Range.AutoFilter Field:=ls_ty�ntekij�.Column, Criteria1:="=", Operator:=xlFilterValues
     
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=sort_lapset_dieetti, SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With

    End With

    ' Ryhm�n lis�ys Codeen
    sht_code.Range("B2").Resize(UBound(Code_ryhm�t) + 1).Value = Application.Transpose(Code_ryhm�t)

    Dim pohjaNimi As String
    ' Lasketaan pohjan parametrit
    Select Case Pohja
    Case "Ma-pe"
        P�iv�t = 5
        Ruokailut = 3
        pohjaNimi = "R_arki"
    Case "Ma-pe, vain ilta"
        P�iv�t = 5
        Ruokailut = 2
        pohjaNimi = "R_ilta"
    Case "Ma-pe + ilta"
        P�iv�t = 5
        Ruokailut = 5
        pohjaNimi = "R_arkikaikki"
    Case "Ma-su + ilta"
        P�iv�t = 7
        Ruokailut = 5
        pohjaNimi = "R_kaikki"
    End Select

    Dim isoFontti As Boolean: isoFontti = False

    Dim p�iv�Poissa As Boolean
    Dim iltaPoissa As Boolean
    Dim vklPoissa As Boolean

    Dim matulorng As Range

    Dim pv As Long
    Dim lnimi As Range

    Dim a1 As String: a1 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_aamupala1.Column).Value2
    Dim a2 As String: a2 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_aamupala2.Column).Value2
    Dim l1 As String: l1 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_lounas1.Column).Value2
    Dim l2 As String: l2 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_lounas2.Column).Value2
    Dim v1 As String: v1 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_v�lipala1.Column).Value2
    Dim v2 As String: v2 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_v�lipala2.Column).Value2
    Dim p1 As String: p1 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_p�iv�llinen1.Column).Value2
    Dim p2 As String: p2 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_p�iv�llinen1.Column).Value2
    Dim i1 As String: i1 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_iltapala1.Column).Value2
    Dim i2 As String: i2 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_iltapala2.Column).Value2

    Dim maTulo As Range
    Dim maL�ht� As Range
    Dim ruokaP�iv� As Long
    Dim tulo As String
    Dim l�ht� As String
    Dim ruokaRivi As Long: ruokaRivi = 4
    Dim tuloArr() As String
    Dim l�ht�Arr() As String
    Dim yhdArr() As String
    Dim kerta As Long
    Dim tyhjienLasku As Long

    ' Poistetaan dieettilistalta lapset, jotka eiv�t ole ruokailuissa mukana
    If tyhjasuodatin(rng_lapset) = False Then
        For Each lnimi In rng_lapset.SpecialCells(xlCellTypeVisible)
            tyhjienLasku = 0
            'Jos kyseess� ei ole ty�ntekij�
            If sht_lapset.Cells(lnimi.Row, ls_ty�ntekij�.Column).Value2 = "" Then
                ' Oletuksena lapset ovat pois, kunnes toisin hoitoajoista todetaan
                vklPoissa = True
                iltaPoissa = True
                p�iv�Poissa = True
                ' K�yd��n l�pi arkip�iv�t
                For pv = 0 To 4
                    Set matulorng = sht_lapset.Cells(lnimi.Row, ls_matulo.Column + 2 * pv)
                    ' Yksitt�inen hoitoaika
                    If IsDate(Format(sht_lapset.Cells(lnimi.Row, ls_matulo.Column + 2 * pv), "h:mm")) = True Then
                        ' Aamupala
                        If matulorng <= a2 And matulorng.Offset(, 1) >= a1 Then p�iv�Poissa = False
                        ' Lounas
                        If matulorng <= l2 And matulorng.Offset(, 1) >= l1 Then p�iv�Poissa = False
                        ' V�lipala
                        If matulorng <= v2 And matulorng.Offset(, 1) >= v1 Then p�iv�Poissa = False
                        ' P�iv�llinen
                        If matulorng <= p2 And matulorng.Offset(, 1) >= p1 Then iltaPoissa = False
                        ' Iltapala
                        If matulorng <= i2 And matulorng.Offset(, 1) >= i1 Then iltaPoissa = False
                
                        ' Useampi hoitoaika
                    ElseIf InStr(matulorng, ",") > 0 Then
                        ' Tehd��n hoitoajoista arrayt
                        tuloArr = Split(matulorng, ",")
                        l�ht�Arr = Split(matulorng.Offset(, 1), ",")
                        ' K�yd��n array l�pi
                        For kerta = 0 To ArrayLen(tuloArr) - 1
                            ' Aamupala
                            If tuloArr(kerta) <= CDate(a2) And l�ht�Arr(kerta) >= CDate(a1) Then p�iv�Poissa = False
                            ' Lounas
                            If tuloArr(kerta) <= CDate(l2) And l�ht�Arr(kerta) >= CDate(l1) Then p�iv�Poissa = False
                            ' V�lipala
                            If tuloArr(kerta) <= CDate(v2) And l�ht�Arr(kerta) >= CDate(v1) Then p�iv�Poissa = False
                            ' P�iv�llinen
                            If tuloArr(kerta) <= CDate(p2) And l�ht�Arr(kerta) >= CDate(p1) Then iltaPoissa = False
                            ' Iltapala
                            If tuloArr(kerta) <= CDate(i2) And l�ht�Arr(kerta) >= CDate(i1) Then iltaPoissa = False
                        Next kerta
                        ' Ei hoitoaikaa tai kirjain
                    Else
                        If matulorng = "" Then tyhjienLasku = tyhjienLasku + 1
               
                    End If
                Next pv
                For pv = 5 To 6
                    Set matulorng = sht_lapset.Cells(lnimi.Row, ls_matulo.Column + 2 * pv)
                    ' Yksitt�inen hoitoaika
                    If IsDate(Format(sht_lapset.Cells(lnimi.Row, ls_matulo.Column + 2 * pv), "h:mm")) = True Then
                        ' Aamupala
                        If matulorng <= a2 And matulorng.Offset(, 1) >= a1 Then vklPoissa = False
                        ' Lounas
                        If matulorng <= l2 And matulorng.Offset(, 1) >= l1 Then vklPoissa = False
                        ' V�lipala
                        If matulorng <= v2 And matulorng.Offset(, 1) >= v1 Then vklPoissa = False
                        ' P�iv�llinen
                        If matulorng <= p2 And matulorng.Offset(, 1) >= p1 Then vklPoissa = False
                        ' Iltapala
                        If matulorng <= i2 And matulorng.Offset(, 1) >= i1 Then vklPoissa = False
                
                        ' Useampi hoitoaika
                    ElseIf InStr(matulorng, ",") > 0 Then
                        ' Tehd��n hoitoajoista arrayt
                        tuloArr = Split(matulorng, ",")
                        l�ht�Arr = Split(matulorng.Offset(, 1), ",")
                        ' K�yd��n array l�pi
                        For kerta = 0 To ArrayLen(tuloArr) - 1
                            ' Aamupala
                            If tuloArr(kerta) <= CDate(a2) And l�ht�Arr(kerta) >= CDate(a1) Then vklPoissa = False
                            ' Lounas
                            If tuloArr(kerta) <= CDate(l2) And l�ht�Arr(kerta) >= CDate(l1) Then vklPoissa = False
                            ' V�lipala
                            If tuloArr(kerta) <= CDate(v2) And l�ht�Arr(kerta) >= CDate(v1) Then vklPoissa = False
                            ' P�iv�llinen
                            If tuloArr(kerta) <= CDate(p2) And l�ht�Arr(kerta) >= CDate(p1) Then vklPoissa = False
                            ' Iltapala
                            If tuloArr(kerta) <= CDate(i2) And l�ht�Arr(kerta) >= CDate(i1) Then vklPoissa = False
                        Next kerta
                    End If
                Next pv
                sht_lapset.Cells(lnimi.Row, ls_arkiPoissa.Column) = ""
                sht_lapset.Cells(lnimi.Row, ls_vklPoissa.Column) = ""
                If p�iv�Poissa And iltaPoissa Then sht_lapset.Cells(lnimi.Row, ls_arkiPoissa.Column) = "Ei sy�"
                If p�iv�Poissa And iltaPoissa = False Then sht_lapset.Cells(lnimi.Row, ls_arkiPoissa.Column) = "P�iv�"
                If p�iv�Poissa = False And iltaPoissa Then sht_lapset.Cells(lnimi.Row, ls_arkiPoissa.Column) = "Ilta"
                If vklPoissa Then sht_lapset.Cells(lnimi.Row, ls_vklPoissa.Column) = "Ei sy�"
                ' Jos ma-pe = "", n�yt� listalla jos ei erikseen estetty
                If tyhjienLasku = 5 And (R_Asetukset = "Pieni fontti - poissa" Or R_Asetukset = "Iso fontti - poissa") Then
                    sht_lapset.Cells(lnimi.Row, ls_arkiPoissa.Column) = ""
                End If
            End If
        Next lnimi
    End If

    ' Suodatetaan ne dieettilapset pois, jotka eiv�t ole ruokailuissa mukana
    If R_Asetukset = "Pieni fontti - poissa" Or R_Asetukset = "Pieni fontti - poissa - tyhj�t" Or _
       R_Asetukset = "Iso fontti - poissa" Or R_Asetukset = "Iso fontti - poissa - tyhj�t" Then
        With tbl_lapset
            ' P�iv�
            If Ruokailut = 3 Then .Range.AutoFilter Field:=ls_arkiPoissa.Column, Criteria1:="<>P�iv�", Criteria2:="<>Ei sy�", Operator:=xlFilterValues
            ' Ilta
            If Ruokailut = 2 Then .Range.AutoFilter Field:=ls_arkiPoissa.Column, Criteria1:="<>Ilta", Criteria2:="<>Ei sy�", Operator:=xlFilterValues
            ' P�iv� & ilta (ei sy�)
            If Ruokailut = 5 Then .Range.AutoFilter Field:=ls_arkiPoissa.Column, Criteria1:="<>Ei sy�", Operator:=xlFilterValues
            ' Viikonloput (ei sy�)
            If P�iv�t = 7 Then .Range.AutoFilter Field:=ls_vklPoissa.Column, Criteria1:="<>Ei sy�", Operator:=xlFilterValues
        End With
    End If

    If R_Asetukset = "Iso fontti" Or R_Asetukset = "Iso fontti - poissa" Or _
       R_Asetukset = "Iso fontti - poissa - tyhj�t" Then isoFontti = True
   
    'sht_lapset.Range.AutoFilter Field:=ls_arkiPoissa.Column, Criteria1:="", Operator:=xlFilterValues
    'sht_lapset.Range.AutoFilter Field:=ls_vklPoissa.Column, Criteria1:="", Operator:=xlFilterValues



    Sarakkeet = 30 / Ruokailut
    If pohjaNimi = "R_ilta" Then Sarakkeet = 10

    ' Montako lasta mahtuu pohjalle. 1. lapulle, 2. lapulle ja molemmille lapuille yhteens�
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

    ' Selvitet��n montako lasta on mill�kin dieettirivill�
    If Dieettilapset <> 0 Then
        If Dieettilapset >= Sarakkeet - 1 Then
            ekaDieettiRivi = Sarakkeet - 1
            If Dieettilapset >= ekaDieettiRivi + 2 * Sarakkeet Then
                tokaDieettiRivi = Sarakkeet
                If Dieettilapset >= ekaDieettiRivi + tokaDieettiRivi + 2 * Sarakkeet Then
                    kolDieettiRivi = Sarakkeet
                    If Dieettilapset > ekaDieettiRivi + tokaDieettiRivi + kolDieettiRivi + 2 * Sarakkeet Then
                        Call Lopetus("O-ou! Ryhm�ss� " & Ryhm�nimi & " on liikaa erityisruokavalioita. Ohjelma tukee max " & maxDieetit & " erityisruokavaliota t�ll� " & Pohja & " -pohjalla." _
                                   & vbCrLf & "Laita minulle viesti� jos tarpeesi on suurempi, niin katsotaan mit� voin tehd� asialle :)" & vbCrLf & "jaakko.haavisto@jyvaskyla.fi", _
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
        .Name = Ryhm�nimi & "_ruoka"
        If R_Yhdistely = vbNullString Then
            .Range("F1").Value = UCase(Ryhm�nimi)
        Else
            .Range("F1").Value = UCase(Join(r_yhdistely_array, ", "))
        End If
        .Range("A1").Value = pvm
        .Range("D1").Value = DateAdd("d", 4, CDate(pvm))
    End With

    Dim Ruokaryhm� As Worksheet: Set Ruokaryhm� = wb.Worksheets(Ryhm�nimi & "_ruoka")

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
    With Ruokaryhm�
        If Dieettilapset <> 0 Then
            For Each rng_dieetti In rng_lapset.SpecialCells(xlCellTypeVisible)
                ' rivi 1
                If Lapsinumero <= rivi1 Then
                    DieettiRivi = 2
                    DieettiSarake = 3 + Ruokailut + Ruokailut * (Lapsinumero - 1)
                
                    ' rivi 2
                ElseIf rivi1 < Lapsinumero And Lapsinumero <= rivi2 Then
                    DieettiRivi = 2 + 2 + P�iv�t
                    DieettiSarake = 3 + Ruokailut * Lapsinumero - Ruokailut * Sarakkeet
           
                    ' rivi 3
                ElseIf rivi2 < Lapsinumero And Lapsinumero <= rivi3 Then
                    DieettiRivi = 23
                    DieettiSarake = 3 + Ruokailut * Lapsinumero - 2 * (Ruokailut * Sarakkeet)
           
                    ' rivi 4
                ElseIf rivi3 < Lapsinumero And Lapsinumero <= rivi4 Then
                    DieettiRivi = 23 + 2 + P�iv�t
                    DieettiSarake = 3 + Ruokailut * Lapsinumero - 3 * (Ruokailut * Sarakkeet)
                End If
            
                ' Nimi
                .Cells(DieettiRivi, DieettiSarake) = UCase(rng_dieetti.Value)
                ' Ruksit
                Call Dieetti(rng_dieetti.Offset(0, 1).Value, Ruokaryhm�, Ryhm�nimi, P�iv�t, Ruokailut, 2 + DieettiRivi, DieettiSarake)
            
                Lapsinumero = Lapsinumero + 1
            Next rng_dieetti
        End If

    End With
    ' Nollaa suodatukset
    tbl_lapset.AutoFilter.ShowAllData

    With tbl_lapset
    
        If R_p�ivystys = "Kyll�" Then
            tbl_lapset.Range.AutoFilter Field:=ls_p�ivystys.Column, Criteria1:=Ryhm�nimi, Operator:=xlFilterValues
        End If
    
        ' Nollaus
        ' Sorttaus valinnan mukaan
        If R_Yhdistely = vbNullString Then
            ' Jos ei yhdistet� mit��n, suodatetaan ryhm�n nimen mukaan
            If R_p�ivystys = "Ei" Then
                tbl_lapset.Range.AutoFilter Field:=ls_ryhm�.Column, Criteria1:=Ryhm�nimi, Operator:=xlFilterValues
            End If
       
        Else
            tbl_lapset.Range.AutoFilter Field:=ls_ryhm�.Column, Criteria1:=r_yhdistely_array, Operator:=xlFilterValues
        
        End If
    End With

    Dim kaikkiHoitoajat As Boolean: kaikkiHoitoajat = True
    ' Ruokien laskeminen
    With Ruokaryhm�
        For Each lnimi In rng_lapset.SpecialCells(xlCellTypeVisible)
            ' Nollataan tyhj�t
            tyhjienLasku = 0
            Set maTulo = sht_lapset.Cells(lnimi.Row, ls_matulo.Column)
            Set maL�ht� = sht_lapset.Cells(lnimi.Row, ls_matulo.Column).Offset(, 1)
            ' arkip�iv�t:  5 pv -> 8
            ' koko viikko: 7 pv -> 12
            For ruokaP�iv� = 0 To (2 * P�iv�t) - 2 Step 2
                ' Yksi hoitoaika
                If IsDate(Format(maTulo.Offset(, ruokaP�iv�).Value2, "h:mm")) = True Then
                    tulo = maTulo.Offset(, ruokaP�iv�).Value2
                    l�ht� = maL�ht�.Offset(, ruokaP�iv�).Value2
                    If Ruokailut = 2 Then        ' ilta
                        If tulo <= p2 And l�ht� >= p1 Then .Cells(ruokaRivi, 3) = .Cells(ruokaRivi, 3) + 1
                        If tulo <= i2 And l�ht� >= i1 Then .Cells(ruokaRivi, 4) = .Cells(ruokaRivi, 4) + 1
                    End If
                    If Ruokailut = 3 Then        ' p�iv�
                        If tulo <= a2 And l�ht� >= a1 Then .Cells(ruokaRivi, 3) = .Cells(ruokaRivi, 3) + 1
                        If tulo <= l2 And l�ht� >= l1 Then .Cells(ruokaRivi, 4) = .Cells(ruokaRivi, 4) + 1
                        If tulo <= v2 And l�ht� >= v1 Then .Cells(ruokaRivi, 5) = .Cells(ruokaRivi, 5) + 1
                    End If
                    If Ruokailut = 5 Then        ' kaikki ruokailut
                        If tulo <= a2 And l�ht� >= a1 Then .Cells(ruokaRivi, 3) = .Cells(ruokaRivi, 3) + 1
                        If tulo <= l2 And l�ht� >= l1 Then .Cells(ruokaRivi, 4) = .Cells(ruokaRivi, 4) + 1
                        If tulo <= v2 And l�ht� >= v1 Then .Cells(ruokaRivi, 5) = .Cells(ruokaRivi, 5) + 1
                        If tulo <= p2 And l�ht� >= p1 Then .Cells(ruokaRivi, 6) = .Cells(ruokaRivi, 6) + 1
                        If tulo <= i2 And l�ht� >= i1 Then .Cells(ruokaRivi, 7) = .Cells(ruokaRivi, 7) + 1
                    End If
                
                    ' Useampi hoitoaika
                ElseIf InStr(maTulo.Offset(, ruokaP�iv�), ",") > 0 Then
                    ' Tehd��n hoitoajoista arrayt
                    tuloArr = Split(maTulo.Offset(, ruokaP�iv�), ",")
                    l�ht�Arr = Split(maL�ht�.Offset(, ruokaP�iv�), ",")
                    ' K�yd��n array l�pi
                    For kerta = 0 To ArrayLen(tuloArr) - 1
                        If Ruokailut = 2 Then    ' ilta
                            If tuloArr(kerta) <= CDate(p2) And l�ht�Arr(kerta) >= CDate(p1) Then .Cells(ruokaRivi, 3) = .Cells(ruokaRivi, 3) + 1
                            If tuloArr(kerta) <= CDate(i2) And l�ht�Arr(kerta) >= CDate(i1) Then .Cells(ruokaRivi, 4) = .Cells(ruokaRivi, 4) + 1
                        End If
                        If Ruokailut = 3 Then    ' p�iv�
                            If tuloArr(kerta) <= CDate(a2) And l�ht�Arr(kerta) >= CDate(a1) Then .Cells(ruokaRivi, 3) = .Cells(ruokaRivi, 3) + 1
                            If tuloArr(kerta) <= CDate(l2) And l�ht�Arr(kerta) >= CDate(l1) Then .Cells(ruokaRivi, 4) = .Cells(ruokaRivi, 4) + 1
                            If tuloArr(kerta) <= CDate(v2) And l�ht�Arr(kerta) >= CDate(v1) Then .Cells(ruokaRivi, 5) = .Cells(ruokaRivi, 5) + 1
                        End If
                        If Ruokailut = 5 Then    ' kaikki ruokailut
                            If tuloArr(kerta) <= CDate(a2) And l�ht�Arr(kerta) >= CDate(a1) Then .Cells(ruokaRivi, 3) = .Cells(ruokaRivi, 3) + 1
                            If tuloArr(kerta) <= CDate(l2) And l�ht�Arr(kerta) >= CDate(l1) Then .Cells(ruokaRivi, 4) = .Cells(ruokaRivi, 4) + 1
                            If tuloArr(kerta) <= CDate(v2) And l�ht�Arr(kerta) >= CDate(v1) Then .Cells(ruokaRivi, 5) = .Cells(ruokaRivi, 5) + 1
                            If tuloArr(kerta) <= CDate(p2) And l�ht�Arr(kerta) >= CDate(p1) Then .Cells(ruokaRivi, 6) = .Cells(ruokaRivi, 6) + 1
                            If tuloArr(kerta) <= CDate(i2) And l�ht�Arr(kerta) >= CDate(i1) Then .Cells(ruokaRivi, 7) = .Cells(ruokaRivi, 7) + 1
                        End If
                    
                    Next kerta
                Else
                    ' Lasketaan tyhji�. Jos ei yht�k��n tyhji�,
                    If maTulo.Offset(, ruokaP�iv�).Value2 = "" Then
                        tyhjienLasku = tyhjienLasku + 1
                        If tyhjienLasku = P�iv�t Or (tyhjienLasku = 5 And ruokaP�iv� = 8) Then kaikkiHoitoajat = False
                    End If
                End If
                ruokaRivi = ruokaRivi + 1
            Next ruokaP�iv�
            ruokaRivi = 4
        Next lnimi
        ' K�ytet��n isoa fonttia jos asetuksissa valittu
        If isoFontti Or kaikkiHoitoajat Then .Range(.Cells(4, 3), .Cells(3 + P�iv�t, 3 + Ruokailut - 1)).Style = "IsoFontti"
    
        ' Jos ei mene 2 sivulle, poistetaan 2. sivun sis�lt� ja asetetaan tulostumaan vain 1. sivu
        If sivu = 1 Then
            .Range(Cells(2 * P�iv�t + 6, 1), Cells(40, 1)).EntireRow.Delete
            If Dieettilapset <= rivi1 Then
                .PageSetup.PrintArea = .Range(Cells(1, 1), Cells(P�iv�t + 3, 2 + Sarakkeet * Ruokailut)).Address
                .Range(Cells(P�iv�t + 4, 1), Cells(40, 1)).EntireRow.Delete
            Else
                .PageSetup.PrintArea = .Range(Cells(1, 1), Cells(2 * P�iv�t + 5, 2 + Sarakkeet * Ruokailut)).Address
            End If
        Else
            If Dieettilapset <= rivi3 Then
                .Range(Cells(23 + 2 + P�iv�t, 1), Cells(40, 1)).EntireRow.Delete
            End If
        End If
        ' Koonti
        If Ruokakoonti Then
            
            Dim k_p�iv� As Integer, k_ruokailu As Integer, iltalis� As Integer: iltalis� = 0
            ' Iltaruokien kohdalla siirret��n focusta iltalis�n verran
            If Ruokailut = 2 Then iltalis� = 3
            
            For k_p�iv� = 0 To P�iv�t - 1
                For k_ruokailu = 0 To Ruokailut - 1
                    koontiruoat(k_p�iv�, k_ruokailu + iltalis�) = koontiruoat(k_p�iv�, k_ruokailu + iltalis�) + .Cells(4 + k_p�iv�, 3 + k_ruokailu)
                Next k_ruokailu
            Next k_p�iv�
        End If
    End With
End Sub

Sub Ryhm�_Lista(Ryhm�nimi As String, Aakkosj�rjestys As String, Ryhm�numero As Long, VL_Tulostus As String, _
                VL_Yhdistely As String, VL_Nimi As String, ByRef Code_ryhm�t() As Variant, VL_p�ivystys As String, _
                ryhrow As Long, VL_Yhd As Long, VL_Vkl As Long)

    ' Kloonataan pohja
    Sheets("Pohja").Copy After:=Sheets("Lapset")

    Dim Ryhm�sivu As Worksheet
    Dim Nimi As Worksheet

    ' Napataan ryhmien nimet talteen, jotta voidaan poistaa ne ensi kerralla
    ' Jos Code_Ryhm�t Array on tyhj�, muuta kooksi 0 ja lis�� ryhm�n nimi
    ' Muutoin lis�� yksi merkint� + ryhm�n nimi
    Dim arrayIsNothing As Boolean
    On Error Resume Next
    arrayIsNothing = IsNumeric(UBound(Code_ryhm�t)) And False
    If Err.Number <> 0 Then arrayIsNothing = True
    On Error GoTo 0

    ' Jos Yhdistettyj� ryhmi�, k�yt� niiden omaa nime�. Muutoin Ryhm�n nime�
    Dim oikeanimi As String
    If VL_Yhd = 0 Or VL_Yhdistely = vbNullString Then
        ActiveSheet.Name = Ryhm�nimi
        Set Nimi = wb.Worksheets(Ryhm�nimi)
        Set Ryhm�sivu = wb.Worksheets(Ryhm�nimi)
        oikeanimi = Ryhm�nimi
        ' Arrayyn ryhm�n nimi
        If arrayIsNothing Then
            ReDim Code_ryhm�t(0)
            Code_ryhm�t(0) = Ryhm�nimi
        Else
            ReDim Preserve Code_ryhm�t(UBound(Code_ryhm�t) + 1)
            Code_ryhm�t(UBound(Code_ryhm�t)) = Ryhm�nimi
        End If
        ' Ryhm�n nimi capseilla yl�laitaan
        Ryhm�sivu.Range("A1").Value = UCase(Ryhm�nimi)
    ElseIf VL_Yhd = 1 Then
        ActiveSheet.Name = VL_Nimi
        Set Nimi = wb.Worksheets(VL_Nimi)
        Set Ryhm�sivu = wb.Worksheets(VL_Nimi)
        oikeanimi = VL_Nimi
        ' Arrayyn yhdistetyn ryhm�n nimi
        If arrayIsNothing Then
            ReDim Code_ryhm�t(0)
            Code_ryhm�t(0) = VL_Nimi
        Else
            ReDim Preserve Code_ryhm�t(UBound(Code_ryhm�t) + 1)
            Code_ryhm�t(UBound(Code_ryhm�t)) = VL_Nimi
        End If
        ' Ryhm�n nimi capseilla yl�laitaan
        Ryhm�sivu.Range("A1").Value = UCase(VL_Nimi)
    End If

    ' Ryhm�n lis�ys Codeen
    sht_code.Range("B2").Resize(UBound(Code_ryhm�t) + 1).Value = Application.Transpose(Code_ryhm�t)

    ' p�iv�m��rien lis��minen
    Dim pvm_vuosi As Long
    Dim pvm_kk As Long
    Dim pvm_pv As Long
    Dim pvm As Date

    ' Haetaan pvm Codesta ja muunnetaan sopivaan muotoon
    pvm_vuosi = Year(Now)
    pvm_pv = sht_code.[C2].Value2
    pvm_kk = sht_code.[C3].Value2

    ' P�iv�m��r�n koonti
    pvm = DateSerial(pvm_vuosi, pvm_kk, pvm_pv)

    ' Lis�t��n pvm:t oikeille p�iville
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

        ' Jos kyseess� p�ivystysryhm�, pidet��n vain p�ivystyslapset
        If VL_p�ivystys = "Kyll�" Then
            tbl_lapset.Range.AutoFilter Field:=ls_p�ivystys.Column, Criteria1:=Ryhm�nimi, Operator:=xlFilterValues
        End If
        ' Ei ty�ntekij�it� listalle
        tbl_lapset.Range.AutoFilter Field:=ls_ty�ntekij�.Column, Criteria1:="=", Operator:=xlFilterValues

        ' Poistetaan koko viikon poissaolevat
        'If VL_Tulostus = "Kyll� - poissaolevat" Then
        '    tbl_lapset.Range.AutoFilter Field:=ls_poissaoleva.Column, Criteria1:="=", Operator:=xlFilterValues
        'End If

        ' Jos ei yhdistettyj� ryhmi�
        If VL_Yhdistely = vbNullString Then
            pituus = 1
            ' Jos ei yhdistet� mit��n, suodatetaan ryhm�n nimen mukaan
            If VL_p�ivystys = "Ei" Then
                .Range.AutoFilter Field:=ls_ryhm�.Column, Criteria1:=Ryhm�nimi, Operator:=xlFilterValues
            End If
            ' Yhdistettyj� ryhmi� l�ytyy
        Else
            ' Lis�t��n Ryhm�nimi yhdistelyyn
            VL_Yhdistely = Ryhm�nimi + ", " + VL_Yhdistely
            ' Tehd��n Array ruokalistan yhdistelyryhmist�
            Dim vl_yhdistely_array() As String
            vl_yhdistely_array = Split(VL_Yhdistely, ",", , vbTextCompare)
            pituus = arrayitems(vl_yhdistely_array)
            .Range.AutoFilter Field:=ls_ryhm�.Column, Criteria1:=vl_yhdistely_array, Operator:=xlFilterValues
        End If

    End With

    ' Lapsiryhm�n sorttaus
    Dim sort_lapset_abc As Range
    If (Aakkosj�rjestys = "Oma j�rjestys") Then
        Set sort_lapset_abc = Range("A1")
    ElseIf (Aakkosj�rjestys = "Kutsumanimi") Then
        Set sort_lapset_abc = Range("B1")
    ElseIf (Aakkosj�rjestys = "Sukunimi") Then
        Set sort_lapset_abc = Range("tbl_lapset[Sukunimi]")
    End If

    With tbl_lapset.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sort_lapset_abc, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    ' Aloitetaan numerointi ykk�sest�
    Dim Lapsinumero As Long: Lapsinumero = 1
    Dim Tj_numero As Long: Tj_numero = 1

    ' Luupataan lapset yksi kerrallaan l�pi (n�kyv�t solut suodatuksen ja filtterin j�lkeen)
    For Each laps In rng_lapset.SpecialCells(xlCellTypeVisible)

        'Aliohjelma Lapsi
        Call Lapsi(sht_lapset.Cells(laps.Row, ls_kutsumanimi.Column).Value, _
                   Lapsinumero, Ryhm�sivu, sht_ryhm�t.Cells(ryhrow, _
                                                            rs_boldaus1.Column).Value, _
                   sht_ryhm�t.Cells(ryhrow, rs_boldaus2.Column).Value, _
                   sht_ryhm�t.Cells(ryhrow, rs_pboldaus1.Column).Value, _
                   sht_ryhm�t.Cells(ryhrow, rs_pboldaus2.Column).Value, _
                   sht_lapset.Cells(laps.Row, ls_matulo.Column), VL_Vkl)

        ' Numerointiin lis�t��n aina yksi (Code-v�lilehte� varten, jotta ryhm�t voidaan my�hemmin poistaa)
        Lapsinumero = Lapsinumero + 1
    Next laps

End Sub

Sub Ryhm�_P�iv�laput(Ryhm�nimi As String, Aakkosj�rjestys As String, Ryhm�numero As Long, VL_Yhdistely As String, _
                     VL_Nimi As String, ByRef Code_ryhm�t() As Variant, VL_p�ivystys As String, VL_p�iv�laput As String, _
                     ryhrow As Long, PL_Yhd As Long, PL_Pohja As String, VL_Yhd As Long, PL_Poissa As Long, _
                     pl_tyhj�t As Integer)

    ' TODO
    ' - poissaolevat
    ' + pienryhm�t
    
    
    Dim Ryhm�sivu As Worksheet
    Dim Nimi As Worksheet

    ' Napataan ryhmien nimet talteen, jotta voidaan poistaa ne ensi kerralla
    ' Jos Code_Ryhm�t Array on tyhj�, muuta kooksi 0 ja lis�� ryhm�n nimi
    ' Muutoin lis�� yksi merkint� + ryhm�n nimi
    Dim arrayIsNothing As Boolean
    On Error Resume Next
    arrayIsNothing = IsNumeric(UBound(Code_ryhm�t)) And False
    If Err.Number <> 0 Then arrayIsNothing = True
    On Error GoTo 0

    ' p�iv�m��rien lis��minen
    Dim pvm_vuosi As Long
    Dim pvm_kk As Long
    Dim pvm_pv As Long
    Dim pvm As Date

    ' Haetaan pvm Codesta ja muunnetaan sopivaan muotoon
    pvm_vuosi = Year(Now)
    pvm_pv = sht_code.[C2].Value2
    pvm_kk = sht_code.[C3].Value2

    ' P�iv�m��r�n koonti
    pvm = DateSerial(pvm_vuosi, pvm_kk, pvm_pv)

    Dim i As Long

    Dim pituus As Long
    Dim listaus As Long: listaus = 1

    With tbl_lapset
        ' Jos kyseess� p�ivystysryhm�, pidet��n vain p�ivystyslapset
        If VL_p�ivystys = "Kyll�" Then
            .Range.AutoFilter Field:=ls_p�ivystys.Column, Criteria1:=Ryhm�nimi, Operator:=xlFilterValues
        End If
        
      
        ' Ei ty�ntekij�it� listalle
        .Range.AutoFilter Field:=ls_ty�ntekij�.Column, Criteria1:="=", Operator:=xlFilterValues

        ' Yhdistellyt ryhm�t
        If Not VL_Yhdistely = vbNullString And PL_Yhd > 0 Then
            ' Lis�t��n Ryhm�nimi yhdistelyyn
            VL_Yhdistely = Ryhm�nimi + ", " + VL_Yhdistely
            ' Tehd��n Array ruokalistan yhdistelyryhmist�
            Dim vl_yhdistely_array() As String
            vl_yhdistely_array = Split(VL_Yhdistely, ",", , vbTextCompare)
    
            ' V�lily�nnit pois arraysta
            Dim v�li As Long
            For v�li = LBound(vl_yhdistely_array) To UBound(vl_yhdistely_array)
                vl_yhdistely_array(v�li) = Trim(vl_yhdistely_array(v�li))
            Next
    
            pituus = arrayitems(vl_yhdistely_array)
            
            ' Sama lista
            If PL_Yhd = 1 Then .Range.AutoFilter Field:=ls_ryhm�.Column, Criteria1:=vl_yhdistely_array, Operator:=xlFilterValues
            ' 2-puoliset
            If PL_Yhd = 2 Then .Range.AutoFilter Field:=ls_ryhm�.Column, Criteria1:=vl_yhdistely_array(0), Operator:=xlFilterValues
            
            ' Vain 1 ryhm�
        Else
            pituus = 1
            ' Jos ei yhdistet� mit��n, suodatetaan ryhm�n nimen mukaan
            If VL_p�ivystys = "Ei" Then
                .Range.AutoFilter Field:=ls_ryhm�.Column, Criteria1:=Ryhm�nimi, Operator:=xlFilterValues
            End If
            
            ' Yhdistettyj� ryhmi� l�ytyy
        End If

    End With

    ' Lapsiryhm�n sorttaus
    Dim sort_lapset_abc As Range
    If (Aakkosj�rjestys = "Oma j�rjestys") Then
        Set sort_lapset_abc = Range("A1")
    ElseIf (Aakkosj�rjestys = "Kutsumanimi") Then
        Set sort_lapset_abc = Range("B1")
    ElseIf (Aakkosj�rjestys = "Sukunimi") Then
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
    ' Jos ei yhdistettyj� ryhmi�, k�ytet��n ryhm�n nime�
    ' Jos Yhdistettyj� ryhmi�, k�yt� niiden omaa nime�
    If PL_Yhd = 0 Then
        oikeanimi = Ryhm�nimi
        ' Lis�t��n arrayhyn (codeen)
        If arrayIsNothing Then
            ReDim Code_ryhm�t(0)
            Code_ryhm�t(0) = Ryhm�nimi
        Else
            ' Jos viimeisin array on ryhm�n nimi -> ei laiteta listalle (ei duplikaatteja)
            ' TODO: tsekkaus, ett� toimii my�s yhdistetyiss� ryhmiss�
            Dim viimeisinRyhm� As Integer
            viimeisinRyhm� = ArrayLen(Code_ryhm�t) - 1
            If Code_ryhm�t(viimeisinRyhm�) = oikeanimi Then
            Else
                ReDim Preserve Code_ryhm�t(UBound(Code_ryhm�t) + 1)
                Code_ryhm�t(UBound(Code_ryhm�t)) = oikeanimi
            End If
        End If
    Else
        oikeanimi = VL_Nimi
        ' Lis�t��n arrayhyn (codeen)
        If arrayIsNothing Then
            If PL_Yhd = 1 Then
                ReDim Code_ryhm�t(0)
                Code_ryhm�t(0) = oikeanimi
            End If
            If PL_Yhd = 2 Then
                ReDim Code_ryhm�t(1)
                Code_ryhm�t(0) = vl_yhdistely_array(0)
                Code_ryhm�t(1) = vl_yhdistely_array(1)
            End If
        Else
            If PL_Yhd = 1 Then
                ReDim Preserve Code_ryhm�t(UBound(Code_ryhm�t) + 1)
                Code_ryhm�t(UBound(Code_ryhm�t)) = oikeanimi
            End If
            If PL_Yhd = 2 Then
                ReDim Preserve Code_ryhm�t(UBound(Code_ryhm�t) + 1)
                Code_ryhm�t(UBound(Code_ryhm�t)) = vl_yhdistely_array(0)
                ReDim Preserve Code_ryhm�t(UBound(Code_ryhm�t) + 1)
                Code_ryhm�t(UBound(Code_ryhm�t)) = vl_yhdistely_array(1)
            End If
        End If
    End If
        
    ' Aloitetaan numerointi ykk�sest�
    Dim Lapsinumero As Long: Lapsinumero = 1
    Dim Tj_numero As Long: Tj_numero = 1
    Dim PL_Pohja2 As String, pl_tyhj�t2 As Integer
    
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
        PL_Pohja2 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(vl_yhdistely_array(1)).Row, rs_plpohja.Column).Value2
        pl_tyhj�t2 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(vl_yhdistely_array(1)).Row, rs_pltyhj�t.Column).Value2
    End If
    
    For Each vko In lappupv
        Call P�iv�lappu(PL_Pohja, PL_Yhd, vko, oikeanimi, Ryhm�nimi, pvm, pl_tyhj�t, PL_Poissa)
        
        ' 2-puoleiset
        If PL_Yhd = 2 Then
            ' Haetaan 2-puolen lapset
            tbl_lapset.Range.AutoFilter Field:=ls_ryhm�.Column, Criteria1:=vl_yhdistely_array(1), Operator:=xlFilterValues
            Call P�iv�lappu(PL_Pohja2, PL_Yhd, vko, "", vl_yhdistely_array(1), pvm, pl_tyhj�t2, PL_Poissa)
            ' Haetaan taas 1-puolen lapset
            If VL_p�ivystys = "Ei" Then
                tbl_lapset.Range.AutoFilter Field:=ls_ryhm�.Column, Criteria1:=Ryhm�nimi, Operator:=xlFilterValues
            Else
                tbl_lapset.Range.AutoFilter Field:=ls_p�ivystys.Column, Criteria1:=Ryhm�nimi, Operator:=xlFilterValues
            End If
            
        End If
    
    Next vko
    
End Sub

Sub P�iv�lappupohja_kopsuri()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' N�ytet��n v�lilehdet
    Dim WSname As Variant
    For Each WSname In Array("pl_pysty", "pl_pystywc", "pl_vaaka", "pl_vaakawc")
        Worksheets(WSname).Visible = True
    Next WSname

    Dim PL_Pohja As String, Kohde As String

    PL_Pohja = ThisWorkbook.Worksheets("P�iv�koti").Range("D8").Value2
    Kohde = ThisWorkbook.Worksheets("P�iv�koti").Range("E8").Value2
    
    If Kohde = "" Then
        MsgBox "Et ole kirjoittanut Ryhm�n nimeksi mit��n."
        'Sheets(Array("pl_pysty", "pl_pystywc", "pl_vaaka", "pl_vaakawc")).Select
        Call Lopetus_Simple
    End If
    
    ' Onko v�lilehti jo olemassa
    Dim i As Integer
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "pl_" & Kohde Then
            MsgBox "pl_" & Kohde & " -niminen v�lilehti l�ytyy jo."
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
    
    ' Piilotetaan v�lilehdet
    Sheets(Array("pl_pysty", "pl_pystywc", "pl_vaaka", "pl_vaakawc")).Select
    ActiveWindow.SelectedSheets.Visible = False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Uusi pohja lis�tty." & vbCrLf & "V�lilehden nimi on pl_" & Kohde


End Sub

Sub P�iv�lappu(PL_Pohja As String, PL_Yhd As Long, vko As Variant, oikeanimi As String, Ryhm�nimi As String, pvm As Date, pl_tyhj�t As Integer, PL_Poissa As Long)
    
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
        Sheets("pl_" & Ryhm�nimi).Copy After:=Sheets("pl_pysty")
    End Select
    
    Dim Ryhm�sivu As Worksheet
    If PL_Yhd = 1 Then
        ActiveSheet.Name = oikeanimi & "_" & vko
        Set Ryhm�sivu = wb.Worksheets(oikeanimi & "_" & vko)
    Else
        ActiveSheet.Name = Ryhm�nimi & "_" & vko
        Set Ryhm�sivu = wb.Worksheets(Ryhm�nimi & "_" & vko)
    End If
    
    Dim Ryhm�Paikka As Range, PvmPaikka As Range, NimiPaikka As Range, VikaSarakePaikka As Range, LoppuRivit As Range
    
    On Error Resume Next
        
    ' Napataan koodit pohjasta
    Dim pl_Ryhm�nimi As Range, pl_Pvm As Range, pl_Nimi As Range, pl_Hoitoaika As Range, pl_Alateksti As Range, pl_Vikasarake As Range, pl_Vikarivi As Range
    With Ryhm�sivu.Range("A1:O200")
        Set pl_Ryhm�nimi = .Find("pl-ryhm�nimi", MatchCase:=False)
        Set pl_Pvm = .Find("pl-pvm", MatchCase:=False)
        Set pl_Nimi = .Find("pl-nimi", MatchCase:=False)
        Set pl_Hoitoaika = .Find("pl-hoitoaika", MatchCase:=False)
        Set pl_Alateksti = .Find("pl-alateksti", MatchCase:=False)
        Set pl_Vikasarake = .Find("pl-vikasarake", MatchCase:=False)
        Set pl_Vikarivi = .Find("pl-vikarivi", MatchCase:=False)
    End With
    On Error GoTo 0
    ' Ryhm�n nimen lis�ys
    If PL_Yhd = 1 Then
        pl_Ryhm�nimi = UCase(oikeanimi)
    Else
        pl_Ryhm�nimi = UCase(Ryhm�nimi)
    End If
    ' P�iv�m��r�n lis�ys
    If vko = "ma" Then pl_Pvm = Format(pvm, "ddd d.m.")
    If vko = "ti" Then pl_Pvm = Format(DateAdd("d", 1, CDate(pvm)), "ddd d.m.")
    If vko = "ke" Then pl_Pvm = Format(DateAdd("d", 2, CDate(pvm)), "ddd d.m.")
    If vko = "to" Then pl_Pvm = Format(DateAdd("d", 3, CDate(pvm)), "ddd d.m.")
    If vko = "pe" Then pl_Pvm = Format(DateAdd("d", 4, CDate(pvm)), "ddd d.m.")
    
    ' K�yd��n ryhm�n lapset l�pi ja lis�t��n nimet ja hoitoajat
    Dim Lapsinumero As Long
    Lapsinumero = 0
    Dim p�iv�apuri As Long, p�iv�apuri2 As Long
    
    Rivi = 0
    ' Ryhm�sivu.Cells(Lapsinumero + 2, 1)
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
    With Ryhm�sivu
        Dim lapsim��r� As Integer: lapsim��r� = 0
        If tyhjasuodatin(rng_lapset) = False Then
            For Each laps In rng_lapset.SpecialCells(xlCellTypeVisible)
                tbl_lapset.Range.AutoFilter Field:=ls_petulo.Column
        
                ' Jos lapsirivit meinaavat menn� alatekstin p��lle, lis�t��n rivi ja otetaan yl�s sen koko.
                ' My�hemmin pienennet��n nimilistaa t�m�n koon perusteella
        
                If pl_Nimi.Row + Rivi = pl_Alateksti.Row Then
                    .Range("A" & pl_Nimi.Row + Rivi).EntireRow.Insert
                    pienennys = pienennys + .Range("A" & pl_Nimi.Row + Rivi).RowHeight
                End If

                ' Kopioidaan tyyli seuraavalle riville (ei ensimm�ist� rivi�)
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
        
                ' Lapsen nimen lis�ys
                'If Dieetti Then
                '    pl_Nimi.Offset(Rivi) = UCase(sht_lapset.Cells(laps.Row, ls_kutsumanimi.Column).Value) + " " + ChrW(&HD83C) & ChrW(&HDF74)
                'Else
                pl_Nimi.Offset(Rivi) = UCase(sht_lapset.Cells(laps.Row, ls_kutsumanimi.Column).Value)
        
                ' Hoitoajan lis�ys
                If vko = "ma" Then
                    p�iv�apuri = ls_matulo.Column
                    p�iv�apuri2 = ls_mal�ht�.Column
                ElseIf vko = "ti" Then
                    p�iv�apuri = ls_titulo.Column
                    p�iv�apuri2 = ls_til�ht�.Column
                ElseIf vko = "ke" Then
                    p�iv�apuri = ls_ketulo.Column
                    p�iv�apuri2 = ls_kel�ht�.Column
                ElseIf vko = "to" Then
                    p�iv�apuri = ls_totulo.Column
                    p�iv�apuri2 = ls_tol�ht�.Column
                ElseIf vko = "pe" Then
                    p�iv�apuri = ls_petulo.Column
                    p�iv�apuri2 = ls_pel�ht�.Column
                End If

                ' Call pl_Lapsi(Lapsinumero, Ryhm�sivu, sht_lapset.Cells(laps.Row, ls_matulo.Column))
                Call pl_Lapsi(Lapsinumero, pl_Hoitoaika, sht_lapset.Cells(laps.Row, p�iv�apuri), sht_lapset.Cells(laps.Row, p�iv�apuri2))
                ' Numerointiin lis�t��n aina yksi (Code-v�lilehte� varten, jotta ryhm�t voidaan my�hemmin poistaa)
                lapsim��r� = lapsim��r� + 1
            Next laps
    
        
            ' 1. sivun viimeinen rivi
            ' Ryhm�sivu.HPageBreaks.Item(1).Location.Row - 1
    
            ' 2. sivun ensimm�inen rivi
            ' Ryhm�sivu.HPageBreaks.Item(1).Location.Row
    
            ' 2. sivun viimeinen rivi:
            ' Ryhm�sivu.HPageBreaks.Item(2).Location.Row - 1
        
            ' TYHJIEN LIS�YS
            ' Lis�t��n tyhji�, kunnes kaikki tyhj�t lis�tty TAI alateksti tulee vastaan
            Dim T_Lis�ys As Integer
            T_Lis�ys = 0
            If lapsim��r� <> 1 Then
                Do Until pl_Nimi.Offset(Rivi).Row = pl_Alateksti.Row Or T_Lis�ys = pl_tyhj�t
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
                        T_Lis�ys = T_Lis�ys + 1
                    End If
                Loop
        
                Dim Nimirivit As Range
                ' Jos lapsirivej� on liikaa, pienennet��n niit�
                If pienennys > 0 Then
                    For Each Nimirivit In Ryhm�sivu.Range(Cells(pl_Nimi.Row, 1), Cells(pl_Nimi.Offset(Rivi).Row - 1, 1))
                        Nimirivit.RowHeight = Application.WorksheetFunction.RoundDown(Nimirivit.RowHeight - (pienennys / Rivi), 1)
                    Next Nimirivit
                    ' Muussa tapauksessa
                End If
     
                ' Poistetaan rivien ja alatekstin v�linen tila JOS ei alateksti ole ihan kiinni riveiss�
                If pl_Nimi.Offset(Rivi).Row = pl_Alateksti.Row Then
                Else
                    .Range(Cells(pl_Nimi.Offset(Rivi).Row, 1), Cells(pl_Alateksti.Row - 1, 1)).Rows.EntireRow.Delete
            
                    ' Silloin on my�s tilaa suurentaa rivej�
                    ' RIVIEN KOON SUURENTAMINEN
                    ' Lasketaan eri osien korkeudet
                    Dim rivikorkeus As Long, LoppuOsa As Range, D_Korkeus As Long
                    D_Korkeus = 0
                    For Each LoppuOsa In Ryhm�sivu.Range(Cells(pl_Vikarivi.Row + 1, 1), Cells(Ryhm�sivu.HPageBreaks.Item(1).Location.Row - 1, 1))
                        rivikorkeus = rivikorkeus + LoppuOsa.Rows.Height
                    Next LoppuOsa
                
                    ' Jaetaan loppuosan tyhj�n tilan korkeus jokaisen lapsirivin kesken (my�s tyhj�t rivit) ja suurennetaan rivien kokoa
                    For Each Nimirivit In Ryhm�sivu.Range(Cells(pl_Nimi.Row, 1), Cells(pl_Nimi.Offset(Rivi).Row - 1, 1))
                        Nimirivit.RowHeight = Nimirivit.RowHeight + (rivikorkeus / Rivi)
                    Next Nimirivit
            
                

                
                End If
                
                On Error Resume Next
                ' TODO
                ' Set Printarea
                ' Ent� 2-puolisena?
                .PageSetup.PrintArea = Ryhm�sivu.Range(Cells(1, 1), Cells(Ryhm�sivu.HPageBreaks.Item(1).Location.Row - 1, pl_Vikasarake.Column)).Address
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

Sub Lapsi(Kutsumanimi As String, Lapsinumero As Long, Ryhm�sivu As Worksheet, tulobold As String, _
          menobold As String, puokkariyks As String, puokkarikaks As String, maTulo As Range, Optional VL_Vkl)

    Dim viiva As String: viiva = "-"
    Dim tuloArr() As String
    Dim l�ht�Arr() As String
    Dim kerta As Long
    Dim arrRivi As Long
    Dim maxArr As Long

    If Lapsinumero = 1 Then Rivi = 2

    With Ryhm�sivu
        ' Lapsen numero ja nimi
        .Cells(Rivi, 1) = Lapsinumero
        .Cells(Rivi, 3) = UCase(Kutsumanimi)
        LapsetSarake = 0
        ListaLis�ys = 0
        ' ma-pe
        P�iv�t = 8
        ' my�s viikonloput
        If VL_Vkl = 1 Then P�iv�t = 12
    
        For LapsetSarake = 0 To P�iv�t Step 2
    
            ' Yksitt�inen hoitoaika
            If IsNumeric(maTulo.Offset(, LapsetSarake).Value2) = True Then
                If Not IsEmpty(maTulo.Offset(, LapsetSarake).Value2) Then
                    ' Muodostetaan hoitoaika
                    ' Puoliy� tulo (muutetaan format stringiksi)
                    If maTulo.Offset(, LapsetSarake).Value2 = 0 Then
                        .Cells(Rivi, 4 + ListaLis�ys).NumberFormat = "@"
                        .Cells(Rivi, 4 + ListaLis�ys) = "0:00"
                    Else
                        .Cells(Rivi, 4 + ListaLis�ys) = maTulo.Offset(, LapsetSarake).Value2
                    End If
                    .Cells(Rivi, 5 + ListaLis�ys) = viiva
                    .Cells(Rivi, 6 + ListaLis�ys) = maTulo.Offset(, LapsetSarake + 1).Value2
                    ' Puokkariboldaukset
                    If maTulo.Offset(, LapsetSarake).Value2 >= puokkariyks And maTulo.Offset(, LapsetSarake).Value2 <= puokkarikaks Then .Cells(Rivi, 4 + ListaLis�ys).Style = "Boldaus3"
                    If maTulo.Offset(, LapsetSarake + 1).Value2 >= puokkariyks And maTulo.Offset(, LapsetSarake + 1).Value2 <= puokkarikaks Then .Cells(Rivi, 6 + ListaLis�ys).Style = "Boldaus3"
                    ' Boldaukset
                    If maTulo.Offset(, LapsetSarake).Value2 <= tulobold Or maTulo.Offset(, LapsetSarake).Value2 >= menobold Then .Cells(Rivi, 4 + ListaLis�ys).Style = "Boldaus2"
                    If maTulo.Offset(, LapsetSarake + 1).Value2 <= tulobold Or maTulo.Offset(, LapsetSarake + 1).Value2 >= menobold Then .Cells(Rivi, 6 + ListaLis�ys).Style = "Boldaus2"
                    ' Puoliy� l�ht� (poistetaan hoitoaika, jos 23:55)
                    If CDate(maTulo.Offset(, LapsetSarake + 1).Value2) = "23:55" Then .Cells(Rivi, 6 + ListaLis�ys) = "23:59"

                 
                End If
                ' Useampia hoitoaikoja
            ElseIf InStr(maTulo.Offset(, LapsetSarake).Value2, ",") > 0 Then
                ' Tehd��n hoitoajoista arrayt
                tuloArr = Split(maTulo.Offset(, LapsetSarake).Value2, ",")
                l�ht�Arr = Split(maTulo.Offset(, LapsetSarake + 1).Value2, ",")
                
                ' Hoitoaikojen m��r� arrayssa
                ' K�yd��n array l�pi
                For kerta = 0 To ArrayLen(tuloArr) - 1
                    ' Eka kerta
                    If kerta = 0 Then
                        arrRivi = Rivi
                    Else
                        arrRivi = arrRivi + 1
                    End If
                    ' Puoliy� tulo (muutetaan format stringiksi)
                    If CDate(tuloArr(kerta)) = "00:00" Then
                        .Cells(Rivi, 4 + ListaLis�ys).NumberFormat = "@"
                        .Cells(Rivi, 4 + ListaLis�ys) = "0:00"
                    Else
                        .Cells(arrRivi, 4 + ListaLis�ys) = tuloArr(kerta)
                    End If
                
                                
                    .Cells(arrRivi, 5 + ListaLis�ys) = viiva
                    .Cells(arrRivi, 6 + ListaLis�ys) = l�ht�Arr(kerta)
                
                    ' Puokkariboldaukset
                    If tuloArr(kerta) >= CDate(puokkariyks) And tuloArr(kerta) <= CDate(puokkarikaks) Then .Cells(arrRivi, 4 + ListaLis�ys).Style = "Boldaus3"
                    If l�ht�Arr(kerta) >= CDate(puokkariyks) And l�ht�Arr(kerta) <= CDate(puokkarikaks) Then .Cells(arrRivi, 6 + ListaLis�ys).Style = "Boldaus3"
                    ' Boldaukset
                    If tuloArr(kerta) <= CDate(tulobold) Or tuloArr(kerta) >= CDate(menobold) Then .Cells(arrRivi, 4 + ListaLis�ys).Style = "Boldaus2"
                    If l�ht�Arr(kerta) <= CDate(tulobold) Or l�ht�Arr(kerta) >= CDate(menobold) Then .Cells(arrRivi, 6 + ListaLis�ys).Style = "Boldaus2"
                
                    ' Puoliy� l�ht� (poistetaan hoitoaika, jos 23:55)
                    If CDate(l�ht�Arr(kerta)) = "23:55" Then .Cells(arrRivi, 6 + ListaLis�ys) = "23:59"
                
                Next kerta
                ' Max arrayn pituus
                If ArrayLen(tuloArr) > maxArr Then maxArr = ArrayLen(tuloArr)
                ' Yksi hoitoaika
            Else
                .Cells(Rivi, 5 + ListaLis�ys) = maTulo.Offset(, LapsetSarake).Value2
            End If
            ListaLis�ys = ListaLis�ys + 3
        Next LapsetSarake

        ' Yl�viiva
        If Lapsinumero <> 1 Then
            If VL_Vkl = 1 Then
                .Range(.Cells(Rivi, 1), .Cells(Rivi, 24)).Style = "Yl�viiva"
            Else
                .Range(.Cells(Rivi, 1), .Cells(Rivi, 18)).Style = "Yl�viiva"
            End If
        End If
    
        ' Harmaatausta
        .Range(.Cells(1, 4), .Cells(Rivi, 6)).Style = "Harmaatausta"
        .Range(.Cells(1, 10), .Cells(Rivi, 12)).Style = "Harmaatausta"
        .Range(.Cells(1, 16), .Cells(Rivi, 18)).Style = "Harmaatausta"
        If VL_Vkl = 1 Then
            .Range(.Cells(1, 22), .Cells(Rivi, 24)).Style = "Harmaatausta"
            .PageSetup.PrintArea = Ryhm�sivu.Range(Cells(1, 1), Cells(Rivi, 24)).Address
        End If
        .Range("C1").EntireColumn.AutoFit
    End With

    ' Lis�t��n rivej� (jos on ollut useampia hoitoaikoja, lis�t��n maksimiarrayn pituuden verran rivej�)
    If maxArr > 0 Then
        Rivi = Rivi + maxArr
    Else
        Rivi = Rivi + 1
    End If

End Sub

Sub pl_Lapsi(Lapsinumero As Long, Ryhm�sivu As Range, maTulo As Range, maL�ht� As Range)

    Dim viiva As String: viiva = "-"
    Dim tuloArr() As String
    Dim l�ht�Arr() As String
    Dim yhdArr() As String
    Dim kerta As Long
    Dim AlkuKorkeus As Long: AlkuKorkeus = 0

    With Ryhm�sivu
        '        Debug.Print (ls_dieetti.Column)
        '       Debug.Print (maTulo.Row)
        ' Lis�t��n kuvake, jos dieettilapsi
        ' TODO: jotain h�ikk��, ei toimi.
'        If Trim(sht_lapset.Cells(maTulo.Row, ls_dieetti.Column).Offset(Rivi).Value2) <> vbNullString Then
        If Trim(sht_lapset.Cells(maTulo.Row, ls_dieetti.Column).Value2) <> vbNullString Then
            .Offset(Rivi, -1) = .Offset(Rivi, -1) + " " + ChrW(&HD83C) & ChrW(&HDF74)
        End If
    
        ' Yksitt�inen hoitoaika
        If IsNumeric(maTulo.Value2) = True Then
            If Not IsEmpty(maTulo.Value2) Then
                ' Muodostetaan hoitoaika
                ' Puoliy� tulo (muutetaan format stringiksi)
                If maTulo.Value2 = 0 Then
                    .Offset(Rivi).NumberFormat = "@"
                    .Offset(Rivi) = "0:00"
                End If
                
                If CDate(maL�ht�.Value2) > "23:54" Then
                    .Offset(Rivi) = Format(maTulo.Value2, "h:mm") & " " & ChrW(8594)
                ElseIf CDate(maTulo.Value2) = "0:00" Then
                    .Offset(Rivi) = " " & ChrW(8594) & " " & Format(maL�ht�.Value2, "h:mm")
                Else
                    .Offset(Rivi) = Format(maTulo.Value2, "h:mm") & " - " & Format(maL�ht�.Value2, "h:mm")
                End If
                
            End If
            ' Useampia hoitoaikoja
        ElseIf InStr(maTulo.Value2, ",") > 0 Then
            AlkuKorkeus = .Offset(1).RowHeight
            ' Tehd��n hoitoajoista arraytnj
            tuloArr = Split(maTulo.Value2, ",")
            l�ht�Arr = Split(maL�ht�.Value2, ",")
            
            ' Hoitoaikojen m��r� arrayssa
            ' K�yd��n array l�pi
            ReDim yhdArr(1)
            For kerta = 0 To ArrayLen(tuloArr) - 1
               
                ' TODO
                ' Puoliy� tulo (muutetaan format stringiksi)
                'If CDate(tuloArr(kerta)) = "00:00" Then
                '    .Offset(Rivi).NumberFormat = "@"
                '     .Offset(Rivi) = "0:00"
                ' End If
               
                ' Array hoitoajoista esim. '11:00 - 14:00'
                'ReDim Preserve yhdArr(kerta)
                
                '                ReDim Preserve Code_ryhm�t(UBound(Code_ryhm�t) + 1)
                '                Code_ryhm�t(UBound(Code_ryhm�t)) = Ryhm�nimi
                ' Y�hoito
                If CDate(l�ht�Arr(kerta)) > "23:54" Then
                    yhdArr(kerta) = tuloArr(kerta) & " " & ChrW(8594)
                ElseIf CDate(tuloArr(kerta)) = "0:00" Then
                    yhdArr(kerta) = " " & ChrW(8594) & " " & l�ht�Arr(kerta)
                Else
                    ' Muutoin vain hoitoaika
                    yhdArr(kerta) = tuloArr(kerta) & " - " & l�ht�Arr(kerta)
                End If
            Next kerta
           
            
            ' Lis�t��n hoitoajat soluun. Hoitoaikojen v�lille rivinvaihto, jotta menev�t samalle solulle
            .Offset(Rivi) = Join(yhdArr, Chr(10))
            ' Koon s��t�
            
            pienennys = pienennys + (.Offset(1).RowHeight - AlkuKorkeus)
            .Offset(Rivi).EntireRow.AutoFit
            
        Else
            .Offset(Rivi) = maTulo.Value2
        End If
    
    End With

    Rivi = Rivi + 1

End Sub

Sub Dieetti(kokonimi As String, Ruokaryhm� As Worksheet, Ryhm�nimi As String, P�iv�t As Long, Ruokailut As Long, Rivi As Long, Sarake As Long)

    ' Ryhm�n ruokailuajat
    Dim a1 As String: a1 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_aamupala1.Column).Value2
    Dim a2 As String: a2 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_aamupala2.Column).Value2
    Dim l1 As String: l1 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_lounas1.Column).Value2
    Dim l2 As String: l2 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_lounas2.Column).Value2
    Dim v1 As String: v1 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_v�lipala1.Column).Value2
    Dim v2 As String: v2 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_v�lipala2.Column).Value2
    Dim p1 As String: p1 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_p�iv�llinen1.Column).Value2
    Dim p2 As String: p2 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_p�iv�llinen1.Column).Value2
    Dim i1 As String: i1 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_iltapala1.Column).Value2
    Dim i2 As String: i2 = sht_ryhm�t.Cells(rs_ryhm�nnimi.Find(Ryhm�nimi).Row, rs_iltapala2.Column).Value2

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
    Dim l�ht�Arr() As String
    Dim yhdArr() As String
    Dim kerta As Long

    With Ruokaryhm�
        For Each lnimi In rng_lapset.SpecialCells(xlCellTypeVisible)
            ' Lopeta jos kyseess� ty�ntekij�
            If sht_lapset.Cells(lnimi.Row, ls_ty�ntekij�.Column).Value2 <> "" Then Exit Sub
            ' K�yd��n l�pi kaikki p�iv�t
            For pv = 0 To P�iv�t - 1
                Set matulorng = sht_lapset.Cells(lnimi.Row, ls_matulo.Column + 2 * pv)
                ' Yksitt�inen hoitoaika
                If IsDate(Format(sht_lapset.Cells(lnimi.Row, ls_matulo.Column + 2 * pv), "h:mm")) = True Then
                    If Ruokailut = 2 Then
                        ' P�iv�llinen
                        If matulorng <= p2 And matulorng.Offset(, 1) >= p1 Then .Cells(Rivi + pv, Sarake).Value = "X"
                        ' Iltapala
                        If matulorng <= i2 And matulorng.Offset(, 1) >= i1 Then .Cells(Rivi + pv, Sarake + 1).Value = "X"
                    
                    ElseIf Ruokailut >= 3 Then
                        ' Aamupala
                        If matulorng <= a2 And matulorng.Offset(, 1) >= a1 Then .Cells(Rivi + pv, Sarake).Value = "X"
                        ' Lounas
                        If matulorng <= l2 And matulorng.Offset(, 1) >= l1 Then .Cells(Rivi + pv, Sarake + 1).Value = "X"
                        ' V�lipala
                        If matulorng <= v2 And matulorng.Offset(, 1) >= v1 Then .Cells(Rivi + pv, Sarake + 2).Value = "X"
                    
                        If Ruokailut = 5 Then
                            ' P�iv�llinen
                            If matulorng <= p2 And matulorng.Offset(, 1) >= p1 Then .Cells(Rivi + pv, Sarake + 3).Value = "X"
                            ' Iltapala
                            If matulorng <= i2 And matulorng.Offset(, 1) >= i1 Then .Cells(Rivi + pv, Sarake + 4).Value = "X"
                        End If
                    End If

                    ' Useampi hoitoaika
                ElseIf InStr(matulorng, ",") > 0 Then
                    ' Tehd��n hoitoajoista arrayt
                    tuloArr = Split(matulorng, ",")
                    l�ht�Arr = Split(matulorng.Offset(, 1), ",")
                    ' K�yd��n array l�pi
                    For kerta = 0 To ArrayLen(tuloArr) - 1
                        If Ruokailut = 2 Then
                            ' P�iv�llinen
                            If tuloArr(kerta) <= CDate(p2) And l�ht�Arr(kerta) >= CDate(p1) Then .Cells(Rivi + pv, Sarake).Value = "X"
                            ' Iltapala
                            If tuloArr(kerta) <= CDate(i2) And l�ht�Arr(kerta) >= CDate(i1) Then .Cells(Rivi + pv, Sarake + 1).Value = "X"
                    
                        ElseIf Ruokailut >= 3 Then
                            ' Aamupala
                            If tuloArr(kerta) <= CDate(a2) And l�ht�Arr(kerta) >= CDate(a1) Then .Cells(Rivi + pv, Sarake).Value = "X"
                            ' Lounas
                            If tuloArr(kerta) <= CDate(l2) And l�ht�Arr(kerta) >= CDate(l1) Then .Cells(Rivi + pv, Sarake + 1).Value = "X"
                            ' V�lipala
                            If tuloArr(kerta) <= CDate(v2) And l�ht�Arr(kerta) >= CDate(v1) Then .Cells(Rivi + pv, Sarake + 2).Value = "X"
                        
                            If Ruokailut = 5 Then
                                ' P�iv�llinen
                                If tuloArr(kerta) <= CDate(p2) And l�ht�Arr(kerta) >= CDate(p1) Then .Cells(Rivi + pv, Sarake + 3).Value = "X"
                                ' Iltapala
                                If tuloArr(kerta) <= CDate(i2) And l�ht�Arr(kerta) >= CDate(i1) Then .Cells(Rivi + pv, Sarake + 4).Value = "X"
                            End If
                        End If
                    Next kerta
                    ' Kirjain tai tyhj�
                Else
                    ' Jos koko viikko tyhj�n�, merkkaa kysymysmerkki ja siirry seuraavaan lapseen
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
    tbl_ryhm�t.Sort.SortFields.Clear
    tbl_ryhm�t.Sort.Apply

    ' Haetaan aamuilta-listojen kellonajat
    Dim aamu1 As String: aamu1 = Replace(sht_p�iv�koti.Range("J5").Value, ",", ".")
    Dim aamu2 As String: aamu2 = Replace(sht_p�iv�koti.Range("K5").Value, ",", ".")
    Dim ilta1 As String: ilta1 = Replace(sht_p�iv�koti.Range("L5").Value, ",", ".")
    Dim ilta2 As String: ilta2 = Replace(sht_p�iv�koti.Range("M5").Value, ",", ".")

    ' Aamulista
    Sheets("Aamuilta_pohja").Copy After:=Sheets("lapset")
    With ActiveSheet
        .Name = "Aamulista"
    End With

    Dim sht_aamulista As Worksheet: Set sht_aamulista = wb.Worksheets("Aamulista")

    ' p�iv�m��rien lis��minen
    Dim pvm_vuosi As Long
    Dim pvm_kk As Long
    Dim pvm_pv As Long
    Dim pvm As Date

    ' Haetaan pvm Codesta ja muunnetaan sopivaan muotoon
    ' P�iv�m��r�n koonti
    pvm = DateSerial(Year(Now), sht_code.[C3].Value, sht_code.[C2].Value)
    Dim viikonp�iv�t(1 To 5) As Range

    Set viikonp�iv�t(1) = Range("tbl_lapset[Ma tulo]")
    Set viikonp�iv�t(2) = Range("tbl_lapset[Ti tulo]")
    Set viikonp�iv�t(3) = Range("tbl_lapset[Ke tulo]")
    Set viikonp�iv�t(4) = Range("tbl_lapset[To tulo]")
    Set viikonp�iv�t(5) = Range("tbl_lapset[Pe tulo]")

    Dim p�iv�nimi(1 To 5) As String
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

    For Each vk In viikonp�iv�t
        tbl_lapset.AutoFilter.ShowAllData
        'sht_aamulista.Range("A1").Offset(rivi - 1).Value = p�iv�nimi(kerta)
        'sht_aamulista.Cells(2, 1 + paivano).value = p�iv�nimi(kerta)
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
            ' Sorttaus: J�rjestyksen mukaan, k��nteisesti
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
                ' Jos tyhj�
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
                        ' sht_aamulista.Range(Cells(rivi + 2, 2), Cells(rivi + 2, 3)).Style = "Yl�viiva"
                        Rivi = Rivi + 1
                        sht_aamulista.Cells(1 + Rivi, 1 + paivano).Value2 = u.Offset(, 6 + paivano).Value2
                        sht_aamulista.Cells(1 + Rivi, 2 + paivano).Value2 = u.Value2
                    End If
                End If
            Next u

            Dim p�iv� As Range
            Set p�iv� = sht_aamulista.Range(Cells(3, 1 + paivano), Cells(3, 1 + paivano).End(xlDown).End(xlDown).End(xlUp))
    
            Dim n As Long: n = 0
            For n = 1 To p�iv�.Rows.Count
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 1, 1) Then p�iv�.Cells(n + 1, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 2, 1) Then p�iv�.Cells(n + 2, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 3, 1) Then p�iv�.Cells(n + 3, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 4, 1) Then p�iv�.Cells(n + 4, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 5, 1) Then p�iv�.Cells(n + 5, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 6, 1) Then p�iv�.Cells(n + 6, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 7, 1) Then p�iv�.Cells(n + 7, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 8, 1) Then p�iv�.Cells(n + 8, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 9, 1) Then p�iv�.Cells(n + 9, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 10, 1) Then p�iv�.Cells(n + 10, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 11, 1) Then p�iv�.Cells(n + 11, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 12, 1) Then p�iv�.Cells(n + 12, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 13, 1) Then p�iv�.Cells(n + 13, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 14, 1) Then p�iv�.Cells(n + 14, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 15, 1) Then p�iv�.Cells(n + 15, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 16, 1) Then p�iv�.Cells(n + 16, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 17, 1) Then p�iv�.Cells(n + 17, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 18, 1) Then p�iv�.Cells(n + 18, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 19, 1) Then p�iv�.Cells(n + 19, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 20, 1) Then p�iv�.Cells(n + 20, 1).Value2 = ""
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
    tbl_ryhm�t.Sort.SortFields.Clear
    tbl_ryhm�t.Sort.Apply

    Sheets("Aamuilta_pohja").Copy After:=Sheets("Aamulista")

    With ActiveSheet
        .Name = "Iltalista"
    End With

    Dim sht_iltalista As Worksheet: Set sht_iltalista = wb.Worksheets("Iltalista")

    Set viikonp�iv�t(1) = Range("tbl_lapset[Ma l�ht�]")
    Set viikonp�iv�t(2) = Range("tbl_lapset[Ti l�ht�]")
    Set viikonp�iv�t(3) = Range("tbl_lapset[Ke l�ht�]")
    Set viikonp�iv�t(4) = Range("tbl_lapset[To l�ht�]")
    Set viikonp�iv�t(5) = Range("tbl_lapset[Pe l�ht�]")

    Rivi = 2
    kerta = 1
    paivano = 0

    sht_iltalista.Range("A2") = pvm
    sht_iltalista.Range("C2") = DateAdd("d", 1, CDate(pvm))
    sht_iltalista.Range("E2") = DateAdd("d", 2, CDate(pvm))
    sht_iltalista.Range("G2") = DateAdd("d", 3, CDate(pvm))
    sht_iltalista.Range("I2") = DateAdd("d", 4, CDate(pvm))

    sht_iltalista.Range("A1").Value2 = "ILTALISTA"

    For Each vk In viikonp�iv�t

        ' Nollataan suodatukset
        tbl_lapset.AutoFilter.ShowAllData

        If paivano = 0 Then
            ' ma
            tbl_lapset.Range.AutoFilter Field:=ls_mal�ht�.Column, Criteria1:=">=" & ilta1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & ilta2, Operator:=xlFilterValues

        ElseIf paivano = 2 Then
            ' ti
            tbl_lapset.Range.AutoFilter Field:=ls_til�ht�.Column, Criteria1:=">=" & ilta1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & ilta2, Operator:=xlFilterValues
        ElseIf paivano = 4 Then
            ' ke
            tbl_lapset.Range.AutoFilter Field:=ls_kel�ht�.Column, Criteria1:=">=" & ilta1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & ilta2, Operator:=xlFilterValues
        ElseIf paivano = 6 Then
            ' to
            tbl_lapset.Range.AutoFilter Field:=ls_tol�ht�.Column, Criteria1:=">=" & ilta1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & ilta2, Operator:=xlFilterValues
        ElseIf paivano = 8 Then
            ' pe
            tbl_lapset.Range.AutoFilter Field:=ls_pel�ht�.Column, Criteria1:=">=" & ilta1, Operator:=xlFilterValues, _
                                        Criteria2:="<=" & ilta2, Operator:=xlFilterValues
        End If

        With tbl_lapset.Sort
            ' Sorttaus: J�rjestyksen mukaan, k��nteisesti
            .SortFields.Clear
            .SortFields.Add Key:=vk, SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With

        On Error GoTo 0
        If tyhjasuodatin(rng_lapset.Offset(, 4)) = True Then
        Else
            For Each u In rng_lapset.SpecialCells(xlCellTypeVisible)
                ' Jos tyhj�
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
                        ' sht_aamulista.Range(Cells(rivi + 2, 2), Cells(rivi + 2, 3)).Style = "Yl�viiva"
                        Rivi = Rivi + 1
                        sht_iltalista.Cells(1 + Rivi, 1 + paivano).Value2 = u.Offset(, 7 + paivano).Value2
                        sht_iltalista.Cells(1 + Rivi, 2 + paivano).Value2 = u.Value2
                    End If
                End If
            Next u
    
            Set p�iv� = sht_iltalista.Range(Cells(3, 1 + paivano), Cells(3, 1 + paivano).End(xlDown).End(xlDown).End(xlUp))
    
    
            n = 0
            For n = 1 To p�iv�.Rows.Count
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 1, 1) Then p�iv�.Cells(n + 1, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 2, 1) Then p�iv�.Cells(n + 2, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 3, 1) Then p�iv�.Cells(n + 3, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 4, 1) Then p�iv�.Cells(n + 4, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 5, 1) Then p�iv�.Cells(n + 5, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 6, 1) Then p�iv�.Cells(n + 6, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 7, 1) Then p�iv�.Cells(n + 7, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 8, 1) Then p�iv�.Cells(n + 8, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 9, 1) Then p�iv�.Cells(n + 9, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 10, 1) Then p�iv�.Cells(n + 10, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 11, 1) Then p�iv�.Cells(n + 11, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 12, 1) Then p�iv�.Cells(n + 12, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 13, 1) Then p�iv�.Cells(n + 13, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 14, 1) Then p�iv�.Cells(n + 14, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 15, 1) Then p�iv�.Cells(n + 15, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 16, 1) Then p�iv�.Cells(n + 16, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 17, 1) Then p�iv�.Cells(n + 17, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 18, 1) Then p�iv�.Cells(n + 18, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 19, 1) Then p�iv�.Cells(n + 19, 1).Value2 = ""
                If p�iv�.Cells(n, 1) = p�iv�.Cells(n + 20, 1) Then p�iv�.Cells(n + 20, 1).Value2 = ""
    
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
    ' P�IV�KOTI
    '************************

    'If sht_code.Range("D2").Value = 0 Then
    '    Lopetus
    '    MsgBox "Hoitoaikoja ei ole lis�tty oikein." & vbCrLf & "K�y lis��m�ss� hoitoajat P�iv�koti-v�lilehdell�. _
    Lue ohjeet. Jos ongelmat jatkuvat, ota yhteytt�: jaakko.haavisto@jyvaskyla.fi", vbExclamation, "Virhe"
    '    End
    'End If

    '************************
    ' AAMU JA ILTALISTAT
    '************************
    Dim o As Range
    ' K�yt�ss�: tyhj� --> "Ei"
    If sht_p�iv�koti.Range("I5") = vbNullString Then sht_p�iv�koti.Range("I5") = "Ei"
    ' Jos K�yt�ss�, kellonajat eiv�t saa olla tyhj�t
    If sht_p�iv�koti.Range("I5") = "Kyll�" Then
        For Each o In sht_p�iv�koti.Range("I5:M5")
            If o = vbNullString Then Call Lopetus("Olet valinnut aamu-ja iltalistat, mutta et ole merkannut kellonaikoja." & vbCrLf & "Lis�� ne P�iv�koti -v�lilehdell�.", vbExclamation, "Virhe")
        Next o
    End If
    ' Alku pit�� olla aikaisempi kuin loppu
    If sht_p�iv�koti.Range("J5") > sht_p�iv�koti.Range("K5") Then Call Lopetus("Aamulistan alku-klo ei voi olla isompi kuin loppu-klo." & vbCrLf & "Tarkasta kellonajat P�iv�koti -v�lilehdelt�.", vbExclamation, "Virhe")
    If sht_p�iv�koti.Range("L5") > sht_p�iv�koti.Range("M5") Then Call Lopetus("Iltalistan alku-klo ei voi olla isompi kuin loppu-klo." & vbCrLf & "Tarkasta kellonajat P�iv�koti -v�lilehdelt�.", vbExclamation, "Virhe")

    '************************
    ' RYHM�T
    '************************

    For Each o In rng_ryhm�t.SpecialCells(xlCellTypeVisible)
        'Jos K�yt�ss� = tyhj� --> "Kyll�"
        If sht_ryhm�t.Cells(o.Row, rs_k�yt�ss�.Column).Value = vbNullString Then sht_ryhm�t.Cells(o.Row, rs_k�yt�ss�.Column).Value = "Kyll�"
    Next o

    ' Suodatetaan K�yt�ss� -mukaan
    'tbl_ryhm�t.Range.AutoFilter 2, "Kyll�"
    With tbl_ryhm�t.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rs_k�yt�ss�, SortOn:=xlSortOnValues, Order:=xlDescending
        .SortFields.Add Key:=rs_j�rjestys, SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .Apply
    End With

    On Error Resume Next
    If tyhjasuodatin(rng_ryhm�t.SpecialCells(xlCellTypeVisible)) = True Then Call Lopetus("Et ole merkinnyt yht�k��n ryhm�� k�yt�ss� olevaksi." & vbCrLf & "Ole hyv� ja merkitse joku ryhmist� k�ytt��n, vaikka tarkoituksenasi olisi tulostaa pelk�st��n aamu- ja iltalistat.", vbExclamation, "Virhe")
    
    On Error GoTo 0
    With sht_ryhm�t
        For Each o In rng_ryhm�t.SpecialCells(xlCellTypeVisible)
            ' Onko k�yt�ss�
            If .Cells(o.Row, rs_k�yt�ss�.Column) = "Kyll�" Then
            
                ' Jos asetukset tyhjin�, merkit��n oletusasetukset
                If .Cells(o.Row, rs_j�rjestys.Column).Value = vbNullString Then .Cells(o.Row, rs_j�rjestys.Column).Value = "1"
                If .Cells(o.Row, rs_aakkosj�rjestys.Column).Value = vbNullString Then .Cells(o.Row, rs_aakkosj�rjestys.Column).Value = "Kutsumanimi"
            
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
                If .Cells(o.Row, rs_v�lipala1.Column).Value = vbNullString Then .Cells(o.Row, rs_v�lipala1.Column).Value = "13:55"
                If .Cells(o.Row, rs_v�lipala2.Column).Value = vbNullString Then .Cells(o.Row, rs_v�lipala2.Column).Value = "14:15"
                If .Cells(o.Row, rs_iltapala1.Column).Value = vbNullString Then .Cells(o.Row, rs_iltapala1.Column).Value = "18:55"
                If .Cells(o.Row, rs_iltapala2.Column).Value = vbNullString Then .Cells(o.Row, rs_iltapala2.Column).Value = "19:15"
            
                If .Cells(o.Row, rs_listatulostus.Column).Value = vbNullString Then .Cells(o.Row, rs_listatulostus.Column).Value = "Ma-pe"
                If .Cells(o.Row, rs_listatulostus.Column).Value = "Kyll�" Then .Cells(o.Row, rs_listatulostus.Column).Value = "Ma-pe"
                
                If .Cells(o.Row, rs_p�ivystys.Column).Value = vbNullString Then .Cells(o.Row, rs_p�ivystys.Column).Value = "Ei"
                If .Cells(o.Row, rs_p�iv�laput.Column).Value = vbNullString Then .Cells(o.Row, rs_p�iv�laput.Column).Value = "Ei"
                If .Cells(o.Row, rs_plabc.Column).Value = vbNullString Then .Cells(o.Row, rs_plabc.Column).Value = "Kutsumanimi"
                If .Cells(o.Row, rs_plpohja.Column).Value = vbNullString Then .Cells(o.Row, rs_plpohja.Column).Value = "Pysty"
                If .Cells(o.Row, rs_pltyhj�t.Column).Value = vbNullString Then .Cells(o.Row, rs_pltyhj�t.Column).Value = "0"
                
                
                If .Cells(o.Row, rs_p�ivystys.Column).Value <> "Ei" And .Cells(o.Row, rs_ruokayhdistys.Column).Value <> vbNullString And .Cells(o.Row, rs_ruokatulostus.Column).Value <> "Ei" Then
                    Call Lopetus(o.Value & " -ryhm� on merkattu p�ivystysryhm�ksi. Siin� on my�s yhdistettyj� ruokatilaus-ryhmi�. " _
                               & vbCrLf & vbCrLf & "Se ei ole tuettu ominaisuus t�ll� hetkell�. Poista yhdistetyt ruokaryhm�t.", vbExclamation, "Virhe")
                End If
                
                If .Cells(o.Row, rs_p�ivystys.Column).Value <> "Ei" And .Cells(o.Row, rs_listayhdistys.Column).Value <> vbNullString And .Cells(o.Row, rs_listatulostus.Column).Value <> "Ei" Then
                    Call Lopetus(o.Value & " -ryhm� on merkattu p�ivystysryhm�ksi. Siin� on my�s yhdistettyj� ryhmi�. " _
                               & vbCrLf & vbCrLf & "Se ei ole tuettu ominaisuus t�ll� hetkell�. Poista yhdistetyt ryhm�t.", vbExclamation, "Virhe")
                End If
                
                If .Cells(o.Row, rs_p�ivystys.Column).Value <> "Ei" And .Cells(o.Row, rs_listayhdistys.Column).Value <> vbNullString And .Cells(o.Row, rs_p�iv�laput.Column).Value <> "Ei" Then
                    Call Lopetus(o.Value & " -ryhm� on merkattu p�ivystysryhm�ksi. Siin� on my�s yhdistettyj� ryhmi�. " _
                               & vbCrLf & vbCrLf & "Se ei ole tuettu ominaisuus t�ll� hetkell�. Poista yhdistetyt ryhm�t.", vbExclamation, "Virhe")
                End If
                
                ' TODO
                ' Yhdistettyj� ruokalistoja, mutta nimi puuttuu --> Ryhmien nimet
   
                ' Yhdistettyjen ryhmien tsekkaus
                Dim validi_array() As Variant
                validi_array = Application.Transpose(sht_code.Range("G2:G" & sht_code.Range("G2").End(xlDown).Row).Value2)
                Dim vl_yhdistely_array() As String
                Dim element As Variant
            
                ' Ruokalapun yhdistetyt ryhm�t pit�� olla valideja
                If .Cells(o.Row, rs_ruokayhdistys.Column).Value <> "" Then
                    vl_yhdistely_array = Split(spaceremove(.Cells(o.Row, rs_listayhdistys.Column).Value), ",", , vbTextCompare)
                    For Each element In vl_yhdistely_array
                        If IsInArray(CStr(element), validi_array) = False Then
                            Call Lopetus(o.Value & " -ryhm�n kanssa yhdistetyn ruokalapun kanssa on ongelma. " & element & _
                                         " ei ole oikea ryhm�." & vbCrLf & vbCrLf & "Mahdollisia ryhmi� ovat: " & _
                                         Join(validi_array, ", "), vbExclamation, "Virhe")
                        End If
                    Next element
                End If
            
                ' Viikkolistan yhdistetyt ryhm�t pit�� olla valideja
                If .Cells(o.Row, rs_listayhdistys.Column).Value <> "" Then
                    vl_yhdistely_array = Split(spaceremove(.Cells(o.Row, rs_listayhdistys.Column).Value), ",", , vbTextCompare)
                    For Each element In vl_yhdistely_array
                        If IsInArray(CStr(element), validi_array) = False Then Call Lopetus(o.Value & " -ryhm�n kanssa yhdistetyn viikkolistan kanssa on ongelma. " & element & _
                                                                                            " ei ole oikea ryhm�." & vbCrLf & vbCrLf & "Mahdollisia ryhmi� ovat: " & _
                                                                                            Join(validi_array, ", "), vbExclamation, "Virhe")
                    Next element
                End If
            
                ' Yhdistettyj� viikkolistoja, mutta listan nimi puuttuu (Viikkolistan tulostus, Ryhmi� valittuna, asetuksissa yhdistetty viikkolista, mutta listan nimi tyhj�)
                If .Cells(o.Row, rs_listatulostus.Column).Value <> "Ei" And _
                                                                .Cells(o.Row, rs_listayhdistys.Column).Value <> "" And _
                                                                Left(.Cells(o.Row, rs_yhdistettytyyli.Column).Value, 10) = "Yhdistetty" And _
                                                                .Cells(o.Row, rs_yhdistettynimi.Column).Value = "" Then
                    Call Lopetus(o.Value & " -ryhm�ll� on yhdistettyj� viikkolistoja, mutta listan nimi puuttuu." & vbCrLf & "K�y lis��m�ss� yhdistetyn listan nimi ja luo listat uudelleen.", vbExclamation, "Virhe")
                End If
            
                ' Jos p�iv�laput ja asetuksista 2-puoliset p�iv�laput, mutta ryhmien yhdist�minen on tyhj� --> Valitus
                If .Cells(o.Row, rs_p�iv�laput.Column).Value <> "Ei" And _
                                                             Right(.Cells(o.Row, rs_yhdistettytyyli.Column).Value, 22) = "2-puoleiset p�iv�laput" And _
                                                             .Cells(o.Row, rs_listayhdistys.Column).Value = vbNullString Then
                    Call Lopetus("Olet valinnut ryhm�lle " & o.Value & " 2-puoleisten p�iv�lappujen tulostamisen," & vbCrLf & "mutta et ole valinnut 2. sivulle tulevaa ryhm��." & vbCrLf & _
                                 "Ole hyv� ja merkitse ryhm�n nimi Ryhmien yhdist�minen-sarakkeeseen Ryhm�t-v�lilehdell�.", vbExclamation, "Virhe")
                End If
            
                Dim arrayIsNothing As Boolean
                On Error Resume Next
                arrayIsNothing = IsNumeric(UBound(vl_yhdistely_array)) And False
                If Err.Number <> 0 Then arrayIsNothing = True
                On Error GoTo 0

                ' Jos yhdistettyjen ryhmien array on tyhj�, ei tehd� alkutarkistusta.
                If arrayIsNothing = False Then
                    ' Jos 2-puoleiset p�iv�laput, saa olla ainoastaan 1 yhdistett�v� ryhm�, jos enemm�n ryhmi� --> Valitus
                    If .Cells(o.Row, rs_p�iv�laput.Column).Value <> "Ei" And _
                                                                 Right(.Cells(o.Row, rs_yhdistettytyyli.Column).Value, 22) = "2-puoleiset p�iv�laput" And _
                                                                 ArrayLen(vl_yhdistely_array) > 1 Then
                        Call Lopetus("Olet valinnut ryhm�lle " & o.Value & " 2-puoleisten p�iv�lappujen tulostamisen," & vbCrLf & "T�m� asetus rajoittaa yhdistett�vien ryhmien m��r�n yhteen." & vbCrLf & _
                                     "K�y poistamassa ylim��r�iset yhdistetyt ryhm�t.", vbExclamation, "Virhe")
                    End If
                End If
           
                ' Jos p�iv�laput + 2-puolinen, tarkistetaan 2.lapun asetukset
                If .Cells(o.Row, rs_p�iv�laput.Column).Value <> "Ei" And _
                                                             Right(.Cells(o.Row, rs_yhdistettytyyli.Column).Value, 22) = "2-puoleiset p�iv�laput" Then
                    Dim tokaryhm� As Integer: tokaryhm� = rs_ryhm�nnimi.Find(vl_yhdistely_array(0)).Row
                    ' Oletusasetukset pohjalle (pysty) ja tyhjien m��r�lle (0
                    If .Cells(tokaryhm�, rs_plpohja.Column).Value2 = vbNullString Then .Cells(tokaryhm�, rs_plpohja.Column).Value2 = "Pysty"
                    If .Cells(tokaryhm�, rs_pltyhj�t.Column).Value2 = vbNullString Then .Cells(tokaryhm�, rs_pltyhj�t.Column).Value2 = 0
                    ' Jos 2. p�iv�lapun pohja on kustomoitu, eik� oikeaa pohjaa l�ydy --> Valitus
                    If .Cells(tokaryhm�, rs_plpohja.Column).Value2 = "Kustomoitu" And WorksheetExists("pl_" & .Cells(tokaryhm�, rs_ryhm�nnimi.Column).Value) = False Then
                        Call Lopetus(o.Value & " -ryhm�ll� on 2-puoleinen p�iv�lappu-asetus. Kuitenkaan 2. ryhm�n kustomoitua p�iv�lappua ei ole olemassa." & vbCrLf & "K�y generoimassa kustomoitu pohja " & sht_ryhm�t.Cells(tokaryhm�, rs_ryhm�nnimi.Column).Value & " -ryhm�lle P�iv�koti-v�lilehdell�.", vbExclamation, "Virhe")
                    End If
                End If
                ' Jos p�iv�laput + kustomoitu pohja, mutta pohjaa ei l�ydy --> Valitus
                If .Cells(o.Row, rs_p�iv�laput.Column).Value <> "Ei" And _
                                                             .Cells(o.Row, rs_plpohja.Column).Value = "Kustomoitu" And _
                                                             WorksheetExists("pl_" & .Cells(o.Row, rs_ryhm�nnimi.Column).Value) = False Then
                    Call Lopetus(o.Value & " -ryhm�ll� on kustomoitu p�iv�lappu, mutta sit� ei ole olemassa." & vbCrLf & "K�y generoimassa kustomoitu pohja P�iv�koti-v�lilehdell�.", vbExclamation, "Virhe")
                End If
                    

            End If
        Next o
    End With

End Sub

Sub Lopetus_Simple()
    ThisWorkbook.Worksheets("P�iv�koti").Select

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    End
End Sub

Sub Lopetus(Optional VirheViesti As String, Optional VirheTyyppi As String, Optional VirheOtsikko As String)

    Dim sort_ryhm�t_j�rjestys As Range
    Set sort_ryhm�t_j�rjestys = Range("tbl_ryhm�t[J�rjestys]")

    ' Sorttaus: j�rjestys alusta loppuun
    With sht_ryhm�t.ListObjects("tbl_ryhm�t").Sort
        .SortFields.Clear
        .SortFields.Add Key:=sort_ryhm�t_j�rjestys, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    ' Lis�t��n suojaukset
    'For Each WSprotect In Array("R_pohja1", "R_pohja2", "Ri_pohja1", "Ri_pohja2", "Pohja")
    '    Worksheets(WSprotect).Protect AllowFormattingCells:=True, AllowFormattingColumns:=True, _
    '        AllowFormattingRows:=True
    'Next WSprotect

    ' Nollataan suodatukset
    Dim sort_lapset_abc As Range
    Set sort_lapset_abc = Range("tbl_lapset[Ryhm�]")

    'tbl_lapset.AutoFilter.ShowAllData
    With tbl_lapset.Sort
        ' Sorttaus: J�rjestyksen mukaan, k��nteisesti
        .SortFields.Clear
        .SortFields.Add Key:=sort_lapset_abc, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    ' Piilotetaan v�lilehdet
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

    wb.Worksheets("P�iv�koti").Select
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
    Set sht_ryhm�t = wb.Worksheets("Ryhm�t")
    Set sht_lapset = wb.Worksheets("Lapset")
    Set tbl_ryhm�t = sht_ryhm�t.ListObjects("tbl_ryhm�t")
    Set tbl_lapset = sht_lapset.ListObjects("tbl_lapset")
    tbl_ryhm�t.ShowAutoFilter = True
    tbl_ryhm�t.AutoFilter.ShowAllData
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

    ' N�ytet��n kaikki v�lilehdet
    Dim WSname As Variant
    For Each WSname In Array("lasna", "Code", "Pohja", "Aamuilta_pohja")
        Worksheets(WSname).Visible = True
    Next WSname

    Set wb = ThisWorkbook
    Set sht_code = wb.Worksheets("Code")

    ' Poistetaan vanhat ryhm�t
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

    ' Piilotetaan v�lilehdet
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

    wb.Worksheets("P�iv�koti").Select

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

    ' N�ytet��n kaikki v�lilehdet
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

    ' Piilotetaan v�lilehdet
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

    wb.Worksheets("P�iv�koti").Select

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

' Tulostaa arrayn lukum��r�n
Function arrayitems(arr As Variant) As Long
    Dim lukum��r� As Long: lukum��r� = UBound(arr) - LBound(arr) + 1
    arrayitems = lukum��r�
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

Sub Ruokatilausten_l�hett�minen()
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
        
    Set sht_p�iv�koti = wb.Worksheets("P�iv�koti")
    
    Dim emailApplication As Object
    Dim emailItem As Object
    Dim strPath As String
    Dim lngPos As Long
    ' Build the PDF file name
    'strPath = ActiveWorkbook.FullName
    strPath = Application.ActiveWorkbook.Path + "\" + sht_p�iv�koti.Range("I11").Value2 + ".pdf"
    'MsgBox strPath
    'lngPos = InStrRev(strPath, ".")
    'strPath = Left(strPath, lngPos) & "pdf"
    ' Export workbook as PDF
    If pdflista.Count > 0 Then
        Sheets(collectionToArray(pdflista)).Select
    Else
        MsgBox "Ei l�ydetty yht��n ruokatilauslappua."
        Call Lopetus_Simple
    End If
    
    ActiveSheet.ExportAsFixedFormat xlTypePDF, strPath, OpenAfterPublish:=False, IgnorePrintAreas:=False
    
    '    ActiveWorkbook.ExportAsFixedFormat xlTypePDF, strPath
    Set emailApplication = CreateObject("Outlook.Application")
    Set emailItem = emailApplication.CreateItem(0)
    ' Now we build the email.
    emailItem.To = sht_p�iv�koti.Range("I9").Value2
    emailItem.Subject = sht_p�iv�koti.Range("I13").Value2
    emailItem.Body = sht_p�iv�koti.Range("I15").Value2
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


