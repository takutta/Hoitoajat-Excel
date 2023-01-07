Attribute VB_Name = "Convert"
Sub Lasnaconvert()
    '1       2         3         4         5         6         7         8         9        10        11        12        13        14
    '2345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678902345
    ' dev
    Set wb = ThisWorkbook
    Set sht_lasna = wb.Worksheets("lasna")
    Set sht_lapset = wb.Worksheets("lapset")
    Set sht_code = wb.Worksheets("Code")
    Set sht_ryhm�t = wb.Worksheets("ryhm�t")
    Set sht_lasna2 = wb.Worksheets("Lasna2")

    Dim ykkosrivi As Range
    Dim kakkosrivi As Range
    Dim KloRng As Range

    Dim testiymparisto As Boolean
    Dim v�lirivit As Long
    testiymparisto = 0

    If testiymparisto = 0 Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    End If

    ' Nollataan suodattimet
    Set tbl_ryhm�t = sht_ryhm�t.ListObjects("tbl_ryhm�t")
    Set tbl_lapset = sht_lapset.ListObjects("tbl_lapset")
    'tbl_ryhm�t.AutoFilter.ShowAllData
    'tbl_lapset.AutoFilter.ShowAllData

   
    ' N�ytet��n kaikki v�lilehdet
    Dim WSname As Variant
    For Each WSname In Array("lasna", "Code")
        Worksheets(WSname).Visible = True
    Next WSname

    sht_lasna.Select
    Range("A1").Select
    ' Tyhjennys
    sht_code.Range("G2:G2000").Value2 = ""
    sht_lapset.Range("H2:Q2000").Value2 = ""
    sht_lasna.Range("A1:V2000").Value2 = ""

    If testiymparisto = True Then
        sht_lasna2.Range("A1:V2000").Copy
        sht_lasna.Range("A1").PasteSpecial
    
    Else
        On Error Resume Next
        ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
                                 False, NoHTMLFormatting:=True
        On Error GoTo 0
    
        Dim pastehaku As Range
        Set pastehaku = sht_lasna.Range("A1:M200").Find("Lasten lajittelu", lookat:=xlWhole)

        If pastehaku Is Nothing Then
            Lopetus
            ThisWorkbook.Worksheets("Code").Range("D2").Value = 0
            MsgBox "Oletko varma, ett� kopioit hoitoajat oikein?" & vbCrLf & "Leikep�yd�lt� ei l�ytynyt viittausta hoitoaikavarauksista. _Kokeile viel� kerran, pari ja jos ongelma toistuu, laita mailia jaakko.haavisto@jyvaskyla.fi", _
                   vbExclamation, "Virhe"
            End
        End If
    End If


    Dim haku As Range
    With sht_lasna
        Set haku = .Range("A:A").Find("LAPSET", MatchCase:=True, LookIn:=xlValues, lookat:=xlWhole)
        ' 1/3 EKA RYHM�
        ' Poistetaan turha alkumatsku
        .Range(Cells(1, 1), Cells(haku.Row - 2, 1)).EntireRow.Delete
    
        ' Selvitet��n l�ytyyk� ryhm�tt�mi� lapsia --> poistoon
        If InStr(.Range("A1").Value2, "-") = 0 Then
            haku.Find ("LAPSET")
            Set haku = .Range("A:A").FindNext(haku)
            .Range(Cells(1, 1), Cells(haku.Row - 2, 1)).EntireRow.Delete
        End If
    
        'Ryhm�n nimen selvitys "- " RYHM�N_NIMI " ("
        .[I4].Value = Split(Split(.[a1].Value, "- ")(1), " (")(0)
    End With

    ' K�yd��n kopioimassa ekan ryhm�n nimi codeen
    sht_code.[G2].Value = Split(Split(sht_lasna.[a1].Value, "- ")(1), " (")(0)

    With sht_lasna
        Dim maPvm As String: maPvm = .[B3].Value
        Dim EkaPiste As String, TokaPiste As String
        Dim loppurivi As Integer
        EkaPiste = InStr(1, maPvm, ".", vbTextCompare)
        TokaPiste = InStr(EkaPiste + 1, maPvm, ".", vbTextCompare)
        ' P�iv� codeen
        sht_code.Range("C2").Value2 = Mid(Left(maPvm, EkaPiste - 1), 3)
        sht_code.Range("C3").Value2 = Mid(Left(maPvm, TokaPiste - 1), EkaPiste + 1)

        ' Poistetaan 3 turhaa rivi�
        .Range(Cells(1, 1), Cells(3, 1)).EntireRow.Delete
    
        ' Poistetaan tuplarivit
        Dim tyhjahaku As Range
        Dim hakukerrat As Integer
        Dim ekahaku As Boolean: ekahaku = True
        hakukerrat = WorksheetFunction.CountIf(.Range(Cells(1, 1), Cells(LastRow_1(sht_lasna), 1)), "")
        Dim p�iv� As Integer
        Dim rivim��r� As Long
        Dim c As Range
        ' Tiivistet��n listaa niin, ett� jokaisesta lapsesta vain yksi rivi.
        ' Jos kaksi rivi, jossa toinen teksti� --> poistetaan teksti
        ' Jos useampi rivi hoitoaikoja, yhdistet��n ja erotellaan pilkulla
        ' hae tyhji� ylh��lt� alas kunnes niit� ei en�� ole
        Set c = .Range(Cells(LastRow_1(sht_lasna), 1), Cells(1, 1)).Find("", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlPrevious)
        If Not c Is Nothing Then
            Do
                Set tyhjahaku = .Range(Cells(LastRow_1(sht_lasna), 1), Cells(1, 1)).Find("", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlPrevious)
        
                Do While IsEmpty(tyhjahaku.Offset(-1).Value) = True
                    Set tyhjahaku = tyhjahaku.Offset(-1)
                Loop
        
                Set kakkosrivi = sht_lasna.Range(Cells(tyhjahaku.Row, 2), Cells(tyhjahaku.Row, 2))
                Set ykkosrivi = kakkosrivi.Offset(-1, 0)
        
                ' K�yd��n kaikki lapsen hoitoajat l�pi, yhdistet��n hoitoaikoja.
                rivim��r� = ykkosrivi.Row + 1
                For p�iv� = 0 To 6
                    ' Poistetaan kaikki solut, joissa n�kyy "Information"
                    If InStr(ykkosrivi.Offset(, p�iv�), "Information") > 0 Then
                        ykkosrivi.Offset(, p�iv�).Value = ""
                        'ykkosrivi.Offset(, p�iv�).Value2 = kakkosrivi.Offset(, p�iv�).Value2
                    End If

                    ' Onko hoitoaikojen 2. rivi kellonaika?
                    If InStr(kakkosrivi.Offset(, p�iv�), "-") <> 0 Then
                        ' Jos ekalla rivill� on my�s kellonaika, yhdistet��n hoitoajat samalle riville
                        If InStr(ykkosrivi.Offset(, p�iv�), "-") <> 0 Then
                            Set KloRng = .Range(Cells(ykkosrivi.Offset(, p�iv�).Row + 1, ykkosrivi.Offset(, p�iv�).Column), _
                                                Cells(Vikatyhja(ykkosrivi.Offset(, -1)), ykkosrivi.Offset(, p�iv�).Column))
                            For Each rng In KloRng
                                If InStr(rng.Value, "-") <> 0 Then
                                    ykkosrivi.Offset(, p�iv�).Value = ykkosrivi.Offset(, p�iv�).Value & "," & rng.Value
                                End If
                                If rivim��r� < rng.Row Then rivim��r� = rng.Row
                            Next rng
                        Else
                            ' Ei kellonaikaa ekalla rivill�, joten 1. rivi = 2. rivi
                            ykkosrivi.Offset(, p�iv�).Value2 = kakkosrivi.Offset(, p�iv�).Value2
                        End If
                    Else
                    
                    End If
            
                Next p�iv�
                ' Poistetaan ylim��r�iset rivit
                .Range(Cells(ykkosrivi.Row + 1, 1), Cells(rivim��r�, 1)).EntireRow.Delete
                ' Uusi haku
                Set c = .Range(Cells(LastRow_1(sht_lasna), 1), Cells(1, 1)).Find("", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlPrevious)
            Loop While Not c Is Nothing
        End If
    
        ' Poistetaan duplikaatit.
        ' Oletetaan, ett� toinen nimi kertoo toisesta sijoituksesta,
        ' joka vaihtuu keskell� viikkoa. Otetaan sen tiedot yl�s.
        Dim sijoitushaku As Range
        Set sijoitushaku = .Range(Cells(1, 1), Cells(LastRow_1(sht_lasna), 1))
        Dim n As Integer: n = 0
        Dim m As Long
        For n = 2 To sijoitushaku.Rows.Count
            If sijoitushaku.Cells(n - 1, 1) = sijoitushaku.Cells(n, 1) Then
                For m = 2 To 8
                    If sijoitushaku.Cells(n, m).Value <> "Sijoitus puuttuu" Then
                        sijoitushaku.Cells(n - 1, m).Value = sijoitushaku.Cells(n, m).Value
                    ElseIf sijoitushaku.Cells(n - 1, m).Value <> "Sijoitus puuttuu" Then
                        sijoitushaku.Cells(n, m).Value = sijoitushaku.Cells(n - 1, m).Value
                    End If
                Next m

                ' poistetaan rivi
                sijoitushaku.Cells(n, 1).EntireRow.Delete
            End If
        Next n
    
        ' 2/3 LOPUT RYHM�T
        ' Etsit��n seuraava LAPSET
        Dim RyhmiaYhteensa As Long
        RyhmiaYhteensa = WorksheetFunction.CountIf(Range("A:A"), "LAPSET")
        Dim ryhm�no As Double: ryhm�no = 2
        Dim Ryhm�looppi As Integer
        For i = 1 To RyhmiaYhteensa
            ryhm�no = ryhm�no + 1
            Set haku = sht_lasna.Range("A:A").Find("LAPSET", MatchCase:=True, LookIn:=xlValues, lookat:=xlWhole)
            Dim rawryhm� As String: rawryhm� = .Cells(haku.Row - 1, 1).Value2
            ' Lis�t��n ryhm�n nimi
            .Cells(haku.Row + 2, 9).Value = Split(Split(rawryhm�, "- ")(1), " (")(0)
            sht_code.Cells(ryhm�no, 7).Value = Split(Split(rawryhm�, "- ")(1), " (")(0)
            ' Poistetaan turhat 3 rivi�
            .Range(Cells(haku.Row - 1, 1), Cells(haku.Row + 1, 1)).EntireRow.Delete
        
        Next i
        

        With .Range(Cells(1, 2), Cells(LastRow_1(sht_lasna), 8))
            .Replace What:="Poissa (P)", Replacement:="P", lookat:=xlPart, _
                     searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                     ReplaceFormat:=False
            .Replace What:="Sairaus (S)", Replacement:="S", lookat:=xlPart, _
                     searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                     ReplaceFormat:=False
            .Replace What:="Ei hoitoaikavarausta", Replacement:="", lookat:=xlPart, _
                     searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                     ReplaceFormat:=False
            .Replace What:="Peruutettu hoitop�iv� (H)", Replacement:="P", lookat:=xlPart, _
                     searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                     ReplaceFormat:=False
            .Replace What:="P�iv�kohtainen v�hennys (D)", Replacement:="P", lookat:=xlPart, _
                     searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                     ReplaceFormat:=False
            .Replace What:="Sijoitus puuttuu", Replacement:="P", lookat:=xlPart, _
                     searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                     ReplaceFormat:=False
            .Replace What:="Loma-ajan poissaoloilmoitus", Replacement:="P", lookat:=xlPart, _
                     searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                     ReplaceFormat:=False
            .Replace What:="�killinen poissaolo", Replacement:="P", lookat:=xlPart, _
                     searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                     ReplaceFormat:=False
            .Replace What:=".", Replacement:=":", lookat:=xlPart, _
                     searchorder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                     ReplaceFormat:=False
        End With
    
        ' Jos hoitoaika alkaa kirjaimella, eik� ole P, niin korvataan hoitoaika tyhj�ll�
        Dim cellturha As Range
        For Each cellturha In .Range(Cells(1, 2), Cells(LastRow_1(sht_lasna), 8))
            If cellturha.Value Like "[a-zA-Z]*" Then
                If cellturha.Value <> "P" Then cellturha.Value = ""
            End If
        Next cellturha

        ' 3/3
        ' Kopioidaan ryhm�n nimi joka riville
        .Range(Cells(1, 9), Cells(LastRow_1(sht_lasna), 9)).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"

        ' II
        Dim LasnaLista As Range
        Set LasnaLista = .Range(Cells(1, 1), Cells(LastRow_1(sht_lasna), 1))
    End With

    Dim LapsiLista As Range
    Set LapsiLista = sht_lapset.Range(sht_lapset.Cells(2, 3), sht_lapset.Cells(LastRow_1(sht_lapset), 3))

    Dim LasnaNimi As Range

    Dim kokonimi As String
    Dim sukunimi As String

    sht_lapset.Select

    Dim onkopoissa As Range
    Dim OnPoissa As Boolean
    Dim hoitoArr() As String
    Dim kerta As Integer
    Dim tulot As String
    Dim menot As String
    Dim LapsetSarake As Integer
    Dim Lis�ys As Integer
    Dim nimirivi As Double

    For Each LasnaNimi In LasnaLista
        ' Kerroin: L�sn�sarake vs Lapsetsarake, jotta p�rj�t��n yhdell� for-loopilla
        Lis�ys = 7
    
        ' VAIN LC 1/3
        If LapsiLista.Find(LasnaNimi.Value2, lookat:=xlWhole) Is Nothing Then
            Dim lisarivi As Double
            lisarivi = LastRow_1(sht_lapset) + 1
            ' Kutsumanimi
            sht_lapset.Cells(lisarivi, 2).Value2 = Split(LasnaNimi.Value2, " ")(0) & " " & Left(Right(LasnaNimi.Value2, Len(LasnaNimi.Value2) - InStrRev(LasnaNimi.Value2, " ")), 1)
            ' Koko nimi
            sht_lapset.Cells(lisarivi, 3).Value2 = LasnaNimi.Value2
            ' Ryhm�n nimi
            sht_lapset.Cells(lisarivi, 4).Value2 = LasnaNimi.Offset(0, 8).Value2
            ' Listan p�ivitys
            Set LapsiLista = sht_lapset.Range(sht_lapset.Cells(2, 3), sht_lapset.Cells(LastRow_1(sht_lapset), 3))
        End If
    
        ' LC & Excel 2/3
        ' P�ivitet��n hoitoajat
        For LapsetSarake = 8 To 20 Step 2
            nimirivi = LapsiLista.Find(LasnaNimi.Value2, lookat:=xlWhole).Row
            ' L�ytyyk� hoitoajoista pilkku (eli 2+ hoitoaikaa)
            If InStr(LasnaNimi.Offset(0, LapsetSarake - Lis�ys).Value2, ",") > 0 Then
                ' Tehd��n hoitoajoista array
                hoitoArr = Split(LasnaNimi.Offset(0, LapsetSarake - Lis�ys).Value2, ",")
                ' Hoitoaikojen m��r� arrayssa
                ' K�yd��n array l�pi
                For kerta = 0 To ArrayLen(hoitoArr) - 1
                    'hoitoArr(kerta) Ekalla kerralla ilman pilkkua
                    If kerta = 0 Then
                        tulot = Left(hoitoArr(kerta), 5)
                        menot = Right(hoitoArr(kerta), 5)
                    
                    Else
                        If Not Left(hoitoArr(kerta), 5) = "" Then
                            tulot = tulot + "," + Left(hoitoArr(kerta), 5)
                            menot = menot + "," + Right(hoitoArr(kerta), 5)
                        End If
                    End If
                Next kerta
            
                sht_lapset.Cells(nimirivi, LapsetSarake).Value2 = tulot
                sht_lapset.Cells(nimirivi, LapsetSarake + 1).Value2 = menot
            
                ' Vain yksitt�inen hoitoaika
            Else
                sht_lapset.Cells(nimirivi, LapsetSarake).Value2 = Left(LasnaNimi.Offset(0, LapsetSarake - Lis�ys).Value2, 5)
                sht_lapset.Cells(nimirivi, LapsetSarake + 1).Value2 = Right(LasnaNimi.Offset(0, LapsetSarake - Lis�ys).Value2, 5)
            End If
            Lis�ys = Lis�ys + 1
        Next LapsetSarake
        
        ' Korvaa jos eri ryhm�n nimi
        If Not sht_lapset.Cells(nimirivi, 4).Value2 = Replace(LasnaNimi.Offset(0, 8).Value2, ".", ":") Then
            sht_lapset.Cells(nimirivi, 4).Value2 = Replace(LasnaNimi.Offset(0, 8).Value2, ".", ":")
        End If
        ' Sukunimi
        sht_lapset.Cells(nimirivi, 22).Value2 = Right(LasnaNimi.Value2, Len(LasnaNimi.Value2) - InStrRev(LasnaNimi.Value2, " "))
    
    Next LasnaNimi


    ' VAIN EXCEL 3/3
    Dim Lapsinimi As Range
    Set tbl_lapset = sht_lapset.ListObjects("tbl_lapset")

    Dim lRivi As Long
    Dim lKokonimi As String
    Dim juttu As Long
    juttu = LapsiLista.Rows.Count + 1
    'For Each Lapsinimi In LapsiLista
    For lRivi = juttu To 1 Step -1
        lKokonimi = sht_lapset.Cells(lRivi, 3).Value
        If lKokonimi <> "Koko nimi" Then
            ' Poista rivi JOS lapsen nime� ei l�ydy L�sn�st� EIK� rivi� ole lukittu
            If LasnaLista.Find(lKokonimi, lookat:=xlWhole) Is Nothing Then
                If sht_lapset.Cells(lRivi, 6).Value2 = "" Then
                    tbl_lapset.ListRows(lRivi - 1).Delete
                End If
            End If
        End If
    Next lRivi

    Dim sort_lapset_ryhm� As Range: Set sort_lapset_ryhm� = Range("tbl_lapset[Ryhm�]")
    Dim sort_lapset_aakkoset As Range: Set sort_lapset_aakkoset = Range("tbl_lapset[Koko nimi]")

    With tbl_lapset
        ' Sorttaus: J�rjestyksen mukaan, k��nteisesti
        With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=sort_lapset_ryhm�, SortOn:=xlSortOnValues, Order:=xlAscending
            .SortFields.Add Key:=sort_lapset_aakkoset, SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With
        ' Suodatus: K�yt�ss� olevat
    End With

    ' Ryhmien synkkaus
    sht_ryhm�t.Select
    Set tbl_ryhm�t = sht_ryhm�t.ListObjects("tbl_ryhm�t")

    Dim excelryhm�t As Range

    Dim lcryhm�t As Range
    Set lcryhm�t = sht_code.Range(sht_code.Cells(2, 7), sht_code.Cells(ryhm�no, 7))

    Dim Ryhm�nimi As Range
    For Each Ryhm�nimi In lcryhm�t
        Set excelryhm�t = sht_ryhm�t.Range(sht_ryhm�t.Cells(2, 3), sht_ryhm�t.Cells(LastRow_1(sht_ryhm�t), 3))
    
        lisarivi = LastRow_1(sht_ryhm�t) + 1
        If excelryhm�t.Find(Ryhm�nimi.Value2) Is Nothing Then
            With sht_ryhm�t
                .Cells(lisarivi, 1).Value2 = Application.WorksheetFunction.Max _
                                             (.Range(sht_ryhm�t.Cells(2, 1), sht_ryhm�t.Cells(lisarivi - 1, 1))) + 1
                .Cells(lisarivi, 2).Value2 = "Kyll�"
                .Cells(lisarivi, 3).Value2 = Ryhm�nimi.Value2
                .Cells(lisarivi, 4).Value2 = "Kutsumanimi"
                .Cells(lisarivi, 5).Value2 = "6:59" 'boldaus 1
                .Cells(lisarivi, 6).Value2 = "17:01" ' boldaus 2
                .Cells(lisarivi, 7).Value2 = "10:00" ' puokkariboldaus 1
                .Cells(lisarivi, 8).Value2 = "14:30" ' puokkariboldaus 2
                .Cells(lisarivi, 9).Value2 = "Ma-pe" ' Ruokalapun tulostus
                .Cells(lisarivi, 10).Value2 = "Pieni fontti" ' Ruokalapun asetukset
                .Cells(lisarivi, 11).Value2 = "" ' Yhdist� ruokalaput
                .Cells(lisarivi, 12).Value2 = "7:50" ' Aamupala 1
                .Cells(lisarivi, 13).Value2 = "8:15" ' Aamupala 2
                .Cells(lisarivi, 14).Value2 = "11:50" ' Lounas 1
                .Cells(lisarivi, 15).Value2 = "12:15" ' Lounas 2
                .Cells(lisarivi, 16).Value2 = "13:55" ' V�lipala 1
                .Cells(lisarivi, 17).Value2 = "14:15" ' V�lipala 2
                .Cells(lisarivi, 18).Value2 = "16:55" ' P�iv�llinen 1
                .Cells(lisarivi, 19).Value2 = "17:15" ' P�iv�llinen 2
                .Cells(lisarivi, 20).Value2 = "18:55" ' Iltapala 1
                .Cells(lisarivi, 21).Value2 = "19:15" ' Iltapala 2
                .Cells(lisarivi, 22).Value2 = "Ma-pe" ' Viikkolistan tulostus
                .Cells(lisarivi, 23).Value2 = "" 'Yhdist� viikkolistat
                .Cells(lisarivi, 24).Value2 = "" ' Yhdistetyn viikkolistan nimi
                .Cells(lisarivi, 25).Value2 = "Viikkolista & p�iv�laput" ' Yhdistetyn listan tyyli"
                .Cells(lisarivi, 26).Value2 = "Ei" ' P�ivystys
                .Cells(lisarivi, 27).Value2 = "Ei" ' P�iv�laput
                .Cells(lisarivi, 28).Value2 = "Kutsumanimi" ' PL aakkosj�rjestys
                .Cells(lisarivi, 29).Value2 = "Pysty" ' PL pohja

            End With
        End If
    Next Ryhm�nimi

    juttu = sht_ryhm�t.Range(sht_ryhm�t.Cells(2, 3), sht_ryhm�t.Cells(LastRow_1(sht_ryhm�t), 3)).Rows.Count + 1
    ' Poistetaan k�ytt�m�tt�m�t ryhm�t
    For lRivi = juttu To 1 Step -1
        lKokonimi = sht_ryhm�t.Cells(lRivi, 3).Value
        If lKokonimi <> "Ryhm�n nimi" Then
            ' Poista rivi JOS ryhm�n nime� ei l�ydy codesta
            If lcryhm�t.Find(lKokonimi, lookat:=xlWhole) Is Nothing Then
                tbl_ryhm�t.ListRows(lRivi - 1).Delete
            End If
        End If
    Next lRivi

    For Each Ryhm�nimi In lcryhm�t
        Set excelryhm�t = sht_ryhm�t.Range(sht_ryhm�t.Cells(2, 3), sht_ryhm�t.Cells(LastRow_1(sht_ryhm�t), 3))
        lisarivi = LastRow_1(sht_ryhm�t) + 1
        If excelryhm�t.Find(Ryhm�nimi.Value2) Is Nothing Then
            With sht_ryhm�t
                .Cells(lisarivi, 1).Value2 = Application.WorksheetFunction.Max _
                                             (.Range(sht_ryhm�t.Cells(2, 1), sht_ryhm�t.Cells(lisarivi - 1, 1))) + 1
                .Cells(lisarivi, 2).Value2 = "Kyll�"
                .Cells(lisarivi, 3).Value2 = Ryhm�nimi.Value2
                .Cells(lisarivi, 4).Value2 = "Kutsumanimi"
                .Cells(lisarivi, 5).Value2 = "7:00"
                .Cells(lisarivi, 6).Value2 = "17:00"

            End With
        End If
    Next Ryhm�nimi

    ' I Taulukon jauhaminen oikeaan muotoon

    ' 1/3 EKA RYHM�
    ' Etsi LAPSET (caps)
    ' Poista A1 -> LAPSET koko rivi (-2 rivi�)
    ' Lis�� ryhm�n nimi kohtaan H4  <- A1 viivan ja sulkeen v�liss� oleva matsku ilman v�lej�
    ' Nappaa talteen pvm (B3) toka piste viel� mukaan
    ' Poista rivit 1-3

    ' 2/3 LOPUT RYHM�T
    ' Etsi seuraava tyhj� sarakkeessa B
    ' Lopeta jos sen rivin sarakkeen A on tyhj�.
    ' Jos ei ole tyhj�, kopioi taas ryhm�n nimi (h+3)
    ' Poista sen rivi --> 3 rivi�

    ' 3/3 HIENOS��T�
    ' - Ympp�� ryhm�n nimet kaikille sarakkeen riveille (jossain oli koodinp�tk� siihen)
    ' - Muuta:
    '       Ei hoitoaikavarausta = ""
    '       Poissa (P) = P
    '       Sairaus (S) = S

    ' II Taulukon yhdist�minen lapsilistaan

    ' taulukon A1 eka nimi (lapsilistalla vaikkapa C1) L�ytyyk� nime�?
    ' K�yd��n kaikki nimet l�pi.

    '   * LC ei Excel --> lis�� uusi rivi taulukon loppuun
    '       - Kopioi ryhm�n nimi ja hoitoajat oikeille kohdille
    '       - Kutsumanimi = 1. sana nimest�
    '       - Oma     = ei
    '       - Dieetti = ei
    '   * LC ja Excel -->
    '       - P�ivit� hoitoajat ma-pe ja ryhm�n nimi
    '   * Excel, ei LC
    '       - Jos oma = ei    --> Poista rivi

    ' III Hoitoaikojen generointi

    ' Lasten hoitoaikoja ei tarvits en�� m�ts�t� erilliselt� listalta Lapset-v�lilehteen,
    ' vaan ne voidaan suoraan hakea Lapset-v�lilehdelt�.

    ' Lapset-v�lilehti.
    ' Filtter�id��n ryhm�n mukaan
    ' Sortataan oikealla tavalla
    ' 1. nimi ma: (lapset) F2 -> (ryhm�) D2-F2
    '         ti: (lapset) G2 -> (ryhm�) G2-I2
    '         ke: (lapset) H2 -> (ryhm�) J2-L2
    '         to: (lapset) I2 -> (ryhm�) M2-O2
    '         pe: (lapset) J2 -> (ryhm�) P2-R2
    ' 2. nimi ma: (lapset) F3 -> (ryhm�) D3-F3
    '         ti: (lapset) G4 -> (ryhm�) G3-I3
    '         ke: (lapset) H5 -> (ryhm�) J3-L3
    '         to: (lapset) I6 -> (ryhm�) M3-O3
    '         pe: (lapset) J7 -> (ryhm�) P3-R3
    ' ... jne.

    ' IV ruokalapun generointi

    ' Lasten hoitoaikoja ei tarvits en�� m�ts�t� erilliselt� listalta Lapset-v�lilehteen,
    ' vaan ne voidaan suoraan hakea Lapset-v�lilehdelt�.
    Sheets(Array("lasna", "Code")).Select
    ActiveWindow.SelectedSheets.Visible = False

    wb.Worksheets("P�iv�koti").Select

    Call PoistaSuodatukset

    MsgBox "Lapsilistat p�ivitetty."

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

Public Function LastRow_1(wS As Worksheet) As Double
    With wS
        If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
            LastRow_1 = .Cells.Find(What:="*", _
                                    After:=.Range("A1"), _
                                    lookat:=xlPart, _
                                    LookIn:=xlFormulas, _
                                    searchorder:=xlByRows, _
                                    searchdirection:=xlPrevious, _
                                    MatchCase:=False).Row
        Else
            LastRow_1 = 1
        End If
    End With
End Function

Public Function LastRow_0(wS As Worksheet) As Double
    On Error Resume Next
    LastRow_0 = wS.Cells.Find(What:="*", _
                              After:=wS.Range("A1"), _
                              lookat:=xlPart, _
                              LookIn:=xlFormulas, _
                              searchorder:=xlByRows, _
                              searchdirection:=xlPrevious, _
                              MatchCase:=False).Row
End Function

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

Sub Lopetus()

    Sheets(Array("Lasna", "Code")).Select
    ActiveWindow.SelectedSheets.Visible = False

    ThisWorkbook.Worksheets("P�iv�koti").Select

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

Function Vikatyhja(rStart As Range) As Long
    NextFilled = rStart.EntireColumn.Find(What:="?*", After:=rStart, LookIn:=xlValues).Row

    If NextFilled <> 0 Then
        If NextFilled = 1 Then
            Vikatyhja = LastRow_1(sht_lasna)
        Else
            Vikatyhja = NextFilled - 1
        End If
    Else
        Vikatyhja = rStart
    End If
End Function


