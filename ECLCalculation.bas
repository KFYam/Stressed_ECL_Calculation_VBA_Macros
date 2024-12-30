Public Sub UpdateStageSEQ()
    Dim S1FieldName As New Scripting.Dictionary
    Dim StressStage2Flag As New Scripting.Dictionary

    Worksheets("Stage1_STAT_StressedECL").Select
    Range("A1").Select

    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    lastCol = Range("A1").End(xlToRight).Column

    For thisCol = 1 To lastCol
        S1FieldName.Add Cells(1, thisCol).Value, thisCol
    Next thisCol

    Dim tmp_sum As Integer
    For thisRow = 2 To lastRow
        tmp_sum = Application.Sum(Range(Cells(thisRow, S1FieldName.Item("SC1_SE1_Stage1_to_2n3_Ind")), Cells(thisRow, S1FieldName.Item("SC2_SE2_Stage1_to_2n3_Ind"))))
        If tmp_sum > 0 Then
            StressStage2Flag.Add Cells(thisRow, 1).Value, 2
        End If
    Next thisRow

    Dim InputDataFieldName As New Scripting.Dictionary
    Worksheets("Input_Data").Select
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    lastCol = Range("A1").End(xlToRight).Column

    For thisCol = 1 To lastCol
        InputDataFieldName.Add Cells(1, thisCol).Value, thisCol
    Next thisCol

    For thisRow = 2 To lastRow
        If Cells(thisRow, InputDataFieldName.Item("FLAG_STAT_STAGE2")).Value = 1 Then
        ElseIf Cells(thisRow, InputDataFieldName.Item("FLAG_STAT_STAGE2")).Value = 2 Then
            If Not StressStage2Flag.Exists(Cells(thisRow, 1).Value) Then
                Cells(thisRow, InputDataFieldName.Item("FLAG_STAT_STAGE2")).ClearContents
            End If
        Else
            If StressStage2Flag.Exists(Cells(thisRow, 1).Value) Then
                Cells(thisRow, InputDataFieldName.Item("FLAG_STAT_STAGE2")).Value = 2
            End If
        End If
    Next thisRow
End Sub

Public Sub Stage2Clear()
    'Clear contents - Stage2_STAT_StressedECL
    Worksheets("Stage2_STAT_StressedECL").Select
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    lastCol = Range("A4").End(xlToRight).Column

    If lastRow > 4 Then
        Range(Range("A5"), Cells(lastRow, lastCol)).Select
        Selection.ClearContents
    End If
End Sub

Public Function DictInputField() As Dictionary
    Dim InputDataFieldName As New Scripting.Dictionary
    Worksheets("Input_Data").Select
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    lastCol = Range("A1").End(xlToRight).Column

    For thisCol = 1 To lastCol
        InputDataFieldName.Add Cells(1, thisCol).Value, thisCol
    Next thisCol
    Set DictInputField = InputDataFieldName
End Function

Public Function DictS2ECLField() As Dictionary
    Dim stage2Field As New Scripting.Dictionary
    Worksheets("Stage2_STAT_StressedECL").Select
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    lastCol = Range("A4").End(xlToRight).Column

    For thisCol = 1 To lastCol
        stage2Field.Add Cells(4, thisCol).Value, thisCol
    Next thisCol
    Set DictS2ECLField = stage2Field
End Function

Public Function DictInputPD() As Dictionary
    'Get Stress PD information
    Dim ratingPD As New Scripting.Dictionary
    With Sheets("Input_PD")
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        For thisRow = 2 To lastRow
            ratingPD.Add .Cells(thisRow, 1).Value, .Cells(thisRow, 11).Value
        Next thisRow
    End With

    Set DictInputPD = ratingPD
End Function

Public Function DictInputPWA() As Dictionary
    Dim arrGND(1 To 3) As String
    'Get Stress PWA information
    arrGND(1) = "GOOD"
    arrGND(2) = "NEUTRAL"
    arrGND(3) = "DOWNTURN"

    Dim PWALookup As New Scripting.Dictionary
    With Sheets("Input_PWA")
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        For thisRow = 2 To lastRow
            For G = 1 To 3
                tmp_PWAkey = .Cells(thisRow, 1).Value & .Cells(thisRow, 2).Value & .Cells(thisRow, 3).Value & arrGND(G)
                PWALookup.Add tmp_PWAkey, .Cells(thisRow, 3 + G).Value
            Next G
        Next thisRow
    End With
    Set DictInputPWA = PWALookup
End Function

Public Function DictLGD_nonretail() As Dictionary
    Dim arrGND(1 To 3) As String
    arrGND(1) = "GOOD"
    arrGND(2) = "NEUTRAL"
    arrGND(3) = "DOWNTURN"
    '--------------------------------------------------------------------------------
    'Get nonretail LGD Multipliers vlookup from input sheet
    Dim LGDMulti As New Scripting.Dictionary
    With Sheets("Input_stressed_LGD_multiplers")
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        For thisRow = 2 To lastRow
            For G = 1 To 3
                tmp_LGDkey = .Cells(thisRow, 1).Value & .Cells(thisRow, 2).Value & .Cells(thisRow, 3).Value & .Cells(thisRow, 4).Value & arrGND(G)
                LGDMulti.Add tmp_LGDkey, .Cells(thisRow, 4 + G).Value
            Next G
        Next thisRow
    End With
    Set DictLGD_nonretail = LGDMulti
End Function

Public Function DictLGD_retail() As Dictionary
    Dim arrGND(1 To 3) As String
    arrGND(1) = "GOOD"
    arrGND(2) = "NEUTRAL"
    arrGND(3) = "DOWNTURN"
    '--------------------------------------------------------------------------------
    'Get retail related LGD value vlookup from input sheet
    Dim retailLGD As New Scripting.Dictionary
    With Sheets("Input_Retail_stressedLGD")
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        For thisRow = 2 To lastRow
            For G = 1 To 3
                tmp_retalLGD_key = .Cells(thisRow, 1).Value & .Cells(thisRow, 2).Value & .Cells(thisRow, 3).Value & arrGND(G)
                retailLGD.Add tmp_retalLGD_key, .Cells(thisRow, 3 + G).Value
            Next G
        Next thisRow
    End With

    Set DictLGD_retail = retailLGD
End Function

Public Sub pasteStage2(sht_src As Worksheet, sht_tar As Worksheet, arr_src As Variant, lRow As Variant)
    For k = LBound(arr_src) To UBound(arr_src)
        sht_src.Select
        Range(Cells(1, arr_src(k)), Cells(1, arr_src(k))).Select
        Range(Selection, Cells(lRow, arr_src(k))).Select
        Selection.Copy
        sht_tar.Select
        Range(Cells(4, k), Cells(4, k)).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Next k
End Sub

Public Sub Stage2_Static(sht_src As Worksheet, sht_tar As Worksheet, InputDataFieldName As Dictionary)

    'Dim sht_src As Worksheet
    'Dim sht_tar As Worksheet
    'Set sht_src = ThisWorkbook.Worksheets("Input_Data")
    'Set sht_tar = ThisWorkbook.Worksheets("Stage2_STAT_StressedECL")
    'Dim InputDataFieldName As New Scripting.Dictionary
    'Worksheets("Input_Data").Select
    'lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    'lastCol = Range("A1").End(xlToRight).Column
    'For thisCol = 1 To lastCol
    '    InputDataFieldName.Add Cells(1, thisCol).Value, thisCol
    'Next thisCol

    sht_src.Select
    lRow = Cells(Rows.Count, 2).End(xlUp).Row
    lCol = Range("A1").End(xlToRight).Column
    Rows("1:1").Select
    Selection.AutoFilter

    ActiveSheet.Range(Range("A1"), Cells(lRow, lCol)).AutoFilter _
        Field:=InputDataFieldName.Item("FLAG_STAT_STAGE2"), _
        Criteria1:="=1", Operator:=xlOr, Criteria2:="=2"

    Dim arrsrc(1 To 11) As Integer
    Dim arrtar(1 To 11) As Integer
    arrsrc(1) = InputDataFieldName.Item("SEQ")
    arrsrc(2) = InputDataFieldName.Item("Exposure Reference")
    arrsrc(3) = InputDataFieldName.Item("RATING_KEY")
    arrsrc(4) = InputDataFieldName.Item("Number of Years in Stage2")
    arrsrc(5) = InputDataFieldName.Item("Expected Life in Year - Stage 2")
    arrsrc(6) = InputDataFieldName.Item("Region Code")
    arrsrc(7) = InputDataFieldName.Item("HKFRS9 PD Model Segment Final")
    arrsrc(8) = InputDataFieldName.Item("Probability Weighted Average - Good")
    arrsrc(9) = InputDataFieldName.Item("Probability Weighted Average - Neutral")
    arrsrc(10) = InputDataFieldName.Item("Probability Weighted Average - Downturn")
    arrsrc(11) = InputDataFieldName.Item("ECL Amount - Statistical - Stage 2")
    Call pasteStage2(sht_src, sht_tar, arrsrc, lRow)
End Sub

Sub Gen_stat_Stage2_ECL()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Dim sht_src As Worksheet
    Dim sht_tar As Worksheet
    Dim dictPD As Dictionary
    Dim dictPWA As New Dictionary
    Dim dictLGDnonretail As Dictionary
    Dim dictLGDretail As Dictionary
    Dim dictInput As Dictionary
    Dim dictS2ECL As Dictionary
    Dim arrGNDECL(1 To 3) As Double
    Dim arrGND(1 To 3) As String
    Dim arrDynSrc(12 To 19) As Integer

    arrGND(1) = "GOOD"
    arrGND(2) = "NEUTRAL"
    arrGND(3) = "DOWNTURN"
    Set sht_src = ThisWorkbook.Worksheets("Input_Data")
    Set sht_tar = ThisWorkbook.Worksheets("Stage2_STAT_StressedECL")
    '-------------------------------------------------------------

    sht_src.Select
    lRow = Cells(Rows.Count, 2).End(xlUp).Row
    Set dictPD = DictInputPD()
    Set dictPWA = DictInputPWA()
    Set dictLGDnonretail = DictLGD_nonretail()
    Set dictLGDretail = DictLGD_retail()
    Set dictInput = DictInputField()
    Set dictS2ECL = DictS2ECLField()

    Call Stage2Clear
    Call UpdateStageSEQ
    Call Stage2_Static(sht_src, sht_tar, dictInput)
    sht_src.Select
    lRow_src = Cells(Rows.Count, 2).End(xlUp).Row
    sht_tar.Select
    lRow_tar = Cells(Rows.Count, 1).End(xlUp).Row

    With Sheets("Stage2_STAT_StressedECL")
        finalindex = 0
        For SC = 1 To 3 ' Scenario
            .Cells(1, 2).Value = "SC" & SC
            For SE = 1 To 3 'Severity
                .Cells(2, 2).Value = "SE" & SE
                For y = 1 To 31 'Year 1 to 30 plus residential
                    .Cells(3, 2).Value = y
                    arrDynSrc(12) = dictInput.Item("EAD Post CCF - Stage 2 - Year 1") + y - 1
                    arrDynSrc(13) = dictInput.Item("ECL Cash Flow Discount Factor - Stage 2 - Year 1") + y - 1
                    arrDynSrc(14) = dictInput.Item("Lifetime PD - Good - Stage 2 - Year 1") + y - 1
                    arrDynSrc(15) = dictInput.Item("Lifetime PD - Neutral - Stage 2 - Year 1") + y - 1
                    arrDynSrc(16) = dictInput.Item("Lifetime PD - Downturn - Stage 2 - Year 1") + y - 1
                    arrDynSrc(17) = dictInput.Item("Realized LGD - Good - Stage 2 - Year 1") + y - 1
                    arrDynSrc(18) = dictInput.Item("Realized LGD - Neutral - Stage 2 - Year 1") + y - 1
                    arrDynSrc(19) = dictInput.Item("Realized LGD - Downturn - Stage 2 - Year 1") + y - 1
                    Call pasteStage2(sht_src, sht_tar, arrDynSrc, lRow_src)
                    For Row = 5 To lRow_tar
                        '----------------------------------------------------------
                        'Derive the partial lifetime
                        If y = 1 Then
                            If .Cells(Row, dictS2ECL.Item("Expected Life in Year - Stage 2")).Value < 1 Then
                                .Cells(Row, dictS2ECL.Item("Partial Expected Life")).Value = .Cells(Row, dictS2ECL.Item("Expected Life in Year - Stage 2")).Value
                            Else
                                .Cells(Row, dictS2ECL.Item("Partial Expected Life")).Value = 1
                            End If
                        ElseIf y > 1 Then
                            If .Cells(Row, dictS2ECL.Item("Expected Life in Year - Stage 2")).Value - y > 0 Then
                                If .Cells(Row, dictS2ECL.Item("Expected Life in Year - Stage 2")).Value - y < 1 Then
                                    .Cells(Row, dictS2ECL.Item("Partial Expected Life")).Value = .Cells(Row, dictS2ECL.Item("Expected Life in Year - Stage 2")).Value - y
                                Else
                                    .Cells(Row, dictS2ECL.Item("Partial Expected Life")).Value = 1
                                End If
                            Else
                                .Cells(Row, dictS2ECL.Item("Partial Expected Life")).Value = 0
                            End If
                        End If

                        For GND = 1 To 3 'For Good, Neutral and Downturn
                            '----------------------------------------------------------
                            ' Set the stressed PD
                            pdkey = "SC" & SC & "SE" & SE & _
                                    .Cells(Row, dictS2ECL.Item("RATING_KEY")).Value & _
                                    y & arrGND(GND)
                            If dictPD.Exists(pdkey) Then
                                .Cells(Row, dictS2ECL.Item("ST_PD_Good") + GND - 1).Value = dictPD.Item(pdkey)
                            Else
                                .Cells(Row, dictS2ECL.Item("ST_PD_Good") + GND - 1).Value = 0
                            End If
                            '----------------------------------------------------------
                            ' Set the non retail stressed LGD
                            nonretailLGDkey = "SC" & SC & "SE" & SE & _
                                              .Cells(Row, dictS2ECL.Item("Region Code")).Value & _
                                              .Cells(Row, dictS2ECL.Item("HKFRS9 PD Model Segment Final")).Value & arrGND(GND)
                            retailLGDkey = "SC" & SC & "SE" & SE & _
                                           .Cells(Row, dictS2ECL.Item("Exposure Reference")).Value & arrGND(GND)

                            If dictLGDnonretail.Exists(nonretailLGDkey) Then
                                .Cells(Row, dictS2ECL.Item("ST_LGD_Good") + GND - 1).Value = .Cells(Row, 17 + GND - 1).Value * dictLGDnonretail.Item(nonretailLGDkey) 'Realized LGD - Good - Stage 2 - Year #
                            ElseIf dictLGDretail.Exists(retailLGDkey) Then
                                .Cells(Row, dictS2ECL.Item("Retail_stressedLGD_Ind")).Value = 1
                                .Cells(Row, dictS2ECL.Item("ST_LGD_Good") + GND - 1).Value = dictLGDretail.Item(retailLGDkey)
                            Else
                                .Cells(Row, dictS2ECL.Item("ST_LGD_Good") + GND - 1).Value = .Cells(Row, 17 + GND - 1).Value
                            End If
                            '----------------------------------------------------------
                            ' Set the PWA
                            PWAkey = "SC" & SC & "SE" & SE & .Cells(Row, dictS2ECL.Item("Region Code")).Value & arrGND(GND)
                            If dictPWA.Exists(PWAkey) Then
                                .Cells(Row, dictS2ECL.Item("ST_ECLPWA_Good") + GND - 1).Value = dictPWA.Item(PWAkey)
                            Else
                                .Cells(Row, dictS2ECL.Item("ST_ECLPWA_Good") + GND - 1).Value = 0
                            End If
                            '----------------------------------------------------------
                            ' Set the ECL delta
                            ' EAD Post CCF - Stage 2* ECL Cash Flow Discount Factor - Stage 2 *
                            ' Partal Expected Life in Year - Stage 2 *Adj - Lifetime PD - Stage 2*
                            ' Realized LGD - Stage 2 * PWA
                            arrGNDECL(GND) = .Cells(Row, 12).Value * .Cells(Row, 13).Value * _
                                             .Cells(Row, dictS2ECL.Item("Partial Expected Life")).Value * .Cells(Row, dictS2ECL.Item("ST_LGD_Good") + GND - 1).Value * _
                                             .Cells(Row, dictS2ECL.Item("ST_LGD_Good") + GND - 1).Value * .Cells(Row, dictS2ECL.Item("ST_ECLPWA_Good") + GND - 1).Value
                        Next GND
                        .Cells(Row, dictS2ECL.Item("ST_ECL_01") + y - 1).Value = arrGNDECL(1) + arrGNDECL(2) + arrGNDECL(3)
                    Next Row
                Next y

                For Row = 5 To lRow_tar
                    tempecl = 0
                    For y = 1 To 31
                        tempecl = tempecl + .Cells(Row, dictS2ECL.Item("ST_ECL_01") + y - 1).Value
                    Next y
                    .Cells(Row, dictS2ECL.Item("SC1_SE1_ECL") + finalindex).Value = tempecl
                Next Row
                finalindex = finalindex + 1
            Next SE
        Next SC
    End With

    Application.ScreenUpdating = True
    Application.Calculation = xlAuto
End Sub








