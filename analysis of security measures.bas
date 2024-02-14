Attribute VB_Name = "Module4"
Option Explicit

Sub Count_Colored_Cells()
    Dim GreenCount As Integer
    Dim OrangeCount As Integer
    Dim YellowCount As Integer
    Dim RedCount As Integer ' ����� ���������� ���� "����� �����" ��� �������� ���������� ������� �����
    Dim Sheet As Worksheet
    Dim lastRow As Long
    Dim lastRowFile As Integer
    Dim Cell As Range
    Dim ws As Worksheet
    Dim targetMonth As Integer
    Dim targetYear As Integer
    ' ����� ��������� ���������� ��� �������, ������� ����� ������
    Dim WhatFilter As String
    Dim OtherWorkbook As Workbook
    Dim FGreenDoneRowCount As Integer
    Dim FGreenInProgressCount As Integer
    Dim FGreenOverdueCount As Integer
    Dim FOrangeDoneRowCount As Integer
    Dim FOrangeInProgressCount As Integer
    Dim FOrangeOverdueCount As Integer
    Dim FYellowDoneRowCount As Integer
    Dim FYellowInProgressCount As Integer
    Dim FYellowOverdueCount As Integer
    ' ����� ������������� ���������� ��� �������� ���-�� F-����� �������� �����
    Dim FRedDoneRowCount As Integer
    Dim FRedInProgressCount As Integer
    Dim FRedOverdueCount As Integer
    Dim TotalGreenCount As Integer
    Dim TotalOrangeCount As Integer
    Dim TotalYellowCount As Integer
    Dim TotalRedCount As Integer
    Dim AverageGreen As Double
    Dim AverageOrange As Double
    Dim AverageYellow As Double
    Dim AverageRed As Double
    Dim FGreenRowCountList As New Collection
    Dim FOrangeRowCountList As New Collection
    Dim FYellowRowCountList As New Collection
    Dim FRedRowCountList As New Collection
    Dim UniqueDates As Object
    Set UniqueDates = CreateObject("Scripting.Dictionary")
    Dim BValue As New Collection
    
    ' ��������� ����, �� ������� ������� �������
    Set ws = ThisWorkbook.Sheets(1)

    ' ��������� ����� �� ����� ������� �� ����� ���������
    If Not IsEmpty(Range("B2").Value) And IsEmpty(Range("C2").Value) And IsEmpty(Range("D2").Value) Then
        ' �������� ����� � ��� �� ������ �������
        targetMonth = Month(ws.Range("B2").Value)
        targetYear = Year(ws.Range("B2").Value)
        WhatFilter = "B"
    ElseIf Not IsEmpty(Range("D2").Value) And IsEmpty(Range("C2").Value) And IsEmpty(Range("B2").Value) Then
        ' �������� ��� �� ������ �������
        targetYear = Year(ws.Range("D2").Value)
        WhatFilter = "D"
    ElseIf Not IsEmpty(Range("C2").Value) And IsEmpty(Range("D2").Value) And IsEmpty(Range("B2").Value) Then
        ' �������� �������� ��� �� ������ C2 � ���� ������
        Dim interval As String
        interval = Range("C2").Value
        ' �������� ������ � ���������� �� 2 ����
        Dim dates() As String
        Dim OneDate As Date
        Dim TwoDate As Date
            
        dates = Split(interval, "-")
        OneDate = DateValue(Trim(dates(0)))
        TwoDate = DateValue(Trim(dates(1)))
        WhatFilter = "C"
    Else
        MsgBox "����������� ����� ������ �� �����"
    End If

    ' �������� ��������
    GreenCount = 0
    OrangeCount = 0
    YellowCount = 0
    RedCount = 0

    FGreenDoneRowCount = 0
    FGreenInProgressCount = 0
    FGreenOverdueCount = 0
    FOrangeDoneRowCount = 0
    FOrangeInProgressCount = 0
    FOrangeOverdueCount = 0
    FYellowDoneRowCount = 0
    FYellowInProgressCount = 0
    FYellowOverdueCount = 0
    FRedDoneRowCount = 0
    FRedInProgressCount = 0
    FRedOverdueCount = 0

    TotalGreenCount = 0
    TotalOrangeCount = 0
    TotalYellowCount = 0
    TotalRedCount = 0

    AverageGreen = 0
    AverageOrange = 0
    AverageYellow = 0
    AverageRed = 0

    ' ������� ������� ���� �������
    lastRowFile = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row ' ��������� ����������� ������ � ������� E

    Dim j As Long
    For j = 2 To lastRowFile ' ������� �� 2 ������ � �� ��������� ����������� ������

        ' ��������� �����, � ������� ����� ������� ������
        Set OtherWorkbook = Workbooks.Open(ws.Cells(j, "E").Value)
        Set Sheet = OtherWorkbook.Sheets(1)
        
        ' Determine the last filled cell in column E
        lastRow = Sheet.Cells(Sheet.Rows.Count, "E").End(xlUp).Row ' ��������� ����������� ������ � ������� E
        
        ' ������� ���������� ������� ����� � ������� E, ��������������� ��������� ���� � ������� B
        Dim i As Long
        For i = 4 To lastRow ' ������� � 4 ������ � �� ��������� ����������� ������
            ' ���������, ��� ������ B �������� ���� � �� �������� ������
            If IsDate(Sheet.Cells(i, "B").Value) Then
                ' ��������� ������ �� ������ �� ������ � ����
                If WhatFilter = "B" Then
                    If Month(Sheet.Cells(i, "B").Value) = targetMonth And Year(Sheet.Cells(i, "B").Value) = targetYear Then
                        '���� ����� ������ ���� ������� ����� ����
                        If Sheet.Cells(i, "E").MergeCells Then
                            'MsgBox "��� ������ ����������."
                            
                            Dim mergedRange As Range
                            Set mergedRange = Sheet.Cells(i, "E").MergeArea ' �������� ������������ ��������
                        
                            ' Check cell color
                            Select Case Sheet.Cells(i, "E").Interior.Color
                                Case RGB(0, 176, 80) ' Green color
                                    GreenCount = GreenCount + 1
                                Case RGB(255, 192, 0) ' Orange color
                                    OrangeCount = OrangeCount + 1
                                Case RGB(255, 255, 0) ' Yellow color
                                    YellowCount = YellowCount + 1
                                Case RGB(255, 0, 0) ' Red color
                                    RedCount = RedCount + 1
                            End Select
                                
                            ' ���������� ������������ ������
                            i = i + mergedRange.Cells.Count - 1
                        Else
                            'MsgBox "��� ������ �� ����������."
                            ' Check cell color
                            Select Case Sheet.Cells(i, "E").Interior.Color
                                Case RGB(0, 176, 80) ' Green color
                                    GreenCount = GreenCount + 1
                                Case RGB(255, 192, 0) ' Orange color
                                    OrangeCount = OrangeCount + 1
                                Case RGB(255, 255, 0) ' Yellow color
                                    YellowCount = YellowCount + 1
                                Case RGB(255, 0, 0) ' Red color
                                    RedCount = RedCount + 1
                            End Select
                        End If
                    End If
                ' ��������� ������ �� ������ �� ����
                ElseIf WhatFilter = "D" Then
                    If Year(Sheet.Cells(i, "B").Value) = targetYear Then
                        '���� ����� ������ ���� ������� ����� ����
                        If Sheet.Cells(i, "E").MergeCells Then
                            'MsgBox "��� ������ ����������."
                            
                            Set mergedRange = Sheet.Cells(i, "E").MergeArea ' �������� ������������ ��������
                        
                            ' Check cell color
                            Select Case Sheet.Cells(i, "E").Interior.Color
                                Case RGB(0, 176, 80) ' Green color
                                    GreenCount = GreenCount + 1
                                Case RGB(255, 192, 0) ' Orange color
                                    OrangeCount = OrangeCount + 1
                                Case RGB(255, 255, 0) ' Yellow color
                                    YellowCount = YellowCount + 1
                                Case RGB(255, 0, 0) ' Red color
                                    RedCount = RedCount + 1
                            End Select
                                
                            ' ���������� ������������ ������
                            i = i + mergedRange.Cells.Count - 1
                        Else
                            'MsgBox "��� ������ �� ����������."
                            ' Check cell color
                            Select Case Sheet.Cells(i, "E").Interior.Color
                                Case RGB(0, 176, 80) ' Green color
                                    GreenCount = GreenCount + 1
                                Case RGB(255, 192, 0) ' Orange color
                                    OrangeCount = OrangeCount + 1
                                Case RGB(255, 255, 0) ' Yellow color
                                    YellowCount = YellowCount + 1
                                Case RGB(255, 0, 0) ' Red color
                                    RedCount = RedCount + 1
                            End Select
                        End If
                    End If
                ' ��������� ������ �� ������ �� �������
                ElseIf WhatFilter = "C" Then
                    ' ��������� ��� ���� � ������� B �������� � ��� ������
                    If Sheet.Cells(i, "B").Value >= OneDate And Sheet.Cells(i, "B").Value <= TwoDate Then
                        '���� ����� ������ ���� ������� ����� ����
                        If Sheet.Cells(i, "E").MergeCells Then
                            'MsgBox "��� ������ ����������."
                            
                            Set mergedRange = Sheet.Cells(i, "E").MergeArea ' �������� ������������ ��������
                        
                            ' Check cell color
                            Select Case Sheet.Cells(i, "E").Interior.Color
                                Case RGB(0, 176, 80) ' Green color
                                    GreenCount = GreenCount + 1
                                Case RGB(255, 192, 0) ' Orange color
                                    OrangeCount = OrangeCount + 1
                                Case RGB(255, 255, 0) ' Yellow color
                                    YellowCount = YellowCount + 1
                                Case RGB(255, 0, 0) ' Red color
                                    RedCount = RedCount + 1
                            End Select
                                
                            ' ���������� ������������ ������
                            i = i + mergedRange.Cells.Count - 1
                        Else
                            'MsgBox "��� ������ �� ����������."
                            ' Check cell color
                            Select Case Sheet.Cells(i, "E").Interior.Color
                                Case RGB(0, 176, 80) ' Green color
                                    GreenCount = GreenCount + 1
                                Case RGB(255, 192, 0) ' Orange color
                                    OrangeCount = OrangeCount + 1
                                Case RGB(255, 255, 0) ' Yellow color
                                    YellowCount = YellowCount + 1
                                Case RGB(255, 0, 0) ' Red color
                                    RedCount = RedCount + 1
                            End Select
                        End If
                    End If
                End If
            End If
        Next i
        
        ' �������� ��������� FCount ��� �������� F ����� �� �������� � �������� �� ���������
        FCount Sheet, targetMonth, targetYear, WhatFilter, ws, lastRow, lastRowFile, OneDate, TwoDate, FGreenDoneRowCount, _
            FGreenInProgressCount, FGreenOverdueCount, FOrangeDoneRowCount, _
            FOrangeInProgressCount, FOrangeOverdueCount, FYellowDoneRowCount, _
            FYellowInProgressCount, FYellowOverdueCount, FRedDoneRowCount, _
            FRedInProgressCount, FRedOverdueCount

        ' �������� ��������� MeanCountColorF ��� �������� F ����� � �������� �� ���������
        MeanCountColorF Sheet, ws, lastRow, TotalGreenCount, TotalOrangeCount, TotalYellowCount, TotalRedCount, AverageGreen, _
                    AverageOrange, AverageYellow, AverageRed, FGreenRowCountList, FOrangeRowCountList, FYellowRowCountList, _
                    FRedRowCountList

        ' �������� ��������� CountDatesB ��� �������� ���������� ��� � �������� �� ���������
        CountDatesB Sheet, ws, lastRow, UniqueDates, BValue, targetMonth, targetYear, WhatFilter, OneDate, TwoDate

        OtherWorkbook.Close ' ��������� ����� ����� �������������
    Next j

    ws.Range("B3").Value = GreenCount
    ws.Range("B4").Value = OrangeCount
    ws.Range("B5").Value = YellowCount
    ws.Range("B6").Value = RedCount

    ws.Range("B7").Value = FGreenDoneRowCount
    ws.Range("B8").Value = FGreenInProgressCount
    ws.Range("B9").Value = FGreenOverdueCount
    
    ws.Range("B10").Value = FOrangeDoneRowCount
    ws.Range("B11").Value = FOrangeInProgressCount
    ws.Range("B12").Value = FOrangeOverdueCount
    
    ws.Range("B13").Value = FYellowDoneRowCount
    ws.Range("B14").Value = FYellowInProgressCount
    ws.Range("B15").Value = FYellowOverdueCount
    
    ws.Range("B16").Value = FRedDoneRowCount
    ws.Range("B17").Value = FRedInProgressCount
    ws.Range("B18").Value = FRedOverdueCount
    
    ' ������������ �������� � ���������
    For i = 1 To FGreenRowCountList.Count
        TotalGreenCount = TotalGreenCount + FGreenRowCountList(i)
    Next i

    ' �������� ������� ��������� � ��������� ����� ��������
    If FGreenRowCountList.Count > 0 Then
        ' ���������� �������� ���������������
        AverageGreen = TotalGreenCount / FGreenRowCountList.Count
    Else
        MsgBox "������� ��������� �����, ���������� ��������� ������� ��������������.", vbExclamation
    End If
    
    ' ������������ �������� � ���������
    For i = 1 To FOrangeRowCountList.Count
        TotalOrangeCount = TotalOrangeCount + FOrangeRowCountList(i)
    Next i

    ' �������� ������� ��������� � ��������� ����� ��������
    If FOrangeRowCountList.Count > 0 Then
        ' ���������� �������� ���������������
        AverageOrange = TotalOrangeCount / FOrangeRowCountList.Count
    Else
        MsgBox "��������� ��������� �����, ���������� ��������� ������� ��������������.", vbExclamation
    End If
    
    ' ������������ �������� � ���������
    For i = 1 To FYellowRowCountList.Count
        TotalYellowCount = TotalYellowCount + FYellowRowCountList(i)
    Next i

    ' �������� ������� ��������� � ��������� ����� ��������
    If FYellowRowCountList.Count > 0 Then
        ' ���������� �������� ���������������
        AverageYellow = TotalYellowCount / FYellowRowCountList.Count
    Else
        MsgBox "������ ��������� �����, ���������� ��������� ������� ��������������.", vbExclamation
    End If
    
    ' ������������ �������� � ���������
    For i = 1 To FRedRowCountList.Count
        TotalRedCount = TotalRedCount + FRedRowCountList(i)
    Next i

    ' �������� ������� ��������� � ��������� ����� ��������
    If FRedRowCountList.Count > 0 Then
        ' ���������� �������� ���������������
        AverageRed = TotalRedCount / FRedRowCountList.Count
    Else
        MsgBox "������� ��������� �����, ���������� ��������� ������� ��������������.", vbExclamation
    End If

    ' ������� ����� ���������� ������� �����������
    ws.Range("B23").Value = TotalGreenCount
    ' ������� ����� ���������� ��������� �����������
    ws.Range("B24").Value = TotalOrangeCount
    ' ������� ����� ���������� ������ �����������
    ws.Range("B25").Value = TotalYellowCount
    ' ������� ����� ���������� ������� �����������
    ws.Range("B26").Value = TotalRedCount

    ' ������� ������� �������� �� ���������� ����� ������� �����
    ws.Range("B19").Value = AverageGreen
    ws.Range("B20").Value = AverageOrange
    ws.Range("B21").Value = AverageYellow
    ws.Range("B22").Value = AverageRed

    ' ����� ���������� ���������� ���
    ws.Range("B27").Value = UniqueDates.Count
    
End Sub

Sub FCount(Sheet As Worksheet, targetMonth As Integer, targetYear As Integer, WhatFilter As String, ws As Worksheet, _
            lastRow As Long, lastRowFile As Integer, OneDate As Date, TwoDate As Date, ByRef FGreenDoneRowCount As Integer, _
            ByRef FGreenInProgressCount As Integer, ByRef FGreenOverdueCount As Integer, ByRef FOrangeDoneRowCount As Integer, _
            ByRef FOrangeInProgressCount As Integer, ByRef FOrangeOverdueCount As Integer, ByRef FYellowDoneRowCount As Integer, _
            ByRef FYellowInProgressCount As Integer, ByRef FYellowOverdueCount As Integer, ByRef FRedDoneRowCount As Integer, _
            ByRef FRedInProgressCount As Integer, ByRef FRedOverdueCount As Integer)

    Dim Cell As Range
    Dim currentDate As Date
    
    ' ������ ���������� ��� �������� ���������� ����� � ������������ ������
    Dim mergedRange As Range
    
    ' ��������� �������� �� ��������� ������ ������������
    If Sheet.Cells(lastRow, "E").MergeCells Then
        Set mergedRange = Sheet.Cells(lastRow, "E").MergeArea ' �������� ������������ ��������
        lastRow = lastRow + mergedRange.Cells.Count - 1 ' ��������� � ������ ��������� ������ ���������� ������������ �����
    End If
    
    ' ��������� ������� ����
    currentDate = Date
    ' ������� ���������� ������� ����� � ������� E, ��������������� ��������� ���� � ������� B
        Dim i As Long
        For i = 4 To lastRow ' ������� � 4 ������ � �� ��������� ����������� ������
            ' ���������, ��� ������ B �������� ���� � �� �������� ������
            If Sheet.Cells(i, "B").MergeCells Then
                If IsDate(Sheet.Cells(i, "B").MergeArea.Cells(1, 1).Value) Then
                    ' ��������� ������ �� ������ �� ������ � ����
                    If WhatFilter = "B" Then
                        If Month(Sheet.Cells(i, "B").MergeArea.Cells(1, 1).Value) = targetMonth And Year(Sheet.Cells(i, "B").MergeArea.Cells(1, 1).Value) = targetYear Then
                            '���� ����� ������ ���� ������� ����� ����
                            Select Case Sheet.Cells(i, "E").Interior.Color
                                Case RGB(0, 176, 80) ' Green color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FGreenDoneRowCount = FGreenDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FGreenInProgressCount = FGreenInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FGreenOverdueCount = FGreenOverdueCount + 1
                                    End If
                                Case RGB(255, 192, 0) ' Orange color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FOrangeDoneRowCount = FOrangeDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FOrangeInProgressCount = FOrangeInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FOrangeOverdueCount = FOrangeOverdueCount + 1
                                    End If
                                Case RGB(255, 255, 0) ' Yellow color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FYellowDoneRowCount = FYellowDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FYellowInProgressCount = FYellowInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FYellowOverdueCount = FYellowOverdueCount + 1
                                    End If
                                Case RGB(255, 0, 0) ' Red color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FRedDoneRowCount = FRedDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FRedInProgressCount = FRedInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FRedOverdueCount = FRedOverdueCount + 1
                                    End If
                            End Select
                        End If
                    ' ��������� ������ �� ������ �� ����
                    ElseIf WhatFilter = "D" Then
                        If Year(Sheet.Cells(i, "B").MergeArea.Cells(1, 1).Value) = targetYear Then
                            '���� ����� ������ ���� ������� ����� ����
                            Select Case Sheet.Cells(i, "E").Interior.Color
                                Case RGB(0, 176, 80) ' Green color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FGreenDoneRowCount = FGreenDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FGreenInProgressCount = FGreenInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FGreenOverdueCount = FGreenOverdueCount + 1
                                    End If
                                Case RGB(255, 192, 0) ' Orange color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FOrangeDoneRowCount = FOrangeDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FOrangeInProgressCount = FOrangeInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FOrangeOverdueCount = FOrangeOverdueCount + 1
                                    End If
                                Case RGB(255, 255, 0) ' Yellow color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FYellowDoneRowCount = FYellowDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FYellowInProgressCount = FYellowInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FYellowOverdueCount = FYellowOverdueCount + 1
                                    End If
                                Case RGB(255, 0, 0) ' Red color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FRedDoneRowCount = FRedDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FRedInProgressCount = FRedInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FRedOverdueCount = FRedOverdueCount + 1
                                    End If
                            End Select
                        End If
                    ' ��������� ������ �� ������ �� �������
                    ElseIf WhatFilter = "C" Then
                        ' ��������� ��� ���� � ������� B �������� � ��� ������
                        If Sheet.Cells(i, "B").MergeArea.Cells(1, 1).Value >= OneDate And Sheet.Cells(i, "B").MergeArea.Cells(1, 1).Value <= TwoDate Then
                            '���� ����� ������ ���� ������� ����� ����
                            Select Case Sheet.Cells(i, "E").Interior.Color
                                Case RGB(0, 176, 80) ' Green color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FGreenDoneRowCount = FGreenDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FGreenInProgressCount = FGreenInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FGreenOverdueCount = FGreenOverdueCount + 1
                                    End If
                                Case RGB(255, 192, 0) ' Orange color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FOrangeDoneRowCount = FOrangeDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FOrangeInProgressCount = FOrangeInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FOrangeOverdueCount = FOrangeOverdueCount + 1
                                    End If
                                Case RGB(255, 255, 0) ' Yellow color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FYellowDoneRowCount = FYellowDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FYellowInProgressCount = FYellowInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FYellowOverdueCount = FYellowOverdueCount + 1
                                    End If
                                Case RGB(255, 0, 0) ' Red color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FRedDoneRowCount = FRedDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FRedInProgressCount = FRedInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FRedOverdueCount = FRedOverdueCount + 1
                                    End If
                            End Select
                        End If
                    End If
                End If
            Else
                If IsDate(Sheet.Cells(i, "B").MergeArea.Cells(1, 1).Value) Then
                    ' ��������� ������ �� ������ �� ������ � ����
                    If WhatFilter = "B" Then
                        If Month(Sheet.Cells(i, "B").Value) = targetMonth And Year(Sheet.Cells(i, "B").Value) = targetYear Then
                            '���� ����� ������ ���� ������� ����� ����
                            Select Case Sheet.Cells(i, "E").Interior.Color
                                Case RGB(0, 176, 80) ' Green color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FGreenDoneRowCount = FGreenDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FGreenInProgressCount = FGreenInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FGreenOverdueCount = FGreenOverdueCount + 1
                                    End If
                                Case RGB(255, 192, 0) ' Orange color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FOrangeDoneRowCount = FOrangeDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FOrangeInProgressCount = FOrangeInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FOrangeOverdueCount = FOrangeOverdueCount + 1
                                    End If
                                Case RGB(255, 255, 0) ' Yellow color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FYellowDoneRowCount = FYellowDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FYellowInProgressCount = FYellowInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FYellowOverdueCount = FYellowOverdueCount + 1
                                    End If
                                Case RGB(255, 0, 0) ' Red color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FRedDoneRowCount = FRedDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FRedInProgressCount = FRedInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FRedOverdueCount = FRedOverdueCount + 1
                                    End If
                            End Select
                        End If
                    ' ��������� ������ �� ������ �� ����
                    ElseIf WhatFilter = "D" Then
                        If Year(Sheet.Cells(i, "B").Value) = targetYear Then
                            '���� ����� ������ ���� ������� ����� ����
                            Select Case Sheet.Cells(i, "E").Interior.Color
                                Case RGB(0, 176, 80) ' Green color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FGreenDoneRowCount = FGreenDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FGreenInProgressCount = FGreenInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FGreenOverdueCount = FGreenOverdueCount + 1
                                    End If
                                Case RGB(255, 192, 0) ' Orange color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FOrangeDoneRowCount = FOrangeDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FOrangeInProgressCount = FOrangeInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FOrangeOverdueCount = FOrangeOverdueCount + 1
                                    End If
                                Case RGB(255, 255, 0) ' Yellow color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FYellowDoneRowCount = FYellowDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FYellowInProgressCount = FYellowInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FYellowOverdueCount = FYellowOverdueCount + 1
                                    End If
                                Case RGB(255, 0, 0) ' Red color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FRedDoneRowCount = FRedDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FRedInProgressCount = FRedInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FRedOverdueCount = FRedOverdueCount + 1
                                    End If
                            End Select
                        End If
                    ' ��������� ������ �� ������ �� �������
                    ElseIf WhatFilter = "C" Then
                        ' ��������� ��� ���� � ������� B �������� � ��� ������
                        If Sheet.Cells(i, "B").Value >= OneDate And Sheet.Cells(i, "B").Value <= TwoDate Then
                            '���� ����� ������ ���� ������� ����� ����
                            Select Case Sheet.Cells(i, "E").Interior.Color
                                Case RGB(0, 176, 80) ' Green color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FGreenDoneRowCount = FGreenDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FGreenInProgressCount = FGreenInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FGreenOverdueCount = FGreenOverdueCount + 1
                                    End If
                                Case RGB(255, 192, 0) ' Orange color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FOrangeDoneRowCount = FOrangeDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FOrangeInProgressCount = FOrangeInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FOrangeOverdueCount = FOrangeOverdueCount + 1
                                    End If
                                Case RGB(255, 255, 0) ' Yellow color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FYellowDoneRowCount = FYellowDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FYellowInProgressCount = FYellowInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FYellowOverdueCount = FYellowOverdueCount + 1
                                    End If
                                Case RGB(255, 0, 0) ' Red color
                                    '������� ������ �� ������ � �������� � F
                                    If UCase(Sheet.Cells(i, "I").Value) = UCase("���������") Then
                                        FRedDoneRowCount = FRedDoneRowCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate < Sheet.Cells(i, "G").Value Then
                                        FRedInProgressCount = FRedInProgressCount + 1
                                    ElseIf IsDate(Sheet.Cells(i, "G").Value) And currentDate > Sheet.Cells(i, "G").Value Then
                                        FRedOverdueCount = FRedOverdueCount + 1
                                    End If
                            End Select
                        End If
                    End If
                End If
            End If
        Next i
End Sub

Sub MeanCountColorF(Sheet As Worksheet, ws As Worksheet, lastRow As Long, ByRef TotalGreenCount As Integer, _
            ByRef TotalOrangeCount As Integer, ByRef TotalYellowCount As Integer, TotalRedCount As Integer, _
            ByRef AverageGreen As Double, ByRef AverageOrange As Double, ByRef AverageYellow As Double, ByRef AverageRed As Double, _
            ByRef FGreenRowCountList As Collection, ByRef FOrangeRowCountList As Collection, ByRef FYellowRowCountList As Collection, _
            ByRef FRedRowCountList As Collection)

    Dim Cell As Range
    Dim mergedRange As Range
    Dim k As Integer
    
    ' ������� ���������� ������� ����� � ������� E, ��������������� ��������� ���� � ������� B
    Dim i As Long
    For i = 4 To lastRow ' ������� � 4 ������ � �� ��������� ����������� ������
        ' ���������, ��� ������ ������������
        If Sheet.Cells(i, "E").MergeCells Then
            Set mergedRange = Sheet.Cells(i, "E").MergeArea ' �������� ������������ ��������
            
            ' ������� ����� ����
            Select Case Sheet.Cells(i, "E").Interior.Color
                Case RGB(0, 176, 80) ' Green color
                    FGreenRowCountList.Add mergedRange.Cells.Count
                Case RGB(255, 192, 0) ' Orange color
                    FOrangeRowCountList.Add mergedRange.Cells.Count
                Case RGB(255, 255, 0) ' Yellow color
                    FYellowRowCountList.Add mergedRange.Cells.Count
                Case RGB(255, 0, 0) ' Red color
                    FRedRowCountList.Add mergedRange.Cells.Count
            End Select
            
            ' ���������� ������������ ������
            i = i + mergedRange.Cells.Count - 1
        Else
            '������� ����� ����
            Select Case Sheet.Cells(i, "E").Interior.Color
                Case RGB(0, 176, 80) ' Green color
                    FGreenRowCountList.Add 1
                Case RGB(255, 192, 0) ' Orange color
                    FOrangeRowCountList.Add 1
                Case RGB(255, 255, 0) ' Yellow color
                    FYellowRowCountList.Add 1
                Case RGB(255, 0, 0) ' Red color
                    FRedRowCountList.Add 1
            End Select
        End If
    Next i
End Sub

Sub CountDatesB(Sheet As Worksheet, ws As Worksheet, lastRow As Long, ByRef UniqueDates As Object, ByRef BValue As Collection, _
                targetMonth As Integer, targetYear As Integer, WhatFilter As String, OneDate As Date, TwoDate As Date)

    ' ��������� ���� ������� � 4 ������ � �� ��������� ����������� ������
    Dim i As Long
    For i = 4 To lastRow
        ' ���������, ��� ������ B �������� ���� � �� �������� ������
        If IsDate(Sheet.Cells(i, "B").Value) Then
            ' ��������� ������ �� ������ �� ������ � ����
            If WhatFilter = "B" Then
                If Month(Sheet.Cells(i, "B").MergeArea.Cells(1, 1).Value) = targetMonth And Year(Sheet.Cells(i, "B").MergeArea.Cells(1, 1).Value) = targetYear Then
                    BValue.Add Sheet.Cells(i, "B").Value
                End If
            ' ��������� ������ �� ������ �� ����
            ElseIf WhatFilter = "D" Then
                If Year(Sheet.Cells(i, "B").MergeArea.Cells(1, 1).Value) = targetYear Then
                    BValue.Add Sheet.Cells(i, "B").Value
                End If
            ' ��������� ������ �� ������ �� �������
            ElseIf WhatFilter = "C" Then
                ' ��������� ��� ���� � ������� B �������� � ��� ������
                If Sheet.Cells(i, "B").MergeArea.Cells(1, 1).Value >= OneDate And Sheet.Cells(i, "B").MergeArea.Cells(1, 1).Value <= TwoDate Then
                    BValue.Add Sheet.Cells(i, "B").Value
                End If
            End If
        End If
    Next i

    ' �������� ���������� ���
    For i = 1 To BValue.Count
        UniqueDates(CDate(BValue(i))) = 1
    Next i

End Sub

