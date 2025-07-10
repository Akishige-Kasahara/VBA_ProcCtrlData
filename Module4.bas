Attribute VB_Name = "Module4"
Option Explicit

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'�q�X�g�O�����̍쐬
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Public Sub MakeHistogram()

    Dim rTarget As Range ' �I��͈�
    Dim Tmp
    Dim sh1 As Worksheet
    Dim sh2 As Worksheet
    Dim sh3 As Worksheet
    
    Dim lSize As Long
    Dim dMax As Double
    Dim dMin As Double
    Dim dMea As Double
    Dim dSd As Double
    Dim dUl As Double
    Dim dLl As Double
    
    Set sh1 = Worksheets("�H���Ǘ��p�f�[�^")
    Set sh2 = Worksheets("�q�X�g�O����")
    Set sh3 = Worksheets("�H���Ǘ��\")

    With sh1
        Set rTarget = .Range(.Cells(2, 3), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3))
    End With
    
    sh2.Cells.Clear

    With Application.WorksheetFunction
    
        lSize = .Count(rTarget)
        
        With sh3
            dMax = .Cells(7, 11)
            dMin = .Cells(8, 11)
            dMea = .Cells(9, 11)
            dSd = .Cells(9, 8)
            dUl = .Cells(7, 8)
            dLl = .Cells(8, 8)
            
            '�s��i�𖳎�����ꍇ�́A�ő�l������K�i�A�ŏ��l�������K�i�ɂ���
            'dMax = dUl
            'dMin = dLl
        End With
        
        '�e��A���S���Y���ɂ��K�����̎Z�o
        Dim dScott As Double, dSturges As Double, dFd As Double, dSqRoot As Double
        
        dScott = 3.5 * dSd / lSize ^ (1 / 3)
        dSturges = (dMax - dMin) / (1 + (.Log(lSize) / .Log(2)))
        dFd = 2 * (.Quartile_Inc(rTarget, 3) - .Quartile_Inc(rTarget, 1)) / lSize ^ (1 / 3)
        dSqRoot = (dMax - dMin) / lSize ^ (1 / 2)
        
        Dim dClassWidth As Double, d1stClass As Double

        ' ��1�K���̉����ƊK�����̌���
'        dClassWidth = BIN_WIDTH(dScott)
'        dClassWidth = BIN_WIDTH(dSturges)
'        dClassWidth = BIN_WIDTH(dFd)
        dClassWidth = BIN_WIDTH(dSqRoot)
        Tmp = .Floor(dMin, dClassWidth)
        If dMin <> Tmp Then
            d1stClass = Tmp
        Else
            d1stClass = Tmp - dClassWidth ' min = h �̂Ƃ��̑�1�K�������̏C��
        End If
    End With
    
    Dim aArry As Variant
    Dim fMax As Long ' �ő�x��
    Dim i As Long
        
    ReDim aArry(1 To ((dMax - d1stClass) / dClassWidth) + 2, 1 To 3)
    
    fMax = 0
    
    For i = 1 To UBound(aArry, 1)
        
        Select Case i
            Case 1 ' ��1�K�������E
                aArry(i, 1) = d1stClass

            Case Else ' ��2�K���ȍ~�̉����E
                aArry(i, 1) = aArry(i - 1, 2)
                                  
        End Select
        
        aArry(i, 2) = aArry(i, 1) + dClassWidth ' �㋫�E
        
        '�x��
        aArry(i, 3) = WorksheetFunction.CountIfs(rTarget, ">=" & aArry(i, 1), rTarget, "<" & aArry(i, 2))
        
        '�ő�x��
        If aArry(i, 3) > fMax Then
            fMax = aArry(i, 3)
        End If

    Next i
    
    
    '���K���z�̎Z�o
    Dim bArry As Variant
    Dim dClassMax As Double, dClassMin As Double

    dClassMin = aArry(1, 1) '�K���̍ŏ��l
    dClassMax = aArry(UBound(aArry, 1), 2)  '�K���̍ő�l
    
    ReDim bArry(1 To 102, 1 To 2)
    For i = 1 To UBound(bArry, 1)
        Select Case i
            Case 1
                bArry(1, 1) = dClassMin
            Case Else
                bArry(i, 1) = bArry(i - 1, 1) + (dClassMax - dClassMin) / 100
        End Select
        
        bArry(i, 2) = WorksheetFunction.Norm_Dist(bArry(i, 1), dMea, dSd, False) * lSize * dClassWidth
    Next i
    
    With sh2
        .Cells(1, 1) = "�ŏ��̊K��"
        .Cells(1, 2) = d1stClass
        .Cells(2, 1) = "�K���̕�"
        .Cells(2, 2) = dClassWidth
        .Cells(3, 1) = "���x�����z�\"
        .Cells(4, 1) = "�����E"
        .Cells(4, 2) = "�㋫�E"
        .Cells(4, 3) = "�x��"
        .Cells(4, 4) = "�ő�x��"
        .Cells(5, 4) = fMax
        .Cells(4, 5) = "���ғx��1"
        .Cells(4, 6) = "���ғx��2"
        
        
        .Range("A5").Resize(UBound(aArry, 1), UBound(aArry, 2)) = aArry
        .Range("E5").Resize(UBound(bArry, 1), UBound(bArry, 2)) = bArry
        
    End With
    
    
    '�����O���t�̏���
    With sh2
        For i = .ChartObjects.Count To 1 Step -1
            .ChartObjects(i).Delete
        Next i
    End With
    
    
    sh2.Activate
    
    '�O���t
            
    Dim gRange As Range, R As Range
    
    '�O���t�Ɏg�p����f�[�^��ݒ�
    lSize = Cells(Rows.Count, 2).End(xlUp).Row
    Range(Cells(5, 3), Cells(lSize, 4)).Select
    Set R = Range("G4:P30")
        
    
    ActiveSheet.Shapes.AddChart(xlColumnClustered).Select ' �W���c�_�O���t���쐬

    With ActiveChart
        .HasLegend = False                      ' �}�Ꮬ��
        .ChartGroups(1).GapWidth = 0            ' �Ԋu=0
        With .SeriesCollection(1)
            .AxisGroup = 2                      ' ����2��
'            .Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent3 ' ���̐F���O���[
            .Border.Color = vbWhite             ' ���̊O�g���F����
            .Border.Weight = xlThin             ' ���O�g�̑���
        End With
        With .SeriesCollection(2)
            .ChartType = xlXYScatter            ' �ő�x�����U�z�}��
            .MarkerStyle = xlMarkerStyleNone    ' �}�[�J�[��s����
        End With
        .Parent.Top = Range("G4").Top           ' �ʒu�����i��[�j
        .Parent.Left = Range("G4").Left         ' �ʒu�����i���[�j
        .Parent.Width = 540         ' �ʒu�����i���j
        .Parent.Height = 300        ' �ʒu�����i�����j
    End With
    
    
    ' ���K���z�Ȑ��̒ǉ�
    lSize = Cells(Rows.Count, 5).End(xlUp).Row
    
    With ActiveChart.SeriesCollection.NewSeries ' Normal_Curve�n��̒ǉ�
        .XValues = Range(Cells(5, 5), Cells(lSize, 5))
        .Values = Range(Cells(5, 6), Cells(lSize, 6))
        .Name = "���K���z�Ȑ�"
        .Border.Color = RGB(50, 50, 50)
        .Format.Line.Weight = 1 'pt
    End With

    ActiveChart.SeriesCollection("���K���z�Ȑ�").ChartType = xlXYScatterSmoothNoMarkers ' ��������

    With ActiveChart.Axes(xlCategory)
        .MinimumScale = dClassMin              ' ���X�P�[�����킹�i�ŏ��l�j
        .MaximumScale = dClassMax              ' ���X�P�[�����킹�i�ő�l�j
        .MajorUnit = 2                 ' ���X�P�[�����킹�i�ڐ���j
        .CrossesAt = dClassMin                 ' ���X�P�[�����킹�i��_�j
    End With
    
    With ActiveChart.Axes(xlValue, xlSecondary)
        .MinimumScale = 0
        .TickLabelPosition = xlNone         ' 2�����x����s����
        .MajorTickMark = xlNone             ' 2���ڐ���s����
    End With
    
    With ActiveSheet.ChartObjects(1).Chart
        .Axes(xlValue).MinimumScale = 0
    End With
    
    
    
    Cells(2, 2).Select
    

End Sub


Private Function BIN_WIDTH(h)
    ' �K�����𒲐�����

    Dim n As Long
    Dim Stp(2) ' �����ߒ�  step1 to 3
    Dim TryVal
    Dim Tmp ' �l

    TryVal = Array(5, 2, 1) ' Mround, Ceiling �̊�l�Ɋ|����E�G�C�g

    With Application.WorksheetFunction
    Select Case h
    Case Is <= 0
        MsgBox "ERROR"
        Exit Function

    Case Is >= 1 ' h��1�ȏ�̏ꍇ�̏���
        n = -1
        Do
            n = n + 1
            Stp(0) = 10 ^ n
            Stp(1) = h / Stp(0)
        Loop Until Stp(1) <= 10
    
        n = 0
        Do
            If n < 2 Then
                Tmp = .MRound(h, Stp(0) * TryVal(n))
            Else
                Tmp = .Ceiling(h, Stp(0) * TryVal(n))
            End If
            n = n + 1
        Loop Until Tmp <> 0

    Case Is < 1 ' h��1��菬���ȏꍇ�̏���
        n = -1
        Do
            n = n + 1
            Stp(0) = 10 ^ n
            Stp(1) = 1 / (Stp(0) * 10)
            Stp(2) = h / Stp(1)
        Loop Until Stp(2) >= 1
    
        n = 0
        Do
            If n < 2 Then
                Tmp = .MRound(h, Stp(1) * TryVal(n))
            Else
                Tmp = .Ceiling(h, Stp(1) * TryVal(n))
            End If
            n = n + 1
        Loop Until Tmp <> 0
        
    
    End Select
    End With

    BIN_WIDTH = Tmp

End Function

Private Sub Err1()
    MsgBox "���o���������I��͈͂ɐ��l�ȊO���܂܂�Ă��܂�"
End Sub

Private Sub Err2()
    MsgBox "�����Ώۂ̃V�[�g���A�N�e�B�u�ɂȂ��Ă��܂���"
End Sub

Private Sub Err3()
    MsgBox "�A�N�e�B�u�ȃO���t������܂���"
End Sub

Private Sub Err4()
    MsgBox "�K�����Ɏw�肳��Ă�����e�����l�ł͂���܂���"
End Sub

Private Sub Err5()
    MsgBox "�ŏ��̊K���̉����E�Ɏw�肳��Ă�����e�����l�ł͂���܂���"
End Sub

Private Sub Err6()
    MsgBox "�ŏ��̊K���̉����E�ɑÓ��łȂ��l���ݒ肳��Ă��܂�" & vbCrLf & _
        "�ϐ��̃����W���J�o�[����ɂ́C���̒l���ŏ��l��菬���Ȓl�ł���K�v������܂�"
End Sub

Private Function Err_Checker1(myRange As Range) As Boolean
' �I��͈͂ɐ��l�ȊO���܂܂�Ă��邩

Dim cc
Err_Checker1 = False

For Each cc In myRange
    If IsNumeric(cc) = False Or cc = "" Then
        Err_Checker1 = True
        Call Err1
    End If
Next

End Function

Private Function Err_Checker2() As Boolean
' �uOP-�v�Ŏn�܂�V�[�g���A�N�e�B�u��

Err_Checker2 = False
If Left(ActiveSheet.Name, 3) <> "OP-" Then
    Err_Checker2 = True
    Call Err2
End If

End Function

Private Function Err_Checker3() As Boolean
' �O���t���A�N�e�B�u�ɂȂ��Ă��邩

If ActiveChart Is Nothing Then
    Err_Checker3 = True
    Call Err3
Else
    Err_Checker3 = False
End If

End Function

Private Function Err_Checker4(s) As Boolean
' ���l���ǂ���

If IsNumeric(s) = False Or _
    s = "" Then
    Err_Checker4 = True
Else
    Err_Checker4 = False
End If

End Function



