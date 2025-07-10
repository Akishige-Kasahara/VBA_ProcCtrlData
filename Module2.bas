Attribute VB_Name = "Module2"
Option Explicit

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' ProcCtrlData()
'   �H���Ǘ��\�̐���
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Sub ProcCtrlData()
    
    Dim sTime As Double
    Dim eTime As Double
    Dim pTime As Double

    sTime = Timer

    Dim aArry As Variant
    Dim lRow As Long
    Dim lCnt As Long
    Dim sh1 As Worksheet

    Set sh1 = Worksheets("Main1")
    
    '�`����~����
    Application.ScreenUpdating = False


    '���M���O�f�[�^��aArry�ɑ������
    With sh1
        .Activate
        .Cells(1, 1).Select
        lRow = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim aArry(1 To lRow, 1 To 6)
'        aArry = .Range(.Cells(13, 1), .Cells(lRow + 1, 6))
        aArry = .Range(.Cells(13, 1), .Cells(lRow, 6))
    End With
    
    lCnt = 1
    
    '�H���Ǘ��p�Ƀ��M���O�f�[�^�𐬌`����
    Call FormData(aArry, lCnt)
        
    '�J�E���^�l�Ń\�[�g
    Call merge_sort2(aArry, 2)
    
    'X-R�f�[�^�V�[�g�̍쐬
    Call FormXRData(aArry)
    'X-R�O���t�̍쐬
    Call DrawXRChart
    
    
    '�H���Ǘ��f�[�^�V�[�g�̍쐬
    Call MakeProcCtrlDataSheet(aArry, lCnt)
    
    '�q�X�g�O�����̍쐬
    Call MakeHistogram
        
    '�H���Ǘ��\�̃O���t�̍쐬
    Call MakeProcCtrlChart
    
    '�H���Ǘ��\�V�[�g�ֈ���t�H�[�}�b�g�ɂ����f�[�^��������
    Call FormProcMeasData4Printing
    
    '�`������ɖ߂�
    Application.ScreenUpdating = True
    
    eTime = Timer
    pTime = eTime - sTime
    
    Debug.Print "��������"; pTime & vbCrLf
    
End Sub

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' FormData(ByRef aArry As Variant, ByRef lCnt As Long)
'   ���M���O�f�[�^���H���Ǘ��p�f�[�^�ɐ��`����
'
'   aArry�F���M���O�f�[�^�ː��`���Ė߂�
'   lCnt�F�����J�E���^
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Private Sub FormData(ByRef aArry As Variant, ByRef lCnt As Long)

    Dim lSizeA As Long
    Dim i As Integer
    Dim j As Integer
    Dim bNG1 As Boolean
    Dim bNG2 As Boolean
    Dim bNG3 As Boolean
    Dim bNG4 As Boolean
    Dim sh As Worksheet

    Set sh = Worksheets("Main1")
    
    With sh
        bNG1 = False
        bNG2 = False
        bNG3 = False
        bNG4 = False
        
        If .Cells(5, 8) = "�܂�" Then bNG1 = True
        If .Cells(6, 8) = "�܂�" Then bNG2 = True
        If .Cells(7, 8) = "�܂�" Then bNG3 = True
        If .Cells(8, 8) = "�܂�" Then bNG4 = True
    
    End With
    
    ' �J�E���^�l�̏�����
    lCnt = 1
    
    lSizeA = UBound(aArry, 1)
    
    For i = 1 To lSizeA
        If i < lSizeA Then
            For j = 1 To 6
                aArry(lCnt, j) = aArry(i, j)
            Next j

        End If
            
        Select Case Mid(aArry(i, 6), 1, 1)
            Case "A"
                If bNG1 Then
                    lCnt = lCnt + 1
                End If
                    
            Case "B"
                If bNG2 Then
                    lCnt = lCnt + 1
                End If
                
            Case "C"
                If bNG3 Then
                    lCnt = lCnt + 1
                End If
            
            Case "D"
                If bNG4 Then
                    lCnt = lCnt + 1
                End If
            Case Else
                lCnt = lCnt + 1
                
        End Select
                
    Next i

    '�z���]�u����
    aArry = WorksheetFunction.Transpose(aArry)
    
    '�폜�����f�[�^���A����l�߂�
    ReDim Preserve aArry(1 To UBound(aArry, 1), 1 To UBound(aArry, 2) - (lSizeA - lCnt) - 1)
    
    '�z���]�u������
    aArry = WorksheetFunction.Transpose(aArry)

End Sub


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' MakeProcCtrlDataSheet(ByRef aArry As Variant, lSize As Long)
'   "�H���Ǘ��p�f�[�^"�V�[�g�̍쐬
'
'   aArry�F�H���Ǘ��p�f�[�^
'   lCnt�F�H���Ǘ��p�f�[�^�������̏����J�E���^�l
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Private Sub MakeProcCtrlDataSheet(ByRef aArry As Variant, lSize As Long)
    
    Dim sh1 As Worksheet
    Dim sh2 As Worksheet
    Dim sNG1 As String
    Dim sNG2 As String
    Dim sNG3 As String
    Dim bArry As Variant
    Dim cArry As Variant
    Dim dArry As Variant
    Dim i As Integer, j As Integer
    Dim lRow As Long
    Dim dUl As Double
    Dim dLl As Double
    Dim dMean As Double
    Dim dMed As Double
    Dim dMax As Double
    Dim dMin As Double
    Dim dSd As Double
    Dim dCpk As Double
    
    Set sh1 = Worksheets("Main1")
    Set sh2 = Worksheets("�H���Ǘ��p�f�[�^")
    
    With sh1
        sNG1 = .Cells(5, 3) '"A:�n�ڐ؂�"
        sNG2 = .Cells(6, 3) '"B:�׏d�s��"
        sNG3 = .Cells(7, 3) '"C:�׏d�ߑ�"
        dUl = .Cells(5, 2)   'dUl�F����K�i�l
        dLl = .Cells(6, 2)   'dLl�F�����K�i�l
    
    End With
        
        
    '�H���Ǘ��p�f�[�^�V�[�g�ւ̓W�J
    
    dArry = WorksheetFunction.Index(aArry, 0, 3)    '�׏d�s�[�N�𒊏o
    dArry = WorksheetFunction.Transpose(dArry)      '1�����z��ɕϊ�
    
    dMean = Application.WorksheetFunction.Average(dArry)    '�׏d�s�[�N�̕���
    dMed = Application.WorksheetFunction.Median(dArry)  '�׏d�s�[�N�̒����l
    dMax = Application.WorksheetFunction.Max(dArry) '�׏d�s�[�N��MAX�l
    dMin = Application.WorksheetFunction.Min(dArry) '�׏d�s�[�N��MIN�l
    dSd = Application.WorksheetFunction.StDevP(dArry)   '�׏d�s�[�N�̕W���΍�
    dCpk = CalcCpk(dUl, dLl, dMean, dSd)    'Cpk�̎Z�o
        
    'dArry�𕽋ςŖ��߂�
    For i = 1 To UBound(dArry)
        dArry(i) = dMean
    Next i
        
    With sh2
    
        .Cells.Delete   '�H���Ǘ��p�f�[�^�V�[�g��������
        
        '�z����V�[�g�ɓW�J����
        .Range("A2").Resize(UBound(aArry, 1), UBound(aArry, 2)) = aArry
        .Columns(6).Insert
       
        dArry = WorksheetFunction.Transpose(dArry)
        .Range("F2").Resize(UBound(dArry, 1), 1) = dArry
        
        '�����̗��11/9 12:34:56�̃t�H�[�}�b�g�Ƃ���
        .Range(.Cells(2, 1), .Cells(lSize + 2, 1)).NumberFormatLocal = "m/d hh:mm:ss"
        
        
        '�\�̌��o��
        bArry = Split("����,�J�E���g,�׏d�s�[�N,����K�i�l,�����K�i�l,���ϒl,�g���u���v��", ",")
        .Range(.Cells(1, 1), .Cells(1, UBound(bArry) + 1)) = bArry
             
        '�׏d�s�[�N�f�[�^�̍ŏI�s���擾
        lRow = .Cells(Rows.Count, 3).End(xlUp).Row
                
        '�Ǘ��f�[�^�̏��o
        ReDim cArry(1 To 5, 1 To 17)
        
        '�V�[�g�֓W�J���邽�߂̏���
        cArry(2, 6) = "����K�i�l"
        cArry(2, 8) = dUl   '����K�i�l
        
        cArry(3, 6) = "�����K�i�l"
        cArry(3, 8) = dLl   '�����K�i�l
        
        cArry(4, 6) = "�W���΍�"
        cArry(4, 8) = dSd   '�W���΍�
        
        cArry(5, 6) = "Cpk"
        cArry(5, 8) = dCpk  'Cpk
        
        cArry(2, 9) = "�ő�l"
        cArry(2, 11) = dMax  '�ő�l
        
        cArry(3, 9) = "�ŏ��l"
        cArry(3, 11) = dMin  '�ŏ��l
        
        cArry(4, 9) = "���ϒl"
        cArry(4, 11) = dMean '����
        
        cArry(5, 9) = "�����l"
        cArry(5, 11) = dMed  '�����l
                
        '�����̃Z�b�g
        Dim lCntNum As Long
        
        lCntNum = UBound(dArry)
        
        cArry(2, 1) = "�J�E���g����"
        cArry(2, 3) = lCntNum
        
        
        ReDim bArry(1 To 3) As Variant
        
        '�g���u���v���R�����g��bArry�ɃZ�b�g
        For i = 1 To UBound(bArry)
            bArry(i) = sh1.Cells(i + 4, 3).Value
        Next i
        
        '�g���u���v����z��dArry�ɃZ�b�g
        dArry = .Range(.Cells(2, 7), .Cells(lRow, 8))
        
        '�f�B�N�V���i���[���쐬
        Dim myDic As Object
        Dim sKey As String
        
        Set myDic = CreateObject("Scripting.Dictionary")
        
        '�g���u���v�����ڂ����[�v
        For i = 1 To UBound(bArry)
            myDic.Add bArry(i), 0 '�L�[�������ɓo�^
        Next
        
        
        '�g���u���v���f�[�^�����[�v
        For i = 1 To UBound(dArry, 1)
            '�����ɓo�^����Ă���ꍇ
            If myDic.Exists(dArry(i, 1)) Then
                myDic(dArry(i, 1)) = myDic(dArry(i, 1)) + 1 '�J�E���g�A�b�v
            End If
        Next
        
        Dim lNGCnt As Long
        
        lNGCnt = WorksheetFunction.Sum(myDic.Items)
        
        '�s�K�����̃Z�b�g
        cArry(1, 14) = "�s�K����"
        cArry(1, 16) = "�s�K����(%)"
        
        cArry(2, 12) = bArry(1)  'A:�n�ڐ؂�"
        cArry(3, 12) = bArry(2)  'B:�׏d�s��
        cArry(4, 12) = bArry(3)  'C:�׏d�ߑ�
        
        cArry(5, 12) = "���v"
        cArry(5, 14) = lNGCnt
        cArry(5, 16) = lNGCnt / lCntNum * 100
        
        For i = 1 To myDic.Count
            cArry(i + 1, 14) = myDic.Item(bArry(i))     '�eNG���̒��o
            cArry(i + 1, 16) = cArry(i + 1, 14) / lCntNum * 100 '�eNG��
        Next i
        
        Set myDic = Nothing
        
        
        
        'OK��
        cArry(4, 1) = "OK��"
        cArry(4, 3) = lCntNum - lNGCnt
        
        '�����񕝒����ɃZ�b�g���Č��h���𐮂���
        .Columns("A:K").AutoFit
            
    End With
        
    Dim sh3 As Worksheet
    
    Set sh3 = Worksheets("�H���Ǘ��\")
        
    '�H���Ǘ��\�V�[�g�̍H���Ǘ��f�[�^������������
    'sh3.Range("A5").Resize(UBound(cArry, 1), UBound(cArry, 2)).ClearContents
    '�H���Ǘ��\�V�[�g�֍H���Ǘ��p�f�[�^����������
    sh3.Range("A6").Resize(UBound(cArry, 1), UBound(cArry, 2)) = cArry
    
    
End Sub



'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' MakeProcCtrlChart()
'   �H���Ǘ��\�̃O���t�쐬
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Private Sub MakeProcCtrlChart()

    Dim sh1 As Worksheet
    Dim i As Integer
    
    Set sh1 = Worksheets("�H���Ǘ��\")

    With sh1
        
        '�f�[�^�G���A�̏�����
        .Range(.Range("A40"), .Range("A" & Cells.Rows.Count)).EntireRow.ClearContents
        .Rows("40:" & Rows.Count).RowHeight = 9.75
        
        
        ' �����O���t�̍폜
    
        If .ChartObjects.Count > 0 Then
            For i = .ChartObjects.Count To 1 Step -1
                .ChartObjects(i).Delete
            Next i
        End If
    
    End With

    '�׏d���茋�ʃ`���[�g�̍쐬
    Call DrawMeasChart
    
    ' �`���[�g�̕���
    Call CopyChart("XR�O���t", "�H���Ǘ��\", 1, 2, "A41", "J41", "J65")
    Call CopyChart("XR�O���t", "�H���Ǘ��\", 2, 3, "A66", "J66", "J89")
    Call CopyChart("�q�X�g�O����", "�H���Ǘ��\", 1, 4, "K41", "P41", "P89")

End Sub


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' FormProcMeasData4Printing()
'   �H���Ǘ��\�̑��茋�ʃf�[�^���̐��`
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Private Sub FormProcMeasData4Printing()
    
    Dim sh1 As Worksheet
    Dim sh2 As Worksheet
    
    Set sh1 = Worksheets("�H���Ǘ��\")
    Set sh2 = Worksheets("�H���Ǘ��p�f�[�^")
    
    Dim aArry As Variant
    Dim bArry As Variant
    Dim lRow As Long, lSize As Long, lCntW As Long, lCntR As Long
    Dim i As Integer, j As Integer, k As Integer

    With sh1
    '�׏d�s�[�N�f�[�^�̃N���A
        .Range(.Cells(40, 1), .Cells(.Cells.Rows.Count, 1)).ClearContents
    End With
    
    With sh2
        '�׏d�s�[�N�f�[�^�̍ŏI�s���擾
        lRow = sh2.Cells(Rows.Count, 3).End(xlUp).Row
      
        '�����A�J�E���g�A�s�[�N��aArry�ɃZ�b�g
        aArry = .Range(.Cells(2, 1), .Cells(lRow, 3))
        '5LP����L�^�A�g���u���v����bArry�ɃZ�b�g
        bArry = .Range(.Cells(2, 7), .Cells(lRow, 7))
    
        'aArry��bArry���}�[�W
        aArry = MergeVariantArraysCol(aArry, bArry)
    
    Dim lChunkRow As Long
    Dim lChunkStage As Long
    Dim lBlockNum As Long
    
    '���50�s��3�i�g
    lChunkRow = 50
    lChunkStage = 4
    
    '��i�Ɋi�[����f�[�^��200
    lBlockNum = lChunkRow * lChunkStage
    
    lSize = Int(UBound(aArry, 1) / lBlockNum + 1) * lBlockNum
    
    aArry = RedimPreserveArray(aArry, lSize)
        
    ReDim bArry(1 To (Int(UBound(aArry, 1) / 2 + Int(lSize / lBlockNum))), 1 To 17)
    
    lCntW = 0   '�i�[��A�h���X
    lCntR = 0   '�Ǎ���A�h���X
    
    For i = 1 To (UBound(aArry, 1) / lBlockNum)
        For j = 1 To lBlockNum
                Select Case j
                
                    Case 1
                        bArry(j + lCntW, 1) = "����"
                        bArry(j + lCntW, 2) = "�J�E���g"
                        bArry(j + lCntW, 3) = "�׏d�s�[�N"
                        bArry(j + lCntW, 4) = "�g���u���v��"
    
                        bArry(j + lCntW, 5) = "����"
                        bArry(j + lCntW, 6) = "�J�E���g"
                        bArry(j + lCntW, 7) = "�׏d�s�[�N"
                        bArry(j + lCntW, 8) = "�g���u���v��"
                    
                        bArry(j + lCntW, 9) = "����"
                        bArry(j + lCntW, 10) = "�J�E���g"
                        bArry(j + lCntW, 11) = "�׏d�s�[�N"
                        bArry(j + lCntW, 12) = "�g���u���v��"
                        
                        bArry(j + lCntW, 13) = "����"
                        bArry(j + lCntW, 14) = "�J�E���g"
                        bArry(j + lCntW, 15) = "�׏d�s�[�N"
                        bArry(j + lCntW, 16) = "�g���u���v��"
                        
                        lCntW = lCntW + 1
                End Select
            
            For k = 1 To UBound(aArry, 2)
                
                Select Case j
                    Case 1 To 50    '1�i�ڂ̏���
                        bArry(j + lCntW, k) = aArry(j + lCntR, k)
                    
                    Case 51 To 100  '2�i�ڂ̏���
                        bArry(j + lCntW - 50, k + 4) = aArry(j + lCntR, k)
                
                    Case 101 To 150 '3�i�ڂ̏���
                        bArry(j + lCntW - 100, k + 8) = aArry(j + lCntR, k)
                
                    Case 151 To 200 '4�i�ڂ̏���
                        bArry(j + lCntW - 150, k + 12) = aArry(j + lCntR, k)
                
                End Select
            Next k
        Next j
        lCntW = lCntW + lChunkRow   'bArry�̊i�[����X�V
        lCntR = lCntR + lBlockNum   'aArry�̓Ǎ�����X�V
    Next i

    End With
    
    With sh1
        '�H���Ǘ��\�V�[�g�ֈ���t�H�[�}�b�g�̑���f�[�^����������
        .Range("A91").Resize(UBound(bArry, 1), UBound(bArry, 2)) = bArry
        
        '�ŏI�s���擾
        lRow = .Cells(Rows.Count, 1).End(xlUp).Row
        
'        For i = 1 To lChunkStage
            
'            With .Range(.Cells(lRow, (i - 1) * 6 + 1), Cells(lRow, (i - 1) * 6 + 4))
            With .Range(.Cells(lRow, 1), Cells(lRow, 16))
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
            End With
        
'        Next i
        
        .Activate
                
        With .PageSetup
            '����͈͂̐ݒ�
            .PrintArea = Range(Cells(1, 1), Cells(lRow, 16)).Address
            .LeftHeader = "" '�����w�b�_�[�F�Ȃ�
            .CenterHeader = Cells(3, 3) '�����w�b�_�[�F���b�g
            .RightHeader = "&D &T"  '�E���w�b�_�[�F���t ����
            .LeftFooter = "" '�����t�b�^�[�F�Ȃ�
            .CenterFooter = "&P/&N" '�����t�b�^�[�F�y�[�W��/���y�[�W��
            .RightFooter = "&""Verdana""&08" & "Confidencial"   '�E�t�b�^�[�FVerdana�t�H���g�A�T�C�Y8�ŁuConfidencial�v
        
        End With
        
        .Cells(1, 1).Select
    End With

End Sub




'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' FormXRData(ByRef xArry As Variant)
'   X-R�Ǘ��p�̃f�[�^�̐��`
'
'   xArry�F���`�Ɏg�p����f�[�^
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Private Sub FormXRData(ByRef xArry As Variant)

    Dim yArry As Variant
    Dim zArry As Variant
    Dim lGun As Long
    Dim i As Long, j As Long, lCnt As Long
    Dim lLenX As Long, lLenY As Long, lLenZ As Long

    Dim dSum1 As Double
    Dim dMax As Double
    Dim dMin As Double

    lGun = Worksheets("Main1").Cells(4, 2)

    lLenX = UBound(xArry, 1)
    lLenY = lLenX / lGun + 1
    lLenZ = lLenY * lGun + lGun
    
    ReDim yArry(1 To lLenY, 1 To 17)
    ReDim zArry(1 To lLenZ, UBound(xArry, 2))
    
    For i = 1 To lLenX
        For j = 1 To UBound(xArry, 2)
            zArry(i, j) = xArry(i, j)
        Next j
    Next i
    
    '�f�[�^���Q�ɕ���U��
    For i = 1 To lLenY
        dSum1 = 0
        dMax = 0
        dMin = 0
        lCnt = 0

        For j = 1 To lGun
            yArry(i, j + 3) = zArry((i - 1) * lGun + j, 3)
            If yArry(i, j + 3) <> "" Then
                lCnt = lCnt + 1
            End If
            '�Q���̍ő�E�ŏ��l�̒��o
            If j = 1 Then
                dMax = yArry(i, j + 3)
                dMin = dMax
            Else
                '�ő�l
                If dMax < yArry(i, j + 3) Then
                    dMax = yArry(i, j + 3)
                End If
                
                '�ŏ��l
                If dMin > yArry(i, j + 3) And yArry(i, j + 3) <> "" Then
                    dMin = yArry(i, j + 3)
                End If
            End If
            
            '�Q���ώZ�o�̂��߂̏���
            dSum1 = dSum1 + yArry(i, j + 3)
        Next j
        
        '���ϒl
        If dSum1 <> 0 Then
            yArry(i, 2) = dSum1 / lCnt
            yArry(i, 1) = i
        End If
        
        '�΍�
        If dMax <> 0 And dMin <> 0 Then
            yArry(i, 3) = dMax - dMin
        End If
        
    Next i

    With Worksheets("XRdata")
        .Range("A21").Resize(UBound(yArry, 1), UBound(yArry, 2)) = yArry
        .Range("A1").Value = "X�Ǘ��}�f�[�^"
        .Range("A2").Value = "�Q����"
        .Range("A3").Value = "X-UCL"
        .Range("A4").Value = "X-LCL"

        .Range("D1").Value = "R�Ǘ��}�f�[�^"
        .Range("D2").Value = "R����"
        .Range("D3").Value = "R-UCL"
        .Range("D4").Value = "R-LCL"

        ReDim zArry(0, 18)
        
        zArry(0, 0) = "�Q"
        zArry(0, 1) = "�Q����"
        zArry(0, 2) = "R"
        
        For i = 1 To 10
            zArry(0, i + 2) = "n" & i
        Next i
        
        zArry(0, 13) = "X����Ǘ��l"
        zArry(0, 14) = "X�����Ǘ��l"
        zArry(0, 15) = "X����"
        zArry(0, 16) = "R����Ǘ��l"
        zArry(0, 17) = "R�����Ǘ��l"
        zArry(0, 18) = "R����"
        
        
        .Range("A20:S20").Value = zArry
        
        ReDim zArry(1 To 11, 1 To 4)
        
        'X-R�W���̃Z�b�g
        zArry = ShewhartConstant
        .Range("T1:W11").Value = zArry
        

        Dim lRow As Long

        lRow = .Cells(Rows.Count, 2).End(xlUp).Row
        
        .Cells(2, 2).Formula = "=AVERAGE(B21:" & Cells(lRow, 2).Address & ")"
        .Cells(2, 5).Formula = "=AVERAGE(C21:" & Cells(lRow, 3).Address & ")"
        .Cells(3, 2).Formula = "=$B$2+VLOOKUP(" & lGun & ",$T$3:$W$11,2)*$E$2"
        .Cells(4, 2).Formula = "=$B$2-VLOOKUP(" & lGun & ",$T$3:$W$11,2)*$E$2"
        .Cells(3, 5).Formula = "=VLOOKUP(" & lGun & ",$T$3:$W$11,4)*$E$2"
        .Cells(4, 5).Formula = "=VLOOKUP(" & lGun & ",$T$3:$W$11,3)*$E$2"
        
        
        .Range(.Cells(21, 14), .Cells(lRow, 14)).Value = .Cells(3, 2).Value
        .Range(.Cells(21, 15), .Cells(lRow, 15)).Value = .Cells(4, 2).Value
        .Range(.Cells(21, 16), .Cells(lRow, 16)).Value = .Cells(2, 2).Value
        
        .Range(.Cells(21, 17), .Cells(lRow, 17)).Value = .Cells(3, 5).Value
        .Range(.Cells(21, 18), .Cells(lRow, 18)).Value = .Cells(4, 5).Value
        .Range(.Cells(21, 19), .Cells(lRow, 19)).Value = .Cells(2, 5).Value
                
        
    End With


End Sub



'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' DrawXRChart()
'  X�Ǘ��O���t��R�Ǘ��O���t�̍쐬
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Private Sub DrawXRChart()

    Dim sh1 As Worksheet
    Dim sh2 As Worksheet
    Dim sh3 As Worksheet
    Dim TrRange As Range
    Dim lRow As Long, i As Long
    Dim R As Range
    Dim aArry As Variant
    
    Dim iMin As Integer, iMax As Integer
    
    Set sh1 = Worksheets("XRData")
    Set sh2 = Worksheets("XR�O���t")
    Set sh3 = Worksheets("�H���Ǘ��\")

        
    '�����O���t�̏���
    With sh2
        For i = .ChartObjects.Count To 1 Step -1
            .ChartObjects(i).Delete
        Next i
    End With
    
    '�O���t�p�̃f�[�^��z��ɓǍ���
    lRow = sh1.Cells(Rows.Count, 1).End(xlUp).Row
    aArry = sh1.Range(sh1.Cells(21, 2), sh1.Cells(lRow, 2))
    
    '�ő�l�ƍŏ��l�̎擾�i�c���̐ݒ�p�j
    iMax = Int(Application.WorksheetFunction.Max(aArry)) + 1
    iMin = Int(Application.WorksheetFunction.Min(aArry)) - 1
    
    
    '========================================================
    '
    ' X�Ǘ��O���t�̍쐬
    '
    '========================================================
    '�O���t�`��G���A�̐ݒ�
    Set R = sh2.Range("A1:Q22")
    
    '�O���t�Ɏg�p����f�[�^��ݒ�
    With sh1
        Set TrRange = Union(.Range(.Cells(20, 2), .Cells(lRow, 2)), .Range(.Cells(20, 14), .Cells(lRow, 16)))
    End With
    
    '�O���t���쐬
    With sh2.ChartObjects.Add(R.Left, R.Top, R.Width, R.Height)

        .Name = "X�Ǘ��}"
        .Chart.ChartType = xlLine
        .Chart.SetSourceData TrRange
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "X�Ǘ��}"
        .Chart.SeriesCollection.Item(1).Format.Line.Weight = 1
        .Chart.SeriesCollection.Item(4).Format.Line.Weight = 2.5
        .Chart.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(124, 124, 124)
        .Chart.SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(20, 20, 200)
    End With
        
    '�c���ő�l�F�����Ƃ���
    sh2.ChartObjects(1).Chart.Axes(xlValue).MaximumScaleIsAuto = True
    '�c���ŏ��l�F�f�[�^�ŏ� - 1
    sh2.ChartObjects(1).Chart.Axes(xlValue).MinimumScale = iMin
    
        
    '========================================================
    '
    ' R�Ǘ��O���t�̍쐬
    '
    '========================================================
    '�O���t�`��G���A�̐ݒ�
    Set R = sh2.Range("A23:Q42")
    
    '�O���t�Ɏg�p����f�[�^��ݒ�
    With sh1
        Set TrRange = Union(.Range(.Cells(20, 3), .Cells(lRow, 3)), .Range(.Cells(20, 17), .Cells(lRow, 19)))
    End With
    
    '�O���t���쐬
    With sh2.ChartObjects.Add(R.Left, R.Top, R.Width, R.Height)

        .Name = "R�Ǘ��}"
        .Chart.ChartType = xlLine
        .Chart.SetSourceData TrRange
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "R�Ǘ��}"
        .Chart.SeriesCollection.Item(1).Format.Line.Weight = 1
        .Chart.SeriesCollection.Item(4).Format.Line.Weight = 2.5
        .Chart.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(124, 124, 124)
        .Chart.SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(20, 20, 200)

    End With
    
    sh2.ChartObjects(2).Chart.Axes(xlValue).MaximumScaleIsAuto = True
    '�c���ŏ��l�F�f�[�^�ŏ� - 1
    sh2.ChartObjects(2).Chart.Axes(xlValue).MinimumScaleIsAuto = True
    
     '�摜�Ƃ��ăR�s�[
'    sh2.ChartObjects(1).CopyPicture
'
'    sh3.Activate
'    sh3.Range("A25:G32").Select '�Z����I��
'    sh3.Paste '�\��t��
    


End Sub


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' DrawMeasChart()
'  ���茋�ʃO���t�̍쐬
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Private Sub DrawMeasChart()

    Dim sh1 As Worksheet
    Dim sh2 As Worksheet
    Dim sh3 As Worksheet
    Dim TrRange As Range
    Dim lRow As Long, i As Long
    Dim R As Range
    Dim iUL As Integer
    
    Dim iMin As Integer, iMax As Integer
    
    Set sh1 = Worksheets("�H���Ǘ��p�f�[�^")
    Set sh2 = Worksheets("�H���Ǘ��\")

    
    iUL = Int(sh1.Cells(2, 5))
    iUL = iUL - 3   '�c���ŏ��l�������K�i�l���Z�o
    
    
    '�����O���t�̏���
    With sh2
        For i = .ChartObjects.Count To 1 Step -1
            .ChartObjects(i).Delete
        Next i
    End With
    
    '========================================================
    '
    ' ���茋�ʃO���t�̍쐬
    '
    '========================================================
    '�O���t�`��G���A�̐ݒ�
    Set R = sh2.Range("A11:P36")
    
    lRow = sh1.Cells(Rows.Count, 1).End(xlUp).Row
    
    '�O���t�Ɏg�p����f�[�^��ݒ�
    With sh1
        Set TrRange = .Range(.Cells(1, 3), .Cells(lRow, 6))
    End With
    
    '�O���t���쐬
    With sh2.ChartObjects.Add(R.Left, R.Top, R.Width, R.Height)

        .Name = "�׏d���茋��"
        .Chart.ChartType = xlLine
        .Chart.SetSourceData TrRange
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "�׏d���茋��"
        .Chart.SeriesCollection.Item(1).Format.Line.Weight = 1
        .Chart.SeriesCollection.Item(4).Format.Line.Weight = 2.5
        .Chart.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(124, 124, 124)
        .Chart.SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(20, 20, 200)
    End With
        
    '�c���ő�l�F�����Ƃ���
    sh2.ChartObjects(1).Chart.Axes(xlValue).MaximumScaleIsAuto = True
    '�c���ŏ��l�F�f�[�^�ŏ� - 1
    sh2.ChartObjects(1).Chart.Axes(xlValue).MinimumScale = iUL

End Sub



'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' CalcCpk(dUl As Double, dLl As Double, dMean As Double, dSd As Double) As Double
'
'  Cpk�̌v�Z
'   dUl�F����Ǘ��l
'   dLl�F�����Ǘ��l
'   dMean�F���ϒl
'   dSd�F�W���΍�
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Private Function CalcCpk(dUl As Double, dLl As Double, dMean As Double, dSd As Double) As Double

    Dim dK As Double
    Dim dCp As Double
    Dim dCpk As Double
    
    dCpk = 0
    
    If dSd Then
        dCp = (dUl - dLl) / (6 * dSd)   'Cp�̌v�Z
        dK = Abs((((dUl + dLl) / 2) - dMean)) / ((dUl - dLl) / 2) 'K�̌v�Z
        dCpk = (1 - dK) * dCp   'Cpk
    End If

    CalcCpk = dCpk
    
End Function

Private Sub CopyChart(ByVal wsSrcName As String, ByVal wsDistName As String, _
                ByVal objNum1 As Integer, ByVal objNum2 As Integer, _
                ByVal objLeft As String, ByVal objTop As String, ByVal objWidth As String)
    
    Dim sh1 As Worksheet
    Dim sh2 As Worksheet
    
    Set sh1 = Worksheets(wsSrcName)
    Set sh2 = Worksheets(wsDistName)
    
    
    '�R�s�[
    sh1.ChartObjects(objNum1).Copy
       
    sh2.Activate
    sh2.Range(objLeft).Select '�Z����I��
    sh2.Paste '�\��t��
    
    sh2.ChartObjects(objNum2).Left = Range(objLeft).Left '���̈ʒu
    sh2.ChartObjects(objNum2).Top = Range(objTop).Top '��̈ʒu
    sh2.ChartObjects(objNum2).Width = Range(objLeft & ":" & objWidth).Width '��
    sh2.ChartObjects(objNum2).Height = Range(objLeft & ":" & objWidth).Height '����
    
End Sub


Private Function ShewhartConstant() As Variant

    Dim aArry As Variant

    aArry = ActiveSheet.Range("T1:W11").Value
    
    ReDim aArry(1 To 11, 1 To 4)
    
    aArry(1, 2) = "n�W��"
    
    aArry(2, 1) = "�Q��"
    aArry(2, 2) = "A2"
    aArry(2, 3) = "D3"
    aArry(2, 4) = "D4"
    
    aArry(3, 1) = 2
    aArry(3, 2) = 1.88
    aArry(3, 3) = 0
    aArry(3, 4) = 3.267
    
    aArry(4, 1) = 3
    aArry(4, 2) = 1.023
    aArry(4, 3) = 0
    aArry(4, 4) = 2.575
    
    aArry(5, 1) = 4
    aArry(5, 2) = 0.729
    aArry(5, 3) = 0
    aArry(5, 4) = 2.282
    
    aArry(6, 1) = 5
    aArry(6, 2) = 0.577
    aArry(6, 3) = 0
    aArry(6, 4) = 2.115
    
    aArry(7, 1) = 6
    aArry(7, 2) = 0.483
    aArry(7, 3) = 0
    aArry(7, 4) = 2.004

    aArry(8, 1) = 7
    aArry(8, 2) = 0.419
    aArry(8, 3) = 0.076
    aArry(8, 4) = 1.924
    
    aArry(9, 1) = 8
    aArry(9, 2) = 0.373
    aArry(9, 3) = 0.136
    aArry(9, 4) = 1.846

    aArry(10, 1) = 9
    aArry(10, 2) = 0.337
    aArry(10, 3) = 0.184
    aArry(10, 4) = 1.816
    
    aArry(11, 1) = 10
    aArry(11, 2) = 0.308
    aArry(11, 3) = 0.223
    aArry(11, 4) = 1.777
    
    ShewhartConstant = aArry

End Function



'--- 2��Variant�z�����������i������j ---'
Private Function MergeVariantArraysCol(vArray1 As Variant, vArray2 As Variant) As Variant
    
    '��������z��̃T�C�Y
    Dim vArray1_row As Long
    Dim vArray1_col As Long
    Dim vArray2_row As Long
    Dim vArray2_col As Long
    vArray1_row = UBound(vArray1, 1)
    vArray1_col = UBound(vArray1, 2)
    vArray2_row = UBound(vArray2, 1)
    vArray2_col = UBound(vArray2, 2)
    
    '������̔z��̃T�C�Y
    Dim newArray_row As Long
    Dim newArray_col As Long
    newArray_row = Application.WorksheetFunction.Max(vArray1_row, vArray2_row)
    newArray_col = vArray1_col + vArray2_col
    
    '������̔z��
    Dim newArray As Variant
    ReDim newArray(1 To newArray_row, 1 To newArray_col)
    
    '�z�����������
    Dim i As Long
    Dim j As Long
    For j = 1 To newArray_col
        If (j <= vArray1_col) Then
            For i = 1 To newArray_row
                If (i <= vArray1_row) Then
                    newArray(i, j) = vArray1(i, j)
                Else
                    newArray(i, j) = Empty
                End If
            Next i
        Else
            For i = 1 To newArray_row
                If (i <= vArray2_row) Then
                    newArray(i, j) = vArray2(i, j - vArray1_col)
                Else
                    newArray(i, j) = Empty
                End If
            Next i
        End If
    Next j
    
    MergeVariantArraysCol = newArray
    
End Function


'�}�[�W�o�u���\�[�g
Private Sub merge_sort2(ByRef arr As Variant, ByVal col As Long)
    Dim irekae As Variant
    Dim indexer As Variant
    Dim tmp1() As Variant
    Dim tmp2() As Variant
    Dim i As Long
    ReDim irekae(LBound(arr, 1) To UBound(arr, 1))
    ReDim indexer(LBound(arr, 1) To UBound(arr, 1))
    ReDim tmp1(LBound(arr, 1) To UBound(arr, 1))
    ReDim tmp2(LBound(arr, 1) To UBound(arr, 1))
    For i = LBound(arr, 1) To UBound(arr, 1) Step 2
        If i + 1 > UBound(arr, 1) Then
            irekae(i) = arr(i, col)
            indexer(i) = i
            Exit For
        End If
        If arr(i + 1, col) < arr(i, col) Then
            irekae(i) = arr(i + 1, col)
            irekae(i + 1) = arr(i, col)
            indexer(i) = i + 1
            indexer(i + 1) = i
        Else
            irekae(i) = arr(i, col)
            irekae(i + 1) = arr(i + 1, col)
            indexer(i) = i
            indexer(i + 1) = i + 1
        End If
    Next
    Dim st1 As Long
    Dim en1 As Long
    Dim st2 As Long
    Dim en2 As Long
    Dim n As Long
    i = 1
    Do While i * 2 <= UBound(arr, 1)
        i = i * 2
        n = 0
        Do While en2 + i - 1 < UBound(arr, 1)
            n = n + 1
            st1 = i * 2 * (n - 1) + LBound(arr, 1)
            en1 = i * 2 * (n - 1) + i - 1 + LBound(arr, 1)
            st2 = en1 + 1
            en2 = IIf(st2 + i - 1 >= UBound(arr, 1), UBound(arr, 1), st2 + i - 1)
            Call merge2(irekae, indexer, tmp1, tmp2, st1, en1, st2, en2)
        Loop
        en2 = 0
    Loop
    Dim ret As Variant
    ReDim ret(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
    For i = LBound(arr, 1) To UBound(arr, 1)
        For n = LBound(arr, 2) To UBound(arr, 2)
            If IsObject(arr(indexer(i), n)) Then
                Set ret(i, n) = arr(indexer(i), n)
            Else
                ret(i, n) = arr(indexer(i), n)
            End If
        Next
    Next
    arr = ret
End Sub

Private Sub merge2(ByRef irekae As Variant, _
ByRef indexer As Variant, _
ByRef tmpArr() As Variant, _
ByRef tmpIndexer() As Variant, _
ByVal st1 As Long, _
ByVal en1 As Long, _
ByVal st2 As Long, _
ByVal en2 As Long)
    Dim j As Long
    Dim n As Long
    Dim i As Long
    For i = st1 To en2
        tmpArr(i) = irekae(i)
        tmpIndexer(i) = indexer(i)
    Next
    j = st1
    n = st2
    Do While (j < en1 + 1 Or n < en2 + 1)
        If n >= en2 + 1 Then
            irekae(j + n - st2) = tmpArr(j)
            indexer(j + n - st2) = tmpIndexer(j)
            j = j + 1
        ElseIf j < en1 + 1 And tmpArr(j) <= tmpArr(n) Then
            irekae(j + n - st2) = tmpArr(j)
            indexer(j + n - st2) = tmpIndexer(j)
            j = j + 1
        Else
            irekae(j + n - st2) = tmpArr(n)
            indexer(j + n - st2) = tmpIndexer(n)
            n = n + 1
        End If
    Loop
End Sub

