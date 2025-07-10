Attribute VB_Name = "Module4"
Option Explicit

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'ヒストグラムの作成
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Public Sub MakeHistogram()

    Dim rTarget As Range ' 選択範囲
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
    
    Set sh1 = Worksheets("工程管理用データ")
    Set sh2 = Worksheets("ヒストグラム")
    Set sh3 = Worksheets("工程管理表")

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
            
            '不具合品を無視する場合は、最大値を上限規格、最小値を下限規格にする
            'dMax = dUl
            'dMin = dLl
        End With
        
        '各種アルゴリズムによる階級幅の算出
        Dim dScott As Double, dSturges As Double, dFd As Double, dSqRoot As Double
        
        dScott = 3.5 * dSd / lSize ^ (1 / 3)
        dSturges = (dMax - dMin) / (1 + (.Log(lSize) / .Log(2)))
        dFd = 2 * (.Quartile_Inc(rTarget, 3) - .Quartile_Inc(rTarget, 1)) / lSize ^ (1 / 3)
        dSqRoot = (dMax - dMin) / lSize ^ (1 / 2)
        
        Dim dClassWidth As Double, d1stClass As Double

        ' 第1階級の下限と階級幅の決定
'        dClassWidth = BIN_WIDTH(dScott)
'        dClassWidth = BIN_WIDTH(dSturges)
'        dClassWidth = BIN_WIDTH(dFd)
        dClassWidth = BIN_WIDTH(dSqRoot)
        Tmp = .Floor(dMin, dClassWidth)
        If dMin <> Tmp Then
            d1stClass = Tmp
        Else
            d1stClass = Tmp - dClassWidth ' min = h のときの第1階級下限の修正
        End If
    End With
    
    Dim aArry As Variant
    Dim fMax As Long ' 最大度数
    Dim i As Long
        
    ReDim aArry(1 To ((dMax - d1stClass) / dClassWidth) + 2, 1 To 3)
    
    fMax = 0
    
    For i = 1 To UBound(aArry, 1)
        
        Select Case i
            Case 1 ' 第1階級下境界
                aArry(i, 1) = d1stClass

            Case Else ' 第2階級以降の下境界
                aArry(i, 1) = aArry(i - 1, 2)
                                  
        End Select
        
        aArry(i, 2) = aArry(i, 1) + dClassWidth ' 上境界
        
        '度数
        aArry(i, 3) = WorksheetFunction.CountIfs(rTarget, ">=" & aArry(i, 1), rTarget, "<" & aArry(i, 2))
        
        '最大度数
        If aArry(i, 3) > fMax Then
            fMax = aArry(i, 3)
        End If

    Next i
    
    
    '正規分布の算出
    Dim bArry As Variant
    Dim dClassMax As Double, dClassMin As Double

    dClassMin = aArry(1, 1) '階級の最小値
    dClassMax = aArry(UBound(aArry, 1), 2)  '階級の最大値
    
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
        .Cells(1, 1) = "最初の階級"
        .Cells(1, 2) = d1stClass
        .Cells(2, 1) = "階級の幅"
        .Cells(2, 2) = dClassWidth
        .Cells(3, 1) = "▼度数分布表"
        .Cells(4, 1) = "下境界"
        .Cells(4, 2) = "上境界"
        .Cells(4, 3) = "度数"
        .Cells(4, 4) = "最大度数"
        .Cells(5, 4) = fMax
        .Cells(4, 5) = "期待度数1"
        .Cells(4, 6) = "期待度数2"
        
        
        .Range("A5").Resize(UBound(aArry, 1), UBound(aArry, 2)) = aArry
        .Range("E5").Resize(UBound(bArry, 1), UBound(bArry, 2)) = bArry
        
    End With
    
    
    '既存グラフの消去
    With sh2
        For i = .ChartObjects.Count To 1 Step -1
            .ChartObjects(i).Delete
        Next i
    End With
    
    
    sh2.Activate
    
    'グラフ
            
    Dim gRange As Range, R As Range
    
    'グラフに使用するデータを設定
    lSize = Cells(Rows.Count, 2).End(xlUp).Row
    Range(Cells(5, 3), Cells(lSize, 4)).Select
    Set R = Range("G4:P30")
        
    
    ActiveSheet.Shapes.AddChart(xlColumnClustered).Select ' 集合縦棒グラフを作成

    With ActiveChart
        .HasLegend = False                      ' 凡例除去
        .ChartGroups(1).GapWidth = 0            ' 間隔=0
        With .SeriesCollection(1)
            .AxisGroup = 2                      ' 柱→2軸
'            .Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent3 ' 柱の色→グレー
            .Border.Color = vbWhite             ' 柱の外枠線色→白
            .Border.Weight = xlThin             ' 柱外枠の太さ
        End With
        With .SeriesCollection(2)
            .ChartType = xlXYScatter            ' 最大度数→散布図へ
            .MarkerStyle = xlMarkerStyleNone    ' マーカーを不可視に
        End With
        .Parent.Top = Range("G4").Top           ' 位置調整（上端）
        .Parent.Left = Range("G4").Left         ' 位置調整（左端）
        .Parent.Width = 540         ' 位置調整（幅）
        .Parent.Height = 300        ' 位置調整（高さ）
    End With
    
    
    ' 正規分布曲線の追加
    lSize = Cells(Rows.Count, 5).End(xlUp).Row
    
    With ActiveChart.SeriesCollection.NewSeries ' Normal_Curve系列の追加
        .XValues = Range(Cells(5, 5), Cells(lSize, 5))
        .Values = Range(Cells(5, 6), Cells(lSize, 6))
        .Name = "正規分布曲線"
        .Border.Color = RGB(50, 50, 50)
        .Format.Line.Weight = 1 'pt
    End With

    ActiveChart.SeriesCollection("正規分布曲線").ChartType = xlXYScatterSmoothNoMarkers ' 平滑線化

    With ActiveChart.Axes(xlCategory)
        .MinimumScale = dClassMin              ' 軸スケール合わせ（最小値）
        .MaximumScale = dClassMax              ' 軸スケール合わせ（最大値）
        .MajorUnit = 2                 ' 軸スケール合わせ（目盛り）
        .CrossesAt = dClassMin                 ' 軸スケール合わせ（交点）
    End With
    
    With ActiveChart.Axes(xlValue, xlSecondary)
        .MinimumScale = 0
        .TickLabelPosition = xlNone         ' 2軸ラベルを不可視に
        .MajorTickMark = xlNone             ' 2軸目盛を不可視に
    End With
    
    With ActiveSheet.ChartObjects(1).Chart
        .Axes(xlValue).MinimumScale = 0
    End With
    
    
    
    Cells(2, 2).Select
    

End Sub


Private Function BIN_WIDTH(h)
    ' 階級幅を調整する

    Dim n As Long
    Dim Stp(2) ' 処理過程  step1 to 3
    Dim TryVal
    Dim Tmp ' 値

    TryVal = Array(5, 2, 1) ' Mround, Ceiling の基準値に掛けるウエイト

    With Application.WorksheetFunction
    Select Case h
    Case Is <= 0
        MsgBox "ERROR"
        Exit Function

    Case Is >= 1 ' hが1以上の場合の処理
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

    Case Is < 1 ' hが1より小さな場合の処理
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
    MsgBox "見出しを除く選択範囲に数値以外が含まれています"
End Sub

Private Sub Err2()
    MsgBox "処理対象のシートがアクティブになっていません"
End Sub

Private Sub Err3()
    MsgBox "アクティブなグラフがありません"
End Sub

Private Sub Err4()
    MsgBox "階級幅に指定されている内容が数値ではありません"
End Sub

Private Sub Err5()
    MsgBox "最初の階級の下境界に指定されている内容が数値ではありません"
End Sub

Private Sub Err6()
    MsgBox "最初の階級の下境界に妥当でない値が設定されています" & vbCrLf & _
        "変数のレンジをカバーするには，この値が最小値より小さな値である必要があります"
End Sub

Private Function Err_Checker1(myRange As Range) As Boolean
' 選択範囲に数値以外が含まれているか

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
' 「OP-」で始まるシートがアクティブか

Err_Checker2 = False
If Left(ActiveSheet.Name, 3) <> "OP-" Then
    Err_Checker2 = True
    Call Err2
End If

End Function

Private Function Err_Checker3() As Boolean
' グラフがアクティブになっているか

If ActiveChart Is Nothing Then
    Err_Checker3 = True
    Call Err3
Else
    Err_Checker3 = False
End If

End Function

Private Function Err_Checker4(s) As Boolean
' 数値かどうか

If IsNumeric(s) = False Or _
    s = "" Then
    Err_Checker4 = True
Else
    Err_Checker4 = False
End If

End Function



