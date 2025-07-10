Attribute VB_Name = "Module2"
Option Explicit

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' ProcCtrlData()
'   工程管理表の生成
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
    
    '描画を停止する
    Application.ScreenUpdating = False


    'ロギングデータをaArryに代入する
    With sh1
        .Activate
        .Cells(1, 1).Select
        lRow = .Cells(Rows.Count, 1).End(xlUp).Row
        ReDim aArry(1 To lRow, 1 To 6)
'        aArry = .Range(.Cells(13, 1), .Cells(lRow + 1, 6))
        aArry = .Range(.Cells(13, 1), .Cells(lRow, 6))
    End With
    
    lCnt = 1
    
    '工程管理用にロギングデータを成形する
    Call FormData(aArry, lCnt)
        
    'カウンタ値でソート
    Call merge_sort2(aArry, 2)
    
    'X-Rデータシートの作成
    Call FormXRData(aArry)
    'X-Rグラフの作成
    Call DrawXRChart
    
    
    '工程管理データシートの作成
    Call MakeProcCtrlDataSheet(aArry, lCnt)
    
    'ヒストグラムの作成
    Call MakeHistogram
        
    '工程管理表のグラフの作成
    Call MakeProcCtrlChart
    
    '工程管理表シートへ印刷フォーマットにしたデータを書込む
    Call FormProcMeasData4Printing
    
    '描画を元に戻す
    Application.ScreenUpdating = True
    
    eTime = Timer
    pTime = eTime - sTime
    
    Debug.Print "処理時間"; pTime & vbCrLf
    
End Sub

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' FormData(ByRef aArry As Variant, ByRef lCnt As Long)
'   ロギングデータを工程管理用データに成形する
'
'   aArry：ロギングデータ⇒成形して戻す
'   lCnt：処理カウンタ
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
        
        If .Cells(5, 8) = "含む" Then bNG1 = True
        If .Cells(6, 8) = "含む" Then bNG2 = True
        If .Cells(7, 8) = "含む" Then bNG3 = True
        If .Cells(8, 8) = "含む" Then bNG4 = True
    
    End With
    
    ' カウンタ値の初期化
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

    '配列を転置する
    aArry = WorksheetFunction.Transpose(aArry)
    
    '削除したデータ分、列を詰める
    ReDim Preserve aArry(1 To UBound(aArry, 1), 1 To UBound(aArry, 2) - (lSizeA - lCnt) - 1)
    
    '配列を転置し直す
    aArry = WorksheetFunction.Transpose(aArry)

End Sub


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' MakeProcCtrlDataSheet(ByRef aArry As Variant, lSize As Long)
'   "工程管理用データ"シートの作成
'
'   aArry：工程管理用データ
'   lCnt：工程管理用データ生成時の処理カウンタ値
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
    Set sh2 = Worksheets("工程管理用データ")
    
    With sh1
        sNG1 = .Cells(5, 3) '"A:溶接切れ"
        sNG2 = .Cells(6, 3) '"B:荷重不足"
        sNG3 = .Cells(7, 3) '"C:荷重過多"
        dUl = .Cells(5, 2)   'dUl：上限規格値
        dLl = .Cells(6, 2)   'dLl：下限規格値
    
    End With
        
        
    '工程管理用データシートへの展開
    
    dArry = WorksheetFunction.Index(aArry, 0, 3)    '荷重ピークを抽出
    dArry = WorksheetFunction.Transpose(dArry)      '1次元配列に変換
    
    dMean = Application.WorksheetFunction.Average(dArry)    '荷重ピークの平均
    dMed = Application.WorksheetFunction.Median(dArry)  '荷重ピークの中央値
    dMax = Application.WorksheetFunction.Max(dArry) '荷重ピークのMAX値
    dMin = Application.WorksheetFunction.Min(dArry) '荷重ピークのMIN値
    dSd = Application.WorksheetFunction.StDevP(dArry)   '荷重ピークの標準偏差
    dCpk = CalcCpk(dUl, dLl, dMean, dSd)    'Cpkの算出
        
    'dArryを平均で埋める
    For i = 1 To UBound(dArry)
        dArry(i) = dMean
    Next i
        
    With sh2
    
        .Cells.Delete   '工程管理用データシートを初期化
        
        '配列をシートに展開する
        .Range("A2").Resize(UBound(aArry, 1), UBound(aArry, 2)) = aArry
        .Columns(6).Insert
       
        dArry = WorksheetFunction.Transpose(dArry)
        .Range("F2").Resize(UBound(dArry, 1), 1) = dArry
        
        '時刻の列は11/9 12:34:56のフォーマットとする
        .Range(.Cells(2, 1), .Cells(lSize + 2, 1)).NumberFormatLocal = "m/d hh:mm:ss"
        
        
        '表の見出し
        bArry = Split("時刻,カウント,荷重ピーク,上限規格値,下限規格値,平均値,トラブル要因", ",")
        .Range(.Cells(1, 1), .Cells(1, UBound(bArry) + 1)) = bArry
             
        '荷重ピークデータの最終行を取得
        lRow = .Cells(Rows.Count, 3).End(xlUp).Row
                
        '管理データの書出
        ReDim cArry(1 To 5, 1 To 17)
        
        'シートへ展開するための準備
        cArry(2, 6) = "上限規格値"
        cArry(2, 8) = dUl   '上限規格値
        
        cArry(3, 6) = "下限規格値"
        cArry(3, 8) = dLl   '下限規格値
        
        cArry(4, 6) = "標準偏差"
        cArry(4, 8) = dSd   '標準偏差
        
        cArry(5, 6) = "Cpk"
        cArry(5, 8) = dCpk  'Cpk
        
        cArry(2, 9) = "最大値"
        cArry(2, 11) = dMax  '最大値
        
        cArry(3, 9) = "最小値"
        cArry(3, 11) = dMin  '最小値
        
        cArry(4, 9) = "平均値"
        cArry(4, 11) = dMean '平均
        
        cArry(5, 9) = "中央値"
        cArry(5, 11) = dMed  '中央値
                
        '総数のセット
        Dim lCntNum As Long
        
        lCntNum = UBound(dArry)
        
        cArry(2, 1) = "カウント総数"
        cArry(2, 3) = lCntNum
        
        
        ReDim bArry(1 To 3) As Variant
        
        'トラブル要因コメントをbArryにセット
        For i = 1 To UBound(bArry)
            bArry(i) = sh1.Cells(i + 4, 3).Value
        Next i
        
        'トラブル要因を配列dArryにセット
        dArry = .Range(.Cells(2, 7), .Cells(lRow, 8))
        
        'ディクショナリーを作成
        Dim myDic As Object
        Dim sKey As String
        
        Set myDic = CreateObject("Scripting.Dictionary")
        
        'トラブル要因項目をループ
        For i = 1 To UBound(bArry)
            myDic.Add bArry(i), 0 'キーを辞書に登録
        Next
        
        
        'トラブル要因データをループ
        For i = 1 To UBound(dArry, 1)
            '辞書に登録されている場合
            If myDic.Exists(dArry(i, 1)) Then
                myDic(dArry(i, 1)) = myDic(dArry(i, 1)) + 1 'カウントアップ
            End If
        Next
        
        Dim lNGCnt As Long
        
        lNGCnt = WorksheetFunction.Sum(myDic.Items)
        
        '不適合数のセット
        cArry(1, 14) = "不適合数"
        cArry(1, 16) = "不適合率(%)"
        
        cArry(2, 12) = bArry(1)  'A:溶接切れ"
        cArry(3, 12) = bArry(2)  'B:荷重不足
        cArry(4, 12) = bArry(3)  'C:荷重過多
        
        cArry(5, 12) = "合計"
        cArry(5, 14) = lNGCnt
        cArry(5, 16) = lNGCnt / lCntNum * 100
        
        For i = 1 To myDic.Count
            cArry(i + 1, 14) = myDic.Item(bArry(i))     '各NG数の抽出
            cArry(i + 1, 16) = cArry(i + 1, 14) / lCntNum * 100 '各NG率
        Next i
        
        Set myDic = Nothing
        
        
        
        'OK数
        cArry(4, 1) = "OK数"
        cArry(4, 3) = lCntNum - lNGCnt
        
        '自動列幅調整にセットして見栄えを整える
        .Columns("A:K").AutoFit
            
    End With
        
    Dim sh3 As Worksheet
    
    Set sh3 = Worksheets("工程管理表")
        
    '工程管理表シートの工程管理データ部分を初期化
    'sh3.Range("A5").Resize(UBound(cArry, 1), UBound(cArry, 2)).ClearContents
    '工程管理表シートへ工程管理用データを書き込む
    sh3.Range("A6").Resize(UBound(cArry, 1), UBound(cArry, 2)) = cArry
    
    
End Sub



'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' MakeProcCtrlChart()
'   工程管理表のグラフ作成
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Private Sub MakeProcCtrlChart()

    Dim sh1 As Worksheet
    Dim i As Integer
    
    Set sh1 = Worksheets("工程管理表")

    With sh1
        
        'データエリアの初期化
        .Range(.Range("A40"), .Range("A" & Cells.Rows.Count)).EntireRow.ClearContents
        .Rows("40:" & Rows.Count).RowHeight = 9.75
        
        
        ' 既存グラフの削除
    
        If .ChartObjects.Count > 0 Then
            For i = .ChartObjects.Count To 1 Step -1
                .ChartObjects(i).Delete
            Next i
        End If
    
    End With

    '荷重測定結果チャートの作成
    Call DrawMeasChart
    
    ' チャートの複製
    Call CopyChart("XRグラフ", "工程管理表", 1, 2, "A41", "J41", "J65")
    Call CopyChart("XRグラフ", "工程管理表", 2, 3, "A66", "J66", "J89")
    Call CopyChart("ヒストグラム", "工程管理表", 1, 4, "K41", "P41", "P89")

End Sub


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' FormProcMeasData4Printing()
'   工程管理表の測定結果データ部の成形
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Private Sub FormProcMeasData4Printing()
    
    Dim sh1 As Worksheet
    Dim sh2 As Worksheet
    
    Set sh1 = Worksheets("工程管理表")
    Set sh2 = Worksheets("工程管理用データ")
    
    Dim aArry As Variant
    Dim bArry As Variant
    Dim lRow As Long, lSize As Long, lCntW As Long, lCntR As Long
    Dim i As Integer, j As Integer, k As Integer

    With sh1
    '荷重ピークデータのクリア
        .Range(.Cells(40, 1), .Cells(.Cells.Rows.Count, 1)).ClearContents
    End With
    
    With sh2
        '荷重ピークデータの最終行を取得
        lRow = sh2.Cells(Rows.Count, 3).End(xlUp).Row
      
        '時刻、カウント、ピークをaArryにセット
        aArry = .Range(.Cells(2, 1), .Cells(lRow, 3))
        '5LP測定記録、トラブル要因をbArryにセット
        bArry = .Range(.Cells(2, 7), .Cells(lRow, 7))
    
        'aArryとbArryをマージ
        aArry = MergeVariantArraysCol(aArry, bArry)
    
    Dim lChunkRow As Long
    Dim lChunkStage As Long
    Dim lBlockNum As Long
    
    '一塊50行が3段組
    lChunkRow = 50
    lChunkStage = 4
    
    '一段に格納するデータ数200
    lBlockNum = lChunkRow * lChunkStage
    
    lSize = Int(UBound(aArry, 1) / lBlockNum + 1) * lBlockNum
    
    aArry = RedimPreserveArray(aArry, lSize)
        
    ReDim bArry(1 To (Int(UBound(aArry, 1) / 2 + Int(lSize / lBlockNum))), 1 To 17)
    
    lCntW = 0   '格納先アドレス
    lCntR = 0   '読込先アドレス
    
    For i = 1 To (UBound(aArry, 1) / lBlockNum)
        For j = 1 To lBlockNum
                Select Case j
                
                    Case 1
                        bArry(j + lCntW, 1) = "時刻"
                        bArry(j + lCntW, 2) = "カウント"
                        bArry(j + lCntW, 3) = "荷重ピーク"
                        bArry(j + lCntW, 4) = "トラブル要因"
    
                        bArry(j + lCntW, 5) = "時刻"
                        bArry(j + lCntW, 6) = "カウント"
                        bArry(j + lCntW, 7) = "荷重ピーク"
                        bArry(j + lCntW, 8) = "トラブル要因"
                    
                        bArry(j + lCntW, 9) = "時刻"
                        bArry(j + lCntW, 10) = "カウント"
                        bArry(j + lCntW, 11) = "荷重ピーク"
                        bArry(j + lCntW, 12) = "トラブル要因"
                        
                        bArry(j + lCntW, 13) = "時刻"
                        bArry(j + lCntW, 14) = "カウント"
                        bArry(j + lCntW, 15) = "荷重ピーク"
                        bArry(j + lCntW, 16) = "トラブル要因"
                        
                        lCntW = lCntW + 1
                End Select
            
            For k = 1 To UBound(aArry, 2)
                
                Select Case j
                    Case 1 To 50    '1段目の処理
                        bArry(j + lCntW, k) = aArry(j + lCntR, k)
                    
                    Case 51 To 100  '2段目の処理
                        bArry(j + lCntW - 50, k + 4) = aArry(j + lCntR, k)
                
                    Case 101 To 150 '3段目の処理
                        bArry(j + lCntW - 100, k + 8) = aArry(j + lCntR, k)
                
                    Case 151 To 200 '4段目の処理
                        bArry(j + lCntW - 150, k + 12) = aArry(j + lCntR, k)
                
                End Select
            Next k
        Next j
        lCntW = lCntW + lChunkRow   'bArryの格納先を更新
        lCntR = lCntR + lBlockNum   'aArryの読込先を更新
    Next i

    End With
    
    With sh1
        '工程管理表シートへ印刷フォーマットの測定データを書き込む
        .Range("A91").Resize(UBound(bArry, 1), UBound(bArry, 2)) = bArry
        
        '最終行を取得
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
            '印刷範囲の設定
            .PrintArea = Range(Cells(1, 1), Cells(lRow, 16)).Address
            .LeftHeader = "" '左側ヘッダー：なし
            .CenterHeader = Cells(3, 3) '中央ヘッダー：ロット
            .RightHeader = "&D &T"  '右側ヘッダー：日付 時刻
            .LeftFooter = "" '左側フッター：なし
            .CenterFooter = "&P/&N" '中央フッター：ページ数/総ページ数
            .RightFooter = "&""Verdana""&08" & "Confidencial"   '右フッター：Verdanaフォント、サイズ8で「Confidencial」
        
        End With
        
        .Cells(1, 1).Select
    End With

End Sub




'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' FormXRData(ByRef xArry As Variant)
'   X-R管理用のデータの成形
'
'   xArry：成形に使用するデータ
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
    
    'データを群に分り振る
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
            '群内の最大・最小値の抽出
            If j = 1 Then
                dMax = yArry(i, j + 3)
                dMin = dMax
            Else
                '最大値
                If dMax < yArry(i, j + 3) Then
                    dMax = yArry(i, j + 3)
                End If
                
                '最小値
                If dMin > yArry(i, j + 3) And yArry(i, j + 3) <> "" Then
                    dMin = yArry(i, j + 3)
                End If
            End If
            
            '群平均算出のための準備
            dSum1 = dSum1 + yArry(i, j + 3)
        Next j
        
        '平均値
        If dSum1 <> 0 Then
            yArry(i, 2) = dSum1 / lCnt
            yArry(i, 1) = i
        End If
        
        '偏差
        If dMax <> 0 And dMin <> 0 Then
            yArry(i, 3) = dMax - dMin
        End If
        
    Next i

    With Worksheets("XRdata")
        .Range("A21").Resize(UBound(yArry, 1), UBound(yArry, 2)) = yArry
        .Range("A1").Value = "X管理図データ"
        .Range("A2").Value = "群平均"
        .Range("A3").Value = "X-UCL"
        .Range("A4").Value = "X-LCL"

        .Range("D1").Value = "R管理図データ"
        .Range("D2").Value = "R平均"
        .Range("D3").Value = "R-UCL"
        .Range("D4").Value = "R-LCL"

        ReDim zArry(0, 18)
        
        zArry(0, 0) = "群"
        zArry(0, 1) = "群平均"
        zArry(0, 2) = "R"
        
        For i = 1 To 10
            zArry(0, i + 2) = "n" & i
        Next i
        
        zArry(0, 13) = "X上限管理値"
        zArry(0, 14) = "X下限管理値"
        zArry(0, 15) = "X平均"
        zArry(0, 16) = "R上限管理値"
        zArry(0, 17) = "R下限管理値"
        zArry(0, 18) = "R平均"
        
        
        .Range("A20:S20").Value = zArry
        
        ReDim zArry(1 To 11, 1 To 4)
        
        'X-R係数のセット
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
'  X管理グラフとR管理グラフの作成
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
    Set sh2 = Worksheets("XRグラフ")
    Set sh3 = Worksheets("工程管理表")

        
    '既存グラフの消去
    With sh2
        For i = .ChartObjects.Count To 1 Step -1
            .ChartObjects(i).Delete
        Next i
    End With
    
    'グラフ用のデータを配列に読込む
    lRow = sh1.Cells(Rows.Count, 1).End(xlUp).Row
    aArry = sh1.Range(sh1.Cells(21, 2), sh1.Cells(lRow, 2))
    
    '最大値と最小値の取得（縦軸の設定用）
    iMax = Int(Application.WorksheetFunction.Max(aArry)) + 1
    iMin = Int(Application.WorksheetFunction.Min(aArry)) - 1
    
    
    '========================================================
    '
    ' X管理グラフの作成
    '
    '========================================================
    'グラフ描画エリアの設定
    Set R = sh2.Range("A1:Q22")
    
    'グラフに使用するデータを設定
    With sh1
        Set TrRange = Union(.Range(.Cells(20, 2), .Cells(lRow, 2)), .Range(.Cells(20, 14), .Cells(lRow, 16)))
    End With
    
    'グラフを作成
    With sh2.ChartObjects.Add(R.Left, R.Top, R.Width, R.Height)

        .Name = "X管理図"
        .Chart.ChartType = xlLine
        .Chart.SetSourceData TrRange
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "X管理図"
        .Chart.SeriesCollection.Item(1).Format.Line.Weight = 1
        .Chart.SeriesCollection.Item(4).Format.Line.Weight = 2.5
        .Chart.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(124, 124, 124)
        .Chart.SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(20, 20, 200)
    End With
        
    '縦軸最大値：自動とする
    sh2.ChartObjects(1).Chart.Axes(xlValue).MaximumScaleIsAuto = True
    '縦軸最小値：データ最小 - 1
    sh2.ChartObjects(1).Chart.Axes(xlValue).MinimumScale = iMin
    
        
    '========================================================
    '
    ' R管理グラフの作成
    '
    '========================================================
    'グラフ描画エリアの設定
    Set R = sh2.Range("A23:Q42")
    
    'グラフに使用するデータを設定
    With sh1
        Set TrRange = Union(.Range(.Cells(20, 3), .Cells(lRow, 3)), .Range(.Cells(20, 17), .Cells(lRow, 19)))
    End With
    
    'グラフを作成
    With sh2.ChartObjects.Add(R.Left, R.Top, R.Width, R.Height)

        .Name = "R管理図"
        .Chart.ChartType = xlLine
        .Chart.SetSourceData TrRange
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "R管理図"
        .Chart.SeriesCollection.Item(1).Format.Line.Weight = 1
        .Chart.SeriesCollection.Item(4).Format.Line.Weight = 2.5
        .Chart.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(124, 124, 124)
        .Chart.SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(20, 20, 200)

    End With
    
    sh2.ChartObjects(2).Chart.Axes(xlValue).MaximumScaleIsAuto = True
    '縦軸最小値：データ最小 - 1
    sh2.ChartObjects(2).Chart.Axes(xlValue).MinimumScaleIsAuto = True
    
     '画像としてコピー
'    sh2.ChartObjects(1).CopyPicture
'
'    sh3.Activate
'    sh3.Range("A25:G32").Select 'セルを選択
'    sh3.Paste '貼り付け
    


End Sub


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' DrawMeasChart()
'  測定結果グラフの作成
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
    
    Set sh1 = Worksheets("工程管理用データ")
    Set sh2 = Worksheets("工程管理表")

    
    iUL = Int(sh1.Cells(2, 5))
    iUL = iUL - 3   '縦軸最小値を下限規格値より算出
    
    
    '既存グラフの消去
    With sh2
        For i = .ChartObjects.Count To 1 Step -1
            .ChartObjects(i).Delete
        Next i
    End With
    
    '========================================================
    '
    ' 測定結果グラフの作成
    '
    '========================================================
    'グラフ描画エリアの設定
    Set R = sh2.Range("A11:P36")
    
    lRow = sh1.Cells(Rows.Count, 1).End(xlUp).Row
    
    'グラフに使用するデータを設定
    With sh1
        Set TrRange = .Range(.Cells(1, 3), .Cells(lRow, 6))
    End With
    
    'グラフを作成
    With sh2.ChartObjects.Add(R.Left, R.Top, R.Width, R.Height)

        .Name = "荷重測定結果"
        .Chart.ChartType = xlLine
        .Chart.SetSourceData TrRange
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "荷重測定結果"
        .Chart.SeriesCollection.Item(1).Format.Line.Weight = 1
        .Chart.SeriesCollection.Item(4).Format.Line.Weight = 2.5
        .Chart.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(124, 124, 124)
        .Chart.SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(20, 20, 200)
    End With
        
    '縦軸最大値：自動とする
    sh2.ChartObjects(1).Chart.Axes(xlValue).MaximumScaleIsAuto = True
    '縦軸最小値：データ最小 - 1
    sh2.ChartObjects(1).Chart.Axes(xlValue).MinimumScale = iUL

End Sub



'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' CalcCpk(dUl As Double, dLl As Double, dMean As Double, dSd As Double) As Double
'
'  Cpkの計算
'   dUl：上限管理値
'   dLl：下限管理値
'   dMean：平均値
'   dSd：標準偏差
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Private Function CalcCpk(dUl As Double, dLl As Double, dMean As Double, dSd As Double) As Double

    Dim dK As Double
    Dim dCp As Double
    Dim dCpk As Double
    
    dCpk = 0
    
    If dSd Then
        dCp = (dUl - dLl) / (6 * dSd)   'Cpの計算
        dK = Abs((((dUl + dLl) / 2) - dMean)) / ((dUl - dLl) / 2) 'Kの計算
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
    
    
    'コピー
    sh1.ChartObjects(objNum1).Copy
       
    sh2.Activate
    sh2.Range(objLeft).Select 'セルを選択
    sh2.Paste '貼り付け
    
    sh2.ChartObjects(objNum2).Left = Range(objLeft).Left '左の位置
    sh2.ChartObjects(objNum2).Top = Range(objTop).Top '上の位置
    sh2.ChartObjects(objNum2).Width = Range(objLeft & ":" & objWidth).Width '幅
    sh2.ChartObjects(objNum2).Height = Range(objLeft & ":" & objWidth).Height '高さ
    
End Sub


Private Function ShewhartConstant() As Variant

    Dim aArry As Variant

    aArry = ActiveSheet.Range("T1:W11").Value
    
    ReDim aArry(1 To 11, 1 To 4)
    
    aArry(1, 2) = "n係数"
    
    aArry(2, 1) = "群数"
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



'--- 2つのVariant配列を結合する（列方向） ---'
Private Function MergeVariantArraysCol(vArray1 As Variant, vArray2 As Variant) As Variant
    
    '結合する配列のサイズ
    Dim vArray1_row As Long
    Dim vArray1_col As Long
    Dim vArray2_row As Long
    Dim vArray2_col As Long
    vArray1_row = UBound(vArray1, 1)
    vArray1_col = UBound(vArray1, 2)
    vArray2_row = UBound(vArray2, 1)
    vArray2_col = UBound(vArray2, 2)
    
    '結合後の配列のサイズ
    Dim newArray_row As Long
    Dim newArray_col As Long
    newArray_row = Application.WorksheetFunction.Max(vArray1_row, vArray2_row)
    newArray_col = vArray1_col + vArray2_col
    
    '結合後の配列
    Dim newArray As Variant
    ReDim newArray(1 To newArray_row, 1 To newArray_col)
    
    '配列を結合する
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


'マージバブルソート
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

