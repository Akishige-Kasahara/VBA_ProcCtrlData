Attribute VB_Name = "Module6"
Option Explicit


Public Sub InitProcCtrlSheet()


    Dim sh As Worksheet
    Dim i As Integer
    Dim lRow As Long
    Dim sTime As Double
    Dim eTime As Double
    Dim pTime As Double
    
    sTime = Timer
    
    Set sh = Worksheets("工程管理表")

    '描画を停止
    Application.DisplayAlerts = False

    With sh
        
        .Cells.FormatConditions.Delete  'シート内の条件付き書式を削除
        
        '測定結果は小数点以下を1桁に設定
        .Range("H7：H8").NumberFormatLocal = "0.0"
        .Range("K7:K8").NumberFormatLocal = "0.0"
        .Range(.Cells(92, 3), .Cells(.Cells.Rows.Count, 3)).NumberFormatLocal = "0.0"
        .Range(.Cells(92, 9), .Cells(.Cells.Rows.Count, 9)).NumberFormatLocal = "0.0"
        .Range(.Cells(92, 15), .Cells(.Cells.Rows.Count, 15)).NumberFormatLocal = "0.0"
    
        '不良率の小数点以下を2桁に設定
        .Range("P7：P10").NumberFormatLocal = "0.00"
        
    
        '時刻の列は12:34:56のフォーマットとする
        .Range(.Cells(92, 1), .Cells(.Cells.Rows.Count, 1)).NumberFormatLocal = "hh:mm:ss"
        .Range(.Cells(92, 7), .Cells(.Cells.Rows.Count, 7)).NumberFormatLocal = "hh:mm:ss"
        .Range(.Cells(92, 13), .Cells(.Cells.Rows.Count, 13)).NumberFormatLocal = "hh:mm:ss"

        lRow = 10086
    
        'データエリアの初期化
        .Range(.Range("A40"), .Range("A" & Cells.Rows.Count)).EntireRow.ClearContents
        .Rows("40:" & Rows.Count).RowHeight = 9.75
        '改ページを挿入
        .Rows(40).PageBreak = xlPageBreakManual
        
        'データエリアのフォントの設定
        Dim font1 As Font

        Set font1 = .Range(Cells(40, 1), Cells(lRow, 17)).Font
        font1.Name = "ＭＳ Ｐゴシック"
        font1.Size = 9
            
        For i = 91 To lRow Step 51
            '表の書式をコピー
            .Range(.Cells(i, 1), .Cells(i + 50, 17)).Copy
            .Range(.Cells(i + 51, 1), .Cells(i + 101, 17)).PasteSpecial _
                    Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            '改ページを挿入
            .Rows(i).PageBreak = xlPageBreakManual
        Next i
        
        '印刷の向きを横にする
        .PageSetup.Orientation = xlLandscape
                
        .Activate
        .Cells(1, 1).Select
        
        End With

        With sh.Range(sh.Cells(92, 1), sh.Cells(lRow, 5))
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND($C92<>"""",$C92<>""荷重ピーク"",OR($C92>$H$7,$C92<$H$8,$E92<>""""))"
            .FormatConditions(1).Interior.Color = vbRed
        End With

        With sh.Range(sh.Cells(92, 7), sh.Cells(lRow, 11))
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND($I92<>"""",$I92<>""荷重ピーク"",OR($I92>$H$7,$I92<$H$8,$K92<>""""))"
            .FormatConditions(1).Interior.Color = vbRed
        End With

        With sh.Range(sh.Cells(92, 13), sh.Cells(lRow, 17))
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:="=AND($O92<>"""",$O92<>""荷重ピーク"",OR($O92>$H$7,$O92<$H$8,$Q92<>""""))"
            .FormatConditions(1).Interior.Color = vbRed
        End With

    
    '描画を元に戻す
    Application.DisplayAlerts = True

    eTime = Timer

    pTime = eTime - sTime
    
    Debug.Print "処理時間"; pTime & vbCrLf
    
End Sub



Public Sub ClearData()

    Dim sh As Worksheet
    Dim i As Long
    
    Set sh = Worksheets("Main1")
    
    '描画を停止
    Application.ScreenUpdating = False
    
    With sh
    
        .Range("B3").Value = ""
        .Range("I3").Value = ""
        .Range("I5").Value = ""
        .Range("B7").Value = ""
        .Range("A13:G30013").Value = ""
    
    End With
    
    Set sh = Worksheets("工程管理表")
    
    With sh
        
        .Range("C1").Value = ""
        .Range("P1").Value = ""
        .Range("P3").Value = ""
        .Range("C7:C9").Value = ""
        .Range("C7:C9").Value = ""
        .Range("H7:H10").Value = ""
        .Range("K7:K10").Value = ""
        .Range("N7:P10").Value = ""
        .Range("A90:Q12000").Value = ""
        
        For i = .ChartObjects.Count To 1 Step -1
            .ChartObjects(i).Delete
        Next i
        
    End With
    
    Set sh = Worksheets("工程管理用データ")
    
    With sh
        .Cells.Delete
    End With
    
    Set sh = Worksheets("XRグラフ")
    
    With sh
        For i = .ChartObjects.Count To 1 Step -1
            .ChartObjects(i).Delete
        Next i
    
    End With
        
    Set sh = Worksheets("XRdata")
    
    With sh
        .Cells.Delete
    End With
    
    Set sh = Worksheets("ヒストグラム")
    
    With sh
        .Cells.Delete
        
        For i = .ChartObjects.Count To 1 Step -1
            .ChartObjects(i).Delete
        Next i
    
    End With
    
    
    '描画をする
    Application.ScreenUpdating = True
    
End Sub



