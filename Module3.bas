Attribute VB_Name = "Module3"
Option Explicit


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' svFile()
'   指令されたフォルダに指定されたファイル名.xlxs形式で保存する
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Public Sub svFile()

    
    
    Dim sTime As Double
    Dim eTime As Double
    Dim pTime As Double

    sTime = Timer


    Dim wb1 As Workbook
    Dim wbNew As Workbook
    Dim sh1 As Worksheet
    Dim FSO As Scripting.FileSystemObject
    Dim svPath As String
    Dim newPath As String
    Dim fileNew As String
    Dim svFile As String
    Dim Msg As String
    
    Set wb1 = ThisWorkbook
    Set sh1 = Worksheets("Main1")
    Set FSO = New Scripting.FileSystemObject

    newPath = sh1.Range("B2").Value
    fileNew = sh1.Range("I3").Value
        
    If newPath = "" Or fileNew = "" Then Exit Sub
    
    fileNew = replaceNGchar(fileNew, "_")
    
    If FSO.FolderExists(newPath) = False Then
        FSO.CreateFolder (newPath)
    End If
    Set FSO = Nothing
    
    svFile = svPath & newPath & "\" & fileNew & "_" & Format(Now(), "yymmdd_hhmm") & ".xlsx"

    If Dir(svFile) <> "" Then
        Msg = svFile & vbCrLf & vbCrLf & "が存在します。上書きしますか？"
        If MsgBox(Msg, vbYesNo) = vbNo Then Exit Sub
    End If
    

    Application.DisplayAlerts = False
    
    Application.StatusBar = "保存中…"
    
    '描画を停止する
    Application.ScreenUpdating = False
    
    Set wbNew = Workbooks.Add
        
    wbNew.SaveAs svFile
    
    ' UserForm2.ProgressBar1.Value = 20

    wb1.Sheets().Copy _
    before:=wbNew.Sheets(1)

    Dim n As Integer
    Dim i As Integer
    
    n = Worksheets.Count

    Dim ws As Worksheet
    
    For Each ws In wbNew.Worksheets
        If ws.Visible Then
            With ws.Cells.Font
            .Name = "ＭＳ Ｐゴシック"
            End With
        End If
    Next ws
    
'    For i = 1 To n
'        wbNew.Worksheets(i).Cells.Font.Name = "ＭＳ Ｐゴシック"
'    Next i

    DoEvents

    Dim sh As Worksheet
    Set sh = Worksheets("Main1")
    
    With sh.Shapes
        For i = .Count To 1 Step -1
            If .Item(i).Type = msoTextBox Then
                .Item(i).Delete
            End If
            
            If .Item(i).Type = msoShapeRectangle Then
                .Item(i).Delete
            End If
            
        Next
    End With
    
    DoEvents
    
    sh.Buttons.Delete
    
    sh.Name = "ログデータ"
    
    
    sh.Activate
    sh.Cells(1, 1).Select
            
    Set sh = Worksheets("工程管理表")
    
    With sh.PageSetup
        .Orientation = xlLandscape '印刷向きを横方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
    End With
    
    If sh.Cells(1, 16).Value = "" Then
        sh.Cells(1, 16).Value = Format(Now(), "yyyy-mm-dd")
    End If
    
    sh.Activate
    sh.Cells(1, 1).Select
        
    wbNew.Save
    
    '描画をする
    Application.ScreenUpdating = True
    
    Application.StatusBar = "保存しました"
        
    Application.DisplayAlerts = True
    
    eTime = Timer
    pTime = eTime - sTime
    
    Debug.Print "処理時間"; pTime & vbCrLf
    
    
End Sub

Sub Column_Print(ByVal wbName As String) 'すべての列を1ページに印刷
    
    Dim wb As Workbook
    Dim sh As Worksheet
    
    Set wb = Workbooks("wbName")
    Set sh = Worksheets("工程管理表")
    
    
    With wb.sh.PageSetup
        .Orientation = xlLandscape '印刷向きを横方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
    End With
End Sub


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' CheckparentFolder(TargetFolder)
'   指定されたパスの親フォルダがあるかチェック。無ければ作成する
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Private Sub CheckparentFolder(TargetFolder)
    Dim parentFolder As String, curFolder As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ''調査対象フォルダの、親フォルダ名を取得する
    parentFolder = FSO.GetParentFolderName(TargetFolder)
    curFolder = TargetFolder
    

    
    If Not FSO.FolderExists(parentFolder) Then
        ''親フォルダが存在しなかったら、
        ''親フォルダを新しい対象フォルダとして
        ''自分自身(Sub CheckparentFolder)を呼び出す
        Call CheckparentFolder(parentFolder)
    End If
    ''新しいフォルダを作る
    FSO.CreateFolder TargetFolder
    Set FSO = Nothing
End Sub


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' LotNameCopy()
'   ロット番号をファイルネームへコピーする
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Private Sub LotNameCopy()

    Dim sh1 As Worksheet
    Dim SlotName As String
    Dim SFileName As String
    
    Set sh1 = Sheets("Main1")
    
    With sh1
        SlotName = .Cells(3, 2)
        SFileName = replaceNGchar(SlotName, "_")
        .Cells(3, 9) = SFileName
    End With
    
End Sub




'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'　 replaceNGchar(ByVal sourceStr As String, _
'        Optional ByVal replaceChar As String = "") As String
'　　ファイル名に使えない文字を置き換える
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Private Function replaceNGchar(ByVal sourceStr As String, _
        Optional ByVal replaceChar As String = "") As String

    Dim tempStr As String
    
    tempStr = sourceStr
    tempStr = Replace(tempStr, "\", replaceChar)
    tempStr = Replace(tempStr, "/", replaceChar)
    tempStr = Replace(tempStr, ":", replaceChar)
    tempStr = Replace(tempStr, "*", replaceChar)
    tempStr = Replace(tempStr, "?", replaceChar)
    tempStr = Replace(tempStr, """", replaceChar)
    tempStr = Replace(tempStr, "<", replaceChar)
    tempStr = Replace(tempStr, ">", replaceChar)
    tempStr = Replace(tempStr, "|", replaceChar)
    tempStr = Replace(tempStr, "[", replaceChar)
    tempStr = Replace(tempStr, "]", replaceChar)

    replaceNGchar = tempStr
End Function


Sub ShowUserForm2()
    '処理中フォームの表示(モーダル)
    UserForm2.Show
End Sub

Sub CloseUserForm2()
    '処理中フォームを消す
    Unload UserForm2
End Sub

