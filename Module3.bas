Attribute VB_Name = "Module3"
Option Explicit


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' svFile()
'   �w�߂��ꂽ�t�H���_�Ɏw�肳�ꂽ�t�@�C����.xlxs�`���ŕۑ�����
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
        Msg = svFile & vbCrLf & vbCrLf & "�����݂��܂��B�㏑�����܂����H"
        If MsgBox(Msg, vbYesNo) = vbNo Then Exit Sub
    End If
    

    Application.DisplayAlerts = False
    
    Application.StatusBar = "�ۑ����c"
    
    '�`����~����
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
            .Name = "�l�r �o�S�V�b�N"
            End With
        End If
    Next ws
    
'    For i = 1 To n
'        wbNew.Worksheets(i).Cells.Font.Name = "�l�r �o�S�V�b�N"
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
    
    sh.Name = "���O�f�[�^"
    
    
    sh.Activate
    sh.Cells(1, 1).Select
            
    Set sh = Worksheets("�H���Ǘ��\")
    
    With sh.PageSetup
        .Orientation = xlLandscape '����������������ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
    End With
    
    If sh.Cells(1, 16).Value = "" Then
        sh.Cells(1, 16).Value = Format(Now(), "yyyy-mm-dd")
    End If
    
    sh.Activate
    sh.Cells(1, 1).Select
        
    wbNew.Save
    
    '�`�������
    Application.ScreenUpdating = True
    
    Application.StatusBar = "�ۑ����܂���"
        
    Application.DisplayAlerts = True
    
    eTime = Timer
    pTime = eTime - sTime
    
    Debug.Print "��������"; pTime & vbCrLf
    
    
End Sub

Sub Column_Print(ByVal wbName As String) '���ׂĂ̗��1�y�[�W�Ɉ��
    
    Dim wb As Workbook
    Dim sh As Worksheet
    
    Set wb = Workbooks("wbName")
    Set sh = Worksheets("�H���Ǘ��\")
    
    
    With wb.sh.PageSetup
        .Orientation = xlLandscape '����������������ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
    End With
End Sub


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' CheckparentFolder(TargetFolder)
'   �w�肳�ꂽ�p�X�̐e�t�H���_�����邩�`�F�b�N�B������΍쐬����
'
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Private Sub CheckparentFolder(TargetFolder)
    Dim parentFolder As String, curFolder As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ''�����Ώۃt�H���_�́A�e�t�H���_�����擾����
    parentFolder = FSO.GetParentFolderName(TargetFolder)
    curFolder = TargetFolder
    

    
    If Not FSO.FolderExists(parentFolder) Then
        ''�e�t�H���_�����݂��Ȃ�������A
        ''�e�t�H���_��V�����Ώۃt�H���_�Ƃ���
        ''�������g(Sub CheckparentFolder)���Ăяo��
        Call CheckparentFolder(parentFolder)
    End If
    ''�V�����t�H���_�����
    FSO.CreateFolder TargetFolder
    Set FSO = Nothing
End Sub


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
' LotNameCopy()
'   ���b�g�ԍ����t�@�C���l�[���փR�s�[����
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
'�@ replaceNGchar(ByVal sourceStr As String, _
'        Optional ByVal replaceChar As String = "") As String
'�@�@�t�@�C�����Ɏg���Ȃ�������u��������
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
    '�������t�H�[���̕\��(���[�_��)
    UserForm2.Show
End Sub

Sub CloseUserForm2()
    '�������t�H�[��������
    Unload UserForm2
End Sub

