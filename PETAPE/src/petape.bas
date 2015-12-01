Attribute VB_Name = "petape"
Option Explicit

'
' 参照ボタンクリック
'
Sub 参照_Click()

    Dim Shell, myPath
    
    Dim defFolder As String
    defFolder = Range("C4").Value2
    
    Set Shell = CreateObject("Shell.Application")
    Set myPath = Shell.BrowseForFolder(&O0, "フォルダを選んでください", &H1 + &H10, "C:\")
    
    If Not myPath Is Nothing Then Range("C4").Value2 = myPath.Items.Item.path
    
    Set Shell = Nothing
    Set myPath = Nothing
End Sub


'
' 実行ボタンクリック
'
Sub 実行_Click()

    Application.ScreenUpdating = False

    ' init compornent
    
    Dim comFSO As FileSystemObject
    Set comFSO = New FileSystemObject
    
    ' get parameters
    
    Dim i As Integer
    Dim path As String, target As String, extentions() As String, rowspan As Integer
    
    path = Range("C4").Value2
    target = Range("C6").Value2
    extentions = Split(Range("C7").Value2, ",")
    rowspan = Range("C8").Value2
    
    For i = LBound(extentions) To UBound(extentions)
        extentions(i) = Trim(extentions(i))
    Next
    
    ' validate parameters
    
    If (Not comFSO.FolderExists(path)) Then
        MsgBox ("Please select 'Image source directory'.")
        Exit Sub
    End If
    
    If (target <> "ディレクトリ直下の画像のみ" And target <> "サブフォルダごとにシート作成") Then
        MsgBox ("Type of execute target is invalid.")
        Exit Sub
    End If
    
    ' create book
    
    Dim newbook As Workbook
    Set newbook = Workbooks.Add
    
    ' paste images
    Dim rootFolder As Folder
    Dim subFolder As Folder
    Dim subFolderPath As String
    Dim appendSheet As Worksheet
    
    Set rootFolder = comFSO.GetFolder(path)
    
    If (target = "ディレクトリ直下の画像のみ") Then
        Call paste_images(newbook.Sheets(1), rootFolder, extentions, rowspan)
        newbook.Worksheets(3).Delete
        newbook.Worksheets(2).Delete
    ElseIf (target = "サブフォルダごとにシート作成") Then
    
        For Each subFolder In rootFolder.SubFolders
            Set appendSheet = newbook.Sheets.Add(After:=newbook.Worksheets(newbook.Worksheets.Count))
            subFolderPath = subFolder
            appendSheet.Name = get_filename(subFolderPath)
            
            Call paste_images(appendSheet, subFolder, extentions, rowspan)
            
        Next subFolder
        Application.DisplayAlerts = False
        newbook.Worksheets(3).Delete
        newbook.Worksheets(2).Delete
        newbook.Worksheets(1).Delete
        Application.DisplayAlerts = True
        
        newbook.Worksheets(1).Select
    End If
    
    Application.ScreenUpdating = True

End Sub

'
' ワークシートに、指定したディレクトリ直下の画像を、垂直方向に並べて貼り付ける
'
Private Sub paste_images(ByRef sh As Worksheet, ByRef targetFolder As Folder, ByRef extentions() As String, ByVal rowspan As Integer)
    
    ' init compornent
    Dim comFSO As FileSystemObject
    Set comFSO = New FileSystemObject
    
    ' init variant
    Dim insertPointCell As Range
    Dim nextInsertPointPx As Long
    
    Set insertPointCell = sh.Range("A2")
    nextInsertPointPx = insertPointCell.Top
     
    ' for each files matching extention
    Dim targetFilepath As String
    Dim addedPicture As Shape
    Dim targetFile As File
    
    ' counter
    Dim i, j As Integer
    
    For Each targetFile In targetFolder.Files
        targetFilepath = targetFile
        
        For i = LBound(extentions) To UBound(extentions)
            If (InStrRev(targetFilepath, extentions(i)) = (Len(targetFilepath) - Len(extentions(i)) + 1)) Then
                ' compute pixel and paste image
                
                ' paste image
                Set addedPicture = sh.Shapes.addPicture(targetFilepath, False, True, 0, insertPointCell.Top, 0, 0)
                With addedPicture
                    .ScaleHeight 1, msoTrue
                    .ScaleWidth 1, msoTrue
                End With
                    
                ' update pixel of next insert point
                nextInsertPointPx = nextInsertPointPx + addedPicture.Height
                
                ' move to next insert point
                Do While (insertPointCell.Top < nextInsertPointPx)
                    Set insertPointCell = insertPointCell.Offset(1, 0)
                Loop
                
                For j = 1 To rowspan
                    Set insertPointCell = insertPointCell.Offset(1, 0)
                Next j
                
                nextInsertPointPx = insertPointCell.Top
            End If
        Next
    Next targetFile
End Sub

'
' フォルダ込みファイルパスから、ファイル名・拡張子を抽出する
'
Private Function get_filename(ByRef path As String) As String
    get_filename = Mid(path, InStrRev(path, "\") + 1)
End Function


