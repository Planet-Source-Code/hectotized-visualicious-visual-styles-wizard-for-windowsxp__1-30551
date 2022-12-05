Attribute VB_Name = "modFileInfo"
Global counter As Integer
Public Function NullPad(strData As String) As String

If strData = "" Then Exit Function
Dim lenData As Long

For i = 1 To Len(strData)
    tempStr = tempStr & Chr(0) & Mid(strData, i, 1)
Next

NullPad = Chr(1) & tempStr


End Function
Public Function ReplaceIt(Original As Variant, Item As String, Replace As String) As String

If InStr(Original, Item) = False Then
    ReplaceIt = Original
    Exit Function
End If

nStage$ = Original
Do Until InStr(nStage$, Item) = 0
    lSide$ = Left$(nStage$, InStr(nStage$, Item) - 1)
    rSide$ = Right$(nStage$, (Len(nStage$) - Len(lSide$) - Len(Item)))
    nStage$ = lSide$ & Replace & rSide$
Loop
ReplaceIt = nStage$


End Function
'keep
Public Function GetCompanyName(strFile As String)

Dim tempFile As String
Dim pos As Long
Dim StartPos As Long, EndPos As Long

fileText$ = "CompanyName"
nextText$ = "FileDescription"

Open strFile For Binary As #1
    tempFile = Space(LOF(1))
    Get #1, , tempFile
Close #1

pos = InStr(tempFile, NullPad("StringFileInfo"))

If pos = 0 Then
    pos = InStr(tempFile, "StringFileInfo")
    If pos = 0 Then pos = 1
    pnStart = InStr(pos, tempFile, fileText$)
    fileLength% = 12
Else
    pnStart = InStr(pos, tempFile, NullPad(fileText$))
    nextText$ = NullPad(nextText$)
    fileLength% = 26
End If

If pnStart > 0 Then
    StartPos = pnStart + fileLength%
    EndPos = InStr(StartPos, tempFile, String(3, Chr(0)))
    
    If InStr(Mid(tempFile, StartPos, EndPos - StartPos), nextText$) <> 0 Then
        For i = 1 To 255
            If CInt(Asc(Mid(tempFile, StartPos + i, 1))) <= 31 Then
                EndPos = StartPos + (i - 1)
                Exit For
            End If
        Next i
        counter = counter + 1
    End If
    
    FileInfo = Mid(tempFile, StartPos, EndPos - StartPos)
    GetCompanyName = ReplaceIt(FileInfo, Chr(0), "")
End If


End Function
'keep
Public Function GetFileDescription(strFile As String)

Dim tempFile As String
Dim pos As Long
Dim StartPos As Long, EndPos As Long

fileText$ = "FileDescription"
nextText$ = "FileVersion"

Open strFile For Binary As #1
    tempFile = Space(LOF(1))
    Get #1, , tempFile
Close #1

pos = InStr(tempFile, NullPad("StringFileInfo"))

If pos = 0 Then
    pos = InStr(tempFile, "StringFileInfo")
    If pos = 0 Then pos = 1
    pnStart = InStr(pos, tempFile, fileText$)
    fileLength% = 16
Else
    pnStart = InStr(pos, tempFile, NullPad(fileText$))
    nextText$ = NullPad(nextText$)
    fileLength% = 34
End If

If pnStart > 0 Then
    StartPos = pnStart + fileLength%
    EndPos = InStr(StartPos, tempFile, String(3, Chr(0)))
    
    If InStr(Mid(tempFile, StartPos, EndPos - StartPos), nextText$) <> 0 Then
        For i = 1 To 255
            If CInt(Asc(Mid(tempFile, StartPos + i, 1))) <= 31 Then
                EndPos = StartPos + i
                Exit For
            End If
        Next i
        counter = counter + 1
    End If
    
    FileInfo = Mid(tempFile, StartPos, EndPos - StartPos)
    GetFileDescription = ReplaceIt(FileInfo, Chr(0), "")
End If


End Function

'keep
Public Function GetProductName(strFile As String)

Dim tempFile As String
Dim pos As Long
Dim StartPos As Long, EndPos As Long

fileText$ = "ProductName"
nextText$ = "ProductVersion"

Open strFile For Binary As #1
    tempFile = Space(LOF(1))
    Get #1, , tempFile
Close #1

pos = InStr(tempFile, NullPad("StringFileInfo"))

If pos = 0 Then
    pos = InStr(tempFile, "StringFileInfo")
    If pos = 0 Then pos = 1
    pnStart = InStr(pos, tempFile, fileText$)
    fileLength% = 12
Else
    pnStart = InStr(pos, tempFile, NullPad(fileText$))
    fileLength% = 26
End If

If pnStart > 0 Then
    StartPos = pnStart + fileLength%
    EndPos = InStr(StartPos, tempFile, String(3, Chr(0)))
    
    If InStr(Mid(tempFile, StartPos, EndPos - StartPos), nextText$) <> 0 Then
        For i = 1 To 255
            If CInt(Asc(Mid(tempFile, StartPos + i, 1))) <= 31 Then
                EndPos = StartPos + (i - 1)
                Exit For
            End If
        Next i
        counter = counter + 1
    End If
    
    FileInfo = Mid(tempFile, StartPos, EndPos - StartPos)
    GetProductName = ReplaceIt(FileInfo, Chr(0), "")
End If


End Function

   
