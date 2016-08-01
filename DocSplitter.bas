Sub FormatDocSplitter()
    Dim Count As Long       'Total number of document split iterations
    Dim Section As Range    'Range for each Section in Doc
    Dim Header As Range     'Range for each Header in a Section
    Dim HeaderNum As Range  'Range for each Number in a Header
    Dim R As Range          'Range for initial doc cleanup
    Dim numTrack As Double  'Stores previous Number in a Header (for error detection)
    Dim startDelimiter As String
    Dim endDelimiter As String
    Dim endDelimiter2 As String
    Dim maxFileName As Integer
    Dim msgBoxResult As Integer
    Dim originalDocName As String
    originalDocName = ActiveDocument.name
    Application.ScreenUpdating = False
    ActiveDocument.Range.ParagraphFormat.SpaceAfter = 10
    
    startDelimiter = "Return To Table Of Contents"  '<- DEFINE: Start delimiter for document splitting
    endDelimiter = "(-{5,})"                        '<- DEFINE: End delimiter for document splitting
    endDelimiter2 = "(_{5,})"                       '<- DEFINE: Another end delimiter for document splitting
    maxFileName = 160                               '<- DEFINE: Sections with larger names will be truncated to this size and have " ..." appended.
     
    Call RemoveAllHyperlinks 'Hyperlinks will not work as delimiters
    Call DeleteShapes 'Remove unnecessary line shapes from document (unnecessary convert to png for html)
    
    'Reformatting Delimiters to remove them from documents...
    With ActiveDocument.Content.Find 'Remove Glossary/Content Index
        .ClearFormatting
        .MatchCase = True
        .MatchWildcards = True
        .Text = "CONTENTS*FOUNTAIN AGRICOUNSEL LLC WEEKLY INDUSTRY NEWS REPORT"
        .Replacement.ClearFormatting
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceOne, Forward:=True, _
        Wrap:=wdFindContinue
    End With
      With ActiveDocument.Content.Find 'Add startDelimiter to first section
        .ClearFormatting
        .MatchCase = True
        .MatchWildcards = True
        .Text = "[^l^13]1.0*[^l^13]"
        .Replacement.ClearFormatting
        .Replacement.Text = startDelimiter & "^13"
        .Execute Replace:=wdReplaceOne, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    With ActiveDocument.Content.Find 'Replace StartDelimiter with whitespace code
        .ClearFormatting
        .MatchCase = False
        .Text = startDelimiter
        .Replacement.ClearFormatting
        .Replacement.Text = "^t^t^t^t"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    With ActiveDocument.Content.Find 'Replace endDelimiter with whitespace code
        .ClearFormatting
        .MatchWildcards = True
        .MatchCase = False
        .Text = endDelimiter
        .Replacement.ClearFormatting
        .Replacement.Text = "^t^l^t"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    With ActiveDocument.Content.Find 'Replace endDelimiter2 with whitespace code
        .ClearFormatting
        .MatchWildcards = True
        .MatchCase = False
        .Text = endDelimiter2
        .Replacement.ClearFormatting
        .Replacement.Text = "^t^l^t"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    With ActiveDocument.Content.Find
        .ClearFormatting
        .MatchWildcards = True
        .MatchCase = False
        .Text = "^12"
        .Replacement.ClearFormatting
        .Replacement.Text = "^13"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    Call URLtoHyperlink '<- Replace all URLs with the relevant hyperlinks
    
    Set Section = ActiveDocument.Range.Duplicate
    'Find Sections
    With Section.Find
        '(*) wildcard separates the two delimiters and selects all of the document between.
        .Text = "^t^t^t^t*^t^l^t"
        .MatchWildcards = True
        While .Execute
            Set Header = Section.Duplicate
            'Find Header within Sections
            With Header.Find
                .ClearFormatting
                .Text = "[^l^13]([1-9]).([1-9])*[^l^13]" '<- Finds text from Header. Ex: "1.2 Some Title Here"
                .MatchWildcards = True
                If .Execute Then
                    .Parent.Bold = True
                    'Find Number within Header (For error checking)
                    Set HeaderNum = Header.Duplicate
                    With HeaderNum.Find
                        .Text = "([1-9]{1,3}).([0-9]{1,3})" '<- Finds number from header. Ex: "1.2"
                        .MatchWildcards = True
                        If .Execute Then
                            If numTrack > 0 And _
                            Left(HeaderNum.Text, InStr(HeaderNum.Text, ".")) _
                            = Left(CStr(numTrack), InStr(CStr(numTrack), ".")) Then '<- Checks that previous header is not a new Integer
                                If Len(HeaderNum.Text) = 3 And CDbl(HeaderNum.Text) - numTrack > 0.11 Then
                                   ' msgBoxResult = MsgBox("SPLIT ERROR: (.1) may have occurred between " & _
                                    numTrack & " and " & HeaderNum.Text & _
                                    " Check that this section was copied.", vbOKCancel, originalDocName & " - No Header Error")
                                    'Select Case msgBoxResult
                                     '   Case vbCancel
                                      '      Debug.Print "ERROR CANCELLED: Ended on " & originalDocName
                                       '     Exit Sub
                                    'End Select
                                    Debug.Print "SPLIT ERROR in " & ActiveDocument.name & " between (" & numTrack & " - " & HeaderNum.Text & ") "
                                ElseIf Len(HeaderNum.Text) = 4 And CDbl(HeaderNum.Text) - numTrack > 0.011 Then
                                   ' msgBoxResult = MsgBox("SPLIT ERROR: (.01) may have occurred between " & _
                                    'numTrack & " and " & HeaderNum.Text & _
                                    " Check that this section was copied.", vbOKCancel, originalDocName & "No Header Error")
                                    'Select Case msgBoxResult
                                     '   Case vbCancel
                                      '      Debug.Print "ERROR CANCELLED: Ended on " & originalDocName
                                       '     Exit Sub
                                    'End Select
                                    Debug.Print "SPLIT ERROR in " & ActiveDocument.name & " between (" & numTrack & " - " & HeaderNum.Text & ")"
                                End If
                            End If
                            numTrack = CDbl(HeaderNum.Text)
                        End If
                    End With
                    .Replacement.Text = "^t^l^t"
                    
                    Call CopyAndSave(Section, Header, maxFileName) '<- Subroutine for saving section as doc
                    Count = Count + 1
                Else
                    'msgBoxResult = MsgBox("ERROR: Could not find header for section " & Count & _
                    ". Section text: " & Section.Text, vbOKCancel, "No Header Error")
                    'Select Case msgBoxResult
                    '    Case vbCancel
                    '        Debug.Print "ERROR CANCELLED"
                    '        Exit Sub
                   'End Select
                    Debug.Print "Error: Coult not find header for section " & Count & " - " & originalDocName
                End If
            End With
            Header.Collapse wdCollapseEnd
            Header.End = Header.Parent.Range.End
        Wend
    End With
    
    'MsgBox "Finished splitting " & originalDocName & " into " & Count & " pieces."
    Debug.Print "!FINISH: " & originalDocName & " into " & Count & " pieces." '& vbCr & vbCr
    Section.Collapse wdCollapseEnd
    Section.End = Section.Parent.Range.End
    Set Header = Section.Duplicate
End Sub

'Save a new file from range Section within the parent file with fileName from Range Header within Section
Private Static Sub CopyAndSave(Section As Range, Header As Range, maxFileName As Integer)
    Dim name As String
    Header.Select
    Selection.ClearFormatting
    Dim D As Document
   
    'File name cannot contain \ / : * ? " < > |
    name = Replace(Header.Text, Chr(13), "")
    name = Application.CleanString(name)
    name = Replace(name, "\", "-")
    name = Replace(name, "/", "-")
    name = Replace(name, ":", "-")
    name = Replace(name, "?", "-")
    name = Replace(name, "*", "")
    name = Replace(name, """", "")
    name = Replace(name, "<", " ")
    name = Replace(name, ">", " ")
    name = Replace(name, "|", " ")
    name = Replace(name, "    ", "")
    name = Replace(name, Chr(10), "")
    name = Replace(name, Chr(13), "")
    name = Replace(name, Chr(9), "")
    name = Replace(name, "‘", "'")
    name = Replace(name, "’", "'")
    name = Replace(name, "“", "'")
    name = Replace(name, "”", "'")
    name = Replace(name, " ", "")
    name = Replace(name, "®", "(R)")
    name = Replace(name, "™", "(TM)")
    name = Replace(name, "™", "(TM)")
    name = Replace(name, "£", "(E)")
    name = Replace(name, "", " ")
    name = Replace(name, "–", "-")
    name = Replace(name, "—", "-")
    name = Trim(name)
    name = StripAccent(name)
    'Debug.Print name
   
    'Reformatting Header
    Header.Select
    Selection.Font.Bold = True
    Selection.Font.Grow
    
    'Truncate File names to under maxFileName chars
    If (Len(name) > maxFileName) Then
        name = Left(name, maxFileName) & " ..."
    End If
    'Debug.Print "Saving: " & name
    
    Header.Copy
    Section.Copy
    
    'Saving Document
    Set D = Documents.Add
    D.Range.PasteAndFormat wdFormatOriginalFormatting
    With D.Content.Find
        .ClearFormatting
        .MatchWildcards = True
        .MatchCase = False
        .Text = "^13([1-9]).([1-9])*^13"
        .Replacement.ClearFormatting
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    D.SaveAs2 FileName:=Section.Parent.Path & Application.PathSeparator & name & ".htm", _
    FileFormat:=wdFormatFilteredHTML
    D.Close
End Sub

'Remove all existing hyperlinks in a document.
Private Static Sub RemoveAllHyperlinks()
    Dim oField As Field
    For Each oField In ActiveDocument.Fields
        If oField.Type = wdFieldHyperlink Then
            oField.Unlink
        End If
    Next
    Set oField = Nothing
End Sub

'Replace all unformatted URLs with a hyperlink to itself.
Private Static Sub URLtoHyperlink()
  Dim f1 As Boolean, f2 As Boolean, f3 As Boolean
  Dim f4 As Boolean, f5 As Boolean, f6 As Boolean
  Dim f7 As Boolean, f8 As Boolean, f9 As Boolean
  Dim f10 As Boolean
  With Options
    ' Save current AutoFormat settings
    f1 = .AutoFormatApplyHeadings
    f2 = .AutoFormatApplyLists
    f3 = .AutoFormatApplyBulletedLists
    f4 = .AutoFormatApplyOtherParas
    f5 = .AutoFormatReplaceQuotes
    f6 = .AutoFormatReplaceSymbols
    f7 = .AutoFormatReplaceOrdinals
    f8 = .AutoFormatReplaceFractions
    f9 = .AutoFormatReplacePlainTextEmphasis
    f10 = .AutoFormatReplaceHyperlinks
    ' Only convert URLs
    .AutoFormatApplyHeadings = False
    .AutoFormatApplyLists = False
    .AutoFormatApplyBulletedLists = False
    .AutoFormatApplyOtherParas = False
    .AutoFormatReplaceQuotes = False
    .AutoFormatReplaceSymbols = False
    .AutoFormatReplaceOrdinals = False
    .AutoFormatReplaceFractions = False
    .AutoFormatReplacePlainTextEmphasis = False
    .AutoFormatReplaceHyperlinks = True
    ' Perform AutoFormat
    ActiveDocument.Content.AutoFormat
    ' Restore original AutoFormat settings
    .AutoFormatApplyHeadings = f1
    .AutoFormatApplyLists = f2
    .AutoFormatApplyBulletedLists = f3
    .AutoFormatApplyOtherParas = f4
    .AutoFormatReplaceQuotes = f5
    .AutoFormatReplaceSymbols = f6
    .AutoFormatReplaceOrdinals = f7
    .AutoFormatReplaceFractions = f8
    .AutoFormatReplacePlainTextEmphasis = f9
    .AutoFormatReplaceHyperlinks = f10
  End With
End Sub

'Replace accended chars with their plaintext alphabet counterparts
Function StripAccent(aString As String)
    Dim A As String * 1
    Dim B As String * 1
    Dim i As Integer
    Const AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
    Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
    For i = 1 To Len(AccChars)
        A = Mid(AccChars, i, 1)
        B = Mid(RegChars, i, 1)
        aString = Replace(aString, A, B)
    Next
    StripAccent = aString
End Function

'Remove all shapes in the Active Document
Function DeleteShapes()
    Dim i As Long
    With ActiveDocument
    For i = .Shapes.Count To 1 Step -1
        With .Shapes(i)
            If .Type = msoAutoShape Then .Delete
        End With
    Next i
End With
End Function
