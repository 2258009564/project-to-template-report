#!/usr/bin/env python3
"""Emit an ASCII-safe Word VBS scaffold driven by a UTF-8 operations TSV."""

from __future__ import annotations

import argparse
from pathlib import Path


VBS = r'''Option Explicit
On Error Resume Next

' Args: template.docx output.docx operations.tsv log.txt
Const wdFindStop = 0
Const wdReplaceAll = 2
Const wdStatisticPages = 2
Const msoTrue = -1

Dim args, templatePath, outputPath, operationsPath, logPath
Set args = WScript.Arguments
If args.Count < 4 Then
    WScript.Echo "usage: cscript //Nologo build_report.vbs template.docx output.docx operations.tsv log.txt"
    WScript.Quit 64
End If
templatePath = args(0)
outputPath = args(1)
operationsPath = args(2)
logPath = args(3)

Dim logStream
Set logStream = CreateObject("ADODB.Stream")
logStream.Type = 2
logStream.Charset = "utf-8"
logStream.Open

Sub LogLine(s)
    logStream.WriteText CStr(s) & vbCrLf
End Sub

Function DecodeValue(s)
    s = Replace(s, "~t~", vbTab)
    s = Replace(s, "~n~", vbCr)
    DecodeValue = s
End Function

Function ReadUtf8(path)
    Dim stm
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    ReadUtf8 = stm.ReadText
    stm.Close
End Function

Function CountExact(doc, needle)
    Dim rng, n
    n = 0
    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = needle
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
    End With
    Do While rng.Find.Execute
        n = n + 1
        rng.Start = rng.End
        rng.End = doc.Content.End
    Loop
    CountExact = n
End Function

Function ReplaceExact(doc, needle, replacement)
    Dim n, rng
    n = CountExact(doc, needle)
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = needle
        .Replacement.Text = replacement
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .Execute , , , , , , , , , , wdReplaceAll
    End With
    ReplaceExact = n
End Function

Function SetBookmarkText(doc, bookmarkName, newText)
    Dim rng
    SetBookmarkText = False
    If Not doc.Bookmarks.Exists(bookmarkName) Then Exit Function
    Set rng = doc.Bookmarks(bookmarkName).Range
    rng.Text = newText
    doc.Bookmarks.Add bookmarkName, rng
    SetBookmarkText = True
End Function

Function SetParagraphText(doc, paragraphIndex, newText, styleName)
    Dim rng
    SetParagraphText = False
    If paragraphIndex < 1 Or paragraphIndex > doc.Paragraphs.Count Then Exit Function
    Set rng = doc.Paragraphs(paragraphIndex).Range.Duplicate
    If rng.End > rng.Start Then rng.End = rng.End - 1
    rng.Text = newText
    If Len(styleName) > 0 Then rng.Style = doc.Styles(styleName)
    SetParagraphText = (Err.Number = 0)
    Err.Clear
End Function

Function InsertImageAtBookmark(doc, bookmarkName, imagePath, widthCm)
    Dim rng, pic
    InsertImageAtBookmark = False
    If Not doc.Bookmarks.Exists(bookmarkName) Then Exit Function
    Set rng = doc.Bookmarks(bookmarkName).Range
    rng.Text = ""
    Set pic = doc.InlineShapes.AddPicture(imagePath, False, True, rng)
    If Not pic Is Nothing Then
        pic.LockAspectRatio = msoTrue
        If widthCm > 0 Then pic.Width = widthCm * 28.3464567
        doc.Bookmarks.Add bookmarkName, pic.Range
        InsertImageAtBookmark = True
    End If
End Function

Sub ApplyOperations(doc, path, ByRef failureCount)
    Dim allText, lines, line, cols, kind, target, val, opt, n, ok
    allText = Replace(ReadUtf8(path), vbCrLf, vbLf)
    allText = Replace(allText, vbCr, vbLf)
    lines = Split(allText, vbLf)
    For Each line In lines
        If Len(Trim(line)) > 0 Then
            If Left(Trim(line), 1) <> "#" Then
                cols = Split(line, vbTab)
                If UBound(cols) < 2 Then
                    failureCount = failureCount + 1
                    LogLine "FAIL malformed operation: " & line
                Else
                    kind = LCase(Trim(cols(0)))
                    target = DecodeValue(cols(1))
                    val = DecodeValue(cols(2))
                    opt = ""
                    If UBound(cols) >= 3 Then opt = DecodeValue(cols(3))
                    If kind = "replace_exact" Then
                        n = ReplaceExact(doc, target, val)
                        LogLine "replace_exact matches=" & n & " target=" & target
                        If n <> 1 Then failureCount = failureCount + 1
                    ElseIf kind = "bookmark_text" Then
                        ok = SetBookmarkText(doc, target, val)
                        LogLine "bookmark_text ok=" & ok & " target=" & target
                        If Not ok Then failureCount = failureCount + 1
                    ElseIf kind = "paragraph_text" Then
                        ok = SetParagraphText(doc, CLng(target), val, opt)
                        LogLine "paragraph_text ok=" & ok & " index=" & target
                        If Not ok Then failureCount = failureCount + 1
                    ElseIf kind = "image_bookmark" Then
                        ok = InsertImageAtBookmark(doc, target, val, CDbl(opt))
                        LogLine "image_bookmark ok=" & ok & " target=" & target
                        If Not ok Then failureCount = failureCount + 1
                    Else
                        failureCount = failureCount + 1
                        LogLine "FAIL unknown operation=" & kind
                    End If
                End If
            End If
        End If
    Next
End Sub

Sub ApplyTaskSpecificEdits(doc, ByRef failureCount)
    ' Patch this procedure for template-specific tables, repeated sections,
    ' caption fields, cross-references, and exact figure placement.
End Sub

Dim fso, word, doc, toc, fld, failures, pages
failures = 0
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(templatePath) Then
    LogLine "FAIL template not found: " & templatePath
    failures = failures + 1
Else
    fso.CopyFile templatePath, outputPath, True
    If Err.Number <> 0 Then
        LogLine "FAIL copy err=" & Err.Number & " " & Err.Description
        failures = failures + 1
        Err.Clear
    End If
End If

If failures = 0 Then
    Set word = CreateObject("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.AutomationSecurity = 3
    Set doc = word.Documents.Open(outputPath, False, False, False)
    If doc Is Nothing Or Err.Number <> 0 Then
        LogLine "FAIL open err=" & Err.Number & " " & Err.Description
        failures = failures + 1
        Err.Clear
    End If
End If

If failures = 0 Then
    ApplyOperations doc, operationsPath, failures
    ApplyTaskSpecificEdits doc, failures
    If failures = 0 Then
    For Each toc In doc.TablesOfContents
        toc.Update
    Next
    For Each fld In doc.Fields
        fld.Update
    Next
    doc.Repaginate
    pages = doc.ComputeStatistics(wdStatisticPages)
    LogLine "pages=" & pages
    LogLine "paragraphs=" & doc.Paragraphs.Count
    LogLine "tables=" & doc.Tables.Count
    LogLine "inline_shapes=" & doc.InlineShapes.Count
    LogLine "toc=" & doc.TablesOfContents.Count
    If Err.Number <> 0 Then
        LogLine "FAIL finalization err=" & Err.Number & " " & Err.Description
        failures = failures + 1
        Err.Clear
    Else
        doc.Save
        If Err.Number <> 0 Then
            LogLine "FAIL save err=" & Err.Number & " " & Err.Description
            failures = failures + 1
            Err.Clear
        Else
            LogLine "saved=" & outputPath
        End If
    End If
    End If
End If

If Not doc Is Nothing Then doc.Close False
If Not word Is Nothing Then word.Quit
logStream.SaveToFile logPath, 2
logStream.Close

If failures > 0 Then
    WScript.Echo "FAILED operations=" & failures
    WScript.Quit 2
End If
WScript.Echo "DONE"
WScript.Quit 0
'''


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", required=True, type=Path)
    args = parser.parse_args()
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(VBS, encoding="ascii", newline="\r\n")
    print(args.out.resolve())


if __name__ == "__main__":
    main()
