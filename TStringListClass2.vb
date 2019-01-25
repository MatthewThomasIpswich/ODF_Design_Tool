
Option Explicit On
Imports System
Imports System.IO
Imports System.Linq

'// Needes M_Useful.vb and TIniFile.vb to be present.

<Serializable()> Public Class TStringList
    '// Class for manipulating files and strings.
    Private _Text As String
    Private _StringList As New Generic.List(Of String)
    Private _ErrorStr As String = ""
    Private SplitChars(2) As Char
    '//
    Private Sub PrepareSplitChar()
        Dim LF_ch As Char = CType(vbLf, Char)
        Dim CRLF_Ch As Char = CType(vbCrLf, Char)
        ''Dim FF_Ch As Char = CType(vbFormFeed, Char)
        ''Dim BK_Ch As Char = CType(vbBack, Char)
        SplitChars(1) = LF_ch
        SplitChars(2) = CRLF_Ch
    End Sub
    '//-----------
    ReadOnly Property Count() As Integer
        Get
            Count = _StringList.Count
        End Get
    End Property

    '//-----------
    Property TextCRLF() As String
        Get
            TextCRLF = WriteOutCRLF()
        End Get
        Set(ByVal value As String)
            ReadInCRLF(value)
        End Set
    End Property

    Property TextLF() As String
        '// as used by rich edit
        Get
            TextLF = WriteOutLF()
        End Get
        Set(ByVal value As String)
            ReadInLF(value)
        End Set
    End Property

    '//-----------
    Property Line(ByVal index As Integer) As String
        Get
            Line = GetLine(index)
        End Get
        Set(ByVal value As String)
            SetLine(index, value)
        End Set
    End Property

    '//-----------
    ReadOnly Property ErrorStr() As String
        Get
            ErrorStr = _ErrorStr
        End Get
    End Property

    Public Sub Add(ByVal Line As String)
        _StringList.Add(Line)
    End Sub  '//-----------

    Public Sub Sort()
        _StringList.Sort()
    End Sub  '//-----------


    Public Sub SetLine(ByVal i As Integer, ByVal line As String)
        If ((i >= 0) And (i <= _StringList.Count)) Then
            _StringList(i) = line
        End If
    End Sub '//-----------

    Public Function GetLine(ByVal i As Integer) As String
        Dim Line As String
        If i <= _StringList.Count Then
            Line = _StringList(i)
        Else
            Line = ""
        End If
        Return Line
    End Function '//-----------

    Public Sub Insert(ByVal i As Integer, ByVal line As String)
        If ((i >= 0) And (i <= _StringList.Count)) Then
            _StringList.Insert(i, line)
        End If
    End Sub '//-----------

    Public Sub Clear()
        _StringList.Clear()
    End Sub '//-----------

    Public Sub ToUpper()
        Dim Line As String
        Dim i As Integer
        '//
        For i = 0 To _StringList.Count - 1
            Line = _StringList.Item(i)
            _StringList.Item(i) = Line.ToUpper
            Application.DoEvents()
        Next i
    End Sub '//-----------

    Public Sub Delete(ByVal i As Integer)
        _StringList.RemoveAt(i)
    End Sub '//-----------

    Public Function LineIndexOf(ByVal line As String) As Integer
        '// If line = ""
        Dim i As Integer
        Dim aLine As String
        Dim Pos As Integer = -1
        '//
        If line.Length > 0 Then
            For i = 0 To _StringList.Count - 1
                aLine = _StringList.Item(i)
                If aLine.Contains(line) Then
                    Pos = i
                    Exit For
                End If
                Application.DoEvents()
            Next
        End If
        Return Pos
    End Function '//-----------

    Public Function LineIndexOf(ByVal line As String, ByVal start As Integer) As Integer
        Dim i As Integer
        Dim aLine As String
        Dim Pos As Integer = -1
        '//
        If line.Length > 0 Then
            If (start < _StringList.Count) And (start > -1) Then
                For i = start To _StringList.Count - 1
                    aLine = _StringList.Item(i)
                    If aLine.Contains(line) Then
                        Pos = i
                        Exit For
                    End If
                    Application.DoEvents()
                Next
            End If
        End If
        Return Pos
    End Function '//-----------

    Public Function LineIndexOfComplete(ByVal line As String) As Integer
        '// If line = ""
        Dim i As Integer
        Dim aLine As String
        Dim Pos As Integer = -1
        '//
        If line.Length > 0 Then
            For i = 0 To _StringList.Count - 1
                aLine = _StringList.Item(i)
                If aLine = line Then
                    Pos = i
                    Exit For
                End If
            Next
        End If
        Return Pos
    End Function '//-----------

    Public Function TxtBetween(ByVal startStr As String, ByVal EndStr As String) As String
        '// include the text at start, but not end.
        Dim StartPos As Integer
        Dim EndPos As Integer
        Dim StrList As New TStringList
        Dim TextBwn As String
        Dim Ok As Boolean
        Dim i As Integer
        '//
        StartPos = _StringList.IndexOf(startStr)
        EndPos = _StringList.IndexOf(EndStr, StartPos + 1)
        '//
        Ok = False
        If StartPos > -1 Then
            If EndPos > -1 Then
                If StartPos < EndPos Then
                    Ok = True
                End If
            Else
                Ok = True
            End If
        End If
        '//
        If Ok Then
            For i = StartPos To (EndPos - 1)
                StrList.Add(_StringList.Item(i))
            Next
            TextBwn = StrList.TextCRLF
        Else
            TextBwn = ""
        End If
        Return TextBwn
    End Function

    '//=================== VISUAL CONTROLS ===================================

    Public Sub WriteToListBox(ByVal aListBox As ListBox)
        Dim i As Integer
        Dim Line As String
        aListBox.Items.Clear()
        For i = 0 To _StringList.Count - 1
            Line = _StringList.Item(i)
            aListBox.Items.Add(Line)
        Next i
    End Sub '//----------------------


    Public Sub WriteToComboBox(ByVal aComboBox As ComboBox)
        Dim i As Integer
        Dim Line As String
        aComboBox.Items.Clear()
        For i = 0 To _StringList.Count - 1
            Line = _StringList.Item(i)
            aComboBox.Items.Add(Line)
        Next i
        If _StringList.Count > 0 Then
            aComboBox.Text = _StringList.Item(0)
        End If
    End Sub '//----------------------

    Public Sub WriteToComboBox2(ByVal aComboBox As ComboBox)
        Dim i As Integer
        Dim Line As String
        aComboBox.Items.Clear()
        For i = 0 To _StringList.Count - 1
            Line = _StringList.Item(i)
            aComboBox.Items.Add(Line)
        Next i
    End Sub '//----------------------

    Public Sub ReadFromListBox(ByVal aListBox As ListBox)
        Dim i As Integer
        Dim Line As String
        _StringList.Clear()
        For i = 0 To aListBox.Items.Count - 1
            Line = aListBox.Items(i).ToString
            _StringList.Add(Line)
        Next i
    End Sub '//----------------------

    Public Sub WriteToRichEdit(ByVal RichText As RichTextBox)
        Dim i As Integer
        PrepareSplitChar()
        Dim Line As String
        Dim Test As String
        RichText.Clear()
        For i = 0 To _StringList.Count - 1
            Line = _StringList.Item(i)
            If Line IsNot Nothing Then
                Line.TrimStart()
                Line = Line & vbLf '// rich text uses vbLF not VbCR
                RichText.AppendText(Line)
                Test = RichText.ToString
            End If
            Application.DoEvents()
        Next i
    End Sub '//----------------------

    Public Sub WriteToRichEdit(ByVal RichText As RichTextBox, ByVal MaxRows As Integer)
        Dim i As Integer
        PrepareSplitChar()
        Dim Line As String
        Dim Test As String
        Dim NumRows As Integer
        '//
        If MaxRows < _StringList.Count Then
            NumRows = MaxRows
        Else
            NumRows = _StringList.Count
        End If
        RichText.Clear()
        For i = 0 To NumRows - 1
            Line = _StringList.Item(i)
            If Line IsNot Nothing Then
                Line.TrimStart()
                Line = Line & vbCr
                RichText.AppendText(Line)
                Test = RichText.ToString
            End If
            Application.DoEvents()
        Next i
    End Sub '//----------------------

    Public Sub ReadInRichEdit(ByVal RichText As RichTextBox)
        Dim i As Integer
        Dim Line As String
        _StringList.Clear()
        For i = 0 To RichText.Lines.Count - 1
            Line = RichText.Lines(i)
            _StringList.Add(Line)
            Application.DoEvents()
        Next i
    End Sub '//----------------------

    Public Sub WriteCSVtoGrid(ByVal Grid As DataGridView)
        '// Write out text to grid. First row always header, CSV format.
        '// Use TTableClass2.vb for more complex interaction
        Dim ColStrs() As String
        Dim Line As String
        Dim NumCol As Integer
        Dim Row As Integer
        '//
        If _StringList.Count < 0 Then
            Exit Sub
        End If
        '// Set number of col and rows
        Line = _StringList.Item(0)
        ColStrs = ParseCSVLine(Line)
        NumCol = ColStrs.GetUpperBound(0) + 1 '// 0,...,N-1. Upper bound = N-1
        Grid.ColumnCount = NumCol
        Grid.RowCount = _StringList.Count - 1
        '//
        WriteHeadersToGrid(ColStrs, Grid)
        '//
        For Row = 0 To Grid.RowCount - 1
            Line = _StringList.Item(Row + 1)
            '//
            WriteLineToGrid(Line, NumCol, Row, Grid)
        Next Row
        Application.DoEvents()
        '//
    End Sub

    Private Sub WriteHeadersToGrid(ByVal ColStrs() As String, ByVal Grid As DataGridView)
        Dim HeaderStr As String
        Dim Col As Integer
        '// write out headers
        For Col = 0 To Grid.ColumnCount - 1
            HeaderStr = ColStrs(Col)
            Grid.Columns.Item(Col).HeaderText = HeaderStr
        Next Col
    End Sub '//---------------------------------------------------

    Private Sub WriteLineToGrid(ByVal LineStr As String, _
                                ByVal NumCol As Integer, _
                                ByVal Row As Integer, _
                                ByVal Grid As DataGridView)
        Dim Col As Integer
        Dim ColStrs() As String
        Dim aCellStr As String
        Dim FoundColNum As Integer
        Dim ColNum As Integer
        '//
        ColStrs = ParseCSVLine(LineStr)
        FoundColNum = ColStrs.GetUpperBound(0) + 1
        ColNum = Math.Min(FoundColNum, NumCol)
        For Col = 0 To ColNum - 1
            aCellStr = ColStrs(Col)
            Grid.Rows(Row).Cells(Col).Value = aCellStr
        Next
    End Sub

    Public Function ParseCSVLine(ByVal LineStr As String) As String()
        Dim aColStrs() As String
        Dim delimiters() As Char = {CType(",", Char)}
        '//
        '// delimiate the columns
        aColStrs = LineStr.Split(delimiters)
        Return aColStrs
        '//
    End Function

    '//------------------- STRING ARRAYs -----------------------------------

    Public Function MakeStrListOf() As List(Of String)
        '// Convert the string list a String list Of.
        Dim NewStrListOf As New List(Of String)
        '//
        For Each line As String In _StringList
            NewStrListOf.Add(line)
        Next
        Return NewStrListOf
    End Function

    Public Sub ReadStrListOf(ByVal InStrList As List(Of String))
        '// read a String list Of into a string list.
        _StringList.Clear()
        For Each line As String In InStrList
            _StringList.Add(line)
        Next
    End Sub

    'Public Sub ReadStrListOfF(ByVal InStrArray As String())
    '    '// read a String list Of into a string list.
    '    _StringList.Clear()
    '    For Each line As String In InStrArray
    '        _StringList.Add(line)
    '    Next
    'End Sub
    '//================== FILEs ===================================================

    Public Function LoadFromFile(ByVal Path As String, ByVal ErrorStr As String) As Boolean
        '// read in atext file on the given path.
        Dim aLine As String
        Dim forReading As IO.StreamReader
        Dim ReadOk As Boolean = False
        '//
        ErrorStr = ""
        If System.IO.File.Exists(Path) Then
            ReadOk = True    '// ***  more checking needed here + exceptions covered.
        Else
            ReadOk = False
            ErrorStr = "Path '" + Path + "' does not exist."
        End If
        '//
        If ReadOk Then
            _StringList.Clear()
            forReading = New IO.StreamReader(Path)
            Try
                '// Read through the file
                aLine = forReading.ReadLine()
                '//
                While (Not aLine Is Nothing)
                    _StringList.Add(aLine)
                    Application.DoEvents()
                    aLine = forReading.ReadLine()
                End While
            Catch FileEx As Exception
                ReadOk = False
                MsgBox(FileEx.Message)
            Finally
                forReading.Close()
            End Try
        End If
        '//
        Return ReadOk
    End Function '//----------------------------------------


    Public Function SaveToFile(ByVal Path As String, ByVal ErrorStr As String) As Boolean
        '// write out a text file on the given path.
        Dim i As Integer
        Dim aLine As String
        Dim FS As FileStream
        Dim SW As StreamWriter
        Dim WriteOk As Boolean = True
        '//
        ErrorStr = ""
        FS = New FileStream(Path, FileMode.Create)
        Try
            SW = New StreamWriter(FS)
            Try
                For i = 0 To _StringList.Count - 1
                    aLine = _StringList(i)
                    SW.WriteLine(aLine)
                    Application.DoEvents()
                Next i
            Catch AMyException As Exception
                ErrorStr = "Problem writing to " & Path & ":" & AMyException.Message
            Finally
                SW.Close() '// could dispose of FS too.
            End Try
        Catch BMyException As Exception
            WriteOk = False
            ErrorStr = "Problem making file stream :" & BMyException.Message
        Finally
            If Not FS Is Nothing Then
                FS.Close()
            End If
        End Try
        '//
        Return WriteOk
    End Function '//----------------------------------------

    Public Function AppendToFile(ByVal Path As String, ByVal ErrorStr As String) As Boolean
        '// append text to a file on the given path.
        Dim i As Integer
        Dim aLine As String
        Dim FS As FileStream
        Dim SW As StreamWriter
        Dim WriteOk As Boolean = True
        '//
        ErrorStr = ""
        FS = New FileStream(Path, FileMode.Append)
        Try
            SW = New StreamWriter(FS)
            Try
                For i = 0 To _StringList.Count - 1
                    aLine = _StringList(i)
                    SW.WriteLine(aLine)
                    Application.DoEvents()
                Next i
            Catch AMyException As Exception
                MsgBox("Problem writing to " & Path & ":" & AMyException.Message)
            Finally
                SW.Close()
            End Try
        Catch BMyException As Exception
            WriteOk = False
            MsgBox("Problem making file stream :" & BMyException.Message)
        Finally
            If Not FS Is Nothing Then
                FS.Close()
            End If
        End Try
        '//
        Return WriteOk
    End Function '//----------------------------------------

    '//=============== SEARCH REPLACE =============================================

    Public Sub RemoveLines(ByVal FindTxt As String)
        '// remove all the lines with the text "FindTxt" in them.
        Dim Pos As Integer
        Pos = 0
        While _StringList.Count > 0
            Pos = LineIndexOf(FindTxt)
            If Pos > -1 Then
                Delete(Pos)
            Else
                Exit While
            End If
        End While
    End Sub

    Public Function GetLines(ByVal FindTxt As String) As TStringList
        '// Get all the lines with the text "FindTxt" in them.
        Dim Pos As Integer
        Dim Start As Integer
        Dim ResultList As New TStringList
        Dim FoundLine As String
        Pos = 0
        Start = 0
        While Start < _StringList.Count
            Pos = LineIndexOf(FindTxt, Start)
            If Pos > -1 Then
                FoundLine = GetLine(Pos)
                ResultList.Add(FoundLine)
                Start = Pos + 1
            Else
                Exit While
            End If
        End While
        Return ResultList
    End Function

    '// ======================== PRIVATE ==============================
    Private Function WriteOutCRLF() As String
        Dim AllText As String = ""
        Dim Line As String
        Dim i As Integer
        '//
        For i = 0 To _StringList.Count - 1
            Line = _StringList.Item(i)
            AllText = AllText & Line & vbCrLf
        Next
        Application.DoEvents()
        Return AllText
        ''If _StringList.Count > 0 Then
        ''    For i = 0 To _StringList.Count - 1
        ''        Line = _StringList.Item(i)
        ''        Line.Replace(vbCrLf, vbNullChar)
        ''        AllText = AllText & Line & vbCrLf
        ''    Next
        ''    Line = _StringList.Item(_StringList.Count - 1)
        ''    AllText = AllText & Line
        ''    Application.DoEvents()
        ''End If
    End Function

    Private Function WriteOutLF() As String
        '// as used by Richeditbox
        Dim AllText As String = ""
        Dim Line As String
        Dim i As Integer
        '//
        For i = 0 To _StringList.Count - 1
            Line = _StringList.Item(i)
            AllText = AllText & Line & vbLf
        Next
        Application.DoEvents()
        Return AllText
    End Function
    '//
    Private Sub ReadInCRLF(ByVal Text As String)
        Dim Lines() As String '// array
        Dim aLine As String
        Dim i As Integer
        '//
        Dim CRLF_Ch As Char = CType(vbCrLf, Char)
        Lines = Text.Split(CRLF_Ch)
        _StringList.Clear()
        '//
        For i = 0 To Lines.Count - 1
            aLine = Lines(i)
            aLine.TrimEnd()
            _StringList.Add(aLine)
        Next
        Application.DoEvents()
    End Sub

    Private Sub ReadInLF(ByVal Text As String)
        '// as used by Richeditbox
        Dim Lines() As String '// array
        Dim aLine As String
        Dim i As Integer
        '//
        Dim LF_Ch As Char = CType(vbLf, Char)
        Lines = Text.Split(LF_Ch)
        _StringList.Clear()
        '//
        For i = 0 To Lines.Count - 1
            aLine = Lines(i)
            aLine.TrimEnd()
            _StringList.Add(aLine)
        Next
        Application.DoEvents()
    End Sub

    Private Sub ReadInQuery(ByVal Query As IEnumerable(Of String))
        '// read a String array into a string list.
        _StringList.Clear()
        For Each line As String In Query
            _StringList.Add(line)
        Next
    End Sub

    '// ===================== SET OPERATIONS ====================================

    Public Sub Concat(ByVal ListA As TStringList, ByVal ListB As TStringList)
        '// String list becomes the concatentaion of the two str lists
        Dim StrArrayA As List(Of String) = ListA.MakeStrListOf
        Dim StrArrayB As List(Of String) = ListB.MakeStrListOf
        Dim ConcatQuery As System.Collections.Generic.IEnumerable(Of String)
        '// 
        ConcatQuery = StrArrayA.Concat(StrArrayB).OrderBy(Function(p) p)
        ReadInQuery(ConcatQuery)
    End Sub


    Public Sub Union(ByVal ListA As TStringList, ByVal ListB As TStringList)
        '// String list becomes the union of the two str lists
        Dim StrArrayA As List(Of String) = ListA.MakeStrListOf
        Dim StrArrayB As List(Of String) = ListB.MakeStrListOf
        Dim ConcatQuery As System.Collections.Generic.IEnumerable(Of String)
        '//
        ConcatQuery = StrArrayA.Union(StrArrayB).OrderBy(Function(p) p)
        ReadInQuery(ConcatQuery)
    End Sub

    Public Sub Intersect(ByVal ListA As TStringList, ByVal ListB As TStringList)
        '// String list becomes the union of the two str lists
        Dim StrArrayA As List(Of String) = ListA.MakeStrListOf
        Dim StrArrayB As List(Of String) = ListB.MakeStrListOf
        Dim ConcatQuery As System.Collections.Generic.IEnumerable(Of String)
        '//
        ConcatQuery = StrArrayA.Intersect(StrArrayB).OrderBy(Function(p) p)
        ReadInQuery(ConcatQuery)
    End Sub

    Public Sub AIntersectNotB(ByVal ListA As TStringList, ByVal ListB As TStringList)
        '// String list becomes the intersection of the two str lists
        Dim StrArrayA As List(Of String) = ListA.MakeStrListOf
        Dim StrArrayB As List(Of String) = ListB.MakeStrListOf
        Dim ConcatQuery As System.Collections.Generic.IEnumerable(Of String)
        '//
        ConcatQuery = StrArrayA.Except(StrArrayB).OrderBy(Function(p) p)
        ReadInQuery(ConcatQuery)
    End Sub

    Public Sub Unique(ByVal ListA As TStringList)
        '// Takes list A and makes self = unique list.
        '// String list removes all duplicates
        Dim StrArrayA As List(Of String) = ListA.MakeStrListOf
        Dim ConcatQuery As System.Collections.Generic.IEnumerable(Of String)
        '//
        ConcatQuery = StrArrayA.Distinct.OrderBy(Function(p) p)
        ReadInQuery(ConcatQuery)
    End Sub

    Public Function SumInt() As Integer
        '// all rows that are integers are summed.
        Dim Line As String
        Dim i As Integer
        Dim Value As Integer
        Dim Sum As Integer
        '//
        Value = 0
        Sum = 0
        For i = 0 To _StringList.Count - 1
            Line = _StringList.Item(i)
            If StrToIntDef(Line, 0, Value) Then
                Sum = Sum + Value
            Else
                _ErrorStr = _ErrorStr + vbCr + "Line " & CType(i, String) _
                & " is not an integer :" & Line
            End If
            Application.DoEvents()
        Next
        Return Sum
    End Function

    Public Function SumReal() As Double
        '// all rows that are integers are summed.
        Dim Line As String
        Dim i As Integer
        Dim Value As Double
        Dim Sum As Double
        '//
        Value = 0
        Sum = 0
        For i = 0 To _StringList.Count - 1
            Line = _StringList.Item(i)
            If StrToRealDef(Line, 0, Value) Then
                Sum = Sum + Value
            Else
                _ErrorStr = _ErrorStr + vbCr + "Line " & CType(i, String) _
                & " is not a number :" & Line
            End If
            Application.DoEvents()
        Next
        Return Sum
    End Function

    Public Function UniqueElements(ByVal ColumnNum As Integer, _
                                   ByVal O_UnElements As TStringList) As Boolean
        '// Make a list of the unique items in the column of self's list.
        '// return false if not suficient columns in all the rows.
        Dim AllElements As New TStringList
        Dim aLine As String
        Dim i As Integer
        Dim aColStrs() As String
        Dim delimiters() As Char = {CType(",", Char)}
        Dim Element As String
        Dim IsError As Boolean
        '//
        AllElements.Clear()
        O_UnElements.Clear()
        IsError = False
        For i = 1 To Count - 1
            aLine = Line(i)
            '//
            aColStrs = aLine.Split(delimiters)
            If ColumnNum < aColStrs.Length Then
                Element = aColStrs(ColumnNum)
                AllElements.Add(Element)
            Else
                Exit For
                IsError = True
            End If
        Next
        O_UnElements.Unique(AllElements)
        Return Not IsError
    End Function


End Class '//===========================================




