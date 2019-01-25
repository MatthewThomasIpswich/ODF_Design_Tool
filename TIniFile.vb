Option Explicit On
Imports System.IO

Module Ini_Useful
    '// Useful functions and constants for Ini file manipulation
    Public EqualChar As Char = CType("=", Char)
    '//----------------
    Function isRow(ByVal line As String) As Boolean
        '// check an equal "=" sign in middle or end of string
        Dim Pos As Integer
        '//
        Pos = line.IndexOf(EqualChar)
        If Pos > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Function isStartOfBlock(ByVal line As String) As Boolean
        '// check an  "<<<" sign in middle or end of string
        Dim Pos As Integer
        '//
        Pos = line.IndexOf("<<<")
        If Pos > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Function isEndOfBlock(ByVal line As String) As Boolean
        '//  ">>>" marks end of block
        '//
        If line = ">>>" Then
            Return True
        Else
            Return False
        End If
    End Function

    '//----------------
    Function IsComment(ByVal Text As String) As Boolean
        Text = Trim(Text)
        Return Text.StartsWith(";")
    End Function

End Module

Public Class TiniRow
    '// A row in a section "key_Str=Stored_value_Str"
    Private _Key As String = "" '// Key e.g. "font size","Path" etc.
    Private _Stored As String = "" '// stored value e.g. "12","C:\new"
    Private _Commnet As Boolean = False
    '//----------------
    Property Key() As String
        Get
            Key = _Key
        End Get
        Set(ByVal value As String)
            _Key = value
            _Commnet = False
        End Set
    End Property
    '//----------------
    Property Value() As String
        Get
            Value = _Stored
        End Get
        Set(ByVal value As String)
            _Stored = value
        End Set
    End Property
    '//----------------
    Property Text() As String
        Get
            If Not _Commnet Then
                Text = _Key & "=" & _Stored
            Else
                Text = _Stored
            End If
        End Get
        '//
        Set(ByVal value As String)
            Dim TempStr As String
            Dim aPairStrs() As String
            TempStr = value
            If Not IsComment(TempStr) Then
                _Commnet = False
                If isRow(TempStr) Then
                    '// delimiate the columns
                    aPairStrs = TempStr.Split(EqualChar)
                    _Key = Trim(aPairStrs(0))
                    _Stored = Trim(aPairStrs(1))
                End If
            Else
                _Commnet = True
                _Key = ""
                _Stored = TempStr
            End If
        End Set
    End Property

End Class '//===================================

Public Class TiniBlock
    '// A block in a section 
    '//key_Str<<<
    '//<Stored_value_Str1"
    '//<Stored_value_Str2"
    '//>>>
    Private _Key As String = "" '// Key e.g. "SQL" etc.
    Private _Stored As New TStringList  '// stored value e.g. rows of text for the SQL, < removed.
    '//----------------
    Property Key() As String
        Get
            Key = _Key
        End Get
        Set(ByVal input As String)
            _Key = input
        End Set
    End Property
    '//----------------
    Property ValueCRLF() As String
        Get
            ValueCRLF = _Stored.TextCRLF
        End Get
        Set(ByVal value As String)
            _Stored.TextCRLF = value
        End Set
    End Property

    Property ValueLF() As String
        Get
            ValueLF = _Stored.TextLF
        End Get
        Set(ByVal value As String)
            _Stored.TextLF = value
        End Set
    End Property

    '//----------------
    ReadOnly Property TextCRLF() As String
        '// is is used to write to a MS formt file(CRLF), not rich text.
        Get
            Dim TextInFile As String
            Dim i As Integer
            Dim OutputLine As String
            '//
            TextInFile = ""
            For i = 0 To _Stored.Count - 1
                OutputLine = _Stored.Line(i)
                OutputLine = OutputLine.TrimStart()
                '//
                TextInFile = TextInFile & "<" & OutputLine & vbCrLf
            Next
            TextCRLF = _Key & "<<<" & vbCrLf & TextInFile & ">>>" & vbCrLf
        End Get
        '//
    End Property

    Public Sub Add(ByVal aline As String)
        '// add to the stored text and strip of leading "<"
        Dim Pos As Integer
        '// only add if start with "<"
        Pos = aline.IndexOf("<")
        If Pos > -1 Then
            If aline.Length >= 2 Then '// "<text"
                aline = aline.Substring(Pos + 1)
                _Stored.Add(aline)
            Else
                _Stored.Add("") '// an empty line "<"
            End If
        End If
    End Sub

End Class '//===================================

Public Class TIniSection
    Private _Title As String '// text inside "[section]"
    Private _Rows As New List(Of TiniRow) '// rows in section.
    Private _Blocks As New List(Of TiniBlock) '// blocks in section.
    Dim EqualChar As Char = CType("=", Char)

    '//----------------
    Property Title() As String
        Get
            Title = _Title
        End Get
        Set(ByVal value As String)
            '// Need to check title is ok.
            _Title = value
        End Set
    End Property

    '//----------------
    Property Text() As String
        Get
            Text = GenerateText()
        End Get
        Set(ByVal value As String)
            Parsetext(value)
        End Set
    End Property

    '//----------------
    Public Function RowRead(ByVal Key As String, ByRef Value As String) As Boolean
        '// read value from data structure.
        Dim index As Integer
        Dim aRow As TiniRow
        ''Dim aBlock As TiniBlock
        '//
        index = RowFind(Key)
        If index > -1 Then
            '// Key specifies a row.
            aRow = _Rows.Item(index)
            Value = aRow.Value
            Return True
        End If
        '//
        Return False
    End Function

    '//----------------
    Public Sub RowWrite(ByVal Key As String, ByVal Value As String)
        '// write value to data structure.
        Dim index As Integer
        Dim aRow As TiniRow
        ''Dim aBlock As TiniBlock
        '//
        index = RowFind(Key)
        If index > -1 Then
            aRow = _Rows.Item(index)
            aRow.Value = Value
        Else
            aRow = New TiniRow
            aRow.Key = Key
            aRow.Value = Value
            _Rows.Add(aRow)
        End If
        '//
    End Sub

    Public Function BlockReadCRLF(ByVal Key As String, ByRef Value As String) As Boolean
        '// read value from data structure.
        Dim index As Integer
        Dim aBlock As TiniBlock
        '//
        index = BlockFind(Key)
        If index > -1 Then
            '// Key specifies a row.
            aBlock = _Blocks.Item(index)
            Value = aBlock.ValueCRLF
            Return True
        End If
        '//
        Return False
    End Function

    Public Sub WriteBlockCRLF(ByVal Key As String, ByVal Value As String)
        '// write value to data structure.
        Dim index As Integer
        Dim aBlock As TiniBlock
        '//
        index = BlockFind(Key)
        If index > -1 Then
            aBlock = _Blocks.Item(index)
            aBlock.ValueCRLF = Value
        Else
            aBlock = New TiniBlock
            aBlock.Key = Key
            aBlock.ValueCRLF = Value
            _Blocks.Add(aBlock)
        End If
    End Sub

    '//----------------
    Private Function RowFind(ByVal Key As String) As Integer
        '// Find first row with key. -1 = not found.
        Dim i As Integer
        Dim aRow As TiniRow
        Dim aKey As String
        Dim Pos As Integer = -1
        '/
        For i = 0 To _Rows.Count - 1
            aRow = _Rows.Item(i)
            aKey = aRow.Key
            If Key = aKey Then
                Pos = i
                Exit For
            End If
        Next
        Return Pos
    End Function

    Private Function BlockFind(ByVal Key As String) As Integer
        '// Find first row with key. -1 = not found.
        Dim i As Integer
        Dim aBlock As TiniBlock
        Dim aKey As String
        Dim Pos As Integer = -1
        '/
        For i = 0 To _Blocks.Count - 1
            aBlock = _Blocks.Item(i)
            aKey = aBlock.Key
            If Key = aKey Then
                Pos = i
                Exit For
            End If
        Next
        Return Pos
    End Function
    '//----------------
    Private Sub Parsetext(ByVal Text As String)
        '// Take text in file and turn into a section object
        Dim aStrList As New TStringList
        Dim aLine As String
        Dim NumLines As Integer
        '//
        aStrList.TextCRLF = Text
        NumLines = aStrList.Count
        If NumLines > 0 Then
            '// Get title
            '// replace with a regular expression
            Dim Start As Integer
            Dim Finish As Integer
            aLine = aStrList.Line(0)
            If aLine.Length > 3 Then
                Start = aLine.IndexOf("[") + 1
                Finish = aLine.IndexOf("]")
                _Title = aLine.Substring(Start, Finish - Start)
                '//
                '// make rows and blocks (and comments) of section
                ProcessBodyofSection(aStrList)
            Else
                aLine = "Short line:" & aLine & ":"
            End If
        End If
        '//
    End Sub

    Private Sub ProcessBodyofSection(ByVal StrList As TStringList)
        '// break the text up into comments, rows and section and store in object.
        Dim i As Integer
        Dim aLine As String
        Dim aRow As TiniRow
        Dim aBlock As TiniBlock = Nothing
        Dim isReadingBlock As Boolean
        Dim Pos_BlockChars As Integer
        '//
        _Rows.Clear()
        _Blocks.Clear()
        isReadingBlock = False
        For i = 1 To StrList.Count - 1
            aLine = StrList.Line(i)
            aLine = aLine.Trim
            '//
            If aLine = "" Then
                Continue For
            End If
            '//
            If isStartOfBlock(aLine) Then
                isReadingBlock = True
                aBlock = New TiniBlock
                '// Get block key
                Pos_BlockChars = aLine.IndexOf("<<<")
                aBlock.Key = Trim(aLine.Substring(0, Pos_BlockChars))
                Continue For
            End If
            '//
            If isEndOfBlock(aLine) Then
                isReadingBlock = False
                _Blocks.Add(aBlock)
                Continue For
            End If
            '//
            If isReadingBlock Then
                '// read in text in block
                aBlock.Add(aLine) '// strip of leading <
                Continue For
            End If
            '//
            If IsComment(aLine) Then
                Continue For
            End If
            '//
            If isRow(aLine) Then
                aRow = New TiniRow
                aRow.Text = aLine
                _Rows.Add(aRow)
                Continue For
            End If
            '//
            If IsComment(aLine) Then
                Continue For
            End If
            '//
            '// something else?
        Next
    End Sub
    '//----------------
    Private Function GenerateText() As String
        '// Make text of the section.
        Dim Text As String
        Dim i As Integer
        Dim aRow As TiniRow
        Dim aBlock As TiniBlock
        '//
        Text = "[" & Title & "]" & vbCrLf
        '// rows
        For i = 0 To _Rows.Count - 1
            aRow = _Rows.Item(i)
            Text = Text & aRow.Text & vbCrLf
        Next
        '// blocks
        For i = 0 To _Blocks.Count - 1
            aBlock = _Blocks.Item(i)
            Text = Text & aBlock.TextCRLF & vbCrLf
        Next
        Text = Text & vbCr '// blank row at the end.
        Return Text
    End Function

End Class '//=====================================================


Public Class TIniFile
    Private _FilePath As String
    'Private _ReadIn As Boolean = False
    Private _SectionList As New List(Of TIniSection)
    Private _Errored As Boolean = True


    Property FilePath() As String
        Get
            FilePath = _FilePath
        End Get
        Set(ByVal value As String)
            _FilePath = value
        End Set
    End Property

    Property Text() As String
        Get
            Text = Serialize()
        End Get
        Set(ByVal value As String)
            _Errored = Build(value)
        End Set
    End Property

    Public Function LoadFile(ByVal Path As String) As Boolean
        Dim StrList As New TStringList
        Dim TextStr As String
        Dim ErrorStr As String = ""
        '//
        _Errored = True
        If Not StrList.LoadFromFile(Path, ErrorStr) Then
            MsgBox("No ini file found: " & ErrorStr)
        Else
            _FilePath = Path
            '//
            TextStr = StrList.TextCRLF
            If Build(TextStr) Then
                _Errored = False
            End If
        End If
        Return _Errored
    End Function

    Public Function SaveToFile(ByVal Path As String) As Boolean
        Dim StrList As New TStringList
        Dim Saved As Boolean
        Dim ErrorStr As String = ""
        Dim FileStatus As New FileInfo(Path)
        '//
        If Not FileStatus.Exists Then
            MsgBox("The init file does not yet exist, it will now be created for you")
        End If
        '//
        StrList.TextCRLF = Serialize()
        If Not StrList.SaveToFile(Path, ErrorStr) Then
            MsgBox("Failed to save ini file: " & ErrorStr)
            Saved = False
        Else
            Saved = True
        End If
        Return Saved
    End Function

    Public Function ReadValue(ByVal Title As String, _
                               ByVal Key As String, _
                               ByVal DefaultValue As String, _
                               ByRef Value As String) As Boolean
        '// read from data structure
        Dim Found As Boolean
        Dim index As Integer
        Dim aSection As TIniSection
        Dim aValue As String = ""
        '//
        Found = False
        Value = DefaultValue
        index = FindSection(Title)
        If index > -1 Then
            aSection = _SectionList.Item(index)
            If aSection.RowRead(Key, aValue) Then
                Value = aValue
                Found = True
            End If
        End If
        '//
        Return Found
    End Function

    Public Sub WriteValue(ByVal Title As String, _
                               ByVal Key As String, _
                               ByVal Value As String)
        '// write to data structure
        Dim index As Integer
        Dim aSection As TIniSection
        '// Fine sectio or make if not there
        index = FindSection(Title)
        If index > -1 Then
            aSection = _SectionList.Item(index)
        Else
            aSection = New TIniSection
            aSection.Title = Title
            _SectionList.Add(aSection)
        End If
        '// write to key or make if not there.
        aSection.RowWrite(Key, Value)
        '//
    End Sub

    Public Function ReadValueBlockCRLF(ByVal Title As String, _
                              ByVal Key As String, _
                              ByVal DefaultValue As String, _
                              ByRef Value As String) As Boolean
        '// read from data structure
        Dim Found As Boolean
        Dim index As Integer
        Dim aSection As TIniSection
        Dim aValue As String = ""
        '//
        Found = False
        Value = DefaultValue
        index = FindSection(Title)
        If index > -1 Then
            aSection = _SectionList.Item(index)
            If aSection.BlockReadCRLF(Key, aValue) Then
                Value = aValue
                Found = True
            End If
        End If
        '//
        Return Found
    End Function

    Public Sub WriteValueBlockCRLF(ByVal Title As String, _
                               ByVal Key As String, _
                               ByVal Value As String)
        '// write to data structure
        Dim index As Integer
        Dim aSection As TIniSection
        '// Fine sectio or make if not there
        index = FindSection(Title)
        If index > -1 Then
            aSection = _SectionList.Item(index)
        Else
            aSection = New TIniSection
            aSection.Title = Title
            _SectionList.Add(aSection)
        End If
        '// write to key or make if not there.
        aSection.WriteBlockCRLF(Key, Value)
        '//
    End Sub

    Private Function Serialize() As String
        Dim i As Integer
        Dim aSection As TIniSection
        Dim Txt As String = ""
        For i = 0 To _SectionList.Count - 1
            aSection = _SectionList.Item(i)
            Txt = Txt & aSection.Text
        Next
        Return Txt
    End Function

    Private Function Build(ByVal Text As String) As Boolean
        '// Read all the text from the file and break into sections.
        Dim start As Integer = 0
        Dim finish As Integer = 0
        Dim sectiontxt As String = ""
        Dim aSection As TIniSection
        Dim isMoreSections As Boolean
        '// Make data structure from text.
        _SectionList.Clear()
        isMoreSections = True
        While isMoreSections
            isMoreSections = ExtractSectionText(start, Text, finish, sectiontxt)
            aSection = New TIniSection
            aSection.Text = sectiontxt
            _SectionList.Add(aSection)
            start = finish
        End While
        Return True
    End Function


    Private Function ExtractSectionText(ByVal StartRow As Integer, _
                                   ByVal FileText As String, _
                                   ByRef FinishRow As Integer, _
                                   ByRef SectionText As String) As Boolean
        '// extract text between the rows each with "[" in them.
        '// [section1 title]
        '//  key = value ...
        '// [section2 title]   .. so get row 1 & 2, not 3
        Dim StrList As New TStringList
        Dim StrList2 As New TStringList
        Dim i As Integer
        Dim aRow As String
        Dim BeginSectionLineNum As Integer
        Dim EndSectionLineNum As Integer
        Dim isMoreSections As Boolean = False
        '//
        StrList.TextCRLF = FileText
        '//
        '// Find start of section (where [] is)
        For i = StartRow To StrList.Count - 1
            aRow = StrList.Line(i)
            '//
            If isLineATile(aRow) Then
                BeginSectionLineNum = i
                Exit For
            End If
        Next
        '//
        If BeginSectionLineNum + 1 >= StrList.Count - 1 Then
            '// there is no room for section text
            FinishRow = -1
            SectionText = ""
            isMoreSections = False
            Return isMoreSections
        End If
        '//

        For i = (BeginSectionLineNum + 1) To StrList.Count - 1
            aRow = StrList.Line(i)
            '//
            If isLineATile(aRow) Then
                EndSectionLineNum = i - 1
                isMoreSections = True
                Exit For
            End If
        Next
        If Not isMoreSections Then
            EndSectionLineNum = StrList.Count - 1
        End If
        '//
        If BeginSectionLineNum >= EndSectionLineNum Then
            aRow = "error"
        End If
        '// Copy approprate lines from sting
        StrList2.Clear()
        For i = BeginSectionLineNum To EndSectionLineNum
            aRow = StrList.Line(i)
            aRow = aRow.Trim
            aRow = aRow.TrimEnd() '// if blank will trim, blanks & line ends by default.
            '//
            StrList2.Add(aRow)
        Next
        '//
        FinishRow = EndSectionLineNum + 1
        SectionText = StrList2.TextCRLF
        Return isMoreSections
    End Function

    Private Function isLineATile(ByVal line As String) As Boolean
        Dim index As Integer
        Dim isLineasTitle As Boolean
        '//
        isLineasTitle = False
        index = line.IndexOf("[")
        If index <= 1 Then '// must be first char in line.
            If line.IndexOf("]") > 1 Then
                isLineasTitle = True
            End If
        End If
        Return isLineasTitle
    End Function

    Private Function FindSection(ByVal Title As String) As Integer
        '// find the section with the given name
        Dim Pos As Integer = -1
        Dim i As Integer
        Dim aSection As TIniSection
        Dim aTitle As String
        '//
        For i = 0 To _SectionList.Count - 1
            aSection = _SectionList.Item(i)
            aTitle = aSection.Title
            If aTitle = Title Then
                Pos = i
                Exit For
            End If
        Next i
        '//
        Return Pos
    End Function

End Class '//=======================================================

Module Ini_Applications
    '//======================== COMBO =============================================

    Public Sub SaveComboToIni(ByVal aComboBox As System.Windows.Forms.ComboBox, _
                               ByVal ListTitle As String, _
                               ByVal aIniFile As TIniFile)
        '// Save list of stings and current sting to an inti file
        Dim Text As String
        Dim NumRows As Integer
        Dim i As Integer
        Dim RowKey As String
        Dim RowStr As String
        '//
        Text = aComboBox.Text
        aIniFile.WriteValue(ListTitle, "Text", Text)
        If aComboBox.Items.IndexOf(aComboBox.Text) = -1 Then
            '// put the latest value into the list too.
            aComboBox.Items.Add(aComboBox.Text)
        End If
        '//
        NumRows = aComboBox.Items.Count
        aIniFile.WriteValue(ListTitle, "NumRows", CType(NumRows, String))
        For i = 0 To NumRows - 1
            RowKey = "Item" & CType(i, String)
            RowStr = aComboBox.Items(i).ToString
            aIniFile.WriteValue(ListTitle, RowKey, RowStr)
        Next
    End Sub



    Public Sub ReadIniToCombo(ByVal aComboBox As System.Windows.Forms.ComboBox, _
                               ByVal ListTitle As String, _
                               ByVal aIniFile As TIniFile)
        '// Read from ini file the list of stings and the current value.
        Dim Text As String = ""
        Dim NumRowsStr As String = ""
        Dim NumRows As Integer
        Dim i As Integer
        Dim RowKey As String
        Dim RowStr As String = ""
        '//
        aIniFile.ReadValue(ListTitle, "Text", "", Text)
        aComboBox.Text = Text
        '//
        aIniFile.ReadValue(ListTitle, "NumRows", "0", NumRowsStr)
        If StrToIntDef(NumRowsStr, 0, NumRows) Then
            aComboBox.Items.Clear()
            For i = 0 To NumRows - 1
                RowKey = "Item" & CType(i, String)
                aIniFile.ReadValue(ListTitle, RowKey, "", RowStr)
                aComboBox.Items.Add(RowStr)
            Next
        End If
    End Sub

    Public Sub AddToCombo(ByVal aComboBox As System.Windows.Forms.ComboBox)
        Dim Current As String
        Dim Pos As Integer
        '//
        Current = aComboBox.Text
        Pos = aComboBox.Items.IndexOf(Current)
        If Pos = -1 Then
            aComboBox.Items.Add(Current)
        End If
    End Sub

    Public Sub DeleteFromCombo(ByVal aComboBox As System.Windows.Forms.ComboBox)
        Dim Current As String
        Dim Pos As Integer
        '//
        Current = aComboBox.Text
        Pos = aComboBox.Items.IndexOf(Current)
        If Pos > -1 Then
            aComboBox.Items.RemoveAt(Pos)
        End If
    End Sub

    '//====================== RichTextBox ===============================


    Public Sub SaveRichEditToIni(ByVal aRichTextBox As System.Windows.Forms.RichTextBox, _
                              ByVal SectionName As String, ByVal BlockName As String, _
                              ByVal aIniFile As TIniFile)
        '// Save  rich text to an ini file
        Dim rtfText As String
        '//
        rtfText = aRichTextBox.Rtf
        aIniFile.WriteValueBlockCRLF(SectionName, BlockName, rtfText)
        '//
    End Sub



    Public Sub ReadIniToRichEdit(ByVal aRichTextBox As System.Windows.Forms.RichTextBox, _
                               ByVal SectionName As String, ByVal BlockName As String, _
                               ByVal aIniFile As TIniFile)
        '// Read rich text from an ini file.
        Dim rtfText As String = ""
        Dim NO_SQL_rtfText As String = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil Consolas;}" & _
        "{\f1\fswiss\fprq2\fcharset0 Arial;}}" & vbCrLf & _
        "{\colortbl ;\red0\green0\blue255;\red0\green128\blue128;" & _
        "\red128\green128\blue128;\red255\green0\blue0;}" & vbCrLf & _
        "\viewkind4\uc1\pard\cf1\f0\fs19 NO SQL\cf0  \par" & vbCrLf & _
        "}"
        '//
        aIniFile.ReadValueBlockCRLF(SectionName, BlockName, NO_SQL_rtfText, rtfText)
        aRichTextBox.Rtf = rtfText
        '//
    End Sub

    '//=========================== Radio Button ==============================

    Public Sub SaveRadioToIni(ByVal aRadioBtn As System.Windows.Forms.RadioButton, _
                               ByVal Title As String, _
                               ByVal aIniFile As TIniFile)
        '// Save  Radio Box value to an ini file
        Dim Status As String
        '//
        If aRadioBtn.Checked Then
            Status = "T"
        Else
            Status = "F"
        End If
        aIniFile.WriteValue(Title, "RadioStatus", Status)
    End Sub

    Public Sub ReadIniToRadio(ByVal aRadioBtn As System.Windows.Forms.RadioButton, _
                              ByVal anOtherRadioBtn As System.Windows.Forms.RadioButton, _
                               ByVal Title As String, _
                               ByVal aIniFile As TIniFile)
        '// Read  Radio Box status (T/F) from an ini file
        Dim Status As String = "T"
        '//
        aIniFile.ReadValue(Title, "RadioStatus", "T", Status)
        If Status = "T" Then
            aRadioBtn.Checked = True
            anOtherRadioBtn.Checked = False
        Else
            aRadioBtn.Checked = False
            anOtherRadioBtn.Checked = True
        End If
    End Sub
End Module