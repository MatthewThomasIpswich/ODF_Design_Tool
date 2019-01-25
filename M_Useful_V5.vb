Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Module M_Useful_V5
    Public Const INFINITY As Integer = 9123
    '//

    Public Function RoundUpInt(ByVal Avalue As Integer, _
                               ByVal Bvalues As Integer) As Integer
        '// round up  10/5 = 2, 10/6 = 2, 10/4 = 3, 10/3 = 4, 10/11 = 1
        '// how many blocks of B needed to satify A
        Dim Remainder As Integer
        Dim roundedUp As Integer
        '//
        If Avalue < Bvalues Then
            roundedUp = 1
        Else
            '//
            Remainder = Avalue Mod Bvalues
            If Remainder = 0 Then
                roundedUp = Avalue \ Bvalues
            Else
                roundedUp = Avalue \ Bvalues + 1
            End If
        End If
        '//
        Return roundedUp
    End Function

    Public Function RoundupReal(ByVal Avalue As Double) As Integer
        Dim aInt As Integer
        aInt = RoundDown(Avalue + 0.5)
        Return aInt
    End Function


    Public Sub ColourList(ByVal aNum As Integer, _
                          ByRef Red As Integer, _
                          ByRef Green As Integer, _
                          ByRef Blue As Integer)
        Dim caseNum As Integer
        Dim LocalNum As Integer
        '//
        '//
        LocalNum = 3 * aNum
        LocalNum = LocalNum Mod 100
        caseNum = aNum Mod 6
        Select Case caseNum
            Case 0
                Red = NextNum(1, LocalNum)
                Green = NextNum(1, LocalNum)
                Blue = NextNum(1, LocalNum)
            Case 1
                Red = NextNum(50, LocalNum)
                Green = NextNum(1, LocalNum)
                Blue = NextNum(100, LocalNum)
            Case 2
                Red = NextNum(1, LocalNum)
                Green = NextNum(50, LocalNum)
                Blue = NextNum(100, LocalNum)
            Case 3
                Red = NextNum(1, LocalNum)
                Green = NextNum(100, LocalNum)
                Blue = NextNum(50, LocalNum)
            Case 4
                Red = NextNum(100, LocalNum)
                Green = NextNum(1, LocalNum)
                Blue = NextNum(50, LocalNum)
            Case 5
                Red = NextNum(100, LocalNum)
                Green = NextNum(80, LocalNum)
                Blue = NextNum(50, LocalNum)
        End Select
    End Sub

    Public Function NextNum(ByVal start As Integer, _
                            ByVal x As Integer) As Integer
        Dim Nextvalue As Integer
        '//
        Nextvalue = (start + x) Mod 100 + 155
        Return Nextvalue
    End Function

    Public Function MoveUp(ByVal Value As Integer, _
                            ByVal max As Integer) As Integer
        '// used to darken a colour
        Dim NewValue As Integer
        '//
        NewValue = Value + 15
        If NewValue > max Then NewValue = max
        Return NewValue
    End Function

    Public Function testForPositiveNumeric(ByRef value As String) As Boolean
        Dim allowedChars As String = "0123456789"
        For i As Integer = value.Length - 1 To 0 Step -1
            If allowedChars.IndexOf(value(i)) = -1 Then
                Return False
            End If
        Next
        Return True
    End Function
    '// Use IsNumeric


    Public Function RoundUpInt(ByVal AReal As Double, _
                               ByVal Bvalues As Integer) As Integer
        '// round up  10/5 = 2, 10/6 = 2, 10/4 = 3, 10/3 = 4, 10/11 = 1
        '// how many blocks of B needed to satify A
        Dim Avalue As Integer
        Dim Remainder As Integer
        Dim roundedUp As Integer
        '//
        If AReal < Bvalues Then
            roundedUp = 1
        Else
            Avalue = RoundUp(AReal)
            '//
            Remainder = Avalue Mod Bvalues
            If Remainder = 0 Then
                roundedUp = Avalue \ Bvalues
            Else
                roundedUp = Avalue \ Bvalues + 1
            End If
        End If
        '//
        Return roundedUp
    End Function

    Public Function RoundUp(ByVal num As Double) As Integer
        Try
            'Temp variable to hold the decimal portion of the parameter
            Dim temp As Double
            'Get the decimal portion
            temp = num - Math.Truncate(num)
            'If there is a decimal portion then we add 1 to force it to round up
            If temp > 0 Then
                num = num + 1
            End If
            'return the truncated version of the double which should be the number rounded up
            Return CType(Math.Truncate(num), Integer)
        Catch ex As Exception
            MessageBox.Show("Error occured in: " & ex.StackTrace & vbCrLf & vbCrLf & "Error: " & ex.Message, _
                "Error.", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Public Function RoundDown(ByVal num As Double) As Integer
        Try
            'return the truncated version of the double which should be the number rounded down
            Return CType(Math.Truncate(num), Integer)
        Catch ex As Exception
            MessageBox.Show("Error occured in: " & ex.StackTrace & vbCrLf & vbCrLf & "Error: " & ex.Message, _
                "Error.", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Public Function StrToIntDef(ByVal LineStr As String, ByVal DefInt As Integer, ByRef Value As Integer) As Boolean
        '// convert a string to an integer - without rasing an exception if it fails.
        Dim i As Integer
        Dim Start As Integer
        Dim ASCIINum As Byte
        Dim EncodeTxt As New System.Text.ASCIIEncoding
        Dim ASCIITxtAr() As Byte
        Dim FoundNumber As Boolean
        Dim Negative As Boolean
        '//
        Value = 0
        Start = 0
        FoundNumber = True
        Negative = False
        LineStr = Trim(LineStr)
        '//
        ASCIITxtAr = EncodeTxt.GetBytes(LineStr)
        If ASCIITxtAr.Length > 0 Then
            ASCIINum = ASCIITxtAr(0)
            If ASCIINum = 45 Then
                Negative = True
                Start = 1
            End If
        Else
            Value = DefInt
            Return False
        End If
        '//
        For i = Start To ASCIITxtAr.Length - 1
            ASCIINum = ASCIITxtAr(i)
            If ((ASCIINum > 47) And (ASCIINum < 58)) Then
                Value = Value * 10 + (ASCIINum - 48)
            Else
                Value = DefInt
                Return False
            End If
        Next
        If Negative Then
            Value = -1 * Value
        End If
        Return FoundNumber
    End Function '//

    Public Function StrToRealDef(ByVal LineStr As String, ByVal DefInt As Double, ByRef Value As Double) As Boolean
        '// convert a string to an integer - without rasing an exception if it fails.
        Dim i As Integer
        Dim Start As Integer
        Dim ASCIINum As Byte
        Dim EncodeTxt As New System.Text.ASCIIEncoding
        Dim ASCIITxtAr() As Byte
        Dim FoundNumber As Boolean
        Dim Negative As Boolean
        Dim BeforeDecPoint As Boolean
        Dim HighDigits As Double = 0
        Dim LowDigits As Double = 0
        Dim PosDecPoint As Integer = -1
        '//
        Value = 0
        Start = 0
        FoundNumber = True
        Negative = False
        BeforeDecPoint = True
        LineStr = Trim(LineStr)
        '//
        If LineStr.Length = 0 Then
            FoundNumber = False
            Value = DefInt
            Return FoundNumber
        End If
        '//
        ASCIITxtAr = EncodeTxt.GetBytes(LineStr)
        If ASCIITxtAr.Length > 1 Then
            ASCIINum = ASCIITxtAr(0)
            If ASCIINum = 45 Then
                Negative = True
                Start = 1
            End If
        End If
        '// move down digits to dec point.
        For i = Start To ASCIITxtAr.Length - 1
            ASCIINum = ASCIITxtAr(i)
            If ((ASCIINum = 46) And BeforeDecPoint) Then '// dec point reached
                BeforeDecPoint = False
                PosDecPoint = i
                Exit For
            Else
                If ((ASCIINum > 47) And (ASCIINum < 58)) Then '// 0..9
                    HighDigits = HighDigits * 10 + (ASCIINum - 48)
                Else
                    FoundNumber = False
                    Value = DefInt
                    Exit For
                End If
            End If
        Next
        '// Move up digits to dec point.
        If FoundNumber And (Not BeforeDecPoint) And (PosDecPoint > -1) Then
            For i = ASCIITxtAr.Length - 1 To PosDecPoint + 1 Step -1
                ASCIINum = ASCIITxtAr(i)
                If ((ASCIINum > 47) And (ASCIINum < 58)) Then '// 0..9
                    LowDigits = (ASCIINum - 48) + LowDigits * 0.1
                Else
                    FoundNumber = False
                    Value = DefInt
                    Exit For
                End If
            Next
            LowDigits = LowDigits * 0.1
        End If
        '//
        Value = HighDigits + LowDigits
        If Negative Then
            Value = -1 * Value
        End If
        Return FoundNumber
    End Function '//

    Public Function ExtractLines(ByVal Path As String, ByVal StartEndStr As String) As String
        '// read in a text file on the given path and put into a grid.
        Dim aLine As String
        Dim forReading As IO.StreamReader
        Dim Extract As String = ""
        Dim ReadOk As Boolean
        Dim ErrorStr As String
        Dim Extracted As New TStringList
        Dim FoundExtraxct As Boolean
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
            '//
            Extracted.Clear()
            FoundExtraxct = False
            forReading = New IO.StreamReader(Path)
            Try
                '// Read through the file
                While forReading.Peek > -1
                    aLine = forReading.ReadLine()
                    If aLine.StartsWith(StartEndStr) Then
                        If Not FoundExtraxct Then
                            FoundExtraxct = True
                        Else
                            FoundExtraxct = False
                        End If
                    End If
                    If FoundExtraxct Then
                        Extracted.Add(aLine)
                    End If
                End While
            Catch FileEx As Exception
                ReadOk = False
                MsgBox(FileEx.Message)
            Finally
                forReading.Close()
            End Try
        End If
        '//
        Extract = Extracted.TextCRLF
        Return Extract
    End Function

    Public Function SafeDivide(ByVal Top As Integer, _
                               ByVal Bottom As Integer, _
                               ByVal defValue As Integer) _
                               As Double
        Const delta As Double = 0.0001
        '// Avoid divde by zero errors
        If Bottom < delta Then
            '// give the default value.
            Return CType(defValue, Double)
        Else
            Return Top / Bottom
        End If
    End Function '//-----------------------------------

    Public Function SafeDivide(ByVal Top As Double, _
                               ByVal Bottom As Double, _
                               ByVal defValue As Double) _
                               As Double
        Const delta As Double = 0.0001
        '// Avoid divde by zero errors
        If Bottom < delta Then
            '// give the default value.
            Return CType(defValue, Double)
        Else
            Return Top / Bottom
        End If
    End Function '//-----------------------------------, _
 

    Public Function SafeRead(ByVal myReader As OleDbDataReader, _
                        ByVal DBCol As String, _
                        ByRef readIn As Boolean) As String
        '// Read - assumes can all be strings
        Dim ReadStr As String = ""
        '//
        If myReader(DBCol) Is DBNull.Value Then
            readIn = False
        Else
            '// do not out in "readIn = true"
            ReadStr = myReader(DBCol).ToString
        End If
        Return ReadStr
    End Function

    Public Function SQLSafeRead(ByVal myReader As SqlDataReader, _
                        ByVal DBCol As String, _
                        ByRef readIn As Boolean) As String
        '// Read in from standard format database
        Dim ReadStr As String = ""
        '//
        If myReader(DBCol) Is DBNull.Value Then
            readIn = False
        Else
            '// do not out in "readIn = true"
            ReadStr = myReader(DBCol).ToString
        End If
        Return ReadStr
    End Function

    Public Function SQLDefaultRead(ByVal myReader As SqlDataReader, _
                        ByVal DBCol As String, _
                        ByRef DefaultStr As String) As String
        '// Read in from standard format database
        Dim ReadStr As String = ""
        '//
        If myReader(DBCol) Is DBNull.Value Then
            ReadStr = DefaultStr
        Else
            '// do not out in "readIn = true"
            ReadStr = myReader(DBCol).ToString
        End If
        Return ReadStr
    End Function

    Public Function SQLQuerySafeRead(ByVal myReader As SqlDataReader, _
                        ByVal DBCol As String, _
                        ByRef readIn As Boolean) As String
        '// Read in from standard format database
        Dim ReadStr As String = ""
        '//
        If myReader(DBCol) Is DBNull.Value Then
            readIn = False
            ReadStr = "NULLStr"
        Else
            '// do not out in "readIn = true"
            ReadStr = myReader(DBCol).ToString
        End If
        Return ReadStr
    End Function



    Public Function Distance(ByVal E1 As Double, _
                              ByVal N1 As Double, _
                              ByVal E2 As Double, _
                              ByVal N2 As Double) _
                              As Double
        '// Distance using Pythagoras.
        Dim Dist As Double
        '//
        Dist = Math.Sqrt((E1 - E2) * (E1 - E2) + (N1 - N2) * (N1 - N2))
        '//
        Return Dist
    End Function

    Public Function AppPath() As String
        Dim path As String
        ''path = System.IO.Path.GetDirectoryName( _
        ''   System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).
        path = System.AppDomain.CurrentDomain.BaseDirectory
        Return path
    End Function

    Public Function DirForFile(ByVal FilePath As String) As String
        Dim LastSlashPos As Integer
        Dim Dir As String
        '//
        LastSlashPos = FilePath.LastIndexOf("\")
        Dir = FilePath.Substring(0, LastSlashPos)
        Return Dir
    End Function

    

    Public Sub GetMidPoint(ByVal aPointArray As Point(), _
                           ByRef MidX As Integer, _
                           ByRef MidY As Integer)
        '// Take in an array of points and find the mid point.
        Dim MinX As Integer = INFINITY
        Dim MaxX As Integer = 0
        Dim MinY As Integer = INFINITY
        Dim MaxY As Integer = 0
        Dim i As Integer
        Dim aPoint As Point
        '//
        For i = 0 To aPointArray.Count - 1
            aPoint = aPointArray(i)
            '//
            If aPoint.X < MinX Then
                MinX = aPoint.X
            End If
            '//
            If aPoint.X > MaxX Then
                MaxX = aPoint.X
            End If
            '//
            If aPoint.Y < MinY Then
                MinY = aPoint.Y
            End If
            '//
            If aPoint.Y > MaxY Then
                MaxY = aPoint.Y
            End If
            '//
        Next
        MidX = (MaxX + MinX) \ 2
        MidY = (MaxY + MinY) \ 2
    End Sub

    Public Sub GetMidPoint(ByVal aPointList As List(Of Point),
                          ByRef MidX As Integer,
                          ByRef MidY As Integer)
        '// Take in a list of points and find the mid point.
        Dim MinX As Integer = INFINITY
        Dim MaxX As Integer = 0
        Dim MinY As Integer = INFINITY
        Dim MaxY As Integer = 0
        Dim i As Integer
        Dim aPoint As Point
        '//
        For i = 0 To aPointList.Count - 1
            aPoint = aPointList(i)
            '//
            If aPoint.X < MinX Then
                MinX = aPoint.X
            End If
            '//
            If aPoint.X > MaxX Then
                MaxX = aPoint.X
            End If
            '//
            If aPoint.Y < MinY Then
                MinY = aPoint.Y
            End If
            '//
            If aPoint.Y > MaxY Then
                MaxY = aPoint.Y
            End If
            '//
        Next
        MidX = (MaxX + MinX) \ 2
        MidY = (MaxY + MinY) \ 2
    End Sub

    Public Function StripOfHeading(Text As String, DividStr As String) As String
        Dim StrippedStr As String
        Dim index As Integer
        '//
        index = Text.IndexOf(DividStr)
        If index > 0 Then
            StrippedStr = Text.Remove(0, index + 1)
        Else
            StrippedStr = Text
        End If
        Return StrippedStr
    End Function

End Module

Public Class TPrintpages
    Public PictureWidth As Integer
    Public PictureHeight As Integer
    Public PageWidth As Integer
    Public PageHeight As Integer

    Private Function NumPages(ByVal PictureDimension As Integer, _
                                ByVal PageDimension As Integer) _
                                As Integer
        Dim Pages As Integer
        Dim Remainder As Integer
        '//
        Remainder = PictureDimension Mod PageDimension
        If Remainder = 0 Then
            Pages = PictureDimension \ PageDimension
        Else
            Pages = PictureDimension \ PageDimension + 1
        End If
        Return Pages
    End Function

    Public Function PosAcross(ByVal pageNum As Integer) As Integer
        If pageNum <= TotalPages Then
            Return (pageNum - 1) Mod PagesAcross
        Else
            Return -1
        End If
    End Function

    Public Function PosDown(ByVal pageNum As Integer) As Integer
        If pageNum <= TotalPages Then
            Return (pageNum - 1) \ PagesAcross
        Else
            Return 0
        End If
    End Function

    ReadOnly Property PagesAcross() As Integer
        Get
            PagesAcross = NumPages(PictureWidth, PageWidth)
        End Get
    End Property
    '//

    ReadOnly Property PagesDown() As Integer
        Get
            PagesDown = NumPages(PictureHeight, PageHeight)
        End Get
    End Property


    ReadOnly Property TotalPages() As Integer
        Get
            TotalPages = PagesAcross * PagesDown
        End Get
    End Property

    ReadOnly Property PageDimensions(ByVal pageNum As Integer) As Rectangle
        Get
            Dim R As New Rectangle
            '//
            If pageNum > TotalPages Then
                Return R
                Exit Property
            End If
            R.X = PosAcross(pageNum) * PageWidth
            R.Y = PosDown(pageNum) * PageHeight
            R.Width = PageWidth
            R.Height = PageHeight
            '//
            Return R
        End Get
    End Property

    

End Class


