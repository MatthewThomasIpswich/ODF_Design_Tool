Imports System.IO

Public Interface IAction

    Sub ActionMethod(FilePath As String, ByRef ErrorStr As String)

End Interface

Public Class TCommonUserInterface

    Public Sub SetPanelColor(aColorPanel As System.Windows.Forms.Panel)
        Dim dlgColor As New ColorDialog
        '// record colour selected. Draw colour in picture.
        dlgColor.Color = aColorPanel.BackColor
        If dlgColor.ShowDialog = DialogResult.OK Then
            '// Draw with new colour
            aColorPanel.BackColor = dlgColor.Color
            Application.DoEvents()
        End If
    End Sub

    Public Function SaveToFile(IniFile As TIniFile, SectionName As String, BlockName As String, FileExtension As String, ByRef aSelectedFilePath As String, Optional subPath As String = "") As Boolean
        Dim aDefaultPath As String
        Dim aDirectoryName As String
        Dim aSaveFileDialog As New SaveFileDialog
        Dim ErrorStr As String = ""
        '//
        If subPath.Length > 0 Then
            aDefaultPath = Application.StartupPath + subPath
        Else
            aDefaultPath = Application.StartupPath
        End If
        IniFile.ReadValue(SectionName, BlockName, aDefaultPath, aSelectedFilePath)
        aSaveFileDialog.Filter = "files (*." + FileExtension + ")|*." + FileExtension + "|All files (*.*)|*.*"
        aSaveFileDialog.FilterIndex = 1
        aDirectoryName = Path.GetDirectoryName(aSelectedFilePath)
        aSaveFileDialog.InitialDirectory = aDirectoryName
        '//
        If aSaveFileDialog.ShowDialog = DialogResult.OK Then
            aSelectedFilePath = aSaveFileDialog.FileName
            IniFile.WriteValue(SectionName, BlockName, aSelectedFilePath)
            IniFile.SaveToFile(IniFile.FilePath)
            '//
            Return True
        End If
        '//
        Return False
    End Function

    Public Function OpenFile(IniFile As TIniFile, SectionName As String, BlockName As String, FileExtension As String, ByRef aSelectedFilePath As String, Optional subPath As String = "") As Boolean
        Dim aDefaultPath As String
        Dim aDirectoryName As String
        Dim aOpenFileDialog As New OpenFileDialog
        '//
        If subPath.Length > 0 Then
            aDefaultPath = Application.StartupPath + subPath
        Else
            aDefaultPath = Application.StartupPath
        End If
        IniFile.ReadValue(SectionName, BlockName, aDefaultPath, aSelectedFilePath)
        aOpenFileDialog.Filter = "files (*." + FileExtension + ")|*." + FileExtension + "|All files (*.*)|*.*"
        aOpenFileDialog.FilterIndex = 1
        aDirectoryName = Path.GetDirectoryName(aSelectedFilePath)
        aOpenFileDialog.InitialDirectory = aDirectoryName
        '//
        If aOpenFileDialog.ShowDialog = DialogResult.OK Then
            aSelectedFilePath = aOpenFileDialog.FileName
            IniFile.WriteValue(SectionName, BlockName, aSelectedFilePath)
            IniFile.SaveToFile(IniFile.FilePath)
            '//
            Return True
        End If
        '//
        Return False
    End Function

    Public Function Add_Model_And_makeLayer(ByRef PossiblLayerName As String, aNewModle As IDataModels, Layers As TLayers) As Boolean
        Return True
    End Function

End Class
