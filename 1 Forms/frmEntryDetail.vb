' Copyright (C) 2015 - 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================

Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices

''' <summary>
''' Form to present the detail for a single entry
''' </summary>
''' <seealso cref="System.Windows.Forms.Form" />
Friend Class frmEntryDetail

    Private _DLLFilename As String = ""
    'Private Const fileprefix As String = "file:///"
    'Private Const fileprefixlength As Integer = 8
    ''' <summary>
    ''' Initializes a new instance of the <see cref="frmEntryDetail"/> class.
    ''' Populate with the detail for the selected addin row
    ''' </summary>
    ''' <param name="pEntryDetail">The p entry detail.</param>
    Protected Friend Sub New(pEntryDetail As AddInEntry)
        Try
            InitializeComponent()
            tbAddInName.Text = pEntryDetail.AddInName
            tbSparxRef.Text = pEntryDetail.SparxEntry
            tbAssemblyName.Text = pEntryDetail.ClassName
            tbClassSource.Text = pEntryDetail.ClassSource
            tbCLSID.Text = pEntryDetail.CLSID
            tbCLSIDSRC.Text = pEntryDetail.CLSIDSource

            tbDLL.Text = pEntryDetail.DLL
            _DLLFilename = pEntryDetail.DLL
            If _DLLFilename <> "" Then
                Dim _filename As String = _DLLFilename

                If Not File.Exists(_filename) Then
                    tbDLL.Text = "FILE DOES NOT EXIST"
                Else
                    tbDLL.Text = _filename
                    _filename = Path.GetFullPath(_filename)
                    ' Get the DLL file details
                    Dim fvi As FileVersionInfo = FileVersionInfo.GetVersionInfo(_filename)
                    ' now this fvi has all the properties for the FileVersion information.
                    tbDLLVersion.Text = fvi.FileVersion ' but other useful properties exist too.
                    Dim _DLLDate As DateTime = File.GetLastWriteTime(_filename)
                    tbDLLDate.Text = _DLLDate.ToString
                End If
            Else
                tbDLL.Text = "FILE DOES NOT EXIST"
            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub
    Private Function cleanFilename(pFilename As String) As String
        Try
            Dim start As Integer = Strings.InStr(pFilename, cFilePrefix)
            Dim _r As String = Strings.Right(pFilename, Len(pFilename) - fileprefixlength - start + 1)
            ' remove all after last \
            Dim _end As Integer = Strings.InStrRev(_r, cForwardSlash)

            Dim _f1 As String = Strings.RTrim(_r)
            _f1 = Strings.Replace(_f1, cForwardSlash, cBackSlash)
            Return _f1
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return pFilename
    End Function
    ''' <summary>
    ''' Handles the Click event of the btClose control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub btClose_Click(sender As Object, e As EventArgs) Handles btClose.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' Handles copy button which capture the details screen and copy to the windows clipboard
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub btCopyDetailToClipboard_Click(sender As Object, e As EventArgs) Handles btCopyDetailToClipboard.Click
        Try
            Dim gfx As Graphics = Me.CreateGraphics()
            Dim bmp As New Bitmap(Me.Width, Me.Height)
            Me.DrawToBitmap(bmp, New Rectangle(0, 0, Me.Width, Me.Height))
            My.Computer.Clipboard.SetImage(bmp)
        Catch ex As Exception
#If DEBUG Then'
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    ' NOT USED
    Private Sub btDLLDetail_Clickold(sender As Object, e As EventArgs)
        Try
            Dim _filename As String = Path.GetFullPath(_DLLFilename)

            'Dim filename As String = _DLLFilename.Replace("file:///", "")
            'filename = filename.Replace("/", "\")
            'filename = _f
            If File.Exists(_filename) Then
                Dim assembly As Assembly = Assembly.LoadFrom(_filename)
                Dim types As Type() = assembly.GetTypes()
                '   Dim s As String = " List of types " & vbCrLf
                Dim _ListOfTypes As New ArrayList
                For Each t As Type In types
                    _ListOfTypes.Add(t)
                Next
                '  MsgBox(s, MsgBoxStyle.OkOnly, "List of public methods for DLL")
                Dim frmMethods As New frmListOfClasses(_ListOfTypes)
                frmMethods.ShowDialog()
            Else
                MessageBox.Show("File " & _filename & " does not exist")
            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
            MessageBox.Show("Error accessing information " & vbCrLf & "(see detail below for some, but not all, the explanation): - " & vbCrLf & "====================================" _
                            & vbCrLf & ex.ToString)
        End Try
    End Sub

    ' This routine explores the selected DLL file using the metadataloadcontext
    ' https://docs.microsoft.com/en-us/dotnet/standard/assembly/inspect-contents-using-metadataloadcontext
    Private Sub btDLLDetail_Click(sender As Object, e As EventArgs) Handles btDLLDetail.Click
        Try
            If _DLLFilename = "" Then
                MsgBox("No DLL file defined", MsgBoxStyle.Exclamation, "No DLL")
                Return
            End If

            Dim _filename As String = Path.GetFullPath(_DLLFilename)

            Try
                Dim runtimeAssemblies As String() = Directory.GetFiles(RuntimeEnvironment.GetRuntimeDirectory(), "*.dll")
                Dim paths = New List(Of String)(runtimeAssemblies)
                paths.Add(_filename)
                Dim resolver = New PathAssemblyResolver(paths)
                Dim mlc = New MetadataLoadContext(resolver)

                Using mlc
                    Dim assembly As Assembly = mlc.LoadFromAssemblyPath(_filename)
                    Dim name As AssemblyName = assembly.GetName()
                    Debug.Print(name.Name & " has following attributes: ")

                    '' can skip these 
                    'For Each attr As CustomAttributeData In assembly.GetCustomAttributesData()

                    '    Try
                    '        Debug.Print(attr.AttributeType.ToString)
                    '    Catch ex As FileNotFoundException
                    '        Debug.Print("Error getting attribute type: " & ex.Message)
                    '    End Try
                    'Next


                    If assembly IsNot Nothing Then
                        Dim abc = assembly.GetExportedTypes()

                        'Dim s As String = "-----"
                        'For Each a In assembly.CustomAttributes
                        '    s += a.ToString & vbCrLf
                        'Next
                        'Debug.Print(s)
                        Dim s1 As String = "-------  Referenced assemblies ----------" & vbCrLf
                        s1 += assembly.GetReferencedAssemblies.ToString
                        For Each g In assembly.GetReferencedAssemblies
                            s1 += g.ToString & vbCrLf
                        Next
                        Debug.Print(s1)

                        Dim types As Type() = assembly.GetExportedTypes() ' GetTypes()
                        Dim _ListOfTypes As New ArrayList
                        For Each t As Type In types
                            Debug.Print(t.FullName)
                            _ListOfTypes.Add(t)
                        Next
                        '  MsgBox(s, MsgBoxStyle.OkOnly, "List of public methods for DLL")
                        Dim frmMethods As New frmListOfClasses(_ListOfTypes)
                        frmMethods.ShowDialog()
                    Else
                        MsgBox("Sorry not available as unable to load assembly for " & vbCrLf & _filename, vbExclamation, "Assembly Information Not Available")
                    End If

                    '==================== NEW TO REVIEW

                    'Dim frmMethods As New frmListOfClasses(_ListOfTypes)
                    'frmMethods.ShowDialog()

                    'For Each t As TypeInfo In assembly.GetTypes()

                    '    Try
                    '        Dim baseType As Type = t.BaseType

                    '        If t.IsClass Then
                    '            Debug.Print("class :" & t.FullName)
                    '        ElseIf t.IsValueType Then

                    '            If String.Equals(baseType?.FullName, "System.Enum", StringComparison.InvariantCulture) Then
                    '                Debug.Print("enum ")
                    '            Else
                    '                Debug.Print("struct ")
                    '            End If
                    '        ElseIf t.IsInterface Then
                    '            Debug.Print("interface ")
                    '        End If

                    '        Debug.Print(t.FullName)

                    '        If t.IsClass AndAlso Not String.Equals(baseType.FullName, "System.Object", StringComparison.InvariantCulture) Then
                    '            Debug.Print(" : " & baseType.FullName)
                    '        End If

                    '        'Console.WriteLine()
                    '    Catch ex As System.IO.FileNotFoundException
                    '        Debug.Print("Error: " & ex.Message)
                    '    End Try
                    'Next
                End Using

            Catch ex As IOException
                Debug.Print("I/O error occured when trying to load assembly: ")
                Debug.Print(ex.ToString())

            Catch ex As UnauthorizedAccessException
                Debug.Print("Access denied when trying to load assembly: ")
                Debug.Print(ex.ToString())
            Catch ex As Exception
                MsgBox("Unable to process DLL file defined", MsgBoxStyle.Exclamation, "Unable to process DLL")

            End Try

            '================OLD=============
            ' Dim tyt As Type
            '
            '   _filename = "C:\Users\EXploringEA\Downloads\dlls\E002_ASimpleEAMenu.DLL"
            '   _filename = "eaDocXAddIn.eaDocX_Addin"
            'If File.Exists(_filename) Then
            '    Debug.Print("DLL Detail: " & _filename)
            '    Dim assembly As Assembly = Nothing
            '    Dim myDllClass As Object = Nothing
            '    Try
            '        My.Settings.DllFolder = Path.GetDirectoryName(_DLLFilename)
            '        My.Settings.Save()
            '        checkDLL(_filename)
            '        assembly = Assembly.ReflectionOnlyLoadFrom(_filename)
            '        Dim abc = assembly.GetExportedTypes() 'Type(tbAssemblyName.Text)

            '        ' Dim b = assembly.GetAssembly("eaDocx.aa")

            '        '  assembly = Assembly.ReflectionOnlyLoadFrom(_filename)
            '        'filecontents = File.ReadAllBytes(_filename)
            '        'assembly = Assembly.Load(filecontents) 'File.ReadAllBytes(_filename))
            '        'myDllClass = assembly.CreateInstance(tbAssemblyName.Text)

            '        ' create a domain
            '        '  Dim aDomain As AppDomain = AppDomain.CreateDomain("DLLDomain")
            '        '  Dim obj As Object = aDomain.CreateInstanceFromAndUnp(_filename, tbAssemblyName.Text)

            '        Dim s As String = "-----"
            '        For Each a In assembly.CustomAttributes
            '            s += a.ToString & vbCrLf
            '        Next

            '        s += "-------  Referenced assemblies ----------" & vbCrLf
            '        s += assembly.GetReferencedAssemblies.ToString
            '        For Each g In assembly.GetReferencedAssemblies
            '            s += g.ToString & vbCrLf
            '        Next

            '        's += myDllClass.GetType.ToString
            '        Debug.Print(s)
            '    Catch bifex As BadImageFormatException
            '        MsgBox("Bad image format exception - which indicates that the DLL may not be a valid assembly at least in terms of loading by the EA Installation Inspector " _
            '               & "it could be due to the DLL being compiler with a later version of the CLR (.NET framework) than this tool. " & vbCrLf _
            '               & "NOTE: it may be this tool failing to load rather than your addin failing !" & vbCrLf _
            '               & "-----------------" & vbCrLf _
            '               & "Windows exception message below which may gives some clues " _
            '               & bifex.ToString, MsgBoxStyle.Exclamation, "Bad Image format exception")
            '        Return
            '    Catch fnfex As FileNotFoundException
            '        MsgBox("File Not found - could be that referenced dll's by your addin aren't present: " & fnfex.ToString, MsgBoxStyle.Exclamation, "File not found exception")
            '        Return
            '    Catch flex As FileLoadException
            '        MsgBox("File not found - could be that referenced dll's by your addin aren't present: " & flex.ToString, MsgBoxStyle.Exclamation, "File Load Exception")
            '        Return
            '    Catch ex As Exception
            '        MsgBox("Exception: " & ex.ToString, MsgBoxStyle.Exclamation, "Other exceptions")
            '        assembly = Nothing
            '    End Try
            '    If assembly IsNot Nothing Then

            '        Dim types As Type() = assembly.GetTypes()
            '        '   Dim s As String = " List of types " & vbCrLf
            '        Dim _ListOfTypes As New ArrayList
            '        For Each t As Type In types
            '            _ListOfTypes.Add(t)
            '        Next
            '        '  MsgBox(s, MsgBoxStyle.OkOnly, "List of public methods for DLL")
            '        Dim frmMethods As New frmListOfClasses(_ListOfTypes)
            '        frmMethods.ShowDialog()
            '    Else
            '        MsgBox("Sorry not available as unable to load assembly for " & vbCrLf & _filename, vbExclamation, "Assembly Information Not Available")
            '    End If

            'Else
            '    MessageBox.Show("File " & _filename & " does not exist")
            'End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
            MsgBox("Error accessing information " & vbCrLf & "(see detail below for some, but not all, the explanation): - " & vbCrLf & "====================================" _
                            & vbCrLf & ex.ToString, vbExclamation, "DLL Detail Not Available")
        End Try
    End Sub

    Private Sub checkDLL(pfilename As String)
        Dim DynamicDomain As AppDomain
        ' Dim _filename As String = cleanFilename(Filename)
        Try
            DynamicDomain = AppDomain.CreateDomain("DynamicDomain")
            Dim DynClass As Object = DynamicDomain.CreateInstanceFromAndUnwrap("C:\Program Files (x86)\E002\E002_ASimpleEAMenu.DLL", "E002_ASimpleEAMenu.ASimpleEAMenu")
            '     Dim MethodInfo As MethodInfo = DynClass.getTypes
            If DynamicDomain IsNot Nothing Then AppDomain.Unload(DynamicDomain)
        Catch ex As Exception

        End Try
    End Sub

End Class