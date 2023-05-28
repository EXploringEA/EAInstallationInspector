Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
' this file contains routines to:
' * read the registry
' * View information on keys
' * export registry entries to create and/or delete the registry entries

Friend Class RegClass

    Private _regfileCreateKeys As RegFileOutput
    Private _regFileDeleteKeys As RegFileOutput

    Private _AddInEntry As AddInEntry
    Private _ClassNamespace As String = ""


    Private Const cHKCU As String = "CU"
    Private Const cHKLM As String = "LM"

    Private Const cSoftwareClasses As String = "SOFTWARE\Classes"
    Private Const cSoftware_Wow6432Node As String = cSoftwareClasses & "\Wow6432Node"
    Private Const cHKLM_Root As String = "HKEY_LOCAL_MACHINE"
    Private Const cHKCU_Root As String = "HKEY_CURRENT_USER"
    Private cRegHeader As String = "Windows Registry Editor Version 5.00"


    Friend Sub New(pEntryDetails As AddInEntry)
        Try
            _AddInEntry = pEntryDetails
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

#Region "Output routines"
    '  Private OutputString As String = ""
    ' routine that outputs the information both to file and concatenates to create an output string
    Sub outputReg(pString As String, pExposedAssembly As Boolean)
        Try
            If _regfileCreateKeys IsNot Nothing Then _regfileCreateKeys.output(pString)
            '  OutputString += pString & vbCrLf
            Dim ki As New RegKeyItem
            ki.KeyName = pString
            ki.Exposed = If(pExposedAssembly, RegKeyItem.KeyType.ExposedAssembly, RegKeyItem.KeyType.Standard)

            ListOfRegStrings.Add(ki)
            ListOfCreateRegistryEntries.Add(ki)
#If DEBUG Then
            Debug.Print(pString)
#End If
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    ''' <summary>
    ''' Routine to output string to Delete Keys file 
    ''' If debug to immediate window
    ''' </summary>
    ''' <param name="pString"></param>
    Sub outputDeleteKey(pString As String, pExposedAssembly As Boolean)
        Try
            If _regFileDeleteKeys IsNot Nothing Then _regFileDeleteKeys.output(pString)
            Dim ki As New RegKeyItem
            ki.KeyName = pString
            ki.Exposed = If(pExposedAssembly, RegKeyItem.KeyType.ExposedAssembly, RegKeyItem.KeyType.Standard)
            ListOfDeleteRegistryEntries.Add(ki)
#If DEBUG Then
            Debug.Print(pString)
#End If
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    ''' <summary>
    ''' Example: @="mscoree.dll"
    ''' </summary>
    ''' <param name="pkey_value"></param>
    Sub OutputDefaultKeyValue(pkey_value As String, pExposed As Boolean)
        outputReg("@=" & Chr(34) & pkey_value & Chr(34), pExposed)
    End Sub

    ''' <summary>
    ''' Example: [HKEY_CURRENT_USER\Software\Sparx Systems\EAAddins\E002_32]
    ''' </summary>
    ''' <param name="pRoot"></param>
    ''' <param name="pRootLocation"></param>
    Sub outputKeyName(pRoot As String, pRootLocation As String, pExposedAssembly As Boolean)
        outputReg("[" & pRoot & cBackSlash & pRootLocation & "]", pExposedAssembly)
        outputDeleteKey("[-" & pRoot & cBackSlash & pRootLocation & "]", pExposedAssembly)
    End Sub

    Sub outputKeyName(pRootLocation As String, pExposedAssembly As Boolean)
        outputReg("[" & pRootLocation & "]", pExposedAssembly)
        outputDeleteKey("[-" & pRootLocation & "]", pExposedAssembly)
    End Sub

    ''' <summary>
    ''' Example: "RuntimeVersion"="v4.0.30319"
    ''' </summary>
    ''' <param name="pKey"></param>
    ''' <param name="pValue"></param>
    Sub outputKeyValuePair(pKey As String, pValue As String, pExposedAssembly As Boolean)
        outputReg(Chr(34) & pKey & Chr(34) & "=" & Chr(34) & pValue & Chr(34), pExposedAssembly)
    End Sub


    Private ListOfRegStrings As New List(Of RegKeyItem)
    Private ListOfCreateRegistryEntries As New List(Of RegKeyItem)
    Private ListOfDeleteRegistryEntries As New List(Of RegKeyItem)

    ''' <summary>
    ''' Outputs: Windows Registry Editor Version 5.00
    ''' When we output header we init the array for strings
    ''' </summary>
    Private Sub OutputRegHeader(pHive As String)
        ListOfRegStrings.Clear()
        ListOfCreateRegistryEntries.Clear()
        ListOfDeleteRegistryEntries.Clear()
        outputReg(cRegHeader, False)
        outputDeleteKey(cRegHeader, False)
    End Sub
    ''' <summary>
    ''' Example: [HKEY_CURRENT_USER\Software\Sparx Systems\EAAddins\E002_32]
    '''        : @="E002_ASimpleEAMenu.ASimpleEAMenu"
    ''' </summary>
    ''' <param name="pSparxLocation"></param>
    ''' <param name="pSparxRegistryLocation"></param>
    Private Sub OutputSparxKey(pSparxLocation As String, pSparxRegistryLocation As String)

        outputkeysAndSubKeysWithValues(pSparxLocation, pSparxRegistryLocation)
    End Sub
    ''' <summary>
    ''' Example: [HKEY_CURRENT_USER\SOFTWARE\Classes\E002_ASimpleEAMenu.ASimpleEAMenu\CLSID]
    '''         : @="{ADD9FA80-D738-3EFE-B8C3-012722A3A2D3}"
    ''' </summary>
    ''' <param name="pHIVE"></param>
    ''' <param name="pClassRegistryLocation"></param>
    Private Sub OutputClassCLISDEntry(pHIVE As String, pClassRegistryLocation As String)
        outputkeysAndSubKeysWithValues(pHIVE, pClassRegistryLocation)
    End Sub


#End Region

    Private Function getNameSpaceFromClassName(pClassname As String) As String
        Try
            If pClassname.Contains(".") Then
                Dim positionOfDot As Integer = Strings.InStr(pClassname, ".")
                Return IIf(positionOfDot > 1, Strings.Left(pClassname, positionOfDot - 1), "")
            Else
                Debug.Print("Classname within 2 parts: " & _AddInEntry.ClassName)
            End If
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return ""
    End Function

    Private listOfFilenames As New List(Of String)

    ''' <summary>
    ''' Handler routine to get information for the currently selected item in the listview (Potentil refactor out)
    ''' Will use the Sparx Entry
    ''' </summary>
    Friend Sub GetRegistryValuesHandler() '_AddInEntry As AddInEntry)
        Try

            ' what do we initialise
            listOfFilenames.Clear()

            'OutputString = ""
            If _AddInEntry.ClassName <> "" Then _ClassNamespace = getNameSpaceFromClassName(_AddInEntry.ClassName)

            Dim _classpath As String = cSoftwareClasses & cBackSlash & _AddInEntry.ClassName
            Dim _64clsidpath As String = cSoftwareClasses & cBackSlash & cCLSID & cBackSlash & _AddInEntry.CLSID
            Dim _32clsidpath As String = IIf(ClassInformation.OS64Bit, cSoftware_Wow6432Node & cBackSlash & cCLSID & cBackSlash & _AddInEntry.CLSID,
                         cSoftwareClasses & cBackSlash & cCLSID & cBackSlash & _AddInEntry.CLSID)

            Select Case _AddInEntry.SparxEntry

                Case AddInEntry.cHKCU32

                    OutputRegHeader(AddInEntry.cHKCU32)
                    OutputSparxKey(cHKCU, cSparxKeys32 & cBackSlash & _AddInEntry.AddInName)
                    OutputClassCLISDEntry(cHKCU, _classpath)
                    Try
                        Dim clsid_key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey(_32clsidpath)
                        If clsid_key IsNot Nothing Then OutputCLSIDKeyAndSubKeys(cHKCU, clsid_key, False) ', True)
                    Catch
                    End Try

                Case AddInEntry.cHKLM32

                    If IsElevated = False Then   ' ONLY possible if elevated
                        MsgBox("Cannot process - application needs to run as Administrator")
                        Return
                    End If
                    OutputRegHeader(AddInEntry.cHKLM32)
                    OutputSparxKey(cHKLM, cSparxKeysWOW32 & cBackSlash & _AddInEntry.AddInName) ' _SparxRoot)
                    OutputClassCLISDEntry(cHKLM, _classpath)
                    Try
                        Dim clsid_key As RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey(_32clsidpath)
                        If clsid_key IsNot Nothing Then OutputCLSIDKeyAndSubKeys(cHKLM, clsid_key, False) ', True)
                    Catch
                    End Try


                Case AddInEntry.cHKCU64
                    OutputRegHeader(AddInEntry.cHKCU64)
                    OutputSparxKey(cHKCU, cSparxKeys64 & cBackSlash & _AddInEntry.AddInName) ' _SparxRoot)
                    OutputClassCLISDEntry(cHKCU, _classpath)
                    Try
                        Dim clsid_key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey(_64clsidpath)
                        If clsid_key IsNot Nothing Then OutputCLSIDKeyAndSubKeys(cHKCU, clsid_key, False) ', True)
                    Catch
                    End Try


                Case AddInEntry.cHKLM64

                    If IsElevated = False Then   ' ONLY possible if elevated
                        MsgBox("Cannot process - application needs to run as Administrator")
                        Return
                    End If
                    OutputRegHeader(AddInEntry.cHKLM64)
                    OutputSparxKey(cHKLM, cSparxKeys64 & cBackSlash & _AddInEntry.AddInName) ' _SparxRoot)
                    OutputClassCLISDEntry(cHKLM, _classpath)
                    Try
                        Dim clsid_key As RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey(_64clsidpath)
                        If clsid_key IsNot Nothing Then OutputCLSIDKeyAndSubKeys(cHKLM, clsid_key, False) ', True)
                    Catch
                    End Try
                Case Else
                    MsgBox("No keys available")

            End Select

            Dim regstrings As New frmDisplayRegistryKeys()
            regstrings.init(ListOfRegStrings, ListOfCreateRegistryEntries, ListOfDeleteRegistryEntries, _AddInEntry.ClassName, _AddInEntry.SparxEntry)
            regstrings.ShowDialog()


            'MsgBox(OutputString)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private clsidtest As String = ""
    Private Sub outputkeysAndSubKeysWithValues(pHive As String, pRootKey As String, Optional pExposed As Boolean = False)
        clsidtest = ""
        Dim keyroot As Microsoft.Win32.RegistryKey = Nothing
        Dim _pathPrefix As String = "["
        Dim _deletePathPrefix As String = "[-"
        Try
            Select Case pHive
                Case cHKCU
                    keyroot = My.Computer.Registry.CurrentUser.OpenSubKey(pRootKey)
                    _pathPrefix += cHKCU_Root
                    _deletePathPrefix += cHKCU_Root
                Case cHKLM
                    keyroot = My.Computer.Registry.LocalMachine.OpenSubKey(pRootKey)
                    _pathPrefix += cHKLM_Root
                    _deletePathPrefix += cHKLM_Root
            End Select

            outputReg(_pathPrefix & cBackSlash & pRootKey & "]", pExposed) ' KEY
            outputDeleteKey(_deletePathPrefix & cBackSlash & pRootKey & "]", pExposed) ' KEY
            If keyroot IsNot Nothing Then

                OutputDefaultKeyValue(Trim(keyroot.GetValue("")), pExposed)

                For Each subkey As String In keyroot.GetSubKeyNames
                    Dim subkey_Path As String = pRootKey & "\" & subkey
                    If subkey_Path IsNot Nothing Then
                        outputReg(_pathPrefix & cBackSlash & pRootKey & "\" & subkey & "]", pExposed) 'KEY
                        outputDeleteKey(_deletePathPrefix & cBackSlash & pRootKey & "\" & subkey & "]", pExposed)
                        Dim sb1 As Microsoft.Win32.RegistryKey = IIf(pHive = cHKCU, My.Computer.Registry.CurrentUser.OpenSubKey(subkey_Path), My.Computer.Registry.LocalMachine.OpenSubKey(subkey_Path))
                        If sb1 IsNot Nothing Then
                            If subkey_Path.Contains("CLSID") Then clsidtest = sb1.GetValue("")
                            OutputDefaultKeyValue(sb1.GetValue(""), pExposed) ' default value for CLSID is the CLSID
                            If sb1.SubKeyCount > 0 Then
                                For Each subkeyname In sb1.GetValueNames
                                    If subkeyname = "" Then Continue For
                                    outputKeyValuePair(subkeyname, sb1.GetValue(subkeyname), pExposed)
                                Next
                                outputkeysAndSubKeysWithValues(pHive, subkey_Path)
                            End If

                        End If
                    End If
                Next
            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    ' outputs information with check for assembly - possibly optionally exclude the getassembly to consolidate the routines
    Private Sub OutputCLSIDKeyAndSubKeys(pHive As String, Keyroot As RegistryKey, pExposed As Boolean) ', pGetAssemblies As Boolean)

        Try
            If Keyroot IsNot Nothing AndAlso Keyroot.Name.Contains("{") Then
                outputKeyName(Keyroot.Name, pExposed)
                For Each keyname In Keyroot.GetValueNames
                    Dim val As String = Trim(Keyroot.GetValue(keyname))
                    Dim val2 As String = val.Replace("\", "\\")
                    If keyname = "" Then
                        OutputDefaultKeyValue(val2, pExposed)
                    Else
                        outputKeyValuePair(keyname, val2, pExposed)
                    End If
                    '                    If _GetAssemblies AndAlso keyname.Contains(cCodeBase) Then
                    If keyname.Contains(cCodeBase) Then
                        Dim _filename As String = val2
                        If Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)
                        If _filename <> "" AndAlso _filename <> cNotSet Then
                            _filename = _filename.Replace("/", "\")
                            If File.Exists(_filename) AndAlso Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)
                            If Not listOfFilenames.Contains(_filename) Then
                                listOfFilenames.Add(_filename)
                                getassemblykeys(_filename, pHive)
                            End If
                        End If
                    End If
                Next

                For Each subkey As String In Keyroot.GetSubKeyNames
                    Dim sk As RegistryKey = Keyroot.OpenSubKey(subkey)
                    If sk IsNot Nothing Then OutputCLSIDKeyAndSubKeys(pHive, sk, pExposed)
                Next
            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub
    ' Get assembly information
    '  Private _GetAssemblies As Boolean = True
    Sub getassemblykeys(pDLLFilename As String, pHive As String)

        Try
            '  outputReg("Get assemblies for " & pDLLFilename & " in hive " & pHive)
            If pDLLFilename = "" Then
                MsgBox("No DLL file defined", MsgBoxStyle.Exclamation, "No DLL")
                Return
            End If

            Dim _filename As String = Path.GetFullPath(pDLLFilename)

            Try
                Dim runtimeAssemblies As String() = Directory.GetFiles(RuntimeEnvironment.GetRuntimeDirectory(), "*.dll")
                Dim paths = New List(Of String)(runtimeAssemblies)
                paths.Add(_filename)
                Dim resolver = New PathAssemblyResolver(paths)
                Dim mlc = New MetadataLoadContext(resolver)

                Using mlc
                    Dim assembly As Assembly = mlc.LoadFromAssemblyPath(_filename)
                    Dim name As AssemblyName = assembly.GetName()
                    If assembly IsNot Nothing Then
                        For Each t As Type In assembly.GetExportedTypes()
                            If t.FullName.Contains(_ClassNamespace) AndAlso Not t.FullName = _AddInEntry.ClassName AndAlso Not Strings.Left(_ClassNamespace, 3) = "EA." Then ' we ignore anything outside of the current namespace - And Not Strings.Left(_ClassNamespace, 3) = "EA." 

                                ' now check if the class exists in the relevant hive otherwise ignore
                                If CheckKeyExists(t.FullName, pHive) Then
                                    Debug.Print("Get asemblies for type " & t.FullName)
                                    ' get class information from classname in pHive &  output class detail
                                    ' THIS will output the class and CLSID
                                    outputkeysAndSubKeysWithValues(pHive, "SOFTWARE\CLASSES\" & t.FullName, True)
                                    If clsidtest <> "" Then
                                        Dim _ClassIDLocation As String = IIf(ClassInformation.OS64Bit, cSoftware_Wow6432Node & cBackSlash & cCLSID & cBackSlash & clsidtest,
                                        cSoftwareClasses & cBackSlash & cCLSID & cBackSlash & clsidtest)
                                        Try
                                            Dim clsid_key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey(_ClassIDLocation)
                                            If clsid_key IsNot Nothing Then OutputCLSIDKeyAndSubKeys(IIf(pHive = cHKCU, cHKCU_Root, cHKLM_Root), clsid_key, True)
                                        Catch
                                        End Try
                                    End If
                                End If
                            End If
                        Next
                    Else
                        MsgBox("Sorry not available as unable to load assembly for " & vbCrLf & _filename, vbExclamation, "Assembly Information Not Available")
                    End If

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
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
            MsgBox("Error accessing information " & vbCrLf & "(see detail below for some, but not all, the explanation): - " & vbCrLf & "====================================" _
                & vbCrLf & ex.ToString, vbExclamation, "DLL Detail Not Available")
#End If
        End Try

        '   _GetAssemblies = False

    End Sub

    ' function to check if class key exists
    Private Function CheckKeyExists(pKey As String, pHive As String) As Boolean
        Try
            Dim _keyroot As RegistryKey = Nothing
            Select Case pHive
                Case cHKCU
                    _keyroot = My.Computer.Registry.CurrentUser.OpenSubKey("SOFTWARE\Classes\" & pKey)

                Case cHKLM
                    _keyroot = My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\Classes\" & pKey)

            End Select
            Return _keyroot IsNot Nothing

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return False
    End Function
End Class
