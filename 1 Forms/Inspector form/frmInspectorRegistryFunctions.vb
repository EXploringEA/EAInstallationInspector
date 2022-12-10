Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

' this file contains routines to:
' * read the registry
' * View information on keys
' * export registry entries to create and/or delete the registry entries
'
Partial Friend Class frmInspector

    Private _regfileCreateKeys As RegFileOutput
    Private _regFileDeleteKeys As RegFileOutput

#Region "Output routines"
    Private OutputString As String = ""
    ' routine that outputs the information both to file and concatenates to create an output string
    Sub outputReg(pString As String)
        Try
            If _regfileCreateKeys IsNot Nothing Then _regfileCreateKeys.output(pString)
            OutputString += pString & vbCrLf
            ListOfRegStrings.Add(pString)
            ListOfCreateRegistryEntries.Add(pString)
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
    Sub outputDeleteKey(pString As String)
        Try
            If _regFileDeleteKeys IsNot Nothing Then _regFileDeleteKeys.output(pString)
            ListOfDeleteRegistryEntries.Add(pString)
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
    Sub OutputDefaultKeyValue(pkey_value As String)
        outputReg("@=" & Chr(34) & pkey_value & Chr(34))
    End Sub

    ''' <summary>
    ''' Example: [HKEY_CURRENT_USER\Software\Sparx Systems\EAAddins\E002_32]
    ''' </summary>
    ''' <param name="pRoot"></param>
    ''' <param name="pRootLocation"></param>
    Sub outputKeyName(pRoot As String, pRootLocation As String)
        outputReg("[" & pRoot & cBackSlash & pRootLocation & "]")
        outputDeleteKey("[-" & pRoot & cBackSlash & pRootLocation & "]")
    End Sub
    ''' <summary>
    ''' Example: "RuntimeVersion"="v4.0.30319"
    ''' </summary>
    ''' <param name="pKey"></param>
    ''' <param name="pValue"></param>
    Sub outputKeyValuePair(pKey As String, pValue As String)
        outputReg(Chr(34) & pKey & Chr(34) & "=" & Chr(34) & pValue & Chr(34))
    End Sub


    Private ListOfRegStrings As New ArrayList
    Private ListOfCreateRegistryEntries As New ArrayList
    Private ListOfDeleteRegistryEntries As New ArrayList

    ''' <summary>
    ''' Outputs: Windows Registry Editor Version 5.00
    ''' When we output header we init the array for strings
    ''' </summary>
    Private Sub OutputRegHeader(pHive As String)
        ListOfRegStrings.Clear()
        ListOfCreateRegistryEntries.Clear()
        ListOfDeleteRegistryEntries.Clear()

        'If _regfileCreateKeys IsNot Nothing Then _regfileCreateKeys.close()
        '_regfileCreateKeys = New RegFileOutput("Regfile_" & pHive & "_" & _Classname)
        outputReg(cRegHeader)


        'If _regFileDeleteKeys IsNot Nothing Then _regFileDeleteKeys.close()
        '_regFileDeleteKeys = New RegFileOutput("DeleteRegFile_" & pHive & "_" & _Classname)
        outputDeleteKey(cRegHeader)



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

    Private _ClassNamespace As String = ""

    ''' <summary>
    ''' Handler routine to get information for the currently selected item in the listview (Potentil refactor out)
    ''' Will use the Sparx Entry
    ''' </summary>
    Friend Sub GetRegistryValuesHandler(pMyEntryDetail As AddInEntry)
        Try
            OutputString = ""

            If pMyEntryDetail.ClassName <> "" Then
                _Classname = pMyEntryDetail.ClassName
                _ClassNamespace = pMyEntryDetail.ClassName
                If _ClassNamespace.Contains(".") Then
                    Dim positionOfDot As Integer = Strings.InStr(_ClassNamespace, ".")
                    _ClassNamespace = IIf(positionOfDot > 1, Strings.Left(_ClassNamespace, positionOfDot - 1), "")
                Else
                    Debug.Print("Classname within 2 parts: " & pMyEntryDetail.ClassName)
                End If
            End If


            Select Case pMyEntryDetail.SparxEntry

                Case AddInEntry.cHKCU32

                    OutputRegHeader(AddInEntry.cHKCU32)
                    OutputSparxKey(cHKCU, cSparxKeys32 & cBackSlash & pMyEntryDetail.AddInName)
                    OutputClassCLISDEntry(cHKCU, cSoftwareClasses & cBackSlash & pMyEntryDetail.ClassName) '_classRoot)
                    Dim _ClassIDLocation As String = IIf(ClassInformation.OS64Bit, cSoftware_Node & cBackSlash & cCLSID & cBackSlash & pMyEntryDetail.CLSID,
                         cSoftwareClasses & cBackSlash & cCLSID & cBackSlash & pMyEntryDetail.CLSID)
                    Try
                        Dim clsid_key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey(_ClassIDLocation)
                        If clsid_key IsNot Nothing Then OutputCLSIDKeyAndSubKeys(cHKCU, clsid_key, True)
                    Catch
                    End Try

                Case AddInEntry.cHKLM32

                    If IsElevated = False Then   ' ONLY possible if elevated
                        MsgBox("Cannot process - application needs to run as Administrator")
                        Return
                    End If
                    OutputRegHeader(AddInEntry.cHKLM32)
                    OutputSparxKey(cHKLM, cSparxKeysWOW32 & cBackSlash & pMyEntryDetail.AddInName) ' _SparxRoot)
                    OutputClassCLISDEntry(cHKLM, cSoftwareClasses & cBackSlash & pMyEntryDetail.ClassName) ' _classRoot)
                    Dim _ClassIDLocation As String = IIf(ClassInformation.OS64Bit, cSoftware_Node & cBackSlash & cCLSID & cBackSlash & pMyEntryDetail.CLSID,
                         cSoftwareClasses & cBackSlash & cCLSID & cBackSlash & pMyEntryDetail.CLSID)
                    Try
                        Dim clsid_key As RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey(_ClassIDLocation)
                        If clsid_key IsNot Nothing Then OutputCLSIDKeyAndSubKeys(cHKLM, clsid_key, True)
                    Catch
                    End Try


                Case AddInEntry.cHKCU64
                    OutputRegHeader(AddInEntry.cHKCU64)
                    OutputSparxKey(cHKCU, cSparxKeys64 & cBackSlash & pMyEntryDetail.AddInName) ' _SparxRoot)
                    OutputClassCLISDEntry(cHKCU, cSoftwareClasses & cBackSlash & pMyEntryDetail.ClassName) ' _classRoot)
                    Dim _ClassIDLocation As String = cSoftwareClasses & cBackSlash & cCLSID & cBackSlash & pMyEntryDetail.CLSID
                    Try
                        Dim clsid_key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey(_ClassIDLocation)
                        If clsid_key IsNot Nothing Then OutputCLSIDKeyAndSubKeys(cHKCU, clsid_key, True)
                    Catch
                    End Try


                Case AddInEntry.cHKLM64

                    If IsElevated = False Then   ' ONLY possible if elevated
                        MsgBox("Cannot process - application needs to run as Administrator")
                        Return
                    End If
                    OutputRegHeader(AddInEntry.cHKLM64)
                    OutputSparxKey(cHKLM, cSparxKeys64 & cBackSlash & pMyEntryDetail.AddInName) ' _SparxRoot)
                    OutputClassCLISDEntry(cHKLM, cSoftwareClasses & cBackSlash & pMyEntryDetail.ClassName)
                    Dim _ClassIDLocation As String = cSoftwareClasses & cBackSlash & cCLSID & cBackSlash & pMyEntryDetail.CLSID
                    Try
                        Dim clsid_key As RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey(_ClassIDLocation)
                        If clsid_key IsNot Nothing Then OutputCLSIDKeyAndSubKeys(cHKLM, clsid_key, True)
                    Catch
                    End Try
                Case Else
                    MsgBox("No keys available")

            End Select

            Dim regstrings As New frmDisplayRegistryKeys()
            regstrings.init(ListOfRegStrings, ListOfCreateRegistryEntries, ListOfDeleteRegistryEntries, pMyEntryDetail.ClassName, pMyEntryDetail.SparxEntry)
            regstrings.ShowDialog()


            'MsgBox(OutputString)
        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub

    Private clsidtest As String = ""
    Private Sub outputkeysAndSubKeysWithValues(pHive As String, pRootKey As String)
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

            outputReg(_pathPrefix & cBackSlash & pRootKey & "]") ' KEY
            outputDeleteKey(_deletePathPrefix & cBackSlash & pRootKey & "]") ' KEY
            If keyroot IsNot Nothing Then

                OutputDefaultKeyValue(Trim(keyroot.GetValue("")))

                For Each subkey As String In keyroot.GetSubKeyNames
                    ' check it exists

                    Dim kk As String = pRootKey & "\" & subkey
                    If kk IsNot Nothing Then
                        outputReg(_pathPrefix & cBackSlash & pRootKey & "\" & subkey & "]") 'KEY
                        outputDeleteKey(_deletePathPrefix & cBackSlash & pRootKey & "\" & subkey & "]")
                        Dim sb1 As Microsoft.Win32.RegistryKey = IIf(pHive = cHKCU, My.Computer.Registry.CurrentUser.OpenSubKey(kk), My.Computer.Registry.LocalMachine.OpenSubKey(kk))
                        If sb1 IsNot Nothing Then
                            If kk.Contains("CLSID") Then clsidtest = sb1.GetValue("")
                            OutputDefaultKeyValue(sb1.GetValue("")) ' default value for CLSID is the CLSID
                            If sb1.SubKeyCount > 0 Then
                                For Each kk1 In sb1.GetValueNames
                                    If kk1 = "" Then Continue For
                                    outputKeyValuePair(kk1, sb1.GetValue(kk1))
                                Next
                                outputkeysAndSubKeysWithValues(pHive, kk)
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

    Const cHKCU As String = "CU"
    Const cHKLM As String = "LM"
    ' outputs information with check for asembly - possibly optionally exclude the getassembly to consolidate the routines
    Private Sub OutputCLSIDKeyAndSubKeys(pHive As String, Keyroot As RegistryKey, pGetAssemblies As Boolean)

        Try
            If Keyroot IsNot Nothing Then


                If Keyroot.Name.Contains("{") Then
                    outputKeyName(IIf(pHive = cHKCU, cHKCU_Root, cHKLM_Root), Keyroot.Name) '_ClassIDLocation)
                    For Each keyname In Keyroot.GetValueNames
                        Dim val As String = Trim(Keyroot.GetValue(keyname))
                        Dim val2 As String = val.Replace("\", "\\")
                        If keyname = "" Then
                            OutputDefaultKeyValue(val2)
                        Else
                            outputKeyValuePair(keyname, val2)
                        End If


                        If pGetAssemblies AndAlso keyname.Contains(cCodeBase) Then
                            Dim _filename As String = val2
                            If Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)
                            If _filename <> "" AndAlso _filename <> cNotSet Then
                                _filename = _filename.Replace("/", "\")
                                If File.Exists(_filename) AndAlso Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)
                                getassemblykeys(_filename, pHive)
                                pGetAssemblies = False
                            End If
                        End If
                    Next
                    pGetAssemblies = False


                    For Each subkey As String In Keyroot.GetSubKeyNames
                        Dim sk As RegistryKey = Keyroot.OpenSubKey(subkey)
                        If sk IsNot Nothing Then
                            OutputCLSIDKeyAndSubKeys(pHive, sk, pGetAssemblies)
                        End If

                    Next
                End If
            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
    End Sub
    ' Get assembly information
    Private Sub getassemblykeys(pDLLFilename As String, pHive As String)

        Try
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
                    '                Debug.Print(name.Name & " has following attributes: ")


                    If assembly IsNot Nothing Then
                        Dim abc = assembly.GetExportedTypes()
                        Dim s1 As String = "-------  Referenced assemblies ----------" & vbCrLf
                        s1 += assembly.GetReferencedAssemblies.ToString
                        For Each g In assembly.GetReferencedAssemblies
                            s1 += g.ToString & vbCrLf
                        Next
                        Debug.Print(s1)

                        Dim types As Type() = assembly.GetExportedTypes() ' GetTypes()
                        Dim _ListOfTypes As New ArrayList
                        Dim totalexported As Integer = 0
                        Dim namespaceexported As Integer = 0
                        For Each t As Type In types
                            totalexported += 1
                            If t.FullName.Contains(_ClassNamespace) And Not t.FullName = _Classname Then ' we ignore anything outside of the current namespace

                                ' now check if the class exists in the relevant hive otherwise ignore

                                If CheckKeyExists(t.FullName, pHive) Then
                                    Debug.Print(t.FullName)
                                    namespaceexported += 1
                                    ' get class information from classname in pHive
                                    ' output class detail
                                    ' THIS will output the class and CLSID
                                    outputkeysAndSubKeysWithValues(pHive, "SOFTWARE\CLASSES\" & t.FullName)
                                    If clsidtest <> "" Then


                                        Dim _ClassIDLocation As String = IIf(ClassInformation.OS64Bit, cSoftware_Node & cBackSlash & cCLSID & cBackSlash & clsidtest,
                                        cSoftwareClasses & cBackSlash & cCLSID & cBackSlash & clsidtest)
                                        Try
                                            Dim clsid_key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey(_ClassIDLocation)
                                            If clsid_key IsNot Nothing Then
                                                'TODO cant use this as it's recursive!!!!! write a special routine
                                                OutputCLSIDKeyAndSubKeys(IIf(pHive = cHKCU, cHKCU_Root, cHKLM_Root), clsid_key, False)
                                            End If
                                        Catch
                                        End Try
                                    End If


                                    ' TODO get the classID and output class detail
                                    _ListOfTypes.Add(t)
                                End If
                            End If
                        Next
                        Debug.Print("Total exported = " & totalexported.ToString & " In namespace = " & namespaceexported.ToString)
                        '  MsgBox(s, MsgBoxStyle.OkOnly, "List of public methods for DLL")
                        '     Dim frmMethods As New frmListOfClasses(_ListOfTypes)
                        '    frmMethods.ShowDialog()

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
