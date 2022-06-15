Imports System.Reflection
Imports Microsoft.Win32


Public Class ClassInformation

    Property ClassSource As String = ""
    Property ClassID As String = ""
    Property DLLSource As String = ""
    Property Version As String = ""
    Property Filename As String = ""
    Property Colour As Color = Color.Magenta



    Friend Sub GetClassInformation(pAddInName As String, pSequence As String())

        getClassID(pAddInName, pSequence)

        If ClassID <> "" Then getClassDLLFilename()

        ' If we find no Class found Then the DLL cannot be found
        ' _ClassSource indicates where the class entry is defined, hence the HIVE where we would expect to find class details
        ' NOTE: this may not always be the case, we have seen cases where the library is put in the "wrong place"
        ' Also depending on whether it is a 32-bit or 64-bit machine will determine where the DLL detail is located
    End Sub



    Private Sub getClassID(pAddInName As String, s As String())

        ClassSource = ""
        ClassID = ""

        For Each loc As String In s
            Select Case loc
                Case cHKCU32
                    ClassID = Registry.GetValue(HKCU_Classes & cBackSlash & pAddInName & cBackSlash & cCLSID, "", cNotFound)
                    ClassSource = cHKCU32
                    If ClassID <> "" Then Return
                Case cHKLM32Wow1
                    ClassID = Registry.GetValue(HKLMWow1_Classes & cBackSlash & pAddInName & cBackSlash & cCLSID, "", cNotFound)
                    ClassSource = cHKLM32Wow1
                    If ClassID <> "" Then Return
                Case cHKLM32Wow2
                    ClassID = Registry.GetValue(HKLMWow2_Classes & cBackSlash & pAddInName & cBackSlash & cCLSID, "", cNotFound)
                    ClassSource = cHKLM32Wow2
                    If ClassID <> "" Then Return
                Case cHKLM32
                    ClassSource = cHKLM32
                    ClassID = Registry.GetValue(HKLM_Classes & cBackSlash & pAddInName & cBackSlash & cCLSID, "", cNotFound)
                    If ClassID <> "" Then Return
                Case cHKCU64
                    ClassSource = cHKCU64
                    ClassID = Registry.GetValue(HKCU_Classes & cBackSlash & pAddInName & cBackSlash & cCLSID, "", cNotFound)
                    If ClassID <> "" Then Return
                Case cHKLM64
                    ClassSource = cHKLM64
                    ClassID = Registry.GetValue(HKLM_Classes & cBackSlash & pAddInName & cBackSlash & cCLSID, "", cNotFound)
                    If ClassID <> "" Then Return
                Case Else
                    Debug.Print("Unknown Addin string " & loc)
            End Select

        Next
        Return
    End Sub

    ''' <summary>
    ''' Get the DLL filename which we assume is in the relevant place for HKCU / HKLM - the hive is provided by the caller
    ''' </summary>
    ''' <remarks>There are some instance where the file would not be found as expected - 
    ''' e.g. the COMServer class but don't think this is a likely case have included some checks</remarks>
    Private Sub getClassDLLFilename()

        Try
            ' 1. Use the Select the HIVE where the classID was found and get the information
            Dim _location As String = ""
            DLLSource = cNotFound
            Select Case ClassSource
                Case cHKCU32
                    _location = HKCU_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                    Filename = Registry.GetValue(_location, cCodeBase, cNotSet) 'using the class try to find the DLL path
                    If Filename IsNot Nothing Then
                        DLLSource = cHKCU32
                    ElseIf (Environment.Is64BitOperatingSystem) Then ' check in Wow6432 
                        _location = HKCU_Classes & cBackSlash & cWow6432Node & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                        Filename = Registry.GetValue(_location, cCodeBase, cNotSet)
                        If Filename IsNot Nothing Then
                            DLLSource = cHKCU32Wow
                        Else
                            _location = HKCU_Classes & cBackSlash & cWow6432Node & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                            Filename = Registry.GetValue(_location, "", cNotSet)
                            If Filename <> cNotSet Then DLLSource = "HKCU In ROOT"
                        End If
                    End If

                Case cHKLM32
                    '' location will depend on whether the AddIn is 32-bit or 64-bit
                    'If pAddInInfo.SparxAddinLocation = cHKLM64 Or pAddInInfo.SparxAddinLocation = cHKCU64 Then
                    '    _location = HKLMClasses & cBackSlash & cCLSID & cBackSlash & pClassID & cBackSlash & cInprocServer32
                    'Else
                    ' check HLKM depending on OS
                    _location = If(Environment.Is64BitOperatingSystem, HKLM_Classes & cBackSlash & cWow6432Node & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32,
                                       HKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)
                    ' End If
                    Filename = Registry.GetValue(_location, cCodeBase, cNotSet) 'using the class try to find the DLL path
                    If Filename IsNot Nothing Then ' found
                        DLLSource = cHKLM32
                    Else ' possibility that it's placed in root of key - have seen this with some classes
                        _location = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                        Filename = Registry.GetValue(_location, "", cNotSet)
                        If Filename IsNot Nothing Then DLLSource = "HKLM ROOT KEY"
                    End If

                Case cHKLM32Wow
                    _location = HKLM_Classes & cBackSlash & cWow6432Node & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                    Filename = Registry.GetValue(_location, cCodeBase, cNotSet) 'using the class try to find the DLL path
                    If Filename IsNot Nothing Then
                        DLLSource = cHKLM32Wow
                    Else
                        _location = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                        Filename = Registry.GetValue(_location, "", cNotSet)
                        If Filename IsNot Nothing Then DLLSource = "HKLM ROOT KEY"
                    End If


                Case cHKCU64
                    _location = HKCU_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                    Filename = Registry.GetValue(_location, cCodeBase, cNotSet) 'using the class try to find the DLL path
                    If Filename IsNot Nothing Then
                        DLLSource = cHKCU64
                    Else
                        _location = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                        Filename = Registry.GetValue(_location, "", cNotSet)
                        If Filename IsNot Nothing Then DLLSource = "HKLM ROOT KEY"
                    End If


                Case cHKLM64
                    _location = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                    Filename = Registry.GetValue(_location, cCodeBase, cNotSet) 'using the class try to find the DLL path
                    If Filename IsNot Nothing Then
                        DLLSource = cHKLM64
                    Else
                        _location = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
                        Filename = Registry.GetValue(_location, "", cNotSet)
                        If Filename IsNot Nothing Then DLLSource = "HKLM ROOT KEY"
                    End If

                Case Else
                    ' ERROR
                    Filename = "Error"
                    DLLSource = "Error"
            End Select
            If Filename <> "Error" Then getDLLAssembly()

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return
    End Sub



    Private Sub getDLLAssembly()

        Dim _filename As String = Filename
        If _filename IsNot Nothing And _filename <> cNotSet Then
            If Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)
            Try
                Dim ass As AssemblyName = AssemblyName.GetAssemblyName(_filename)
                Version = ass.Version.ToString
            Catch ex As Exception
                Version = "Unable to determine"
            End Try
            If _filename <> "" Then _filename = _filename.Replace("/", "\")
            Filename = _filename
            ' now depending on where items were found flag accordingly
            If ClassSource <> DLLSource Then
                Colour = Color.Red
            ElseIf Filename = cNotSet Then
                Colour = Color.Yellow
            Else
                Colour = If(DLLexists(Filename), Color.LightGreen, Color.Cyan)  ' File does not exist
            End If

        End If

    End Sub


    ''' <summary>
    ''' Dls the lexists.
    ''' </summary>
    ''' <param name="pFilePath">DLL Filename path.</param>
    ''' <returns>True if exists else false</returns>
    Private Function DLLexists(pFilePath As String) As Boolean
        Try
            ' remove file from front of string
            If Strings.InStr(pFilePath, "file:///") Then
                Dim myNewFN As String = Strings.Right$(pFilePath, Len(pFilePath) - 8)
                If System.IO.File.Exists(myNewFN) Then Return True
            ElseIf Strings.InStr(pFilePath, ":\") Or Strings.InStr(pFilePath, ":/") Then
                If System.IO.File.Exists(pFilePath) Then Return True
            End If

        Catch ex As Exception

        End Try
        Return False
    End Function

    ' used by treeviewer
    ' function to get class information from the registry location and classID
    ' HIVE NEEDS TO be HKCU32, HKCU64
    Friend Shared Function OLDgetClassInformation(pHIVE As String, pID As String) As ClassRegistryInformation
        Dim myClassInfo As New ClassRegistryInformation
        Try

            Dim _keylocation As String = cNotSet
            Select Case pHIVE
                Case cHKCU32
                    _keylocation = HKCU_Classes & cBackSlash & cCLSID & cBackSlash & pID

                Case cHKLM32
                    _keylocation = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & pID

                Case cHKLM32Wow
                    _keylocation = HKLMWow2_Classes & cBackSlash & cCLSID & cBackSlash & pID

                Case cHKLM32Wow1
                    _keylocation = HKLMWow1_Classes & cBackSlash & cWow6432Node & cBackSlash & cCLSID & cBackSlash & pID

                Case cHKLM32Wow2
                    _keylocation = HKLMWow2_Classes & cBackSlash & cWow6432Node & cBackSlash & cCLSID & cBackSlash & pID


                Case cHKCU64
                    _keylocation = HKCU_Classes & cBackSlash & cCLSID & cBackSlash & pID
                Case cHKLM64
                    _keylocation = HKLM_Classes & cBackSlash & cCLSID & cBackSlash & pID
                Case Else
                    Debug.Print("HIVE - " & pHIVE)
            End Select
            myClassInfo.HIVE = pHIVE
            If _keylocation = cNotSet Then
                myClassInfo.CodeBase = cNotSet
                myClassInfo.Assembly = cNotSet
                myClassInfo.ClassName = cNotSet
                myClassInfo.RunTimeVersion = cNotSet
                myClassInfo.ProgID = cNotSet

            Else
                myClassInfo.CodeBase = Registry.GetValue(_keylocation & cBackSlash & cInprocServer32, cCodeBase, cNotSet)
                myClassInfo.Assembly = Registry.GetValue(_keylocation & cBackSlash & cInprocServer32, cAssembly, cNotSet)
                myClassInfo.ClassName = Registry.GetValue(_keylocation & cBackSlash & cInprocServer32, cClass, cNotSet)
                myClassInfo.RunTimeVersion = Registry.GetValue(_keylocation & cBackSlash & cInprocServer32, cRuntimeVersion, cNotSet)
                myClassInfo.ProgID = Registry.GetValue(_keylocation & cBackSlash & cProgID, "", cNotSet)

            End If

        Catch ex As Exception
#If DEBUG Then
            Debug.Print(ex.ToString)
#End If
        End Try
        Return myClassInfo

    End Function



End Class


Friend Class ClassRegistryInformation
    Friend HIVE As String = ""
    Friend ClassName As String = ""
    Friend Assembly As String = ""
    Friend CodeBase As String = ""
    Friend RunTimeVersion As String = ""
    Friend ProgID As String = ""
End Class

