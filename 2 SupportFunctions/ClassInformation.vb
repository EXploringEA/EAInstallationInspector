' Copyright (C) 2022 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================
'

Imports System.IO
Imports System.Reflection
Imports Microsoft.Win32

''' <summary>
''' Class to capture information about a specific class entry - treat multiple class entriese.g. 32-bit 64-bti as separate entries
''' Only looks for information for a single AddIn based on it's Sparx key location
''' </summary>
Public Class ClassInformation

#Region "Current user"


    ' Registry class locations
    ''' <summary>
    ''' HKCU Classes for 
    ''' * 32-bit on 32-bit OS
    ''' * 64-bit on 64-bit OS
    ''' </summary>'
    '''  

    Private Const cHKCU_ClassesCLSID As String = "HKEY_CURRENT_USER\SOFTWARE\Classes\CLSID"

    '''' <summary>
    '''' HKCU Classes for 
    '''' * 32-bit on 64-bit OS
    '''' </summary>
    Private Const cHKCUWOW_ClassesCLSID As String = "HKEY_CURRENT_USER\SOFTWARE\Classes\Wow6432Node\CLSID"
#End Region

#Region "Local machine"

    ''' <summary>
    ''' HKLM Classes
    ''' * 32-bit on 32-bit OS
    ''' * 64-bit on 64-bit OS
    ''' </summary>
    '''  

    Private Const cHKLM_ClassesCLSID As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID"
    ''' <summary>
    ''' HKLM Classes for 
    ''' * 32-bit on 64-bit OS
    ''' </summary>
    Private Const cHKLMWow_ClassesCLSID As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node\CLSID"
    Private Const cHKLMWow_ClassesCLSID2 As String = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes\CLSID"

    Private Const cHKCR_ClassesRoot As String = "HKEY_CLASSES_ROOT\"
    Private Const cMainClass As String = "main class" ' CI
#End Region


    Private Const cImplementedCategories As String = "Implemented Categories"
    Private Const cProgID As String = "ProgID"
    Private Const cInprocServer32 As String = "InprocServer32"
    Private Const cWow6432Node As String = "Wow6432Node"

    ' Friend cClass As String = "Class"
    Private Const cAssembly As String = "Assembly"
    Private Const cRuntimeVersion As String = "RuntimeVersion"
    Private Const cThreadingModel As String = "ThreadingModel"


    'Classification of issues
    ' use booleans to flag 
    Private SparxKeyExists As Boolean = False
    Private ClassIDExists As Boolean = False
    Private MatchedHives As Boolean = False
    Private DLLExistsInExpectedLocation As Boolean = False


    Property ClassSource As String = ""


    ' Information from Spaxr registry keys
    ' AddIn name - the name of the Key in Sparx entry
    Property AddInClassName As String = ""
    ' Location of the AddIn key within Sparx keys
    Property AddInSource As String = ""

    'Property LibraryType As String = "" ' 64-bit / 32-bit /both


    ' Class information
    ' ClassID - captured from HKCR - note Class ID will be the same for the same named class even if multiple verssion 32-bit/ 64-bit
    Property ClassID As String = ""
    Property DLLSource As String = ""
    Property DLLVersion As String = ""
    ' run time version for DLL
    Property RunTimeVersion As String = ""




    Property DisplayFilename As String
        Get
            Return cleanFilename(Filename)
        End Get
        Set(value As String)

        End Set
    End Property
    Property ThreadingModel As String = ""
    Property AssemblyAsString As String = ""

    ' Location of key in registry
    Property RegistryLocation As String = ""

    Friend Shared OS64Bit As Boolean = False


    ' Private testnames As Boolean = True
    ' Filename as captured from HKCR
    Private Property Filename As String = ""
    Private AddInClassNameMatch As Boolean = False
    '  Private AddInCheckName As String = ""
    Private IntegrationAddIn As Boolean = False
    ''' <summary>
    ''' Constructor used to save Sparx information and get the class ID
    ''' </summary>
    ''' <param name="pAddInClassName"></param>
    ''' <param name="pAddInSource"></param>
    Sub New(pAddInClassName As String, pAddInSource As String) ', ptest As Boolean)
        ' testnames = ptest
        AddInClassName = pAddInClassName
        AddInSource = pAddInSource
        If Environment.Is64BitOperatingSystem Then OS64Bit = True
        If AddInSource <> "" Then SparxKeyExists = True
        getClassInfo()
        Debug.Print("Addin " & AddInClassName & " Class ID = " & ClassID)
    End Sub

    ' we have the addin name and location of Sparx key
    ' Note that the Sparx Key does not mean that the class will be registered correctly
    ' For 64-bit OS - 32bit classes are stored under WOW6432Node and 64-bit under HKCU,HKLM
    ' 
    ' Interesting note that there are 
    ' \HKEY_CURRENT_USER\SOFTWARE\WOW6432Node\ -> Not a lot under this so no need to check
    ' \HKEY_CURRENT_USER\SOFTWARE\Classes\WOW6432Node\CLSID
    ' The follow are mirrors?
    ' \HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node\CLSID
    ' \HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes\CLSID

    ' for 64-bit on 64-bit os and 32-bit on 32-bit os
    ' HKEY_CURRENT_USER\SOFTWARE\Classes\CLSID
    ' HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID
    ' so to search 
    ' returns classID found from HKCR - the classID is then used to search for the relevant key elsewhere

    ''' <summary>
    ''' Returns the classID based on the Classname from the HKEY_CLASSES_ROOT
    ''' This hive is a merge of the others so will produce a classID for a classname regardless if it is 32-bit or 64-bit
    ''' Within HKCR there are CLSID entries and on 64-bit systems WOW6432NODE CLSID
    ''' The classID is then used to check in the relevant hives starting with that where the Sparx entry is defined and then if not found checked in the other HIVE
    ''' NOTE: 32-biut and 64-bit are treated separately
    ''' </summary>
    ''' <remarks>
    ''' The logic steps are:
    ''' * Get the classID from HKCR
    ''' * Based on the AddInSource HIVE - check the relative HIVE for a class entry (using ClassID) and check the classname matches the AddInName
    ''' * If the entry is correct then check if there is also an entry in the alternative HIVE
    ''' * If the entry is absent then look for the full entry
    ''' </remarks>
    Private Sub getClassInfo()
        ' if we have HKCU and HKLM that's fine as the current user will take precendence
#If classtrace Then
        logger.logger.write(vbCrLf & "=== getClassID ==== " & AddInSource & ":" & AddInClassName)
#End If

        ClassSource = ""
        ClassID = ""
        ' for 64-bit OS we would expect the class information to be in the relevant hive - we could get from HKCR
        ' so use the Spaxr AddIn location PLUS the OS to identify where we would expect to find or not - get the information from 
        ' get all information from HKCR
        ' 64-bit OS
        ClassSource = cNotFound
        Debug.Print(AddInClassName)

        ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInClassName & cBackSlash & cCLSID, "", cNotFound)
        If ClassID Is Nothing Then Return  ' no action possible

        Select Case AddInSource

            Case AddInEntry.cHKCU32   '32-bit add in
                ' check if we have a HKCU32 entry - expected
                Dim HKCU32Key As KeyInfo = CheckKeyInfo(ClassID, KeyLocation(AddInEntry.cHKCU32, ClassID))
                Dim HKLM32key As KeyInfo = CheckKeyInfo(ClassID, KeyLocation(AddInEntry.cHKLM32, ClassID))

                If HKCU32Key.Found Then ' check in primary location for 32-bit CU
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKCU32
                    MatchedHives = True
                    Filename = HKCU32Key.DLLFilename
                    DLLVersion = HKCU32Key.DLLVersion
                    DLLSource = AddInEntry.cHKCU32
                    DLLExistsInExpectedLocation = HKCU32Key.DLLExistsInLocation
                    RegistryLocation = HKCU32Key.DLLRegistryLocation
                    RunTimeVersion = HKCU32Key.DLLRunTimeVersion
                    ThreadingModel = HKCU32Key.DLLThreadingModel
                    AssemblyAsString = HKCU32Key.AssemblyAsString
                    AddInClassNameMatch = CheckAddInClassNamesMatch(AddInClassName, HKCU32Key.ClassName)

                    If HKCU32Key.DefaultKey.ToLower = cMainClass Then
                        DLLSource = "Default key:" & AddInEntry.cHKCU32
                        IntegrationAddIn = True
                    End If

                    If HKLM32key.Found Then 'we also found an entry in HKLM
                        ClassSource += "," & AddInEntry.cHKLM32
                    End If

                ElseIf HKLM32key.Found Then ' No HKCU but HKLM found there is it in 32-bit HKLM
                    MatchedHives = False
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKLM32
                    DLLExistsInExpectedLocation = False ' HKLM32key.DLLExistsInLocation
                    Filename = HKLM32key.DLLFilename ' this is the primary filename saved for later
                    DLLSource = AddInEntry.cHKLM32
                    DLLVersion = HKLM32key.DLLVersion
                    RegistryLocation = HKLM32key.DLLRegistryLocation
                    RunTimeVersion = HKLM32key.DLLRunTimeVersion
                    ThreadingModel = HKLM32key.DLLThreadingModel
                    AssemblyAsString = HKLM32key.AssemblyAsString
                    AddInClassNameMatch = CheckAddInClassNamesMatch(AddInClassName, HKLM32key.ClassName)
                End If

            Case AddInEntry.cHKLM32

                Dim HKCU32Key As KeyInfo = CheckKeyInfo(ClassID, KeyLocation(AddInEntry.cHKCU32, ClassID))
                Dim HKLM32key As KeyInfo = CheckKeyInfo(ClassID, KeyLocation(AddInEntry.cHKLM32, ClassID))


                If HKLM32key.Found Then
                    Filename = HKLM32key.DLLFilename
                    DLLExistsInExpectedLocation = True
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKLM32
                    MatchedHives = True
                    DLLVersion = HKLM32key.DLLVersion
                    DLLExistsInExpectedLocation = True
                    DLLSource = AddInEntry.cHKLM32
                    RegistryLocation = HKLM32key.DLLRegistryLocation
                    RunTimeVersion = HKLM32key.DLLRunTimeVersion
                    ThreadingModel = HKLM32key.DLLThreadingModel
                    AssemblyAsString = HKLM32key.AssemblyAsString
                    AddInClassNameMatch = CheckAddInClassNamesMatch(AddInClassName, HKLM32key.ClassName)

                    If HKLM32key.DefaultKey.ToLower = cMainClass Then
                        DLLSource = "Default key:" & AddInEntry.cHKLM32
                        IntegrationAddIn = True
                    End If

                    If HKCU32Key.Found Then
                        ClassSource += "," & AddInEntry.cHKCU32
                    End If

                ElseIf HKCU32Key.Found Then
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKCU32
                    MatchedHives = False
                    DLLExistsInExpectedLocation = False
                    Filename = HKCU32Key.DLLFilename ' this is the primary filename saved for later
                    DLLSource = AddInEntry.cHKCU32
                    DLLVersion = HKCU32Key.DLLVersion
                    RegistryLocation = HKCU32Key.DLLRegistryLocation
                    RunTimeVersion = HKCU32Key.DLLRunTimeVersion
                    ThreadingModel = HKCU32Key.DLLThreadingModel
                    AssemblyAsString = HKCU32Key.AssemblyAsString
                    AddInClassNameMatch = CheckAddInClassNamesMatch(AddInClassName, HKCU32Key.ClassName)
                End If

                    '64-bit addins
            Case AddInEntry.cHKCU64
                Dim HKCUKey As KeyInfo = CheckKeyInfo(ClassID, KeyLocation(AddInEntry.cHKCU64, ClassID))
                Dim HKLMkey As KeyInfo = CheckKeyInfo(ClassID, KeyLocation(AddInEntry.cHKLM64, ClassID))

                If HKCUKey.Found Then
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKCU64
                    MatchedHives = True
                    Filename = HKCUKey.DLLFilename
                    DLLSource = AddInEntry.cHKCU64
                    DLLExistsInExpectedLocation = True
                    DLLVersion = HKCUKey.DLLVersion
                    RegistryLocation = HKCUKey.DLLRegistryLocation
                    RunTimeVersion = HKCUKey.DLLRunTimeVersion
                    ThreadingModel = HKCUKey.DLLThreadingModel
                    AssemblyAsString = HKCUKey.AssemblyAsString

                    AddInClassNameMatch = CheckAddInClassNamesMatch(AddInClassName, HKCUKey.ClassName)

                    If HKCUKey.DefaultKey.ToLower = cMainClass Then
                        DLLSource = "Default key:" & AddInEntry.cHKCU64
                        IntegrationAddIn = True
                    End If


                    If HKLMkey.Found Then
                        ClassSource += "," & AddInEntry.cHKLM64
                        '     DLLSource += "and MORE"
                    End If

                ElseIf HKLMkey.Found Then
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKLM64
                    MatchedHives = False

                    Filename = HKLMkey.DLLFilename ' this is the primary filename saved for later
                    DLLSource = AddInEntry.cHKLM64
                    DLLExistsInExpectedLocation = False
                    DLLVersion = HKLMkey.DLLVersion
                    RegistryLocation = HKLMkey.DLLRegistryLocation
                    RunTimeVersion = HKLMkey.DLLRunTimeVersion
                    ThreadingModel = HKLMkey.DLLThreadingModel
                    AssemblyAsString = HKLMkey.AssemblyAsString

                    AddInClassNameMatch = CheckAddInClassNamesMatch(AddInClassName, HKLMkey.ClassName)

                End If

            Case AddInEntry.cHKLM64

                Dim HKCUKey As KeyInfo = CheckKeyInfo(ClassID, KeyLocation(AddInEntry.cHKCU64, ClassID))
                Dim HKLMkey As KeyInfo = CheckKeyInfo(ClassID, KeyLocation(AddInEntry.cHKLM64, ClassID))

                If HKLMkey.Found Then
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKLM64
                    MatchedHives = True
                    Filename = HKLMkey.DLLFilename
                    DLLSource = AddInEntry.cHKLM64
                    DLLVersion = HKLMkey.DLLVersion
                    DLLExistsInExpectedLocation = True
                    RegistryLocation = HKLMkey.DLLRegistryLocation
                    RunTimeVersion = HKLMkey.DLLRunTimeVersion
                    ThreadingModel = HKLMkey.DLLThreadingModel
                    AssemblyAsString = HKLMkey.AssemblyAsString
                    AddInClassNameMatch = CheckAddInClassNamesMatch(AddInClassName, HKLMkey.ClassName)

                    If HKLMkey.DefaultKey.ToLower = cMainClass Then
                        DLLSource = "Default key:" & AddInEntry.cHKLM64
                        IntegrationAddIn = True
                    End If

                    If HKCUKey.Found Then
                        ClassSource += "," & AddInEntry.cHKCU64
                        '   DLLSource += "and MORE"
                    End If

                ElseIf HKCUKey.Found Then

                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKCU64
                    MatchedHives = False

                    Filename = HKCUKey.DLLFilename ' this is the primary filename saved for later
                    DLLSource = AddInEntry.cHKCU64
                    DLLExistsInExpectedLocation = False
                    DLLVersion = HKCUKey.DLLVersion

                    RegistryLocation = HKCUKey.DLLRegistryLocation
                    RunTimeVersion = HKCUKey.DLLRunTimeVersion
                    ThreadingModel = HKCUKey.DLLThreadingModel
                    AssemblyAsString = HKCUKey.AssemblyAsString

                    AddInClassNameMatch = CheckAddInClassNamesMatch(AddInClassName, HKCUKey.ClassName)

                End If

            Case Else
                MsgBox("Unable to get class ID for " & AddInClassName & " Source " & AddInSource)
        End Select
        Return
    End Sub

    'Sub setInfoValues(True, AddInEntry.cHKLM64,
    '    ClassIDExists = True
    '    ClassSource = AddInEntry.cHKLM64
    '    MatchedHives = True
    '    Filename = HKLMkey.Filename
    '    DLLSource = "Default key:" & AddInEntry.cHKLM64
    '    DLLVersion = HKLMkey.DLLVersion
    '    RegistryLocation = HKLMkey.RegistryLocation
    '    RunTimeVersion = HKLMkey.RunTimeVersion
    '    ThreadingModel = HKLMkey.ThreadingModel
    '    AssemblyAsString = HKLMkey.AssemblyAsString
    '    AddInNameMatch = CheckAddInNamesMatch(AddInName, HKLMkey.AddInName)
    'End Sub
    Private Function GUIDsMatch(pGUID1 As String, pGUID2 As String) As Boolean
        Try
            Return (String.Compare(Trim(pGUID1), Trim(pGUID2)) = 0)
        Catch ex As Exception

        End Try
        Return False

    End Function
    Friend Shared Function CheckAddInClassNamesMatch(pAdd1 As String, pAdd2 As String) As Boolean
        Try
            If pAdd1 Is Nothing Then Return False
            If pAdd2 Is Nothing Then Return False
            Dim a1 As String = Trim(pAdd1.ToLower)
            Dim a2 As String = Trim(pAdd2.ToLower)
            '  Return a2.Contains(a1)
            Return a1 = a2
        Catch ex As Exception

        End Try
        Return False

    End Function
    ' ----- CLASSID check for DLL file functions
    ' checkHKCU32
    ' checkHKLM32
    ' checkHKCU
    ' checkHKLM


    Friend Shared Function KeyLocation(pKeySource As String, ClassID As String) As String
        Dim _location As String = ""
        Try
            Select Case pKeySource
                Case AddInEntry.cHKCU32
                    _location = IIf(OS64Bit,
                               cHKCUWOW_ClassesCLSID & cBackSlash & ClassID & cBackSlash,
                               cHKCU_ClassesCLSID & cBackSlash & ClassID & cBackSlash)
                Case AddInEntry.cHKLM32
                    _location = IIf(OS64Bit,
                                 cHKLMWow_ClassesCLSID & cBackSlash & ClassID & cBackSlash,
                                 cHKLM_ClassesCLSID & cBackSlash & ClassID & cBackSlash)
                Case AddInEntry.cHKCU64
                    _location = cHKCU_ClassesCLSID & cBackSlash & ClassID & cBackSlash
                Case AddInEntry.cHKLM64
                    _location = cHKLM_ClassesCLSID & cBackSlash & ClassID & cBackSlash

            End Select
        Catch ex As Exception

        End Try
        Return _location
    End Function




    Friend Shared Function CheckKeyInfo(pClassID As String, pLocation As String) As KeyInfo
        Dim _info As New KeyInfo
        Try
            If pLocation <> "" Then
                _info.ClassRootLocation = pLocation
                Dim prefix As String = ""
                _info.ClassName = Registry.GetValue(pLocation, "", "")
                Dim _filename As String = Registry.GetValue(pLocation & cInprocServer32, cCodeBase, "") 'using the class try to find the DLL path
                If _filename Is Nothing Then Return _info
                _info.DLLFilename = _filename
                _info.Found = False
                If _filename = "" Then ' see if there is a default key - possible addin
                    _info.DefaultKey = Registry.GetValue(pLocation, "", "")
                    If _info.DefaultKey <> "" Then prefix = "Default key: "
                    _info.DLLFilename = Registry.GetValue(pLocation & cInprocServer32, "", "")
                    _info.ClassID = pClassID
                    _info.DLLExistsInLocation = False
                    _info.DLLVersion = getDLLVersionFromAssembly(cleanFilename(_filename))
                    _info.Found = True
                    _info.DLLRunTimeVersion = Registry.GetValue(pLocation & cInprocServer32, cRuntimeVersion, cNotSet)
                    _info.DLLThreadingModel = Registry.GetValue(pLocation & cInprocServer32, cThreadingModel, cNotSet)
                    _info.ClassName = Registry.GetValue(pLocation & "VersionIndependentProgID", "", "")
                Else
                    If (_filename <> "") Then
                        _info.Found = True
                        _info.ClassID = pClassID
                        _info.DLLExistsInLocation = True
                        _info.DLLVersion = getDLLVersionFromAssembly(cleanFilename(_filename))
                        _info.DLLRegistryLocation = pLocation & cInprocServer32
                        _info.DLLRunTimeVersion = Registry.GetValue(pLocation & cInprocServer32, cRuntimeVersion, cNotSet)
                        _info.DLLThreadingModel = Registry.GetValue(pLocation & cInprocServer32, cThreadingModel, cNotSet)
                        _info.AssemblyAsString = Registry.GetValue(pLocation & cInprocServer32, cAssembly, cNotSet)
                    End If
                End If

            End If ' location
        Catch ex As Exception

        End Try
        Return _info
    End Function




    ''' <summary>
    ''' Get information for DLL file
    ''' * Assembly
    ''' * Version
    ''' 
    ''' </summary>
    Private Shared Function getDLLVersionFromAssembly(_filename As String) As String
        '   Dim _filename As String = cleanFilename(Filename)
        If _filename IsNot Nothing AndAlso _filename <> "" AndAlso _filename <> cNotSet Then
#If classtrace Then
            logger.logger.write(" getDLLAssembly " & _filename)
#End If
            If File.Exists(_filename) Then
                If Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)
                Try
                    Dim ass As AssemblyName = AssemblyName.GetAssemblyName(_filename)
                    Return ass.Version.ToString
                Catch ex As Exception
                    Return "Unable to determine"
                End Try
            End If
        End If
        Return ""
    End Function


    ' was used to ensure the comparison between filenames was fair - however all entries that are checked are from registry so unless we need to 
    ' use the address other than checking OK to leave raw
    Private Shared Function cleanFilename(_Filename As String) As String

        If Strings.Left(_Filename, fileprefixlength) = cFilePrefix Then _Filename = Strings.Right(_Filename, _Filename.Length - fileprefixlength)
        If _Filename <> "" Then _Filename = _Filename.Replace("/", "\")
        Return _Filename
    End Function

    ''' <summary>
    ''' does the DLL for the given filename exist
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

    ' function to set line colours
    ' uses the properties of the entry to define the colour that should be set
    ' precendence of colours
    ' SPARX present
    ' CLASSID
    ' Mismatched HIVES
    ' DLL exists
    ' DLL wrong location

    Friend Function getLineColour() As Color
        Dim c As Color = Color.White
        Try
            If Not SparxKeyExists Then Return Color.FromArgb(255, 0, 0)  'Color.Red
            If Not ClassIDExists Then Return Color.FromArgb(255, 105, 180)
            If Not DLLExistsInExpectedLocation Or Not MatchedHives Then Return Color.FromArgb(186, 252, 186)
            If Not MatchedHives Then Return Color.FromArgb(235, 250, 180)
            If IntegrationAddIn = True Then Return Color.FromArgb(255, 255, 54)
            If Not AddInClassNameMatch Then Return Color.FromArgb(255, 230, 153)
            Return Color.FromArgb(0, 255, 0) ' All OK

        Catch ex As Exception

        End Try
        Return Color.Red
    End Function

End Class

Class KeyInfo

    Public Found As Boolean = False ' entry for classid found
    Public ClassID As String = "" ' classid
    Public ClassName As String = "" ' name of AddIn Class
    Public ClassRootLocation As String = "" ' location for key root 
    Public DLLFilename As String = "" ' DLL filename
    ' .net information
    Public DLLRunTimeVersion As String = "" '.net
    Public DLLThreadingModel As String = "" '.net
    ' Assembly information
    Public AssemblyAsString As String = ""
    Public DLLRegistryLocation As String = "" ' location of assembly key
    Public DLLExistsInLocation As Boolean = False
    Public DLLVersion As String
    Public DefaultKey As String = ""



End Class



