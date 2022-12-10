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
    Private DLLPathExists As Boolean = False
    Private MismatchedHives As Boolean = False
    Private DLLExistsInSpecifiedLocation As Boolean = False


    Property ClassSource As String = ""


    ' Information from Spaxr registry keys
    ' AddIn name - the name of the Key in Sparx entry
    Property AddInName As String = ""
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
    ' Filename as captured from HKCR
    Private Property Filename As String = ""
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
    ''' <summary>
    ''' Constructor used to save Sparx information and get the class ID
    ''' </summary>
    ''' <param name="pAddInName"></param>
    ''' <param name="pAddInSource"></param>
    Sub New(pAddInName As String, pAddInSource As String)
        AddInName = pAddInName
        AddInSource = pAddInSource
        If Environment.Is64BitOperatingSystem Then OS64Bit = True
        If AddInSource <> "" Then SparxKeyExists = True
        getClassID()
        Debug.Print("Addin " & AddInName & " Class ID = " & ClassID)
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
    Private Sub getClassID()
        'TODO - review the conditions
        ' REVIEW logic of secondary checks
        ' is what taskes precendence
        ' if we have HKCU and HKLM that's fine as the current user will take precendence


        ClassSource = ""
        ClassID = ""
        ' for 64-bit OS we would expect the class information to be in the relevant hive - we could get from HKCR
        ' so use the Spaxr AddIn location PLUS the OS to identify where we would expect to find or not - get the information from 
        ' get all information from HKCR
        ' 64-bit OS
        ClassSource = cNotFound
        Debug.Print(AddInName)
        Select Case AddInSource
               '32-bit add in
            Case AddInEntry.cHKCU32
                ' check if we have a HKCU32 entry - expected

                ' check if we have a HKCU32 entry - expected
                ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)

                If ClassID IsNot Nothing And ClassID <> "" Then ' we know we have a copy lets checkif it exists in the required HIVE
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKCR
                    ' now check if in HKCU as expected
                    Dim CLSID = Registry.GetValue(cHKCU_Classes & cBackSlash & AddInName & cBackSlash & cCLSID, "", cNotFound)
                    If (CLSID IsNot Nothing And CLSID <> "") Then
                        ClassSource = AddInEntry.cHKCU32
                        'GetHKCU32 Class Information
                        If CheckHKCU32(True) <> "" Then
                            ClassSource = AddInEntry.cHKCU32
                            CheckHKLM32(False)
                        End If

                        '?? Also check if there is a HKLM entry
                        Exit Select
                    End If
                    CLSID = Registry.GetValue(cHKLM_Classes & cBackSlash & AddInName & cBackSlash & cCLSID, "", cNotFound)
                    If (CLSID IsNot Nothing And CLSID <> "") Then
                        'GetHKLM32 Class Information
                        If CheckHKLM32(True) <> "" Then ClassSource = AddInEntry.cHKLM32
                        Exit Select
                    End If

                End If

            Case AddInEntry.cHKLM32
                ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)

                If ClassID IsNot Nothing And ClassID <> "" Then ' we know we have a copy lets checkif it exists in the required HIVE
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKCR
                    ' now check if in HKCU as expected
                    Dim CLSID = Registry.GetValue(cHKLM_Classes & cBackSlash & AddInName & cBackSlash & cCLSID, "", cNotFound)
                    If (CLSID IsNot Nothing And CLSID <> "") Then
                        ClassSource = AddInEntry.cHKLM32
                        'GetHKLM32 Class Information
                        If CheckHKLM32(True) <> "" Then
                            ClassSource = AddInEntry.cHKLM32
                            CheckHKCU32(False)
                        End If

                        Exit Select
                    End If

                    CLSID = Registry.GetValue(cHKCU_Classes & cBackSlash & AddInName & cBackSlash & cCLSID, "", cNotFound)
                    If (CLSID IsNot Nothing And CLSID <> "") Then

                        'GetHKCU32 Class Information
                        If CheckHKCU32(True) <> "" Then ClassSource = AddInEntry.cHKCU32
                        Exit Select
                    End If
                End If



                    '64-bit addins
            Case AddInEntry.cHKCU64
                ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)

                If ClassID IsNot Nothing And ClassID <> "" Then ' we know we have a copy lets checkif it exists in the required HIVE
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKCR64
                    ' now check if in HKCU as expected
                    Dim CLSID = Registry.GetValue(cHKLM_Classes & cBackSlash & AddInName & cBackSlash & cCLSID, "", cNotFound)
                    If (CLSID IsNot Nothing And CLSID <> "") Then
                        ClassSource = AddInEntry.cHKCU64
                        'GetHKCU Class Information
                        If CheckHKCU(True) <> "" Then ClassSource = AddInEntry.cHKCU64
                        Exit Select
                    End If

                    CLSID = Registry.GetValue(cHKLM_Classes & cBackSlash & AddInName & cBackSlash & cCLSID, "", cNotFound)
                    If (CLSID IsNot Nothing And CLSID <> "") Then
                        'GetHKLM Class Information
                        If CheckHKLM(True) <> "" Then ClassSource = AddInEntry.cHKLM64
                        Exit Select
                    End If
                End If


            Case AddInEntry.cHKLM64
                ClassID = Registry.GetValue(cHKCR_ClassesRoot & AddInName & cBackSlash & cCLSID, "", cNotFound)


                If ClassID IsNot Nothing And ClassID <> "" Then ' we know we have a copy lets checkif it exists in the required HIVE
                    ClassIDExists = True
                    ClassSource = AddInEntry.cHKCR
                    ' now check if in HKCU as expected
                    Dim CLSID = Registry.GetValue(cHKLM_Classes & cBackSlash & AddInName & cBackSlash & cCLSID, "", cNotFound)
                    If (CLSID IsNot Nothing And CLSID <> "") Then
                        ClassSource = AddInEntry.cHKLM64
                        'GetHKLM Class Information
                        If CheckHKLM(True) <> "" Then ClassSource = AddInEntry.cHKLM64
                        Exit Select
                    End If

                    CLSID = Registry.GetValue(cHKCU_Classes & cBackSlash & AddInName & cBackSlash & cCLSID, "", cNotFound)
                    If (CLSID IsNot Nothing And CLSID <> "") Then

                        'GetHKCU Class Information
                        If CheckHKCU(True) <> "" Then ClassSource = AddInEntry.cHKCU64
                        Exit Select
                    End If
                End If

            Case Else
                MsgBox("Unable to get class ID for " & AddInName & " Source " & AddInSource)
        End Select
        Return
    End Sub
    ' ----- CLASSID check for DLL file functions


    ' checkHKCU32
    ' checkHKLM32
    ' checkHKCU
    ' checkHKLM

    ' each function gets the filename from the expected location derived from the classID entry and compares with the 
    ' get location if present
    ' * check for filename of dll
    ' * if dll filename exists then get dll details
    ' * then also check if there is a other location DLL HLCU32 <> HKLM32 or HKCU <> HKLM
    ' * 
    ' routines return location found or "" if not found

    ' * Primary check - used to get all information for CLassID
    ' * Secondary check - used to check if filename for the location is the same as existing - if so ...
    ' TODO check if the secondary is a reasonable check


    ''' <summary>
    ''' Check if there is an HKCU entry the current classID
    ''' </summary>
    ''' <param name="pPrimaryCheck">used to flag that this is a secondary check - so we may already have an entry</param>
    ''' <returns>HKCR if yes else ""</returns>
    Private Function CheckHKCU32(pPrimaryCheck As Boolean) As String
        Dim _location As String = IIf(OS64Bit, cHKCUWOW_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32,
                               cHKCU_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)
        If _location <> "" Then
            Dim prefix As String = ""
            ' get name of DLL based on location
            Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path
            If _filename Is Nothing Then
                Return ""
            Else
                If _filename = "" Then ' see if there is a default key
                    _filename = Registry.GetValue(_location, "", "")
                    If _filename <> "" Then prefix = "Default key: "
                Else

                    '1st pass we check we have a filename 
                    If pPrimaryCheck Then
                        If (_filename <> "") Then
                            Filename = _filename ' this is the primary filename
                            DLLSource = AddInEntry.cHKCU32
                            getDLLAssembly()
                            RegistryLocation = _location
                            RunTimeVersion = Registry.GetValue(RegistryLocation, cRuntimeVersion, cNotSet)
                            ThreadingModel = Registry.GetValue(RegistryLocation, cThreadingModel, cNotSet)
                            AssemblyAsString = Registry.GetValue(RegistryLocation, cAssembly, cNotSet)
                            Return DLLSource
                        Else
                            DLLSource = "No HKCU32 filename"
                        End If
                    Else ' secondary check
                        ' we do not have a filename have another entry
                        If ((_filename <> "" And _filename IsNot Nothing) And _filename = Filename) Then ' wee have a match so this is a possible entry
                            Debug.Print("CheckHKCU32() same filename")
                            DLLSource += " : " & AddInEntry.cHKCU32
                        End If
                    End If
                End If ' filename ""
            End If ' filename nothing
        End If ' location

        Return ""
    End Function
    ''' <summary>
    ''' Check if there are any HKLM entries for the current classID
    ''' </summary>
    ''' <returns>HKLM if yes else ""</returns>
    Private Function CheckHKLM32(pPrimaryCheck As Boolean) As String
        ' there are 2 locations for WOW6432 in HKLM - it is believed they mirror each other but not sure so check both
        Dim _location As String = IIf(OS64Bit, cHKLMWow1_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32,
                       cHKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)
        If _location = "" Then ' try other location
            _location = IIf(OS64Bit, cHKLMWow2_Classes & cBackSlash & ClassID & cBackSlash & cInprocServer32,
                       cHKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32)
        End If
        If _location <> "" Then
            Dim prefix As String = ""

            ' get name of DLL based on location
            Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path
            If _filename Is Nothing Then
                Return ""
            Else

                If _filename = "" Then
                    _filename = Registry.GetValue(_location, "", "")
                    If _filename <> "" Then prefix = "Default key: "
                End If
                '1st pass we check we have a filename 
                If pPrimaryCheck Then
                    If (_filename <> "" And _filename IsNot Nothing) Then
                        Filename = _filename ' this is the primary filename
                        DLLSource = prefix & AddInEntry.cHKLM32
                        getDLLAssembly()
                        RegistryLocation = _location
                        RunTimeVersion = Registry.GetValue(RegistryLocation, cRuntimeVersion, cNotSet)
                        ThreadingModel = Registry.GetValue(RegistryLocation, cThreadingModel, cNotSet)
                        AssemblyAsString = Registry.GetValue(RegistryLocation, cAssembly, cNotSet)
                        Return DLLSource
                    Else
                        DLLSource = "No HKLM filename"
                        Return DLLSource
                    End If
                Else ' secondary check
                    ' we do not have a filename have another entry
                    If ((_filename <> "" And _filename IsNot Nothing) And _filename = Filename) Then ' wee have a match so this is a possible entry
                        Debug.Print("CheckHKLM32() same filename")
                        DLLSource += " : " & AddInEntry.cHKLM32
                        Return DLLSource
                    End If
                End If
            End If


        End If
        Return ""
    End Function


    Private Function CheckHKCU(pPrimaryCheck As Boolean) As String
        Dim _location As String = cHKCU_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
        If _location <> "" Then
            Dim prefix As String = ""
            ' get name of DLL based on location
            Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path
            If _filename Is Nothing Then
                Return ""
            Else

                If _filename = "" Then
                    _filename = Registry.GetValue(_location, "", "")
                    If _filename <> "" Then prefix = "Default key: "
                End If
                '1st pass we check we have a filename 
                If pPrimaryCheck Then
                    If (_filename <> "" And _filename IsNot Nothing) Then
                        Filename = _filename ' this is the primary filename
                        DLLSource = AddInEntry.cHKCU64
                        getDLLAssembly()
                        RegistryLocation = _location
                        RunTimeVersion = Registry.GetValue(RegistryLocation, cRuntimeVersion, cNotSet)
                        ThreadingModel = Registry.GetValue(RegistryLocation, cThreadingModel, cNotSet)
                        AssemblyAsString = Registry.GetValue(RegistryLocation, cAssembly, cNotSet)
                        Return DLLSource
                    Else
                        DLLSource = "No HKCU filename"
                        Return DLLSource
                    End If
                Else ' secondary check
                    ' we do not have a filename have another entry
                    If ((_filename <> "" And _filename IsNot Nothing) And _filename = Filename) Then ' wee have a match so this is a possible entry
                        Debug.Print("CheckHKCU64() same filename")
                        DLLSource += " : " & AddInEntry.cHKCU64
                        Return DLLSource
                    End If
                End If
            End If

        End If
        Return ""
    End Function
    Private Function CheckHKLM(pPrimaryCheck As Boolean) As String
        Dim _location As String = cHKLM_Classes & cBackSlash & cCLSID & cBackSlash & ClassID & cBackSlash & cInprocServer32
        Dim prefix As String = ""

        ' get name of DLL based on location
        Dim _filename As String = Registry.GetValue(_location, cCodeBase, "") 'using the class try to find the DLL path
        If _filename Is Nothing Then
            Return ""
        Else

            If _filename = "" Then
                _filename = Registry.GetValue(_location, "", "")
                If _filename <> "" Then prefix = "Default key: "
            End If

            '1st pass we check we have a filename 
            If pPrimaryCheck Then
                If (_filename <> "" And _filename IsNot Nothing) Then
                    Filename = _filename ' this is the primary filename
                    DLLSource = AddInEntry.cHKLM64
                    getDLLAssembly()
                    RegistryLocation = _location
                    RunTimeVersion = Registry.GetValue(RegistryLocation, cRuntimeVersion, cNotSet)
                    ThreadingModel = Registry.GetValue(RegistryLocation, cThreadingModel, cNotSet)
                    AssemblyAsString = Registry.GetValue(RegistryLocation, cAssembly, cNotSet)
                    Return DLLSource
                Else
                    DLLSource = "No HKCU filename"
                    Return DLLSource
                End If
            Else ' secondary check
                ' we do not have a filename have another entry
                If ((_filename <> "" And _filename IsNot Nothing) And _filename = Filename) Then ' wee have a match so this is a possible entry
                    Debug.Print("CheckHKCU32() same filename")
                    DLLSource += " : " & AddInEntry.cHKLM64
                    Return DLLSource
                End If
            End If
        End If

        Return ""

    End Function


    ''' <summary>
    ''' Get information for DLL file
    ''' * Assembly
    ''' * Version
    ''' 
    ''' </summary>
    Private Sub getDLLAssembly()

        Dim _filename As String = cleanFilename(Filename)

        If _filename IsNot Nothing AndAlso _filename <> "" AndAlso _filename <> cNotSet Then
            If File.Exists(_filename) Then
                If Strings.Left(_filename, fileprefixlength) = cFilePrefix Then _filename = Strings.Right(_filename, _filename.Length - fileprefixlength)
                Try
                    Dim ass As AssemblyName = AssemblyName.GetAssemblyName(_filename)
                    DLLVersion = ass.Version.ToString
                Catch ex As Exception
                    DLLVersion = "Unable to determine"
                End Try
                Filename = cleanFilename(_filename)
            End If

            MismatchedHives = ClassSource <> DLLSource
            DLLPathExists = IIf(Filename = cNotSet, False, True)
            If DLLPathExists Then DLLExistsInSpecifiedLocation = DLLexists(Filename)

        End If

    End Sub


    ' was used to ensure the comparison between filenames was fair - however all entries that are checked are from registry so unless we need to 
    ' use the address other than checking OK to leave raw
    Private Function cleanFilename(_Filename As String) As String

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
    ' ?loads
    ' Mismatched HIVES
    ' DLL exists
    ' DLL wrong location
    ' TODO Can we detect other stuff??
    ' TODO presedence on colouring

    Friend Function getLineColour() As Color
        Dim c As Color = Color.White
        Try
            If Not SparxKeyExists Then Return Color.Red
            If Not ClassIDExists Then Return Color.HotPink
            If MismatchedHives Then Return Color.Wheat
            If Not DLLExistsInSpecifiedLocation Then Return Color.Pink
            Return Color.Lime

        Catch ex As Exception

        End Try
        Return Color.Red
    End Function

End Class



