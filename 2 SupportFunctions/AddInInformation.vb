Imports Microsoft.Win32

Public Class AddInInformation

    ''' <summary>
    ''' Gets a list of addin entries from Windows Registry
    ''' - start with HKCU as they have precendence over the HKLM
    ''' Assume that the registry keys in the following locations identify 32-bit AddIns
    '''	"HKCU\SOFTWARE\Sparx Systems\EAAddins" - HKCU
    '''	"HKLM\ SOFTWARE\Sparx Systems\EAAddins" - HKLM, although if the addin Is running on a 64-bit operating system this would be "HKLM\ SOFTWARE\WOW6432NODE\Sparx Systems\EAAddins"  HKLMWow
    '''	 Assume that the registry keys in the following locations identify 64-bit AddIns
    '''	 "HKCU\SOFTWARE\Sparx Systems\EAAddins64" - HKCU64
    '''	 "HKLM\ SOFTWARE\Sparx Systems\EAAddins64" – HKLM64
    ''' </summary>
    ''' <returns>List of AddIn entries</returns>
    Function getListOfEAAddinEntries() As ArrayList

        ' for all of the keys which are not nothing we add them to an array
        Dim myAddInEntries As New ArrayList
        For Each entry As AddInEntry In getListof32BitHKCUAddinEntries()
            myAddInEntries.Add(entry)
        Next

        For Each entry As AddInEntry In getListof32BitHKLMAddinEntries()
            myAddInEntries.Add(entry)
        Next

        For Each entry As AddInEntry In getListof64BitHKCUAddinEntries()
            myAddInEntries.Add(entry)
        Next
        For Each entry As AddInEntry In getListof64BitHKLMAddinEntries()
            myAddInEntries.Add(entry)
        Next

        Return myAddInEntries

    End Function


    ''' <summary>
    ''' 32 bit AddIns when install for current users only
    ''' Sparx keys (32-bit apps) can be in HKCU\Software\SparxSystems\EAAddIns
    ''' </summary>
    ''' <returns>List of 32-bit AddIns for current user</returns>
    Function getListof32BitHKCUAddinEntries() As ArrayList
        Dim my32HKCUAddInEntries As New ArrayList
        Dim myCUKey As RegistryKey = Registry.CurrentUser.OpenSubKey(SparxKeys32) ' get 32-bit AddIn for the current user
        If myCUKey IsNot Nothing Then
            For Each pEntryKey In myCUKey.GetSubKeyNames
                Dim AddInInfo As New AddInEntry
                AddInInfo.AddInName = pEntryKey
                AddInInfo.ClassName = Registry.GetValue(eaHKCU32AddInKeys & cBackSlash & pEntryKey, "", cNotSet)
                AddInInfo.SparxAddinLocation = cHKCU32
                my32HKCUAddInEntries.Add(AddInInfo)
            Next
        End If
        Return my32HKCUAddInEntries
    End Function
    ''' <summary>
    ''' 32 bit AddIns when installed local machine (for all users)
    ''' Sparx keys (32-bit apps) can be in HKLM\Software\SparxSystems\EAAddIns
    ''' BUT NOTE for 64-bit operating systems the keys will be located in HKLM\Software\Wow6432Node\SparxSystems\EAAddIns
    ''' </summary>
    ''' <returns>List of 32-bit Addins for local machine</returns>
    Function getListof32BitHKLMAddinEntries() As ArrayList

        Dim my32HKLMAddInEntries As New ArrayList
        Dim myLMKey As RegistryKey = IIf(Environment.Is64BitOperatingSystem, Registry.LocalMachine.OpenSubKey(SparxKeysWOW32), Registry.LocalMachine.OpenSubKey(SparxKeys32))
        If myLMKey IsNot Nothing Then
            For Each pEntryKey In myLMKey.GetSubKeyNames
                Dim AddInInfo As New AddInEntry
                AddInInfo.AddInName = pEntryKey ' registry location for addin name 
                AddInInfo.ClassName = Registry.GetValue(eaHKLM32AddInKey32 & cBackSlash & pEntryKey, "", cNotSet)
                AddInInfo.SparxAddinLocation = If(Environment.Is64BitOperatingSystem, cHKLM32Wow, cHKLM32)
                my32HKLMAddInEntries.Add(AddInInfo)
            Next
        End If
        Return my32HKLMAddInEntries
    End Function
    ''' <summary>
    ''' 64 bit AddIns when install for current users only
    ''' HKCU\Software\SparxSystems\EAAddins64
    ''' </summary>
    ''' <returns>List of current user 64-bit addin entries</returns>
    Function getListof64BitHKCUAddinEntries() As ArrayList
        Dim my64HKUCAddInEntries As New ArrayList
        Dim myCUKey64 As RegistryKey = Registry.CurrentUser.OpenSubKey(SparxKeys64)
        If myCUKey64 IsNot Nothing Then
            For Each pEntryKey In myCUKey64.GetSubKeyNames
                Dim AddInInfo As New AddInEntry
                AddInInfo.AddInName = pEntryKey
                AddInInfo.ClassName = Registry.GetValue(eaHKCU64AddInKey & cBackSlash & pEntryKey, "", cNotSet)
                AddInInfo.SparxAddinLocation = cHKCU64
                my64HKUCAddInEntries.Add(AddInInfo)
            Next
        End If
        Return my64HKUCAddInEntries
    End Function
    ''' <summary>
    ''' 64 bit AddIns when installed for all users
    ''' HKLM\Software\Wow6432Node\SparxSystems\EAAddIns64
    ''' </summary>
    ''' <returns>List of 64-bit addins for Local Machine (all users)</returns>
    Function getListof64BitHKLMAddinEntries() As ArrayList
        Dim my64KLMAddInEntries As New ArrayList
        Dim myLMKey64 As RegistryKey = Registry.LocalMachine.OpenSubKey(SparxKeys64) '"SOFTWARE\Sparx Systems\EAAddins64") 'SparxKeys)
        If myLMKey64 IsNot Nothing Then
            For Each pEntryKey In myLMKey64.GetSubKeyNames
                Dim AddInInfo As New AddInEntry
                AddInInfo.AddInName = pEntryKey ' registry location for addin name 
                AddInInfo.ClassName = Registry.GetValue(eaHKLM64AddInKey & cBackSlash & pEntryKey, "", cNotSet) ' look for the class name CLSID name
                AddInInfo.SparxAddinLocation = cHKLM64
                my64KLMAddInEntries.Add(AddInInfo)
            Next
        End If
        Return my64KLMAddInEntries
    End Function

    ''' <summary>
    ''' AddIn entry summary
    ''' </summary>
    Private Class AddInEntry

        '' AddIn Name | Class | Source | CLSID | Source | DLL
        Property AddInName As String = ""
        Property ClassName As String = "" ' Assembly.Class
        Property SparxAddinLocation As String = "" ' Values are HKCU, HKLM, HKCU64, HKLM64


    End Class

End Class


''' <summary>
''' AddIn entry details
''' </summary>
Friend Class AddInDetail

    '' AddIn Name | Class | Source | CLSID | Source | DLL
    ''' <summary>
    ''' Gets or sets the name of the add in.
    ''' </summary>
    ''' <value>
    ''' The AddIn Name
    ''' </value>
    Property AddInName As String
    ''' <summary>
    ''' Gets or sets the class definition.
    ''' </summary>
    ''' <value>
    ''' The class name i.e. Assembly.Class
    ''' </value>
    Property ClassDefinition As String
    ''' <summary>
    ''' Gets or sets the sparx entry.
    ''' </summary>
    ''' <value>
    ''' The location of the SparxEntry in registry - e.g. HKCU or HKLM
    ''' </value>
    Property SparxEntry As String
    ''' <summary>
    ''' Gets or sets the class source.
    ''' </summary>
    ''' <value>
    ''' Location of class in registry  - e.g. HKCU or HKLM
    ''' </value>
    Property ClassSource
    ''' <summary>
    ''' Gets or sets the CLSID.
    ''' </summary>
    ''' <value>
    ''' Class ID (GUID)
    ''' </value>
    Property CLSID
    ''' <summary>
    ''' Gets or sets the CLSID source.
    ''' </summary>
    ''' <value>
    ''' Class ID Source in registry  - e.g. HKCU or HKLM
    ''' </value>
    Property CLSIDSource
    ''' <summary>
    ''' Gets or sets the DLL.
    ''' </summary>
    ''' <value>
    ''' The DLL full file pathname
    ''' </value>
    Property DLL

End Class
