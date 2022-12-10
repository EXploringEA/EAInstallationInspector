Friend Class AddInEntry

    'Registry related strings defines the source of the Sparx AddIn key and OS which indicates where the class information should be located

    Friend Const cHKCR32 As String = "HKCR32"
    Friend Const cHKCR64 As String = "HKCR64"
    Friend Const cHKCR As String = "HKCR" ' set when not clear which 

    ' 32-bit AddIns on 32-bit EA
    Friend Const cHKCU32 As String = "HKCU32" ' 32-bit Addins on 32-bit OS 
    Friend Const cHKLM32 As String = "HKLM32" ' 32-bit Addins on 32-bit OS

    ' 32-bit AddIns on 64-bit OS
    Friend Const cHKLM32Wow As String = "HKLM32Wow" ' 32-bit AddIn on 64-bit OS ?? NT SURE WE NEED THIS as the Sparx key is detached from the class registration

    ' 64-bit AddIns
    Friend Const cHKCU64 As String = "HKCU64" ' 64-bit AddIn on 64-bit OS
    Friend Const cHKLM64 As String = "HKLM64" ' 64-bit AddIn on 64-bit OS

    '' AddIn Name | Class | Source | CLSID | Source | DLL
    ''' <summary>
    ''' Gets or sets the name of the add in.
    ''' </summary>
    ''' <value>
    ''' The AddIn Name
    ''' </value>
    Property AddInName As String = ""
    ''' <summary>
    ''' Gets or sets the class definition.
    ''' </summary>
    ''' <value>
    ''' The class name i.e. Assembly.Class
    ''' </value>
    Property ClassName As String = "" ' Assembly.Class
    Property SparxAddinLocation As String = "" ' Values are below
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