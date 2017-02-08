' Copyright (C) 2015 - 2017 Adrian LINCOLN, EXploringEA - All Rights Reserved
' You may use, distribute and modify this code under the terms of the 3-Clause BSD License
'
' You should have received a copy of the 3-Clause BSD License with this file. 
' If not, please email: adrian@EXploringEA.co.uk 
'=====================================================================================

Module Constants


    ''' <summary>
    ''' Root to look in HKCU
    ''' </summary>
    Friend Const EA As String = "HKEY_CURRENT_USER\Software\Sparx Systems\EA400\EA"


    ''' <summary>
    ''' The sparx keys - x86
    ''' </summary>
    Friend Const SparxKeys As String = "Software\Sparx Systems\EAAddins"
    ''' <summary>
    ''' The sparx keys - x64
    ''' </summary>
    Friend Const WSparxKeys As String = "Software\Wow6432Node\Sparx Systems\EAAddins"

    ''' <summary>
    ''' HKCU Keys
    ''' </summary>
    Friend Const HKCUfullKey As String = "HKEY_CURRENT_USER\" & SparxKeys
    ''' <summary>
    ''' HKLM Keys
    ''' </summary>
    Friend Const HKLMfullKey As String = "HKEY_LOCAL_MACHINE\" & SparxKeys

    ''' <summary>
    ''' HKCR Classes
    ''' </summary>
    Friend Const HKCRClasses As String = "HKEY_CURRENT_USER\SOFTWARE\Classes"
    ''' <summary>
    ''' HKLM Classes
    ''' </summary>
    Friend Const HKLMClasses As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes"
    ''' <summary>
    ''' Text not set
    ''' </summary>
    Friend Const cNotSet As String = "NOT SET"


End Module
