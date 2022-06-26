'   Copyright (C) 2017, 2018 Adrian LINCOLN, EXploringEA - All Rights Reserved
'
'   This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

'   This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
' =============================================================================================================================================
Imports System.Text.RegularExpressions
Namespace util


    ''' <summary>
    ''' Class to provide functions to convert between arrays, Strings (CSV), and Lists
    ''' </summary>
    Public Class CSVArrayListConversion

        Private Const cComma As String = ","
        Private Const cDoubleQuote As String = """"

#Region "List conversion functions"

        ''' <summary>
        ''' Convert an ArrayList into a CSV
        ''' </summary>
        ''' <param name="pItems">Arraylist of items</param>
        ''' <returns>CSV representation of Array list</returns>
        Public Shared Function convertArrayListIntoCSV(pItems As ArrayList) As String
            If pItems IsNot Nothing Then
                Try
                    Dim s As String = ""
                    If pItems IsNot Nothing Then
                        If pItems.Count > 0 Then
                            For Each a As String In pItems
                                s += a & cComma
                            Next
                            If Len(s) > 0 Then Return Mid$(s, 1, Len(s) - 1)
                        End If
                    End If
                    Return s
                Catch ex As Exception
#If DEBUG Then
                    Debug.Print(ex.ToString)
#End If
                End Try

            End If
            Return ""
        End Function


        ''' <summary>
        ''' Converts the CSV to string arrat
        ''' </summary>
        ''' <param name="pCSV">CSV string to convert.</param>
        ''' <returns>String Array of items</returns>
        Public Shared Function convertCSV2Array(pCSV) As Array
            Dim A() As String = Nothing
            Try
                A = Split(pCSV, ",")
            Catch ex As Exception
#If DEBUG Then
                Debug.Print(ex.ToString)
#End If
            End Try
            Return A

        End Function




        ''' <summary>
        ''' Converts the csv string to array list
        ''' </summary>
        ''' <param name="pCSV">CSV string to convert.</param>
        ''' <returns>Array list of items</returns>
        Public Shared Function convertCSVtoArrayList(pCSV As String) As ArrayList
            Dim a As New ArrayList
            Try
                ' extract the inidividual items from string
                Dim myStringLength As Integer = Len(pCSV)
                If myStringLength > 0 Then
                    For Each item As String In convertCSV2Array(pCSV)
                        a.Add(item)
                    Next

                End If

            Catch ex As Exception
#If DEBUG Then
                Debug.Print(ex.ToString)
#End If
            End Try
            Return a
        End Function



        ''' <summary>
        ''' Extract tokens from CSV and put into an arraylist
        ''' </summary>
        ''' <param name="pCSV">Complex CSV</param>
        ''' <returns>Array list of items</returns>
        ''' <remarks>
        ''' Basic string is a,b,c
        ''' to support complex tokens use [a,YYIk],[sggh],[sddh]
        ''' We use a combination of regex to extract the tokens and then the string and prefix characters that may exists 
        ''' as well as ignore empty strings
        ''' </remarks>
        Public Shared Function convertComplexCSVtoArrayList(pCSV As String) As ArrayList
            Dim myArray As New ArrayList

            Try
                ' if the start and end of the string has "" then tokens are in "token" - and we parse the string and seperate on tokens that may include ,

                If Len(pCSV) = 0 Then Return myArray
                If Strings.Left(pCSV, 1) = "[" Then
                    Dim pattern As String = "\[*\]"
                    For Each result As String In Regex.Split(pCSV, pattern)
                        If Strings.Left(result, 1) = "," Then result = Strings.Right(result, Len(result) - 1)
                        If Strings.Left(result, 1) = "[" Then result = Strings.Right(result, Len(result) - 1)
                        If result <> "" Then myArray.Add(result)
                    Next
                Else
                    pCSV = pCSV.Replace(cDoubleQuote, "")
                    Dim mysplit() As String = Split(pCSV, cComma)

                    For Each s As String In mysplit
                        ' remove "" from the string
                        '    s = s.Replace(cDoubleQuote, "")
                        If s <> "" Then myArray.Add(s)
                    Next
                End If
                Return myArray
            Catch ex As Exception
#If DEBUG Then
                Debug.Print(ex.ToString)
#End If
            End Try
            Return Nothing ' i.e. there are no STs if we have a "" then there will always be one which we have to exclude if we then add others
        End Function


        ''' <summary>
        ''' Converts the ListBox items to CSV
        ''' </summary>
        ''' <param name="pListview">ListView</param>
        ''' <returns>CSV of items</returns>
        Public Shared Function convertListBoxItemsToCSV(pListview As System.Windows.Forms.ListBox) As String
            Dim s As String = ""
            Try
                If pListview.SelectedItems.Count > 0 Then
                    For Each item As String In pListview.SelectedItems
                        s += item & cComma
                    Next
                    If Len(s) > 0 Then Return Mid$(s, 1, Len(s) - 1)
                Else
                    Return ""
                End If
            Catch ex As Exception
#If DEBUG Then
                Debug.Print(ex.ToString)
#End If

            End Try
            Return ""

        End Function

        ''' <summary>
        ''' Converts the ListBox items to CSV
        ''' </summary>
        ''' <param name="pListview">ListView</param>
        ''' <returns>Array List of items</returns>
        Public Shared Function convertListBoxItemsToArrayList(pListview As System.Windows.Forms.ListBox) As ArrayList
            Dim myArrayList As New ArrayList
            Try
                If pListview.SelectedItems.Count > 0 Then
                    For Each item As String In pListview.SelectedItems
                        myArrayList.Add(item)
                    Next
                    Return myArrayList
                Else
                    Return Nothing
                End If
            Catch ex As Exception

#If DEBUG Then
                Debug.Print(ex.ToString)
#End If
            End Try
            Return Nothing

        End Function


#End Region

        Private Sub New()

        End Sub
    End Class

End Namespace
