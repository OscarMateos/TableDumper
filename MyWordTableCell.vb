Imports System.Text.RegularExpressions
''' <summary>
'''     Clase WordTableCell, celda de una WordTable
''' </summary>
''' <remarks>
'''     Se corresponde con cada celda de una tabla de un documento fuente Word, donde puede haber celdas de dimensiones mayores que 1.
''' </remarks>
Public Class MyWordTableCell
    Public RowNumber As Integer
    Public ColumnNumber As Integer
    Public RowSpan As Integer
    Public ColumnSpan As Integer
    Public Text As [String]
    Public BgColor As WdColor

    ''' <summary>
    '''     Funcion que devuelve Verdadero si el contenido de la celda Word que se le pasa es un texto con subindices, 
    '''     pues este debe ser codificado de forma espcial.
    ''' </summary>
    ''' <param name="WordTableCell">
    '''     Microsoft.Office.Interop.Word.Cell
    ''' </param>
    ''' <returns>
    '''     Se usa en MyWordTable.leerCeldas(Table)
    ''' </returns>
    Public Function ContainsSubscript(ByRef WordTableCell As Cell) As Boolean
        Dim CellRange As Range = WordTableCell.Range
        Dim LastCharIndex As Integer = Left(CellRange.Text, CellRange.Text.IndexOf(Chr(13))).Length

        ' Para detectar las celdas sin contenido
        Dim TextLength = CellRange.Text.Replace(vbCr & ChrW(7), "").Length
        If TextLength > 0 Then
            If CellRange.Characters(LastCharIndex).Font.Subscript = -1 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    ''' <summary>
    '''     Funcion para codificar de forma especial la parte de una cadena que deba ser representada o interpretada como un subindice.
    '''     Elimina asi la ambiguedad posible al codificar todo al mismo nivel, con literales que pudiesen existir, por ejemplo:
    '''         - Ant(i), parte subindice entre parentesis se codificaria sin desambiguar como Anti, que puede coincidir con el 
    '''           prefijo "anti" (aunque aqui la "i" es numero romano) y que se codificaria como Ant_csub1.
    ''' </summary>
    ''' <param name="WordTableCell">
    '''     Microsoft.Office.Interop.Word.Cell
    ''' </param>
    Public Sub DisambiguateSubscript(ByRef WordTableCell As Cell)
        Dim CellRange As Range = WordTableCell.Range
        Dim CellText As String = Left(CellRange.Text, CellRange.Text.IndexOf(Chr(13)))
        Dim FirstCharIndex As Integer = -1
        Dim LastCharIndex As Integer = CellText.Length
        Dim rgxRomanChars As New Regex("^[MmDdCcLlXxVvIi0]+")

        For i = LastCharIndex To 1 Step -1
            If CellRange.Characters(i).Font.Subscript = -1 Then
                FirstCharIndex = i
            Else
                Exit For
            End If
        Next

        Dim SubscriptSubstring As String = CellText.Substring(FirstCharIndex - 1)

        If rgxRomanChars.IsMatch(SubscriptSubstring) Then
            Dim Substring As String = Left(CellText, FirstCharIndex - 1)
            Me.Text = Substring + "_csub" + Roman2Num(SubscriptSubstring).ToString
        End If
    End Sub

    ''' <summary>
    '''     Convierte una representacion de numero romano a numero arabigo.
    ''' </summary>
    ''' <param name="Roman">
    '''     String, representacion de un numero romano.
    ''' </param>
    ''' <returns></returns>
    Function Roman2Num(Roman As String) As Long
        Dim Roman2 As String
        Dim Char1 As String, Char2 As String
        Dim Number As Long

        Roman2 = UCase(Roman)

        Do While Len(Roman2)
            Char1 = Left(Roman2, 1)
            Char2 = Mid(Roman2, 2, 1)
            Roman2 = Right(Roman2, Len(Roman2) - 1)
            Select Case Char1
                Case "M"
                    Number = Number + 1000

                Case "D"
                    Number = Number + 500
                Case "C"
                    Select Case Char2
                        Case "M"
                            Number = Number + 900
                            Roman2 = Right(Roman2, Len(Roman2) - 1)
                        Case "D"
                            Number = Number + 400
                            Roman2 = Right(Roman2, Len(Roman2) - 1)
                        Case Else
                            Number = Number + 100

                    End Select
                Case "L"
                    Number = Number + 50
                Case "X"
                    Select Case Char2
                        Case "M"
                            Number = Number + 990
                            Roman2 = Right(Roman2, Len(Roman2) - 1)
                        Case "D"
                            Number = Number + 490
                            Roman2 = Right(Roman2, Len(Roman2) - 1)
                        Case "C"
                            Number = Number + 90
                            Roman2 = Right(Roman2, Len(Roman2) - 1)
                        Case "L"
                            Number = Number + 40
                            Roman2 = Right(Roman2, Len(Roman2) - 1)
                        Case Else
                            Number = Number + 10
                    End Select

                Case "V"
                    Number = Number + 5
                Case "I"
                    Select Case Char2
                        Case "M"
                            Number = Number + 999
                            Roman2 = Right(Roman2, Len(Roman2) - 1)
                        Case "D"
                            Number = Number + 499
                            Roman2 = Right(Roman2, Len(Roman2) - 1)
                        Case "C"
                            Number = Number + 99
                            Roman2 = Right(Roman2, Len(Roman2) - 1)
                        Case "L"
                            Number = Number + 49
                            Roman2 = Right(Roman2, Len(Roman2) - 1)
                        Case "X"
                            Number = Number + 9
                            Roman2 = Right(Roman2, Len(Roman2) - 1)
                        Case "V"
                            Number = Number + 4
                            Roman2 = Right(Roman2, Len(Roman2) - 1)
                        Case Else
                            Number = Number + 1
                    End Select
            End Select
        Loop

        Return Number
    End Function
End Class
