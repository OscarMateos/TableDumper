Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Tools.Word
Imports System.Collections
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Xml

''' <summary>
'''     Clase WordTable para procesado de Tablas de un documento Word
''' </summary>
''' <remarks></remarks>
Public Class MyWordTable
    Property Name As String
    Property [Structure] As Integer(,)
    Property Cells As ArrayList '(Of MyWordTableCell)
    Property HeaderRow As HashSet(Of String)
    Property Data As HashSet(Of String)()
    Property RepeatedRowValues As Boolean

    ''' <summary>
    '''     Constructor 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        Cells = New ArrayList '(Of MyWordTableCell)
    End Sub

    ''' <summary>
    '''     Asigna tamanyo y memoria para los campos Cabecera y Datos 
    ''' </summary>
    ''' <param name="DataSize">
    '''     Tamanyo de los datos de la tabla.
    ''' </param>
    ''' <param name="HasHeader">
    '''     [Opcional] Para determinar si se debe reservar tamanyo para la cabecera de la tabla, por defecto Si (True).
    ''' </param>
    ''' <remarks></remarks>
    Private Sub ManageDataSize(ByVal DataSize As Integer, Optional ByVal HasHeader As Boolean = True)
        If DataSize = 0 Then
            Exit Sub
        End If

        'Tiene Cabecera (HeaderRow)
        If HasHeader Then
            Me.HeaderRow = New HashSet(Of String)
            ReDim Data(DataSize - 2)

            For i = 0 To DataSize - 2
                Data(i) = New HashSet(Of String)
            Next
        Else
            ' No tiene Cabecera
            ReDim Data(DataSize - 1)

            For i = 0 To DataSize - 1
                Data(i) = New HashSet(Of String)
            Next
        End If
    End Sub

    ''' <summary>
    '''     Funcion que devuelve Verdadero si la tabla pasada como argumento tiene cabecera.
    ''' </summary>
    ''' <param name="WordTable">
    '''     Microsoft.Office.Interop.Word.Table
    ''' </param>
    ''' <returns></returns>
    '''     Boolean
    ''' <remarks></remarks>
    Private Function HasHeader(ByRef WordTable As Table) As Boolean
        With WordTable
            Dim C As Cell = .Cell(1, 1)
            Dim BgColor As WdColor
            Dim HeaderTest As Boolean = True

            While C.RowIndex = 1
                BgColor = C.Range.Shading.BackgroundPatternColor

                If BgColor <> WdColor.wdColorBlack Then
                    HeaderTest = False
                    Exit While
                End If
                C = C.Next
            End While

            Return HeaderTest
        End With
    End Function

    ''' <summary>
    '''     Asigna valor para el campo RepeatedRowValues.
    ''' </summary>
    ''' <remarks>
    '''     RepeatedRowValues: Sera Verdadero si para una misma fila de la tabla se repite alguna 
    '''                        celda en diferentes columnas. Usado en getLayOut.           
    ''' </remarks>
    Private Sub SetRepeatedRowValues()
        Dim RTest As HashSet(Of Integer) = New HashSet(Of Integer)
        Dim RValues As List(Of String) = New List(Of String)
        RepeatedRowValues = False

        ' Para cada fila se guardan en una lista el valor de cada celda
        For i = 0 To UBound([Structure], 1)

            For j = 0 To UBound([Structure], 2)
                RTest.Add([Structure](i, j))
            Next

            ' Los valores de las celdas de la fila
            For Each CellIndex As Integer In RTest
                If CellIndex <> -1 Then
                    Dim CellText As String = Cells(CellIndex).Text
                    If Not String.IsNullOrWhiteSpace(CellText) Then
                        RValues.Add(Cells(CellIndex).Text)
                    End If
                End If
            Next

            ' Si hay algun valor repetido se fija a True
            Dim DuplicateExists = RValues.GroupBy(Function(n) n).Any(Function(g) g.Count() > 1)
            If DuplicateExists Then
                RepeatedRowValues = True
            End If

            RTest.Clear()
            RValues.Clear()
        Next
    End Sub

    ''' <summary>
    '''     Obtiene la distribucion de las celdas (ya que pueden ser irregulares) sobre el espacio  
    '''     que forman las filas y columnas de las tablas y se la asigna al campo Estuctura.     
    ''' </summary>
    ''' <param name="WordTable">
    '''     Microsoft.Office.Interop.Word.Table
    ''' </param>
    ''' <remarks></remarks>
    Private Sub GetLayOut(ByRef WordTable As Table)
        If Not IsNothing(Cells) Then
            With WordTable
                'Array con la estructura de la tabla
                Dim RowsCount As Integer = .Rows.Count
                Dim ColumnsCount As Integer = .Columns.Count
                ReDim [Structure](RowsCount - 1, ColumnsCount - 1)
                For i = 0 To RowsCount - 1
                    For j = 0 To ColumnsCount - 1
                        [Structure](i, j) = -1
                    Next
                Next

                'Distribución de las celdas en la tabla
                For i = 0 To Cells.Count - 1
                    Dim C As MyWordTableCell = Cells(i)
                    For j = C.RowNumber - 1 To (C.RowNumber - 1) + (C.RowSpan - 1)
                        For k = C.ColumnNumber - 1 To (C.ColumnNumber - 1) + (C.ColumnSpan - 1)
                            If [Structure](j, k) = -1 Then
                                [Structure](j, k) = i
                            Else
                                If k + 1 > UBound([Structure], 2) Then
                                    ReDim Preserve [Structure](RowsCount - 1, ColumnsCount)
                                    For l = 0 To UBound([Structure], 1)
                                        [Structure](l, UBound([Structure], 2)) = -1
                                    Next
                                End If

                                [Structure](j, k + 1) = i
                            End If
                        Next
                    Next
                Next

                'Determinar si existen valores repetidos a nivel de fila
                SetRepeatedRowValues()
            End With
        End If
    End Sub

    ''' <summary>
    '''     Lee las celdas de la tabla (como MyWordTableCell) y las asigna al campo Cells.
    ''' </summary>
    ''' <param name="WordTable">
    '''     Microsoft.Office.Interop.Word.Table
    ''' </param>
    ''' <remarks>
    '''     Recorre la tabla segun la estructura Xml de la misma en el fichero Word.
    ''' </remarks>
    Private Sub ReadCells(ByRef WordTable As Table)
        Dim WordCell As Cell = WordTable.Cell(1, 1)
        Dim S As [String] = WordTable.Range.XML
        Dim XMLDoc As New XmlDocument()
        XMLDoc.LoadXml(S)
        Dim NsMgr As New XmlNamespaceManager(XMLDoc.NameTable)
        NsMgr.AddNamespace("w", "http://schemas.microsoft.com/office/word/2003/wordml")

        While WordCell IsNot Nothing
            Dim Cell As New MyWordTableCell()
            Cell.RowNumber = WordCell.RowIndex
            Cell.ColumnNumber = WordCell.ColumnIndex

            'ColSpan: Alto de celda combinada vertical
            Dim ColSpan As Integer
            Dim ExactCell As Xml.XmlNode = XMLDoc.SelectNodes("//w:tr[" + WordCell.RowIndex.ToString() + "]/w:tc[" + WordCell.ColumnIndex.ToString() + "]/w:tcPr/w:gridSpan", NsMgr)(0)

            ColSpan = If(Not IsNothing(ExactCell), Convert.ToInt16(ExactCell.Attributes("w:val").Value), 1)

            'RowSpan: Ancho de celda combinada horizontal
            Dim RowSpan As Integer = 1
            Dim EndRows As [Boolean] = False
            Dim NextRows As Integer = WordCell.RowIndex + 1
            Dim ExactCellVMerge As Xml.XmlNode = XMLDoc.SelectNodes("//w:tr[" + WordCell.RowIndex.ToString() + "]/w:tc[" + WordCell.ColumnIndex.ToString() + "]/w:tcPr/w:vmerge", NsMgr)(0)

            If (ExactCellVMerge Is Nothing) OrElse (ExactCellVMerge IsNot Nothing AndAlso ExactCellVMerge.Attributes("w:val") Is Nothing) Then
                RowSpan = 1
            Else
                While NextRows <= WordTable.Rows.Count AndAlso Not EndRows
                    Dim NextCellMerge As Xml.XmlNode = XMLDoc.SelectNodes("//w:tr[" + NextRows.ToString() + "]/w:tc[" + WordCell.ColumnIndex.ToString() + "]/w:tcPr/w:vmerge", NsMgr)(0)
                    If NextCellMerge IsNot Nothing AndAlso (NextCellMerge.Attributes("w:val") Is Nothing) Then
                        NextRows += 1
                        RowSpan += 1
                        Continue While
                    Else
                        EndRows = True
                    End If
                End While
            End If

            'Anyadimos celda procesada actual
            Cell.RowSpan = RowSpan
            Cell.ColumnSpan = ColSpan

            ''''''''''''''''''''''''''''SUBSCRIPT DETECT & PARSE
            ''Contiene subindices? Son Validos (un solo grupo al final de la palabra)?
            'For k = 1 To contenidoCelda.Length - 1
            '    Dim MyRangeN = MyRange.Characters(k).Font.Subscript
            '    Dim MyRangeN2 = MyRange.Characters(k)

            '    'obtener cadena substring de los caracteres subscript si:
            '    'solo hay una palabra
            '    'Y
            '    'pos ultimo caracter cadena-subscript coincide con pos ultimo caracter cadena 
            '    'Y 
            '    'pos primer caracter cadena-subscript coincide con pos ((ultimo caracter cadena - longitud(cadena-subscript)) +1

            '    'Entonces interpretar y convertir. Si resultado, sustituir. si no nada.
            'Next
            ''''''''''''''''''''''''''''SUBSCRIPT DETECT & PARSE

            If Cell.ContainsSubscript(WordCell) Then
                Cell.DisambiguateSubscript(WordCell)
            Else
                Cell.Text = WordCell.Range.Text.Trim(" ")

                'Eliminar STX, caracter especial referencia
                Cell.Text = Cell.Text.Replace(ChrW(2), "")

                'Eliminar BEL, caracter especial
                Cell.Text = Cell.Text.Replace(ChrW(7), "")

                'Eliminar vbCr, salto de linea
                Cell.Text = Cell.Text.Trim(vbCr)
            End If

            Cell.BgColor = WordCell.Range.Shading.BackgroundPatternColor
            Cells.Add(Cell)

            'Siguiente Celda
            WordCell = WordCell.[Next]
        End While
    End Sub

    ''' <summary>
    '''     Lee la tabla llamando a las funciones anetiores.
    '''     Asigna valor a los campos Cabecera (si existe) y Datos en el formato de salida.
    ''' </summary>
    ''' <param name="WordTable">
    '''     Microsoft.Office.Interop.Word.Table
    ''' </param>
    ''' <remarks></remarks>
    Public Sub ReadTable(ByRef WordTable As Table)
        'Obtenemos el nombre de la tabla
        Dim pName As String = WordTable.Cell(1, 1).Range.Previous(WdUnits.wdParagraph, 1).Text.Trim(vbCr)
        Name = pName.Replace(":", " -")
        Console.WriteLine(vbTab + vbTab + Name)

        'Obtenemos la lista con todas las celdas de la tabla
        ReadCells(WordTable)

        'Obtenemos la estructura de la tabla
        GetLayOut(WordTable)

        'Leemos los datos de la tabla [WordTable es la tabla del documento Word]
        With WordTable
            Dim numFilas, numColumnas As Integer
            Dim valor As String = Nothing
            Dim Cabecera As Boolean = HasHeader(WordTable)
            Dim RowItems As New Dictionary(Of String, Integer)()

            'Asignar tamaño para los datos
            numFilas = .Rows.Count
            numColumnas = .Columns.Count

            ManageDataSize(numFilas, Cabecera)

            'Rellenamos cabecera y datos 
            For i = 0 To numFilas - 1
                RowItems.Clear()

                'Recuperar el contenido de la celda teniendo en cuenta la estructura de la tabla
                For j = 0 To UBound([Structure], 2)
                    Dim indiceCelda As Integer = [Structure](i, j)

                    If indiceCelda <> -1 Then
                        Dim contenidoCelda As String = Cells(indiceCelda).Text.Trim(" "c)

                        'Celdas especiales 
                        If (String.IsNullOrEmpty(contenidoCelda) And Cells(indiceCelda).BgColor = WdColor.wdColorGray25) Then
                            contenidoCelda = "$TOP"
                        ElseIf String.IsNullOrEmpty(contenidoCelda) And Cells(indiceCelda).BgColor = WdColor.wdColorBlack Then
                            contenidoCelda = "$CRLF"
                        End If

                        'Contenido
                        If i = 0 And Cabecera Then
                            'Primera fila con cabecera
                            Me.HeaderRow.Add(contenidoCelda)
                        ElseIf i = 0 And Not Cabecera Then
                            'Sin Cabecera
                            If Not RepeatedRowValues Then
                                Data(i).Add(contenidoCelda)
                            Else
                                If RowItems.ContainsKey(contenidoCelda) Then
                                    RowItems(contenidoCelda) += 1
                                    Data(i).Add(contenidoCelda & " _[" & RowItems(contenidoCelda) & "]_")
                                Else
                                    RowItems.Add(contenidoCelda, 0)
                                    Data(i).Add(contenidoCelda)
                                End If
                            End If
                        ElseIf i > 0 And Cabecera Then
                            'Resto: Datos con Cabecera
                            If Not RepeatedRowValues Then
                                Data(i - 1).Add(contenidoCelda)
                            Else
                                If RowItems.ContainsKey(contenidoCelda) Then
                                    RowItems(contenidoCelda) += 1
                                    Data(i - 1).Add(contenidoCelda & " _[" & RowItems(contenidoCelda) & "]_")
                                Else
                                    RowItems.Add(contenidoCelda, 0)
                                    Data(i - 1).Add(contenidoCelda)
                                End If
                            End If
                        ElseIf i > 0 And Not Cabecera Then
                            'Resto: Datos sin Cabecera
                            If Not RepeatedRowValues Then
                                Data(i).Add(contenidoCelda)
                            Else
                                If RowItems.ContainsKey(contenidoCelda) Then
                                    RowItems(contenidoCelda) += 1
                                    Data(i).Add(contenidoCelda & " _[" & RowItems(contenidoCelda) & "]_")
                                Else
                                    RowItems.Add(contenidoCelda, 0)
                                    Data(i).Add(contenidoCelda)
                                End If
                            End If
                        End If
                    End If
                Next
            Next
        End With
    End Sub
End Class
