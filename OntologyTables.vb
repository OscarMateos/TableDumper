Imports Microsoft.Office.Interop.Word
Imports System.Collections
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>
'''     Clase OntologyTables
''' </summary>
''' <remarks></remarks>
Public Class OntologyTables
    Property Name As String
    Property TablesIndex As Integer()
    Property OntologyTables As MyWordTable()
    Property FootNotes As HashSet(Of String())

    ''' <summary>
    '''     Constructor, hay que pasarle el nombre de la Ontologia como parametro
    ''' </summary>
    ''' <param name="Name">
    '''     String, nombre de la Ontologia.
    ''' </param>
    Sub New(ByVal Name As String)
        Me.Name = Name
    End Sub


    ''' <summary>
    '''     Se filtran las tablas no válidas que estén en el rango de tablas seleccionado 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SelectTables(ByRef objDoc As Document, ByVal ParamArray RangoTablas() As Integer)
        Dim Picking As New ArrayList(RangoTablas)

        With objDoc
            Dim UpperLimit As Integer = Picking.Count - 1

            For i = 0 To UpperLimit
                If i > UpperLimit Or i < 0 Or UpperLimit < 0 Then
                    Exit For
                End If

                'Si el nombre contiene el literal "axiom" es tabla de axiomas, se omite
                Dim Name As String = .Tables(Picking(i)).Cell(1, 1).Range.Previous(WdUnits.wdParagraph, 1).Text.Trim(vbCr)

                'Si son tablas con horientacion vertical (tablas de axiomas) las omitimos
                Dim Orientation As WdTextOrientation = .Tables(Picking(i)).Range.Orientation

                'Si son tablas de imágenes o gráficos las omitimos
                Dim FirstCell As String = .Tables(Picking(i)).Cell(1, 1).Range.Text

                'Se suprimen caracteres especiales y/o de control
                FirstCell = FirstCell.Remove(FirstCell.IndexOf(vbCr), FirstCell.Length - FirstCell.IndexOf(vbCr))

                'Se determina si es tabla válida: No tabla de axiomas ni tabla de figuras o imágenes
                If Name.Contains("axiom") Or Orientation = WdTextOrientation.wdTextOrientationUpward Or FirstCell = "/" Or String.IsNullOrEmpty(FirstCell.Replace("", "")) Then
                    Picking.Remove(Picking(i))
                    i = i - 1
                    UpperLimit = UpperLimit - 1
                End If
            Next
        End With

        'Se guardan los indices de las tablas omitiendo las de axiomas
        Me.TablesIndex = CType(Picking.ToArray(GetType(Integer)), Integer())
    End Sub

    ''' <summary>
    '''     Submetodo para leer el contenido de las tablas de un documento fuente Word.
    ''' </summary>
    ''' <param name="objDoc">
    '''     Referencia al objeto del documento Word origen.
    ''' </param>
    ''' <param name="TablesIndex">
    '''     Indices de las tablas del documento a extraer separadas por comas.
    '''      0 -> Todas las tablas del documento.
    '''     -1 -> Ninguna tabla del documento.
    ''' </param>
    ''' <remarks></remarks>
    Public Sub ReadTables(ByRef objDoc As Document, ByVal ParamArray TablesIndex() As Integer)
        'Se determina y comprueba RangoTablas
        If TablesIndex.Count = 1 And TablesIndex(0) = 0 Then
            TablesIndex = Util.Seq(1, objDoc.Tables.Count)
        Else
            For i = 0 To UBound(TablesIndex) + 1
                If TablesIndex(i) = 0 Or TablesIndex(i) > objDoc.Tables.Count Then
                    System.Console.WriteLine("[ERROR]   Indice de tabla fuera de rango ( " + TablesIndex(i).ToString + " )")
                    Exit Sub
                End If
            Next
        End If

        'Se filtran las tablas no válidas que estén en el rango de tablas seleccionado 
        SelectTables(objDoc, TablesIndex)
        ReDim OntologyTables(Me.TablesIndex.Count - 1)

        'Procesado de las tablas
        Console.WriteLine(vbTab + "LEYENDO TABLAS:")
        For i = 0 To UBound(Me.TablesIndex)
            Dim Table As New MyWordTable
            Table.ReadTable(objDoc.Tables(Me.TablesIndex(i)))
            OntologyTables(i) = Table
        Next
    End Sub

    ''' <summary>
    '''     Submetodo para leer las notas a pie de un documento fuente Word junto con sus referencias.
    ''' </summary>
    ''' <param name="objDoc">
    '''     Microsoft.Office.Interop.Word.Document
    ''' </param>
    Public Sub ReadFootNotes(ByRef objDoc As Document)
        Console.WriteLine(vbCrLf + vbTab + "LEYENDO NOTAS A PIE")
        Dim Data As HashSet(Of String()) = New HashSet(Of String())
        Dim NotesCount As Integer = objDoc.Footnotes.Count

        If NotesCount = 0 Then
            Console.WriteLine(vbTab + vbTab + "El documento no contiene notas a pie de pagina")
        Else
            For index As Integer = 1 To NotesCount
                Dim FN As Footnote = objDoc.Footnotes(index)

                'Entrada de nueva Nota para el conjunto de Notas
                Dim d(3) As String

                'Extraccion de la nota
                Dim FootNoteText As String = FN.Range.Text

                'Extraccion de las referencias
                Dim FootNoteReference As String = Nothing, NewReference As String = Nothing
                Dim R As Range = FN.Reference
                Dim WithinTableIndex As Integer = R.Tables.Count

                If WithinTableIndex <> 0 Then
                    'Si la referencia está dentro de una tabla, es directo y será todo el texto de la celda
                    FootNoteReference = R.Cells(WithinTableIndex).Range.Text
                    'Eliminar STX, caracter especial referencia
                    FootNoteReference = FootNoteReference.Replace(ChrW(2), "")
                    'Eliminar vbCr, salto de linea
                    FootNoteReference = FootNoteReference.Replace(vbCr, "")
                    'Eliminar BEL, caracter especial
                    FootNoteReference = FootNoteReference.Replace(ChrW(7), "")

                    'Tercer campo para marcar que no requiere procesamiento del texto de la referencia posterior
                    d(1) = "false"
                Else
                    'Si no hay que extraer la frase que contiene la referencia
                    R.MoveStartUntil(Cset:=".", Count:=WdConstants.wdBackward)
                    FootNoteReference = objDoc.Range(R.Start, R.End + 1).Text

                    'Eliminar STX, caracter especial referencia
                    FootNoteReference = FootNoteReference.Replace(ChrW(2), "")
                    'Eliminar vbCr, salto de linea
                    FootNoteReference = FootNoteReference.Replace(vbCr, "")
                    'Eliminar BEL, caracter especial
                    FootNoteReference = FootNoteReference.Replace(ChrW(7), "")
                    'Eliminar espacios en blanco al principio y final
                    FootNoteReference = FootNoteReference.Trim(" "c, "."c, ","c, ";"c)

                    'Acotamos el texto de la referencia para reducir ambiguedad al extraer NEs
                    Dim idxAnd, idxOr, idxAndOr, idxComma, idxSemiColon As Integer

                    idxAnd = FootNoteReference.LastIndexOf(" and ")
                    idxOr = FootNoteReference.LastIndexOf(" or ")
                    idxAndOr = FootNoteReference.LastIndexOf(" and/or ")
                    idxComma = FootNoteReference.LastIndexOf(",")
                    idxSemiColon = FootNoteReference.LastIndexOf(";")

                    Dim newStart As Integer = Util.FindMax(idxAnd, idxOr, idxAndOr, idxComma, idxSemiColon)
                    NewReference = FootNoteReference.Substring(newStart).TrimEnd(","c) + "."

                    'Tercer campo para marcar que no requiere procesamiento del texto de la referencia posterior
                    d(1) = "true"
                End If

                'Se anade nueva nota a pie al conjunto de notas
                d(0) = index
                d(2) = If(IsNothing(NewReference), FootNoteReference, NewReference)
                d(3) = FootNoteText
                Data.Add(d)
            Next

            FootNotes = Data
        End If
    End Sub

    ''' <summary>
    '''     Submetodo para exportar los datos extraidos y transformados a ficheros de texto.
    '''     Se creara un fichero por cada tabla y otro para todas las notas a pie.
    ''' </summary>
    ''' <param name="OutPath">
    '''     String, ruta de salida de los ficheros.
    ''' </param>
    Public Async Sub ExportFile(ByVal OutPath As String)
        If Not Dir(OutPath, vbDirectory) = vbNullString Then
            My.Computer.FileSystem.CreateDirectory(OutPath & "\" & Me.Name)

            Dim rgxVertTab As New Regex("\s*")
            Dim rgxRepId As New Regex("\s+_\[\d+\]_")
            Dim rgxFileName As New Regex("\s*|\\+|/+|:+\s*|\*+|\?+|<+|>+")

            Dim FileName = Nothing, Directory = Nothing, FullPath As String
            Dim sb As New StringBuilder()

            'TABLAS
            If Not IsNothing(OntologyTables) Then
                For Each T As MyWordTable In OntologyTables
                    'Nombre
                    sb.AppendLine(T.Name.Trim(" "c, vbCr))

                    'Cabecera
                    If Not IsNothing(T.HeaderRow) Then
                        For Each valor As String In T.HeaderRow
                            sb.Append(valor.Trim(" "c, vbCr) & "~~")
                        Next
                    End If
                    sb.AppendLine()

                    'Datos
                    For Each linea As HashSet(Of String) In T.Data
                        For Each Value As String In linea
                            If Not Value.StartsWith(vbCr) And Not Value.EndsWith(vbCr) Then
                                Value = Value.Replace(vbCr, "%")
                            End If

                            If Value.Contains("_[") Then
                                Value = rgxRepId.Replace(Value, "")
                            End If

                            sb.Append(rgxVertTab.Replace(Value, "%").Trim(" "c, vbCr) & "~~")
                        Next
                        sb.AppendLine()
                    Next

                    'Se escribe el fichero
                    FileName = rgxFileName.Replace(T.Name.Trim(" "c, vbCr), " - ") & ".txt"
                    Directory = rgxFileName.Replace(Me.Name, " - ")
                    FullPath = OutPath & "\" & Directory & "\" & FileName
                    Using outfile As New StreamWriter(FullPath, False)
                        Await outfile.WriteAsync(sb.ToString())
                    End Using
                    sb.Clear()
                Next
            End If

            'NOTAS A PIE 
            If Not IsNothing(FootNotes) Then
                'Nombre
                Dim Nombre As String = "Table FN - FootNotes related to " + Me.Name.Trim(" "c, vbCr)
                sb.AppendLine(Nombre)

                'Cabecera
                sb.AppendLine()

                'Datos
                For Each d As String() In FootNotes
                    ' Si hay salto de linea en el texto de la nota es multi-nota
                    If Not IsNothing(d(3)) Then
                        If d(3).Trim(" "c, vbCr).Contains(vbCr) Then
                            d(3) = d(3).Replace(vbCr, "%")
                        End If
                    End If

                    For Each Value As String In d
                        If Not IsNothing(Value) Then
                            sb.Append(Value.Trim(" "c, vbCr) & "~~")
                        Else
                            sb.Append((" " & "~~"))
                        End If
                    Next
                    sb.AppendLine()
                Next

                'Se escribe el fichero
                FileName = rgxFileName.Replace(Nombre.Trim(" "c, vbCr), " - ") & ".txt"
                Directory = rgxFileName.Replace(Me.Name, " - ")
                FullPath = OutPath & "\" & Directory & "\" & FileName
                Using outfile As New StreamWriter(FullPath, False)
                    Await outfile.WriteAsync(sb.ToString())
                End Using
                sb.Clear()
            End If
        Else
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("[ERROR]" + vbTab + "No existe el directorio ( " + OutPath + " )")
            Console.ForegroundColor = ConsoleColor.White
            Exit Sub
        End If
    End Sub
End Class
