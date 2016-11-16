Imports Microsoft.Office.Interop.Word
Imports System.Collections
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>
'''     MODULO PRINCIPAL PARA EXTRAER LOS DATOS DEL DOCUMENTO WORD FUENTE CON LA ESPECIFICACION DE LA ONTOLOGIA
''' </summary>
Module TableDumper
    ''' <summary>
    '''     Metodo principal de la aplicacion.
    '''     Se debe llamar desde linea de comandos con los argumentos separados por espacios:
    '''         (1) Ruta del fichero Word a procesar.
    '''         (2) Nombre de la Ontologia.
    '''         (3) Indice de las tablas a procesar.
    '''         (4) "True" | "False" para que se procesen las notas a pie. 
    '''         (5) Ruta de salida de los ficheros procesados.
    ''' </summary>
    Sub Main()
        Dim objWord As Application
        Dim m As Integer = 0

        ' Impresion de Titulo y mensajes iniciales
        ConsoleApp.HandleTitle()
        ConsoleApp.HandleArgumentData()

        ' Argumentos pasados por linea de comandos
        Dim Arguments As String() = Environment.GetCommandLineArgs()
        Dim DocPath As String = Arguments(1)
        Dim Name As String = Arguments(2)
        Dim RangoTablas As Integer() = Util.GetParamFrom3rdArg(Arguments(3))
        Dim FootNotes As Boolean = Boolean.Parse(Arguments(4))
        Dim OutPath As String = Arguments(5)

        'Abrir Word
        objWord = CreateObject("Word.Application")

        ' Impresion de mensajes de procesamiento
        ConsoleApp.HandleProcessing()

        With objWord
            .Visible = False
            Dim objDoc As Document = .Documents.Open(DocPath)
            Dim OntologyData As New OntologyTables(Name)

            'Tablas
            If objDoc.Tables.Count > 0 And Not RangoTablas.Contains(-1) Then
                OntologyData.ReadTables(objDoc, RangoTablas)
            End If

            'Notas a Pie
            If FootNotes = True Then
                OntologyData.ReadFootNotes(objDoc)
            End If

            'Exportar ficheros
            OntologyData.ExportFile(OutPath)
            objDoc.Close()
            objDoc = Nothing

            ' Impresion de mensajes de fin
            ConsoleApp.HandleEnding()
            .Quit()
        End With
        objWord.Quit()
        objWord = Nothing
    End Sub
End Module
