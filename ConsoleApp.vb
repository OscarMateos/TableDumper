''' <summary>
'''     MODULO PARA GESTIONAR LA IMPRESION DE MENSAJES DE LA INTERFAZ PARA LINEA DE COMANDOS
''' </summary>
Public Class ConsoleApp
    ''' <summary>
    '''     Imprime una linea de guiones horizontal a lo largo de todo el ancho de pantalla 
    ''' </summary>
    Public Shared Sub PrintHorizontalLine()
        For i = 1 To Console.WindowWidth
            Console.Write("-")
        Next
    End Sub

    ''' <summary>
    '''     Gestiona el titulo de la ventana, las dimensiones, posicion inicial, e imprime titulo
    ''' </summary>
    Public Shared Sub HandleTitle()
        Dim Width As Integer = 3 * Console.LargestWindowWidth / 4
        Dim Height As Integer = 3 * Console.LargestWindowHeight / 4
        Console.WindowHeight = Height
        Console.WindowWidth = Width
        Console.SetBufferSize(Width, Height)
        Console.SetWindowPosition(0, 0)

        Dim title As String = "WORD DOCUMENT TABLE DUMPER"
        Console.Clear()
        PrintHorizontalLine()
        Console.CursorLeft = Console.WindowWidth / 2 - title.Length / 2
        Console.Title = title
        Console.WriteLine(title)
        PrintHorizontalLine()
    End Sub

    ''' <summary>
    '''     Imprime y valida el valor de los argumentos de entrada pasados por linea de comandos a la aplicacion
    ''' </summary>
    Public Shared Sub HandleArgumentData()
        Dim Arguments As String() = Environment.GetCommandLineArgs()
        If Arguments.Length <> 6 Then
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("[ERROR]" + vbTab + "PARAMETROS INCORRECTOS.")
            Console.WriteLine("Llame a la aplicacion con los siguientes parametros:" + vbCrLf +
                    vbTab + "(1) Ruta del fichero Word a procesar." + vbCrLf +
                    vbTab + "(2) Nombre de la Ontologia." + vbCrLf +
                    vbTab + "(3) Indice de las tablas a procesar:" + vbCrLf +
                    vbTab + "     - Se permiten intervalos separados por guion (""-"")." + vbCrLf +
                    vbTab + "     - ""notab"" si solo se desean extraer las notas a pie." + vbCrLf +
                    vbTab + "     - 0 si se desean extraer todas las tablas." + vbCrLf +
                    vbTab + "     - -1 si no se desea extraer ninguna tabla." + vbCrLf +
                    vbTab + "(4) ""True"" | ""False"" para que se procesen las notas a pie." + vbCrLf +
                    vbTab + "(5) Ruta de salida de los ficheros procesados." + vbCrLf +
                    vbCrLf + "SINTAXIS:" + vbCrLf +
                    vbTab + ".> tabledumper [1] [2] [3] [4] [5]")
            Console.ReadKey()
            Environment.Exit(-1)
        Else
            Console.WriteLine("DATOS DEL TRABAJO:")

            If My.Computer.FileSystem.FileExists(Arguments(1)) Then
                Console.WriteLine(vbTab + "Fichero de entrada: " + Arguments(1))
            Else
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("[ERROR]" + vbTab + "Fichero de entrada no encontrado ( " + Arguments(1) + " )")
                Console.ReadKey()
                Environment.Exit(-1)
            End If

            Console.WriteLine(vbTab + "Nombre de la Ontologia: " + Arguments(2))

            Dim intValue As Integer
            If Integer.TryParse(Arguments(3), intValue) And intValue = 0 Then
                Console.WriteLine(vbTab + "Tablas a procesar: TODAS las tablas del documento.")
            ElseIf Arguments(3).ToLower.Equals("notab") Then
                Console.WriteLine(vbTab + "Tablas a procesar: NINGUNA tabla del documento.")
            ElseIf Not IsNothing(Util.GetParamFrom3rdArg(Arguments(3))) Then
                Dim TableList As String = Arguments(3).Replace(" ", "").Replace(",", ", ")
                Console.WriteLine(vbTab + "Tablas a procesar: " + TableList)
            Else
                Console.ReadKey()
                Environment.Exit(-1)
            End If
        End If

        Dim FootNotes As Boolean
        If Boolean.TryParse(Arguments(4), FootNotes) Then
            Console.WriteLine(vbTab + "Procesar notas a pie: " + FootNotes.ToString)
        Else
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("[ERROR]" + vbTab + "El cuarto parametro debe ser ""True""|""False""")
            Console.ReadKey()
            Environment.Exit(-1)
        End If

        If My.Computer.FileSystem.DirectoryExists(Arguments(5)) Then
            Console.WriteLine(vbTab + "Ruta de salida: " + Arguments(5))
        Else
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("[ERROR]" + vbTab + "Ruta de salida no encontrada ( " + Arguments(5) + " )")
            Console.ReadKey()
            Environment.Exit(-1)
        End If

        PrintHorizontalLine()
    End Sub

    ''' <summary>
    '''     Imprime el inicio de la seccion de procesamiento principal
    ''' </summary>
    Public Shared Sub HandleProcessing()
        Console.WriteLine("PROCESANDO DOCUMENTO, ESPERE, PUEDE TARDAR VARIOS MINUTOS:")
    End Sub

    ''' <summary>
    '''     Imprime la seccion de finalizacion, emite un sonido y espera para salir
    ''' </summary>
    Public Shared Sub HandleEnding()
        PrintHorizontalLine()
        Console.WriteLine("EXTRACCION FINALIZADA")
        PrintHorizontalLine()
        Console.Beep()
        Console.ReadKey()
        Environment.Exit(-1)
    End Sub
End Class
