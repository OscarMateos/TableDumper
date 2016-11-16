Public Class Util
    ''' <summary>
    '''     Devuelve un array de enteros con la sucesion de numeros desde LowBound hasta UpBound, 
    '''     siendo (LowBound menor o igual que UpBound).
    ''' </summary>
    ''' <param name="LowBound">
    '''     Integer, limite inferior.
    ''' </param>
    ''' <param name="UpBound">
    '''     Integer, limite superior.
    ''' </param>
    ''' <returns>
    '''     Integer() 
    ''' </returns>
    Public Shared Function Seq(ByVal LowBound As Integer, ByVal UpBound As Integer) As Integer()
        If LowBound = UpBound Then
            Return New Integer() {LowBound}
        Else
            Dim Size As Integer = UpBound - LowBound
            Dim Sequence(Size) As Integer

            For i = 0 To Size
                Sequence(i) = LowBound + i
            Next

            Return Sequence
        End If
    End Function

    ''' <summary>
    '''     Funcion para devolver el minimo entero de un array.
    ''' </summary>
    ''' <param name="args">
    '''     Integer, array de enteros a encontrar el menor.
    ''' </param>
    ''' <returns>
    '''     Integer
    ''' </returns>
    Public Shared Function FindMin(ByVal ParamArray args() As Integer) As Integer
        Dim myMin As Double
        Dim i As Long
        myMin = Integer.MaxValue
        For i = LBound(args) To UBound(args)
            If args(i) < myMin Then
                myMin = args(i)
            End If
        Next i
        FindMin = myMin
    End Function

    ''' <summary>
    '''  Funcion para devolver el maximo entero de un array.
    ''' </summary>
    ''' <param name="args">
    '''     Integer, array de enteros a encontrar el mayor.
    ''' </param>
    ''' <returns>
    '''     Integer
    ''' </returns>
    Public Shared Function FindMax(ByVal ParamArray args() As Integer) As Integer
        Dim myMax As Double
        Dim i As Long
        myMax = 0
        For i = LBound(args) To UBound(args)
            If args(i) > myMax Then
                myMax = args(i)
            End If
        Next i
        FindMax = myMax
    End Function

    ''' <summary>
    '''     Funcion para obtener el valor del parametro RangoTablas (indices de las tablas que se 
    '''     porcesaran de un documento Word).
    '''     
    '''     Es el tercer argumento pasado por linea de comandos en la llamada a la aplicacion. 
    ''' </summary>
    ''' <param name="ThirdArg"></param>
    ''' <returns>
    '''     Integer(),
    '''     Si ThirdArg = "notab"; RangoTablas = {-1}
    '''     Si ThirdArg contiene un intervalo[a,a]; RangoTablas = {a}
    '''     Si ThirdArg contiene un intervalo[a,b]; RangoTablas = {a ... b}
    '''     Si ThirdArg contiene una lista de valores (x,y,z,...); RangoTablas = {x,y,z,...}
    '''     
    '''     Nothing en caso contrario.
    ''' </returns>
    Public Shared Function GetParamFrom3rdArg(ByVal ThirdArg As String) As Integer()
        Dim Parts As String()

        'Determinar tipo
        If ThirdArg.ToLower.Equals("notab") Then
            Return Seq(-1, -1)
        End If
        If ThirdArg.Contains("-") Then
            'Intervalo
            Parts = ThirdArg.Replace(" ", "").Split("-")
            Dim LowBound As Integer, UpBound As Integer

            If Integer.TryParse(Parts(0), LowBound) And Integer.TryParse(Parts(1), UpBound) Then
                Return Seq(LowBound, UpBound)
            Else
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("[ERROR]" + vbTab + "Intervalo de indices de tabla no valido ( " + ThirdArg + ")")
                Console.ForegroundColor = ConsoleColor.White
                Return Nothing
            End If
        Else
            'Lista
            Parts = ThirdArg.Replace(" ", "").Split(",")
            Dim Size As Integer = Parts.Length
            Dim Sequence(Size - 1) As Integer

            For i = 0 To Size - 1
                Dim Value As Integer
                If Integer.TryParse(Parts(i), Value) Then
                    Sequence(i) = Value
                Else
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("[ERROR]" + vbTab + "Indice de tabla no valido ( " + Parts(i) + " )")
                    Console.ForegroundColor = ConsoleColor.White
                    Return Nothing
                End If
            Next

            Return Sequence
        End If
    End Function
End Class
