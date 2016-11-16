''' <summary>
'''     Clase tPoint para las coordenadas de las celdas de las tablas Word (x,y)
''' </summary>
''' <remarks></remarks>
Public Class tPoint
    Private dimX, dimY As Integer

    ''' <summary>
    '''     Getters y Setters para la propiedad X
    ''' </summary>
    ''' <returns></returns>
    Public Property X As Integer
        Get
            Return dimX
        End Get
        Set(value As Integer)
            dimX = value
        End Set
    End Property

    ''' <summary>
    '''     Getters y Setters para la propiedad Y
    ''' </summary>
    ''' <returns></returns>
    Public Property Y As Integer
        Get
            Return dimY
        End Get
        Set(value As Integer)
            dimY = value
        End Set
    End Property

    ''' <summary>
    '''     Constructor
    ''' </summary>
    ''' <param name="_x"></param>
    ''' <param name="_y"></param>
    Public Sub New(_x As Integer, _y As Integer)
        X = _x
        Y = _y
    End Sub
End Class
