''' <summary>
'''     Clase tPointEqualityComparer para la definir por sobrecarga la comparación  
'''     de puntos para evitar inserción de repetidos en el HashSet 
''' </summary>
''' <remarks></remarks>
Class tPointEqualityComparer
    Implements IEqualityComparer(Of tPoint)

    ''' <summary>
    '''    Sobrecarga del metodo generico de igualdad para puntos
    ''' </summary>
    ''' <param name="b1"></param>
    ''' <param name="b2"></param>
    ''' <returns></returns>
    Public Overloads Function Equals(ByVal b1 As tPoint, ByVal b2 As tPoint) _
                   As Boolean Implements IEqualityComparer(Of tPoint).Equals

        If b1.X = b2.X And b1.Y = b2.Y Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    '''     Sobrecarga del metodo generico de obtencion del codigo hash para insertar en HashSet para puntos
    ''' </summary>
    ''' <param name="bx"></param>
    ''' <returns></returns>
    Public Overloads Function GetHashCode(ByVal bx As tPoint) _
                As Integer Implements IEqualityComparer(Of tPoint).GetHashCode
        Dim hCode As Integer = bx.X Xor bx.Y
        Return hCode.GetHashCode()
    End Function
End Class