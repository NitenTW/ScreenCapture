Public Class Screen

#Region "各種設定"
    Private _X As Integer
    Private _Y As Integer
    Public _screenBMP As Bitmap
    Public _selectBMP As Bitmap

    Public WriteOnly Property X() As Integer
        Set(value As Integer)
            _X = value
        End Set
    End Property

    Public WriteOnly Property Y() As Integer
        Set(value As Integer)
            _Y = value
        End Set
    End Property

    Public Property ScreenBMP() As Bitmap
        Get
            Return _screenBMP
        End Get
        Set(value As Bitmap)
            _screenBMP = value
        End Set
    End Property

    Public Property SelectBMP() As Bitmap
        Get
            Return _selectBMP
        End Get
        Set(value As Bitmap)
            _selectBMP = value
        End Set
    End Property
#End Region

    ''' <summary>
    ''' 關閉時釋放資源
    ''' </summary>
    Public Sub Close()
        If Not IsNothing(screenBMP) Then screenBMP.Dispose()
        If Not IsNothing(selectBMP) Then selectBMP.Dispose()
    End Sub

    Public Function MakeRect(ByVal X As Integer, ByVal Y As Integer) As Rectangle
        Dim tmpX As Integer = Me._X
        Dim tmpY As Integer = Me._Y

        X = FixZero(FixX(X))
        Y = FixZero(FixY(Y))

        Swap(tmpX, X)
        Swap(tmpY, Y)

        Return New Rectangle(tmpX, tmpY, X - tmpX, Y - tmpY)
    End Function

    ''' <summary>
    ''' 修正小於零
    ''' </summary>
    Private Function FixZero(ByVal s As Integer) As Integer
        If s < 0 Then Return 0
        Return s
    End Function

    ''' <summary>
    ''' 修正 X 值大於畫面長度
    ''' </summary>
    Private Function FixX(ByVal s As Integer) As Integer
        If s > screenBMP.Width Then Return screenBMP.Width
        Return s
    End Function

    ''' <summary>
    ''' 修正 Y 值大於畫面長度
    ''' </summary>
    Private Function FixY(ByVal s As Integer) As Integer
        If s > screenBMP.Height Then Return screenBMP.Height
        Return s
    End Function

    ''' <summary>
    ''' 如果 X 值大於 Y 值，則互相交換
    ''' </summary>
    Private Sub Swap(ByRef x As Integer, ByRef y As Integer)
        If x > y Then
            Dim t As Integer = x
            x = y
            y = t
        End If
    End Sub
End Class