'' sub_Class.vb
'' © Шульжицкий Владимир, 2014 (boxa.shu@gmail.com)

Public Class X_kolvo
    Public x As Double
    Public kol As Integer

End Class

Public Class GL_POLY
    Public x As Double
    Public y As Double
    Public z As Double
End Class


Public Class GL_KNOT
    Public N As Integer
    Public x As Double
    Public y As Double
    Public z As Double
End Class

Public Class GL_ELEM
    Public d1 As Integer

    Public N As Integer
    Public p1 As Integer
    Public p2 As Integer
    Public p3 As Integer
    Public p4 As Integer

    Public Asn_X As Double
    Public Asn_Y As Double
    Public Asn_Z As Double

    Public Asv_X As Double
    Public Asv_Y As Double
    Public Asv_Z As Double
End Class

Public Class GL_ARM
    Public d1 As Integer
    Public d2 As Integer

    Public x As Double
    Public y As Double
    Public z As Double

    Public Asn_X As Double
    Public Asn_Y As Double
    Public Asn_Z As Double

    Public Asv_X As Double
    Public Asv_Y As Double
    Public Asv_Z As Double
End Class