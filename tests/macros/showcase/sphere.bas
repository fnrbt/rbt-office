' Source: https://junmitani.hatenablog.com/entry/20130614 (post title: "the wrong way to use Excel")
' Author: Jun Mitani, on his Hatena blog
' Licence: none stated
' What: ray-traced red sphere with diffuse shading and a specular highlight, painted into 101x101 cells, one cell per pixel.
' Entry: drawSphere
' Before: Setup
' Run: the code does not resize cells; make columns and rows small and square first, or the sphere comes out stretched.
' Uses a Private Type (UDT) Vec3 passed ByRef, Sqr, ^ 20, ByVal Integer parameters given Double values (implicit rounding),
' Cells.Clear, Cells.Interior.Color, Cells(r, c).Interior.Color = RGB(...). No animation.
' Code below is verbatim from the source.
Option Explicit

Private Type Vec3
    x As Double
    y As Double
    z As Double
End Type

Sub VectorNormalize(ByRef vec As Vec3)
    Dim length As Double
    length = Sqr(getDotProduct(vec, vec))
    vec.x = vec.x / length
    vec.y = vec.y / length
    vec.z = vec.z / length
End Sub

Function getDotProduct(ByRef vec0 As Vec3, ByRef vec1 As Vec3) As Double
    getDotProduct = vec0.x * vec1.x + vec0.y * vec1.y + vec0.z * vec1.z
End Function

Sub drawPixcel(ByVal x As Integer, ByVal y As Integer, ByVal origin As Integer, _
    ByVal color_R As Integer, ByVal color_G As Integer, ByVal color_B As Integer)
    Dim px As Integer
    Dim py As Integer
    
    px = x + origin
    py = y + origin
    If color_R > 255 Then color_R = 255
    If color_R < 0 Then color_R = 0
    If color_G > 255 Then color_G = 255
    If color_G < 0 Then color_G = 0
    If color_B > 255 Then color_B = 255
    If color_B < 0 Then color_B = 0
    If px > 0 And py > 0 Then Cells(py, px).Interior.Color = RGB(color_R, color_G, color_B)        
End Sub

Sub drawSphere()
    Cells.Clear
    Cells.Interior.Color = RGB(255, 255, 255)
    
    Dim Radius As Integer
    Radius = 50
    
    Dim lightVec As Vec3
    lightVec.x = 1
    lightVec.y = -1
    lightVec.z = 0
    VectorNormalize lightVec
    
    Dim rayVec As Vec3
    rayVec.x = 0
    rayVec.y = 0
    rayVec.z = 1
    
    Dim normalVec As Vec3
    Dim refrectedVec As Vec3
    Dim y As Integer
    Dim x As Integer
    For y = -Radius To Radius
        For x = -Radius To Radius
            If x * x + y * y < Radius * Radius Then
                Dim z As Double
                z = Sqr(Radius * Radius - x * x - y * y)
                normalVec.x = x
                normalVec.y = y
                normalVec.z = z
                VectorNormalize normalVec
                
                Dim intensity As Double
                intensity = 0.5 * getDotProduct(lightVec, normalVec)
                
                refrectedVec.x = normalVec.x * 2 - lightVec.x
                refrectedVec.y = normalVec.y * 2 - lightVec.y
                refrectedVec.z = normalVec.z * 2 - lightVec.z
                VectorNormalize refrectedVec
                
                Dim intensity2 As Double
                intensity2 = getDotProduct(refrectedVec, rayVec)
                If intensity2 > 0 Then
                    intensity2 = intensity2 ^ 20
                    intensity = intensity + intensity2
                End If
                
                Call drawPixcel(x, y, Radius, 128 + intensity * 128, intensity * 128, intensity * 128)
            End If
        Next
    Next
End Sub
