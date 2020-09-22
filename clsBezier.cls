VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBezier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Code by Cyborg (contact at vbforums.com)
' 30/3 - 2009

Option Explicit

Private Type Point2D
    X As Double
    Y As Double
End Type

Dim Points() As Point2D

Public Sub InitCurve(numVerts As Long)
    ReDim Points(numVerts - 1)  'set size of array
End Sub

Public Sub RemoveCurve()
    Erase Points 'free mem
End Sub

'find a point in close proximity to a coordinate
Public Function FindPoint(ByVal X As Double, ByVal Y As Double, Precision As Double) As Long
    Dim i As Long
    
    For i = 0 To UBound(Points)
        With Points(i)
            If X > .X - Precision And X < .X + Precision And Y > .Y - Precision And Y < .Y + Precision Then
                FindPoint = i
                Exit Function
            End If
        End With
    Next
    FindPoint = -1
End Function

'return values
Public Function GetPointX(i As Long) As Double
    GetPointX = Points(i).X
End Function

Public Function GetPointY(i As Long) As Double
    GetPointY = Points(i).Y
End Function

Public Function GetNumVerts() As Long
    GetNumVerts = UBound(Points) + 1
End Function

Public Sub SetPointCoords(i As Long, ByVal X As Double, ByVal Y As Double)
    Points(i).X = X
    Points(i).Y = Y
End Sub

'returns x and y coords for the curve
Public Sub GetBezierPoint(t As Double, ByRef X As Double, ByRef Y As Double)
    Dim Point As Point2D
    Point = TraceCurve(Points, UBound(Points) + 1, t)
    X = Point.X
    Y = Point.Y
End Sub

'recursive function to trace out the curve.
'for each call it connects the lines from the previous call at distance t from each point.
'each call will produce one less point(or line), and it will keep calling itself until there's only one point left.
'the last point will finally be returned
Private Function TraceCurve(P() As Point2D, ByVal numPoints As Long, t As Double) As Point2D
    Dim i As Long
    Dim newP() As Point2D 'holds the new points
    ReDim newP(numPoints - 2)
    
    For i = 0 To numPoints - 2
        'get a new point at fraction t between 2 other points
        newP(i).X = t * (P(i + 1).X - P(i).X) + P(i).X
        newP(i).Y = t * (P(i + 1).Y - P(i).Y) + P(i).Y
    Next
    
    If numPoints > 2 Then
        'send the new points to TraceCurve
        TraceCurve = TraceCurve(newP, numPoints - 1, t)
    ElseIf numPoints = 2 Then
        'if the function is at it's last stop i.e. there is only one new point.
        'return that point instead of calling TraceCurve again
        TraceCurve = newP(0)
    End If
    
    Erase newP 'free mem
End Function
































