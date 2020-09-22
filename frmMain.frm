VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bezier Curves"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   10920
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkPoints 
      Caption         =   "Show Points"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   7680
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkLines 
      Caption         =   "Show Lines"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   7440
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.PictureBox picDisp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   479
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   711
      TabIndex        =   0
      Top             =   120
      Width           =   10695
   End
   Begin VB.Label lblNumPoints 
      Caption         =   "3"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "NumPoints:"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   7440
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code by Cyborg (contact at vbforums.com)
' 30/3 - 2009

Option Explicit

Dim Curve As clsBezier
Dim SelectedPoint As Long

Private Sub Render()
    Dim t As Double
    Dim X As Double
    Dim Y As Double
    Dim i As Long
    
    picDisp.Cls
    
    For t = 0 To 1 Step 1 / 1000    '1000 iterations (increase for a smoother curve)
        Curve.GetBezierPoint t, X, Y 'find the positions for the curve
        picDisp.PSet (X, Y), vbRed  'draw curve
    Next
    
    'draw handles
    If chkPoints.Value Then
        For i = 0 To Curve.GetNumVerts - 1
            picDisp.Circle (Curve.GetPointX(i), Curve.GetPointY(i)), 3, vbBlue
        Next
    End If
    
    'draw lines
    If chkLines.Value Then
        For i = 0 To Curve.GetNumVerts - 2
            picDisp.Line (Curve.GetPointX(i), Curve.GetPointY(i))-(Curve.GetPointX(i + 1), Curve.GetPointY(i + 1)), vbBlue
        Next
    End If
    
    picDisp.Refresh
End Sub

Private Sub chkLines_Click()
    Render
End Sub

Private Sub chkPoints_Click()
    Render
End Sub

Private Sub Form_Load()
    Set Curve = New clsBezier
    
    'initiate a curve with 3 points
    Curve.InitCurve 3
    
    'set initial positions for the points
    Curve.SetPointCoords 0, 50, 200
    Curve.SetPointCoords 1, 20, 50
    Curve.SetPointCoords 2, 350, 160
    
    Render
    
    Me.Show
    MsgBox "Drag the handles with the left mouse button" & vbNewLine & "Click with the right mouse button on points to split them", vbOKOnly, "Bezier Curves"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'free up memory
    Curve.RemoveCurve
    Set Curve = Nothing
End Sub

Private Sub picDisp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then
        SelectedPoint = Curve.FindPoint(X, Y, 5) 'find which point we are close to clicking on
        picDisp_MouseMove Button, Shift, X, Y
    End If
End Sub

Private Sub picDisp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If SelectedPoint > -1 Then
            Curve.SetPointCoords SelectedPoint, X, Y 'move the selected point
            Render
        End If
    End If
End Sub

Private Sub picDisp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tempX() As Double
    Dim tempY() As Double
    Dim i As Long
    Dim numVerts As Long
    
    If Button = 2 Then  'add new vertecies by creating a new line and replacing the old one
        If SelectedPoint > 0 And SelectedPoint < Curve.GetNumVerts - 1 Then
            numVerts = Curve.GetNumVerts
            ReDim tempX(numVerts - 1) 'for storing the old coords
            ReDim tempY(numVerts - 1)
            
            'store old coords
            For i = 0 To numVerts - 1
                tempX(i) = Curve.GetPointX(i)
                tempY(i) = Curve.GetPointY(i)
            Next
            
            'we're adding two new points
            numVerts = numVerts + 2
            
            're-initialize the curve, since we're adding points
            Curve.InitCurve numVerts
            
            'set coords for all points before the addition
            For i = 0 To SelectedPoint - 1
                Curve.SetPointCoords i, tempX(i), tempY(i)
            Next
            
            'calculate and set the new coords
            Curve.SetPointCoords SelectedPoint, (tempX(SelectedPoint) - tempX(SelectedPoint - 1)) / 2 + tempX(SelectedPoint - 1), _
                                                (tempY(SelectedPoint) - tempY(SelectedPoint - 1)) / 2 + tempY(SelectedPoint - 1)
            
            'this is the same as the one which was clicked on, except it's placement in the array.
            Curve.SetPointCoords SelectedPoint + 1, tempX(SelectedPoint), tempY(SelectedPoint)
            
            Curve.SetPointCoords SelectedPoint + 2, (tempX(SelectedPoint + 1) - tempX(SelectedPoint)) / 2 + tempX(SelectedPoint), _
                                                    (tempY(SelectedPoint + 1) - tempY(SelectedPoint)) / 2 + tempY(SelectedPoint)
            
            'set coords for all points after the addition
            For i = SelectedPoint + 3 To numVerts - 1
                Curve.SetPointCoords i, tempX(i - 2), tempY(i - 2)
            Next
            
            lblNumPoints.Caption = numVerts
            
            SelectedPoint = -1
            
            Render
        End If
    End If
    
    Erase tempX
    Erase tempY
End Sub



















