VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FireWorks"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Public xMin, xMax, yMin, yMax As Long

Const PI = 3.14159265358979 'Without this forget it,
                            'necessary in all math with angles
Dim mEnd As Boolean
Dim CurrentPoint As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  'Exit Program
  mEnd = True
  '------------
End Sub

Private Sub Form_Load()
'Run all preparation steps before jumping into the main loop
Setup
Calculate
'-----------------------------------------------------------
Draw
End Sub

Sub Setup()

  Me.ScaleMode = vbPixels
  Me.BackColor = vbBlack
  Me.DrawWidth = 1
  
  mEnd = False
  
  'Center Form
  Dim CenterX As Long, CenterY As Long

  CenterX = ((Screen.Width / Screen.TwipsPerPixelX) / 2) - _
             (Me.ScaleWidth / 2)
  CenterY = ((Screen.Height / Screen.TwipsPerPixelY) / 2) - _
             (Me.ScaleHeight / 2)

 'Left and Top uses twips, convert values to twips
  Me.Left = ScaleX(CenterX, vbPixels, vbTwips)
  Me.Top = ScaleY(CenterY, vbPixels, vbTwips)
  '------------------------------------------------
  '-----------
  
  Me.Show

End Sub

Sub Calculate()

Dim mRand As Long
Dim tmp As Long
  
For tmp = 1 To UBound(Points)
  Points(tmp).Pos(1).x = Me.ScaleWidth / 2
  Points(tmp).Pos(1).y = Me.ScaleHeight - 1
  Points(tmp).Pos(1).Angle = 90
  Points(tmp).Pos(1).Length = 0
  Points(tmp).Pos(1).Position = 2
  
  'Set AngleChange
  mRand = Randomizer(41)
  
  Select Case mRand 'Tosses fourtyone sided dice
    Case 1
      Points(tmp).Pos(1).AngleChange = -0.66
    Case 2
      Points(tmp).Pos(1).AngleChange = -0.63
    Case 3
      Points(tmp).Pos(1).AngleChange = -0.6
    Case 4
      Points(tmp).Pos(1).AngleChange = -0.56
    Case 5
      Points(tmp).Pos(1).AngleChange = -0.53
    Case 6
      Points(tmp).Pos(1).AngleChange = -0.5
    Case 7
      Points(tmp).Pos(1).AngleChange = -0.46
    Case 8
      Points(tmp).Pos(1).AngleChange = -0.43
    Case 9
      Points(tmp).Pos(1).AngleChange = -0.4
    Case 10
      Points(tmp).Pos(1).AngleChange = -0.36
    Case 11
      Points(tmp).Pos(1).AngleChange = -0.33
    Case 12
      Points(tmp).Pos(1).AngleChange = -0.3
    Case 13
      Points(tmp).Pos(1).AngleChange = -0.26
    Case 14
      Points(tmp).Pos(1).AngleChange = -0.23
    Case 15
      Points(tmp).Pos(1).AngleChange = -0.2
    Case 16
      Points(tmp).Pos(1).AngleChange = -0.16
    Case 17
      Points(tmp).Pos(1).AngleChange = -0.13
    Case 18
      Points(tmp).Pos(1).AngleChange = -0.1
    Case 19
      Points(tmp).Pos(1).AngleChange = -0.06
    Case 20
      Points(tmp).Pos(1).AngleChange = -0.03
    Case 21
      Points(tmp).Pos(1).AngleChange = 0
    Case 22
      Points(tmp).Pos(1).AngleChange = 0.03
    Case 23
      Points(tmp).Pos(1).AngleChange = 0.06
    Case 24
      Points(tmp).Pos(1).AngleChange = 0.1
    Case 25
      Points(tmp).Pos(1).AngleChange = 0.13
    Case 26
      Points(tmp).Pos(1).AngleChange = 0.16
    Case 27
      Points(tmp).Pos(1).AngleChange = 0.2
    Case 28
      Points(tmp).Pos(1).AngleChange = 0.23
    Case 29
      Points(tmp).Pos(1).AngleChange = 0.26
    Case 30
      Points(tmp).Pos(1).AngleChange = 0.3
    Case 31
      Points(tmp).Pos(1).AngleChange = 0.33
    Case 32
      Points(tmp).Pos(1).AngleChange = 0.36
    Case 33
      Points(tmp).Pos(1).AngleChange = 0.4
    Case 34
      Points(tmp).Pos(1).AngleChange = 0.43
    Case 35
      Points(tmp).Pos(1).AngleChange = 0.46
    Case 36
      Points(tmp).Pos(1).AngleChange = 0.5
    Case 37
      Points(tmp).Pos(1).AngleChange = 0.53
    Case 38
      Points(tmp).Pos(1).AngleChange = 0.56
    Case 39
      Points(tmp).Pos(1).AngleChange = 0.6
    Case 40
      Points(tmp).Pos(1).AngleChange = 0.63
    Case 41
      Points(tmp).Pos(1).AngleChange = 0.66
  End Select
  '---------------

  'Set Color
  mRand = Randomizer(6)
  
  Select Case mRand 'Tosses four dice
    Case 1
      Points(tmp).Pos(1).Color = vbRed
    Case 2
      Points(tmp).Pos(1).Color = vbGreen
    Case 3
      Points(tmp).Pos(1).Color = vbBlue
    Case 4
      Points(tmp).Pos(1).Color = vbYellow
    Case 5
      Points(tmp).Pos(1).Color = vbWhite
    Case 6
      Points(tmp).Pos(1).Color = vbCyan
  End Select
  '---------
Next

End Sub

'My function let me code more efficient by shortening overall code length
Function Randomizer(mCount As Long) As Long
  Randomize
  
  Dim tmp As Long
  tmp = (Rnd * (mCount - 1)) + 1
  Randomizer = tmp
End Function
'------------------------------------------------------------------------

Sub Draw()
On Error Resume Next

'Set Values for firework area
xMin = (Me.ScaleWidth / 7) * 2
xMax = (Me.ScaleWidth / 7) * 5

yMin = Me.ScaleHeight
yMax = (Me.ScaleHeight / 2)
'----------------------------

Dim Position As Long

Position = 2


Dim tmp As Long
Do
  
  Dim PointsTmp As Long
    
  For tmp = 1 To UBound(Points) 'Go through all points
        
    Position = Points(tmp).Pos(1).Position
    
    Points(tmp).Pos(Position).Angle = Points(tmp).Pos(Position - 1).Angle + _
                                      Points(tmp).Pos(1).AngleChange
                                        
    xAngle = Cos((Points(tmp).Pos(Position).Angle) * (PI / 180))
    yAngle = Sin((Points(tmp).Pos(Position).Angle) * (PI / 180))
          
    Points(tmp).Pos(Position).x = Points(tmp).Pos(Position - 1).x + xAngle
    Points(tmp).Pos(Position).y = Points(tmp).Pos(Position - 1).y - yAngle
    Points(tmp).Pos(1).Length = Points(tmp).Pos(1).Length + 1
    
    'When Point reaches max value clear all point data (black)
    If Points(tmp).Pos(Position).y < yMax And _
       Points(tmp).Pos(Position).y > yMax - 1.1 Or _
       Points(tmp).Pos(Position).x < xMin And _
       Points(tmp).Pos(Position).x > xMin - 1.1 Or _
       Points(tmp).Pos(Position).x > xMax And _
       Points(tmp).Pos(Position).x < xMax + 1.1 Or _
       Points(tmp).Pos(Position).y > yMin And _
       Points(tmp).Pos(Position).y < yMin + 1.1 Then

        For PointsTmp = 1 To Points(tmp).Pos(1).Length
          SetPixelV Me.hdc, Points(tmp).Pos(PointsTmp).x, Points(tmp).Pos(PointsTmp).y, vbBlack
        Next

    End If
    '---------------------------------------------------------

    'Clears out point before and set new point with color (make no tail), single pixel
    If Points(tmp).Pos(Position).y < yMax - 1.1 Or _
       Points(tmp).Pos(Position).x < xMin - 1.1 Or _
       Points(tmp).Pos(Position).x > xMax + 1.1 Or _
       Points(tmp).Pos(Position).y > yMin + 1.1 Then
       
       SetPixelV Me.hdc, Points(tmp).Pos(Position - 1).x, Points(tmp).Pos(Position - 1).y, vbBlack
       SetPixelV Me.hdc, Points(tmp).Pos(Position).x, Points(tmp).Pos(Position).y, _
                                              Points(tmp).Pos(1).Color
    End If
    '---------------------------------------------------------------------------------
      
    'Do this if point hasnt reached max values yet.
    If Points(tmp).Pos(Position).y > yMax And _
       Points(tmp).Pos(Position).y < yMin Or _
       Points(tmp).Pos(Position).x > xMin And _
       Points(tmp).Pos(Position).x < xMax Then
      
      SetPixelV Me.hdc, Points(tmp).Pos(Position).x, Points(tmp).Pos(Position).y, _
                                         Points(tmp).Pos(1).Color
    End If
    '----------------------------------------------
    
    'When Point reaches form boarders clear point and generate a new one
    If Points(tmp).Pos(Position).y > Me.ScaleHeight Or _
       Points(tmp).Pos(Position).x > Me.ScaleWidth Or _
       Points(tmp).Pos(Position).x < -1 Or _
       Points(tmp).Pos(Position).y < -1 Then
    
      CurrentPoint = tmp
      Point_NewData
    End If
    '-------------------------------------------------------------------
    
    Points(tmp).Pos(1).Position = Points(tmp).Pos(1).Position + 1
  Next
      
DoEvents
Loop Until mEnd = True
Err.Clear
Unload Me
End
End Sub

Sub Point_NewData()
  
  Points(CurrentPoint).Pos(1).Position = 1
  
  'Calculate new data for firework (new random)
  Points(CurrentPoint).Pos(1).x = Me.ScaleWidth / 2
  Points(CurrentPoint).Pos(1).y = Me.ScaleHeight - 1
  Points(CurrentPoint).Pos(1).Angle = 90
  Points(CurrentPoint).Pos(1).Length = 1
  
  'Set AngleChange
  mRand = Randomizer(41)
  
  Select Case mRand 'Tosses fourtyone sided dice
    Case 1
      Points(CurrentPoint).Pos(1).AngleChange = -0.66
    Case 2
      Points(CurrentPoint).Pos(1).AngleChange = -0.63
    Case 3
      Points(CurrentPoint).Pos(1).AngleChange = -0.6
    Case 4
      Points(CurrentPoint).Pos(1).AngleChange = -0.56
    Case 5
      Points(CurrentPoint).Pos(1).AngleChange = -0.53
    Case 6
      Points(CurrentPoint).Pos(1).AngleChange = -0.5
    Case 7
      Points(CurrentPoint).Pos(1).AngleChange = -0.46
    Case 8
      Points(CurrentPoint).Pos(1).AngleChange = -0.43
    Case 9
      Points(CurrentPoint).Pos(1).AngleChange = -0.4
    Case 10
      Points(CurrentPoint).Pos(1).AngleChange = -0.36
    Case 11
      Points(CurrentPoint).Pos(1).AngleChange = -0.33
    Case 12
      Points(CurrentPoint).Pos(1).AngleChange = -0.3
    Case 13
      Points(CurrentPoint).Pos(1).AngleChange = -0.26
    Case 14
      Points(CurrentPoint).Pos(1).AngleChange = -0.23
    Case 15
      Points(CurrentPoint).Pos(1).AngleChange = -0.2
    Case 16
      Points(CurrentPoint).Pos(1).AngleChange = -0.16
    Case 17
      Points(CurrentPoint).Pos(1).AngleChange = -0.13
    Case 18
      Points(CurrentPoint).Pos(1).AngleChange = -0.1
    Case 19
      Points(CurrentPoint).Pos(1).AngleChange = -0.06
    Case 20
      Points(CurrentPoint).Pos(1).AngleChange = -0.03
    Case 21
      Points(CurrentPoint).Pos(1).AngleChange = 0
    Case 22
      Points(CurrentPoint).Pos(1).AngleChange = 0.03
    Case 23
      Points(CurrentPoint).Pos(1).AngleChange = 0.06
    Case 24
      Points(CurrentPoint).Pos(1).AngleChange = 0.1
    Case 25
      Points(CurrentPoint).Pos(1).AngleChange = 0.13
    Case 26
      Points(CurrentPoint).Pos(1).AngleChange = 0.16
    Case 27
      Points(CurrentPoint).Pos(1).AngleChange = 0.2
    Case 28
      Points(CurrentPoint).Pos(1).AngleChange = 0.23
    Case 29
      Points(CurrentPoint).Pos(1).AngleChange = 0.26
    Case 30
      Points(CurrentPoint).Pos(1).AngleChange = 0.3
    Case 31
      Points(CurrentPoint).Pos(1).AngleChange = 0.33
    Case 32
      Points(CurrentPoint).Pos(1).AngleChange = 0.36
    Case 33
      Points(CurrentPoint).Pos(1).AngleChange = 0.4
    Case 34
      Points(CurrentPoint).Pos(1).AngleChange = 0.43
    Case 35
      Points(CurrentPoint).Pos(1).AngleChange = 0.46
    Case 36
      Points(CurrentPoint).Pos(1).AngleChange = 0.5
    Case 37
      Points(CurrentPoint).Pos(1).AngleChange = 0.53
    Case 38
      Points(CurrentPoint).Pos(1).AngleChange = 0.56
    Case 39
      Points(CurrentPoint).Pos(1).AngleChange = 0.6
    Case 40
      Points(CurrentPoint).Pos(1).AngleChange = 0.63
    Case 41
      Points(CurrentPoint).Pos(1).AngleChange = 0.66
  End Select
  '---------------
          
  'Set Color
  mRand = Randomizer(6)
  
  Select Case mRand 'Tosses four dice
    Case 1
      Points(CurrentPoint).Pos(1).Color = vbRed
    Case 2
      Points(CurrentPoint).Pos(1).Color = vbGreen
    Case 3
      Points(CurrentPoint).Pos(1).Color = vbBlue
    Case 4
      Points(CurrentPoint).Pos(1).Color = vbYellow
    Case 5
      Points(CurrentPoint).Pos(1).Color = vbWhite
    Case 6
      Points(CurrentPoint).Pos(1).Color = vbCyan
  End Select
  '---------

End Sub

Private Sub Form_Unload(Cancel As Integer)
mEnd = True
End Sub
