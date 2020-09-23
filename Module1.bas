Attribute VB_Name = "Module1"
Option Base 1

Public xAngle, yAngle As Double

Private Type mPOS
  'Double and single gives presicly calculations
  x As Double
  y As Double
  Color As Long
  Angle As Single
  AngleChange As Single
  Length As Long
  Position As Long
End Type

Private Type mPoints
  Pos(1 To 700) As mPOS
End Type

Public Points(100) As mPoints

Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


