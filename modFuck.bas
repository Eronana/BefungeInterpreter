Attribute VB_Name = "modFuck"
Option Explicit

Enum Direction
    D_UP = 0
    D_DOWN = 1
    D_LEFT = 2
    D_RIGHT = 3
End Enum
Type POINT
    X As Long
    Y As Long
End Type

Sub moveIP(IP As POINT, drc As Direction)
    Select Case drc
        Case D_UP
            IP.Y = IP.Y - 1
        Case D_DOWN
            IP.Y = IP.Y + 1
        Case D_LEFT
            IP.X = IP.X - 1
        Case D_RIGHT
            IP.X = IP.X + 1
    End Select
End Sub
