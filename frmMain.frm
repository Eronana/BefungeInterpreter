VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Befunge -- 一个更艹蛋的语言 By:笹瀬川佐々美"
   ClientHeight    =   4932
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   12636
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4932
   ScaleWidth      =   12636
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Run Code"
      Height          =   612
      Left            =   6000
      TabIndex        =   4
      Top             =   360
      Width           =   2532
   End
   Begin VB.Frame Frame2 
      Caption         =   "Console"
      Height          =   3492
      Left            =   5520
      TabIndex        =   2
      Top             =   1320
      Width           =   6972
      Begin VB.TextBox txtConsole 
         Height          =   3132
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   6732
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Code"
      Height          =   4692
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5292
      Begin VB.TextBox txtCode 
         Height          =   4212
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   5052
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function fuckBefunge(code As String)
    Dim bStack() As Long
    Dim sCode() As String
    Dim lStrLen() As Long, sUBound As Long
    Dim IP As POINT
    Dim SP As Long, drc As Direction
    Dim c As String
    Dim moved As Boolean
    Dim vSwap As Long
    Dim i As Long
    
    drc = D_RIGHT
    SP = 100
    sCode = Split(code, vbCrLf)
    sUBound = UBound(sCode)
    ReDim lStrLen(sUBound)
    For i = 0 To sUBound
        lStrLen(i) = Len(sCode(i))
    Next
    
    ReDim bStack(10000)
    
    Do
        If IP.Y > sUBound Then IP.Y = 0
        If IP.Y < 0 Then IP.Y = 0
        If IP.X < 0 Then IP.X = 0
        If IP.X >= lStrLen(IP.Y) Then IP.X = 0
        c = Mid(sCode(IP.Y), IP.X + 1, 1)
        moved = False
        Select Case c
            Case "0" To "9"
                SP = SP + 1
                bStack(SP) = Val(c)
            Case "+"
                bStack(SP - 1) = bStack(SP - 1) + bStack(SP)
                SP = SP - 1
            Case "-"
                bStack(SP - 1) = bStack(SP - 1) - bStack(SP)
                SP = SP - 1
            Case "*"
                bStack(SP - 1) = bStack(SP - 1) * bStack(SP)
                SP = SP - 1
            Case "/"
                bStack(SP - 1) = Int(bStack(SP - 1) / bStack(SP))
                SP = SP - 1
            Case "%"
                bStack(SP - 1) = Int(bStack(SP - 1) Mod bStack(SP))
                SP = SP - 1
            Case "!"
                bStack(SP) = IIf(bStack(SP) = 0, 1, 0)
            Case "`"
                bStack(SP - 1) = IIf(bStack(SP - 1) > bStack(SP), 1, 0)
                SP = SP - 1
            Case ">"
                drc = D_RIGHT
            Case "<"
                drc = D_LEFT
            Case "^"
                drc = D_UP
            Case "v"
                drc = D_DOWN
            Case "?"
                drc = Int(Rnd() * 4)
            Case "_"
                drc = IIf(bStack(SP) = 0, D_RIGHT, D_LEFT)
                SP = SP - 1
            Case "|"
                drc = IIf(bStack(SP) = 0, D_DOWN, D_UP)
                SP = SP - 1
            Case """"
                Do
                    moveIP IP, drc
                    c = Mid(sCode(IP.Y), IP.X + 1, 1)
                    If c = """" Then Exit Do
                    SP = SP + 1
                    bStack(SP) = Asc(c)
                Loop
            Case ":"
                SP = SP + 1
                bStack(SP) = bStack(SP - 1)
            Case "\"
                vSwap = bStack(SP)
                bStack(SP) = bStack(SP - 1)
                bStack(SP - 1) = vSwap
            Case "$"
                SP = SP - 1
            Case "."
                putInt (bStack(SP))
                SP = SP - 1
            Case ","
                putchar (bStack(SP))
                SP = SP - 1
            Case "#"
                moveIP IP, drc
            Case "p"
                Mid(sCode(bStack(SP)), bStack(SP - 1) + 1, 1) = Chr(bStack(SP - 2))
                SP = SP - 3
            Case "g"
                bStack(SP - 1) = Asc(Mid(sCode(bStack(SP)), bStack(SP - 1) + 1, 1))
                SP = SP - 1
            Case "&"
                SP = SP + 1
                bStack(SP) = getInt()
            Case "~"
                SP = SP + 1
                bStack(SP) = getchar()
            Case "@"
                Exit Do
        End Select
        If Not moved Then moveIP IP, drc
    Loop
End Function

Sub putchar(ByVal c As Byte)
    writeLog Chr(c)
End Sub
Sub putInt(ByVal c)
    writeLog Str(c)
End Sub
Function getchar() As Byte
'整个程序,就这里写的白痴了..
'不过小程序而已..这样偷懒也没什么问题
    Dim l As Long
    l = Len(txtConsole.Text)
    txtConsole.Locked = False
    txtConsole.SetFocus
    Do Until Len(txtConsole.Text) > l
        DoEvents
    Loop
    txtConsole.Locked = True
    getchar = Asc(Mid(txtConsole.Text, l + 1, 1))
    If getchar = 13 Then getchar = 10
End Function
Function getInt() As Byte
    Dim l As Long
    l = Len(txtConsole.Text)
    txtConsole.Locked = False
    txtConsole.SetFocus
    Do Until Len(txtConsole.Text) > l
        DoEvents
    Loop
    txtConsole.Locked = True
    getInt = Int(Mid(txtConsole.Text, l + 1))
End Function
Sub writeLog(s As String)
    txtConsole.SelStart = Len(txtConsole.Text)
    txtConsole.SelText = s
End Sub

Private Sub Command1_Click()
    fuckBefunge txtCode.Text
End Sub

Private Sub Form_Load()
    Randomize
End Sub

