VERSION 5.00
Begin VB.Form frmPaper 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Sheet Layout"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin PrintFormCodeCreator.pfTiResize Gripper 
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
   End
   Begin PrintFormCodeCreator.Line sLine 
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin VB.TextBox TextBox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox PictureBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   3600
      ScaleHeight     =   495
      ScaleWidth      =   1455
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmPaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'       NAME: frmPaper
'    PURPOSE: This is a graphical representation of a piece
'             of paper.  It is the form that controls will
'             be placed on to create a form.
'       DATE: Thursday, June 9, 2003
'     AUTHOR: Neil Haverlandt
'14 Jun - Rewrite of grid lines to aliviate ghost lines and blinking during refresh
Option Explicit
Public GridOn        As Boolean
Public BoxesOn       As Boolean
Public InchScale     As Boolean

Private Sub Form_Load()

    InchScale = True
    
    'make paper 20" x 20" so drawing will work the first time drawn
    Me.Show
    DoEvents
    With Me
        .Width = 28800
        .Height = 28800
'        .Left = ((Me.ScaleWidth - .Width) / 2)
        .GridOn = False
        .BoxesOn = False
        .Toggle_Grid
        .Toggle_Boxes
    End With 'Me
    DoEvents
End Sub

Private Sub Gripper_Remove(Control As Object)

    On Error Resume Next
    Unload Control
    Gripper.Visible = False
    On Error GoTo 0

End Sub

Private Sub Image1_Click(Index As Integer)

    Set Gripper.ResizeControl = Image1(Index)

End Sub

Private Sub PictureBox_MouseDown(Index As Integer, _
                                 Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 y As Single)

    Set Gripper.ResizeControl = PictureBox(Index)

End Sub

Private Sub sLine_ControlChanged(Index As Integer)

'SUSPEND_CODE_FIXER_ON
   Me.Refresh
'SUSPEND_CODE_FIXER_OFF

End Sub

Private Sub sLine_MoveToBottom(Index As Integer)

    sLine(Index).ZOrder 1

End Sub

Private Sub sLine_MoveToTop(Index As Integer)

    sLine(Index).ZOrder 0

End Sub

Private Sub sLine_Remove(Index As Integer)

    On Error Resume Next
    Unload sLine(Index)
    On Error GoTo 0

End Sub

Private Sub TextBox_MouseDown(Index As Integer, _
                              Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              y As Single)

    Set Gripper.ResizeControl = TextBox(Index)

End Sub

Public Sub Toggle_Boxes()

  Dim Control As VB.Control

    If BoxesOn Then
        BoxesOn = False
        Gripper.Visible = False
     Else
        BoxesOn = True
    End If
    For Each Control In frmPaper.Controls
        If Control.Name = "TextBox" Or Control.Name = "PictureBox" Then
            If Control.Index <> 0 Then
                If BoxesOn Then
                    Control.BorderStyle = 1
                 Else
                    Control.BorderStyle = 0
                End If
            End If
        End If
        If Control.Name = "sLine" Then
            If Control.Index <> 0 Then
                If BoxesOn Then
                    Control.IsVisible = True
                 Else
                    Control.IsVisible = False
                End If
            End If
        End If
    Next Control

End Sub

Public Sub Toggle_Grid()

  Dim I As Long

    'Dim II As Long
    On Error Resume Next
    If GridOn Then
        If InchScale Then
        '*******   Inch Scale Draw white
            frmPaper.ForeColor = vbWhite
            For I = 1 To Me.Width Step 720
                If I >= 360 Then
                    If I <= Me.Width - 360 Then
                        Me.DrawStyle = vbSolid
                        Me.Line (I, 360)-(I, Me.Height - 360), vbWhite
                        With frmPaper
                            .CurrentY = Me.Height - 200
                            .CurrentX = I - 100
                            frmPaper.Print CStr(Int(I / 1440))
                            .CurrentY = 0
                            .CurrentX = I - 100
                            frmPaper.Print CStr(Int(I / 1440))
                        End With 'frmPaper
                    End If
                End If
                I = I + 720
                If I >= 360 Then
                    With Me
                        If I <= .Width - 360 Then
                            .DrawStyle = vbDot
                            Me.Line (I, 360)-(I, Me.Height - 360), vbWhite
                        End If
                    End With 'Me
                End If
            Next I
            For I = 1 To Me.Height Step 720
                If I >= 360 Then
                    If I <= Me.Height - 360 Then
                        Me.DrawStyle = vbSolid
                        Me.Line (360, I)-(Me.Width - 360, I), vbWhite
                        With frmPaper
                            .CurrentY = I - 100
                            If Int(I / 1440) < 10 Then
                                .CurrentX = Me.Width - 100
                            Else
                                .CurrentX = Me.Width - 200
                            End If
                            frmPaper.Print CStr(Int(I / 1440))
                            .CurrentY = I - 100
                            .CurrentX = 0
                            frmPaper.Print CStr(Int(I / 1440))
                        End With 'frmPaper
                    End If
                End If
                I = I + 720
                If I >= 360 Then
                    With Me
                        If I <= .Height - 360 Then
                            .DrawStyle = vbDot
                            Me.Line (360, I)-(Me.Width - 360, I), vbWhite
                        End If
                    End With 'Me
                End If
            Next I
            '***** End InchScale Draw White
         Else
            '***** Start Centemeter Draw White
            frmPaper.ForeColor = vbWhite
            For I = 1 To Me.Width Step 285
                If I >= 360 Then
                    If I <= Me.Width - 360 Then
                        Me.DrawStyle = vbSolid
                        Me.Line (I, 360)-(I, Me.Height - 360), vbWhite
                        With frmPaper
                            .CurrentY = Me.Height - 200
                            .CurrentX = I - 100
                            frmPaper.Print CStr(Int(I / 570))
                            .CurrentY = 0
                            .CurrentX = I - 100
                            frmPaper.Print CStr(Int(I / 570))
                        End With 'frmPaper
                    End If
                End If
                I = I + 285
                If I >= 360 Then
                    With Me
                        If I <= .Width - 360 Then
                            .DrawStyle = vbDot
                            Me.Line (I, 360)-(I, Me.Height - 360), vbWhite
                        End If
                    End With 'Me
                End If
            Next I
            For I = 1 To Me.Height Step 285
                If I >= 360 Then
                    If I <= Me.Height - 360 Then
                        Me.DrawStyle = vbSolid
                        Me.Line (360, I)-(Me.Width - 360, I), vbWhite
                        With frmPaper
                            .CurrentY = I - 100
                            If Int(I / 570) < 10 Then
                                .CurrentX = Me.Width - 100
                            Else
                                .CurrentX = Me.Width - 200
                            End If
                            frmPaper.Print CStr(Int(I / 570))
                            .CurrentY = I - 100
                            .CurrentX = 0
                            frmPaper.Print CStr(Int(I / 570))
                        End With 'frmPaper
                    End If
                End If
                I = I + 285
                If I >= 360 Then
                    With Me
                        If I <= .Height - 360 Then
                            .DrawStyle = vbDot
                            Me.Line (360, I)-(Me.Width - 360, I), vbWhite
                        End If
                    End With 'Me
                End If
            Next I
            '***** End Centemeter Scale Draw White
        End If
        '***** Draw Margin Box white
        Me.DrawStyle = vbDot
        Me.Line (360, 360)-(Me.Width - 360, Me.Height - 360), vbWhite, B
        GridOn = False
     Else
        If InchScale Then
            frmPaper.ForeColor = vbBlack
            For I = 1 To Me.Width Step 720
                If I >= 360 Then
                    If I <= Me.Width - 360 Then
                        Me.DrawStyle = vbSolid
                        Me.Line (I, 360)-(I, Me.Height - 360), &HE0E0E0
                        With frmPaper
                            .CurrentY = Me.Height - 200
                            .CurrentX = I - 100
                            frmPaper.Print CStr(Int(I / 1440))
                            .CurrentY = 0
                            .CurrentX = I - 100
                            frmPaper.Print CStr(Int(I / 1440))
                        End With 'frmPaper
                    End If
                End If
                I = I + 720
                If I >= 360 Then
                    With Me
                        If I <= .Width - 360 Then
                            .DrawStyle = vbDot
                            Me.Line (I, 360)-(I, Me.Height - 360), &HE0E0E0
                        End If
                    End With 'Me
                End If
            Next I
            For I = 1 To Me.Height Step 720
                If I >= 360 Then
                    If I <= Me.Height - 360 Then
                        Me.DrawStyle = vbSolid
                        Me.Line (360, I)-(Me.Width - 360, I), &HE0E0E0
                        With frmPaper
                            .CurrentY = I - 100
                            If Int(I / 1440) < 10 Then
                                .CurrentX = Me.Width - 100
                            Else
                                .CurrentX = Me.Width - 200
                            End If
                            frmPaper.Print CStr(Int(I / 1440))
                            .CurrentY = I - 100
                            .CurrentX = 0
                            frmPaper.Print CStr(Int(I / 1440))
                        End With 'frmPaper
                    End If
                End If
                I = I + 720
                If I >= 360 Then
                    With Me
                        If I <= .Height - 360 Then
                            .DrawStyle = vbDot
                            Me.Line (360, I)-(Me.Width - 360, I), &HE0E0E0
                        End If
                    End With 'Me
                End If
            Next I
         Else
            frmPaper.ForeColor = vbBlack
            For I = 1 To Me.Width Step 285
                If I >= 360 Then
                    If I <= Me.Width - 360 Then
                        Me.DrawStyle = vbSolid
                        Me.Line (I, 360)-(I, Me.Height - 360), &HE0E0E0
                        With frmPaper
                            .CurrentY = Me.Height - 200
                            .CurrentX = I - 100
                            frmPaper.Print CStr(Int(I / 570))
                            .CurrentY = 0
                            .CurrentX = I - 100
                            frmPaper.Print CStr(Int(I / 570))
                        End With 'frmPaper
                    End If
                End If
                I = I + 285
                If I >= 360 Then
                    With Me
                        If I <= .Width - 360 Then
                            .DrawStyle = vbDot
                            Me.Line (I, 360)-(I, Me.Height - 360), &HE0E0E0
                        End If
                    End With 'Me
                End If
            Next I
            For I = 1 To Me.Height Step 285
                If I >= 360 Then
                    If I <= Me.Height - 360 Then
                        Me.DrawStyle = vbSolid
                        Me.Line (360, I)-(Me.Width - 360, I), &HE0E0E0
                        With frmPaper
                            .CurrentY = I - 100
                            If Int(I / 570) < 10 Then
                                .CurrentX = Me.Width - 100
                            Else
                                .CurrentX = Me.Width - 200
                            End If
                            frmPaper.Print CStr(Int(I / 570))
                            .CurrentY = I - 100
                            .CurrentX = 0
                            frmPaper.Print CStr(Int(I / 570))
                        End With 'frmPaper
                    End If
                End If
                I = I + 285
                If I >= 360 Then
                    With Me
                        If I <= .Height - 360 Then
                            .DrawStyle = vbDot
                            Me.Line (360, I)-(Me.Width - 360, I), &HE0E0E0
                        End If
                    End With 'Me
                End If
            Next I
        End If
        'Draw grid
        Me.DrawStyle = vbDot
        Me.Line (360, 360)-(Me.Width - 360, Me.Height - 360), vbBlack, B
        GridOn = True
    End If
    On Error GoTo 0

End Sub

':) Roja's VB Code Fixer V1.0.99 (6/18/2003 8:12:38 PM) 12 + 303 = 315 Lines Thanks Ulli for inspiration and lots of code.

