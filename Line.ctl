VERSION 5.00
Begin VB.UserControl Line 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   315
      Left            =   3480
      Top             =   1200
      Width           =   345
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   345
      Left            =   360
      Top             =   1200
      Width           =   345
   End
   Begin VB.Line sLine 
      X1              =   600
      X2              =   3720
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Height          =   345
      Left            =   3480
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Height          =   345
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   345
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuMenuItem 
         Caption         =   "Move to Top"
         Index           =   0
      End
      Begin VB.Menu mnuMenuItem 
         Caption         =   "Move to Bottom"
         Index           =   1
      End
      Begin VB.Menu mnuMenuItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuMenuItem 
         Caption         =   "Remove"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'       NAME: Line.ctl
'    PURPOSE: This is a line control that allows the line to be moved resized
'       DATE: Thursday, June 9, 2003
'     AUTHOR: Neil Haverlandt
Option Explicit
Private MoveIt          As Boolean
Private OldX            As Integer
Private OldY            As Integer
Private ControlName     As String
Public Event ControlChanged()
Public Event Remove()
Public Event MoveToTop()
Public Event MoveToBottom()

Public Property Get DataField()

    DataField = ControlName

End Property

Public Property Let DataField(cName As Variant)

    ControlName = cName

End Property

Public Property Get hwnd()

    hwnd = UserControl.hwnd

End Property

Public Property Let IsVisible(Vis As Boolean)

    If Vis Then
        Label1.BackStyle = 1
        Label2.BackStyle = 1
        Shape1.BorderStyle = 1
        Shape2.BorderStyle = 1
     Else
        Label1.BackStyle = 0
        Label2.BackStyle = 0
        Shape1.BorderStyle = 0
        Shape2.BorderStyle = 0
    End If
    UserControl.Refresh

End Property

Private Sub Label1_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

    Select Case Shift
     Case 3 'Shift Ctrl
        Label2.Top = Label1.Top
        Shape2.Top = Shape1.Top
     Case 5 'Shift Alt
        Label2.Left = Label1.Left
        Shape2.Left = Shape1.Left
     Case Else
        OldX = x
        OldY = y
        If Button = 1 Then
            MoveIt = True
         ElseIf Button = 2 Then
            PopupMenu mnuMenu
        End If
        Exit Sub
    End Select
    With sLine
        .X1 = Label1.Left + 50
        .Y1 = Label1.Top + 50
        .X2 = Label2.Left + 50
        .Y2 = Label2.Top + 50
    End With 'sLine

End Sub

Private Sub Label1_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

    If MoveIt Then
        Select Case Shift
         Case 1 'Shift
            Label2.Top = Label2.Top + y - OldY
            Label2.Left = Label2.Left + x - OldX
            Label1.Top = Label1.Top + y - OldY
            Label1.Left = Label1.Left + x - OldX
            Shape2.Top = Shape2.Top + y - OldY
            Shape2.Left = Shape2.Left + x - OldX
            Shape1.Top = Shape1.Top + y - OldY
            Shape1.Left = Shape1.Left + x - OldX
         Case 4 'Ctrl
            Label1.Top = Label1.Top + y - OldY
            Shape1.Top = Shape1.Top + y - OldY
         Case 2 'Alt
            Label1.Left = Label1.Left + x - OldX
            Shape1.Left = Shape1.Left + x - OldX
         Case Else
            Label1.Top = Label1.Top + y - OldY
            Label1.Left = Label1.Left + x - OldX
            Shape1.Top = Shape1.Top + y - OldY
            Shape1.Left = Shape1.Left + x - OldX
        End Select
        With sLine
            .X1 = Label1.Left + 50
            .Y1 = Label1.Top + 50
            .X2 = Label2.Left + 50
            .Y2 = Label2.Top + 50
        End With 'sLine
        'UserControl.Refresh
        RaiseEvent ControlChanged
    End If

End Sub

Private Sub Label1_MouseUp(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)

    MoveIt = False

End Sub

Private Sub Label2_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

    Select Case Shift
     Case 3 'Shift Ctrl
        Label1.Top = Label2.Top
        Shape1.Top = Shape2.Top
     Case 5 'Shift Alt
        Label1.Left = Label2.Left
        Shape1.Left = Shape2.Left
     Case Else
        OldX = x
        OldY = y
        If Button = 1 Then
            MoveIt = True
         ElseIf Button = 2 Then
            PopupMenu mnuMenu
        End If
        Exit Sub
    End Select
    With sLine
        .X1 = Label1.Left + 50
        .Y1 = Label1.Top + 50
        .X2 = Label2.Left + 50
        .Y2 = Label2.Top + 50
    End With 'sLine

End Sub

Private Sub Label2_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

    If MoveIt Then
        Select Case Shift
         Case 1 'Shift
            Label2.Top = Label2.Top + y - OldY
            Label2.Left = Label2.Left + x - OldX
            Label1.Top = Label1.Top + y - OldY
            Label1.Left = Label1.Left + x - OldX
            Shape2.Top = Shape2.Top + y - OldY
            Shape2.Left = Shape2.Left + x - OldX
            Shape1.Top = Shape1.Top + y - OldY
            Shape1.Left = Shape1.Left + x - OldX
         Case 4 'Ctrl
            Label2.Top = Label2.Top + y - OldY
            Shape2.Top = Shape2.Top + y - OldY
         Case 2 'Alt
            Label2.Left = Label2.Left + x - OldX
            Shape2.Left = Shape2.Left + x - OldX
         Case Else
            Label2.Top = Label2.Top + y - OldY
            Label2.Left = Label2.Left + x - OldX
            Shape2.Top = Shape2.Top + y - OldY
            Shape2.Left = Shape2.Left + x - OldX
        End Select
        With sLine
            .X1 = Label1.Left + 50
            .Y1 = Label1.Top + 50
            .X2 = Label2.Left + 50
            .Y2 = Label2.Top + 50
        End With 'sLine
        'UserControl.Refresh
        RaiseEvent ControlChanged
    End If

End Sub

Private Sub Label2_MouseUp(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)

    MoveIt = False

End Sub

Private Sub mnuMenuItem_Click(Index As Integer)

    Select Case Index
     Case 0
        RaiseEvent MoveToTop
     Case 1
        RaiseEvent MoveToBottom
     Case 3
        RaiseEvent Remove
    End Select

End Sub

Private Sub UserControl_GotFocus()

    RaiseEvent ControlChanged

End Sub

Private Sub UserControl_Initialize()

    Label1.Move 100, 100, 100, 100
    Label2.Move 1000, 100, 100, 100
    Shape1.Move 100, 100, 100, 100
    Shape2.Move 1000, 100, 100, 100
    With sLine
        .X1 = Label1.Left + 50
        .Y1 = Label1.Top + 50
        .X2 = Label2.Left + 50
        .Y2 = Label2.Top + 50
    End With 'sLine

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)

    OldX = x
    OldY = y
    If Button = 1 Then
        MoveIt = True
     ElseIf Button = 2 Then
        PopupMenu mnuMenu
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)

    If MoveIt Then
        Select Case Shift
         Case 4 'Ctrl
            Label2.Top = Label2.Top + y - OldY
            Label1.Top = Label1.Top + y - OldY
            Label2.Top = Label2.Top + y - OldY
            Label1.Top = Label1.Top + y - OldY
            Shape2.Top = Shape2.Top + y - OldY
            Shape1.Top = Shape1.Top + y - OldY
            Shape2.Top = Shape2.Top + y - OldY
            Shape1.Top = Shape1.Top + y - OldY
         Case 2 'Alt
            Label2.Left = Label2.Left + x - OldX
            Label1.Left = Label1.Left + x - OldX
            Shape2.Left = Shape2.Left + x - OldX
            Shape1.Left = Shape1.Left + x - OldX
         Case Else
            Label2.Top = Label2.Top + y - OldY
            Label2.Left = Label2.Left + x - OldX
            Label1.Top = Label1.Top + y - OldY
            Label1.Left = Label1.Left + x - OldX
            Shape2.Top = Shape2.Top + y - OldY
            Shape2.Left = Shape2.Left + x - OldX
            Shape1.Top = Shape1.Top + y - OldY
            Shape1.Left = Shape1.Left + x - OldX
        End Select
        With sLine
            .X1 = Label1.Left + 50
            .Y1 = Label1.Top + 50
            .X2 = Label2.Left + 50
            .Y2 = Label2.Top + 50
        End With 'sLine
        OldX = x
        OldY = y
        'UserControl.Refresh
        RaiseEvent ControlChanged
    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

    MoveIt = False

End Sub

Public Property Get X1()

    X1 = sLine.X1

End Property

Public Property Let X1(x As Variant)

    sLine.X1 = x
    Label1.Left = x - 50
    Shape1.Left = x - 50

End Property

Public Property Get X2()

    X2 = sLine.X2

End Property

Public Property Let X2(x As Variant)

    sLine.X2 = x
    Label2.Left = x - 50
    Shape2.Left = x - 50

End Property

Public Property Get Y1()

    Y1 = sLine.Y1

End Property

Public Property Let Y1(y As Variant)

    sLine.Y1 = y
    Label1.Top = y - 50
    Shape1.Top = y - 50

End Property

Public Property Get Y2()

    Y2 = sLine.Y2

End Property

Public Property Let Y2(y As Variant)

    sLine.Y2 = y
    Label2.Top = y - 50
    Shape2.Top = y - 50

End Property

':) Roja's VB Code Fixer V1.0.99 (6/18/2003 8:12:40 PM) 14 + 330 = 344 Lines Thanks Ulli for inspiration and lots of code.

