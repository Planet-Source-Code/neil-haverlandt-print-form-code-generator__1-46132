VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Print Form Code Creator"
   ClientHeight    =   8820
   ClientLeft      =   -15
   ClientTop       =   195
   ClientWidth     =   11970
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer timControl 
      Interval        =   100
      Left            =   480
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8820
      Left            =   10050
      ScaleHeight     =   8820
      ScaleWidth      =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   1920
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   22
         Top             =   6720
         Width           =   1455
         Begin VB.OptionButton optOrient 
            Caption         =   "Portrait"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optOrient 
            Caption         =   "Landscape"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Orientation"
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.Frame fraScale 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   5880
         Width           =   1455
         Begin VB.OptionButton optScale 
            Caption         =   "Centimeter"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optScale 
            Caption         =   "Inches"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Grid Scale"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.ComboBox cboPaperSize 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Text            =   "PaperSize"
         Top             =   5520
         Width           =   1575
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox txtHeightY2 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox txtWidthX2 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txtLeftX1 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtTopY1 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Paper Size"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label lblLeftX1 
         Caption         =   "Left/X1"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblWidthX2 
         Caption         =   "Width/X2"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblHeightY2 
         Caption         =   "Height/Y2"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label lblPath 
         Caption         =   "Picture Path"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label lblTopY1 
         Caption         =   "Top/Y1"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblType 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Control Type"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Control Name"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":015E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":02B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":03F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1048
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D54
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3092
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   3  'Align Left
      Height          =   8820
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   15558
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TextBox"
            Description     =   "Text Box"
            Object.ToolTipText     =   "Text Box"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "StraitLine"
            Object.ToolTipText     =   "Line"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PictureBox"
            Object.ToolTipText     =   "Picture Box"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
            Object.Width           =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Create Code"
            Object.ToolTipText     =   "Generate Code"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grid"
            Object.ToolTipText     =   "Toggle Grid"
            ImageIndex      =   5
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Boxes"
            Object.ToolTipText     =   "Toggle Place Holders"
            ImageIndex      =   11
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Close"
            Object.ToolTipText     =   "New"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Load"
            Object.ToolTipText     =   "Load Form"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Form"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'       NAME: frmMain
'    PURPOSE: This is the main form that will control most aspects of the program
'       DATE: Thursday, June 9, 2003
'     AUTHOR: Neil Haverlandt
' 14 Jun 2003 Allow changing of paper size
Option Explicit
Private ControlNumber        As Integer
Private m_ActiveControl      As VB.Control
Private CurrentFileName      As String
Private Portrait             As Boolean

Private Sub cboPaperSize_Click()

    Get_PaperSize

End Sub

Private Sub optOrient_Click(Index As Integer)
    frmPaper.GridOn = True
    frmPaper.Toggle_Grid
    If optOrient(0) Then
        Portrait = True
    Else
        Portrait = False
    End If
    Get_PaperSize
    'frmPaper.Toggle_Grid
End Sub

Private Sub optScale_Click(Index As Integer)

    frmPaper.GridOn = True
    frmPaper.Toggle_Grid
    If optScale(0) Then
        frmPaper.InchScale = True
     Else
        frmPaper.InchScale = False
    End If
    frmPaper.Toggle_Grid

End Sub

Private Sub GenerateCode()

  Dim Control As VB.Control

    On Error Resume Next
    Load frmCode
    frmCode.txtCode = "With Printer" & vbNewLine
    If Not Portrait Then
        frmCode.txtCode = frmCode.txtCode & vbTab & ".Orientation = vbPRORLandscape" & vbNewLine & vbNewLine
    End If
    For Each Control In frmPaper.Controls
        With Control
            If .Name <> "Gripper" Then
                If .Index <> 0 Then
                    Select Case .Name
                     Case "TextBox"
                        frmCode.txtCode = frmCode.txtCode & vbTab & "'TextBox '" & .DataField & "'" & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbTab & ".CurrentY = " & .Top & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbTab & ".CurrentX = " & .Left & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbTab & ".Font.Name = """ & .Font.Name & """" & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbTab & ".Font.Size = " & .Font.Size & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbTab & ".Font.Bold = " & .Font.Bold & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbTab & ".Font.Italic = " & .Font.Italic & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbTab & ".Font.Underline = " & .Font.Underline & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbTab & ".Font.Strikethrough = " & .Font.Strikethrough & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbTab & "Printer.Print """ & .Text & """" & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbNewLine
                     Case "sLine"
                        frmCode.txtCode = frmCode.txtCode & vbTab & "'Line '" & .DataField & "'" & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbTab & "Printer.Line(" & .X1 - 2 & "," & .Y1 & ")-(" & .X2 + 2 & "," & .Y2 & "),vbBlack" & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbNewLine
                     Case "PictureBox"
                        frmCode.txtCode = frmCode.txtCode & vbTab & "'Picture '" & .DataField & "'" & vbNewLine
                        frmCode.txtCode = frmCode.txtCode & vbTab & "Printer.PaintPicture LoadPicture(""" & .DataMember & """)," & .Left & "," & .Top & vbNewLine
                        Printer.EndDoc
                    End Select
                End If
            End If
        End With 'Control
    Next Control
    With frmCode
        .txtCode = .txtCode & vbTab & ".EndDoc" & vbNewLine
        .txtCode = .txtCode & "End With" & vbNewLine
        .Show 1
    End With 'frmCode
    On Error GoTo 0

End Sub

Private Sub Get_PaperSize()

  Dim PaperWidth  As Long
  Dim PaperHeight As Long
  Dim TempPaper   As Long
  Dim Control     As VB.Control

    '
    frmPaper.GridOn = True
    frmPaper.Toggle_Grid
    Select Case cboPaperSize.ListIndex
     Case 0  'Letter
        PaperWidth = 8.5 * 1440
        PaperHeight = 11 * 1440
        'frmPaper.InchScale = True
     Case 1  'Leagal
        PaperWidth = 8.5 * 1440
        PaperHeight = 14 * 1440
        'frmPaper.InchScale = True
     Case 2 'A4
        PaperWidth = 210 * 57
        PaperHeight = 297 * 57
        'frmPaper.InchScale = False
     Case 3 'Envelope #10
        PaperWidth = 9.5 * 1440
        PaperHeight = 4.125 * 1440
        'frmPaper.InchScale = True
    End Select
    
    If Not Portrait Then
        TempPaper = PaperWidth
        PaperWidth = PaperHeight
        PaperHeight = TempPaper
    End If
    
    With frmPaper
        .Width = PaperWidth 'Printer.Width
        .Height = PaperHeight
        .Left = ((Me.ScaleWidth - frmPaper.Width) / 2)
        .GridOn = False
        .BoxesOn = False
        .Toggle_Grid
        .Toggle_Boxes
        For Each Control In .Controls
            If Control.Name = "sLine" Then
                Control.Top = 0
                Control.Left = 0
                Control.Width = .Width
                Control.Height = .Height
            End If
        Next Control
    End With 'frmPaper

End Sub

Public Sub Load_Form(OpenFileName As String)

  Dim Control As VB.Control
  Dim FileIn  As Variant

    On Error Resume Next
    'On Error Resume Next
    frmPaper.BoxesOn = False
    frmPaper.GridOn = False
    ControlNumber = 1
    For Each Control In frmPaper.Controls
        With Control
            If .Name <> "Gripper" Then
                If .Index <> 0 Then
                    Unload Control
                End If
            End If
        End With
    Next Control
    Open OpenFileName For Input As #1
    Do While Not EOF(1)
        Input #1, FileIn
        If FileIn = "TextBox" Then
            With frmPaper
                Load .TextBox(ControlNumber)
                Input #1, FileIn
                .TextBox(ControlNumber).Top = FileIn
                Input #1, FileIn
                .TextBox(ControlNumber).Left = FileIn
                Input #1, FileIn
                .TextBox(ControlNumber).Width = FileIn
                Input #1, FileIn
                .TextBox(ControlNumber).Height = FileIn
                Input #1, FileIn
                .TextBox(ControlNumber).Font.Name = FileIn
                Input #1, FileIn
                .TextBox(ControlNumber).Font.Size = FileIn
                Input #1, FileIn
                .TextBox(ControlNumber).Font.Bold = FileIn
                Input #1, FileIn
                .TextBox(ControlNumber).Font.Italic = FileIn
                Input #1, FileIn
                .TextBox(ControlNumber).Font.Underline = FileIn
                Input #1, FileIn
                .TextBox(ControlNumber).Font.Strikethrough = FileIn
                Input #1, FileIn
                .TextBox(ControlNumber).Text = FileIn
                Input #1, FileIn
                .TextBox(ControlNumber).DataField = FileIn
                .TextBox(ControlNumber).Visible = True
                If .BoxesOn Then
                    .TextBox(ControlNumber).BorderStyle = 1
                 Else
                    .TextBox(ControlNumber).BorderStyle = 0
                End If
            End With
         ElseIf FileIn = "sLine" Then
            With frmPaper
                Load .sLine(ControlNumber)
                .sLine(ControlNumber).Left = 0
                .sLine(ControlNumber).Top = 0
                .sLine(ControlNumber).Width = Me.ScaleWidth
                .sLine(ControlNumber).Height = Me.ScaleHeight
                Input #1, FileIn
                .sLine(ControlNumber).X1 = FileIn
                Input #1, FileIn
                .sLine(ControlNumber).Y1 = FileIn
                Input #1, FileIn
                .sLine(ControlNumber).X2 = FileIn
                Input #1, FileIn
                .sLine(ControlNumber).Y2 = FileIn
                Input #1, FileIn
                .sLine(ControlNumber).DataField = FileIn
                .sLine(ControlNumber).Visible = True
                If .BoxesOn Then
                    .sLine(ControlNumber).IsVisible = True
                 Else
                    .sLine(ControlNumber).IsVisible = False
                End If
            End With 'frmPaper
         ElseIf FileIn = "PictureBox" Then
            With frmPaper
                Load .PictureBox(ControlNumber)
                Input #1, FileIn
                .PictureBox(ControlNumber).Top = FileIn
                Input #1, FileIn
                .PictureBox(ControlNumber).Left = FileIn
                Input #1, FileIn
                .PictureBox(ControlNumber).Width = FileIn
                Input #1, FileIn
                .PictureBox(ControlNumber).Height = FileIn
                Input #1, FileIn
                .PictureBox(ControlNumber).DataMember = FileIn
                Input #1, FileIn
                .PictureBox(ControlNumber).DataField = FileIn
                .PictureBox(ControlNumber).Visible = True
                .PictureBox(ControlNumber).Picture = LoadPicture(.PictureBox(ControlNumber).DataMember)
                If .BoxesOn Then
                    .PictureBox(ControlNumber).BorderStyle = 1
                 Else
                    .PictureBox(ControlNumber).BorderStyle = 0
                End If
            End With
        End If
        ControlNumber = ControlNumber + 1
    Loop
    Close #1
    frmPaper.Toggle_Boxes
    frmPaper.Toggle_Grid
    On Error GoTo 0

End Sub

Private Sub MDIForm_Load()

    On Error Resume Next
    Load frmPaper
    
    With cboPaperSize
        .AddItem "Letter"
        .AddItem "Leagal"
        .AddItem "A4"
        .AddItem "#10 Envelope"
        .ListIndex = 0
    End With 'cboPaperSize
    
    optScale(0) = True
    
    frmPaper.InchScale = True
    
    Get_PaperSize
    CurrentFileName = "*.pfm"
    On Error GoTo 0

End Sub

Public Sub Save_Form(SaveFileName As String)

  Dim Control As VB.Control

    frmPaper.BoxesOn = False
    frmPaper.GridOn = False
    Open SaveFileName For Output As #1
    For Each Control In frmPaper.Controls
        With Control
            If .Name <> "Gripper" Then
                If .Index <> 0 Then
                    Select Case .Name
                     Case "TextBox"
                        Print #1, "TextBox"
                        Print #1, .Top
                        Print #1, .Left
                        Print #1, .Width
                        Print #1, .Height
                        Print #1, .Font.Name
                        Print #1, .Font.Size
                        Print #1, .Font.Bold
                        Print #1, .Font.Italic
                        Print #1, .Font.Underline
                        Print #1, .Font.Strikethrough
                        Print #1, .Text
                        Print #1, .DataField
                     Case "sLine"
                        Print #1, "sLine"
                        Print #1, .X1
                        Print #1, .Y1
                        Print #1, .X2
                        Print #1, .Y2
                        Print #1, .DataField
                     Case "PictureBox"
                        Print #1, "PictureBox"
                        Print #1, .Top
                        Print #1, .Left
                        Print #1, .Width
                        Print #1, .Height
                        Print #1, .DataMember
                        Print #1, .DataField
                    End Select
                End If
            End If
        End With 'Control
    Next Control
    Close #1
    frmPaper.Toggle_Grid
    frmPaper.Toggle_Boxes

End Sub

Private Sub timControl_Timer()

    If frmPaper.ActiveControl Is Nothing Then
        txtLeftX1.Enabled = False
        txtTopY1.Enabled = False
        txtWidthX2.Enabled = False
        txtHeightY2.Enabled = False
        txtName.Enabled = False
        txtPath.Enabled = False
        Exit Sub
    End If
    With frmPaper.ActiveControl
        If .Name <> "Gripper" Then
            If .Index <> 0 Then
                Set m_ActiveControl = frmPaper.ActiveControl
                lblType.Caption = m_ActiveControl.Name
                txtName.Text = m_ActiveControl.DataField
                If m_ActiveControl.Name = "sLine" Then
                    txtLeftX1.Text = m_ActiveControl.X1
                    txtTopY1.Text = m_ActiveControl.Y1
                    txtWidthX2.Text = m_ActiveControl.X2
                    txtHeightY2.Text = m_ActiveControl.Y2
                    txtPath.Enabled = False
                 ElseIf m_ActiveControl.Name = "TextBox" Then
                    txtLeftX1.Text = m_ActiveControl.Left
                    txtTopY1.Text = m_ActiveControl.Top
                    txtWidthX2.Text = m_ActiveControl.Width
                    txtHeightY2.Text = m_ActiveControl.Height
                    txtPath.Enabled = False
                 ElseIf m_ActiveControl.Name = "PictureBox" Then
                    txtLeftX1.Text = m_ActiveControl.Left
                    txtTopY1.Text = m_ActiveControl.Top
                    txtWidthX2.Text = m_ActiveControl.Width
                    txtHeightY2.Text = m_ActiveControl.Height
                    txtPath.Text = m_ActiveControl.DataMember
                    txtPath.Enabled = True
                End If
                txtLeftX1.Enabled = True
                txtTopY1.Enabled = True
                txtWidthX2.Enabled = True
                txtHeightY2.Enabled = True
                txtName.Enabled = True
                Exit Sub
            End If
        End If
    End With 'Control
    Set m_ActiveControl = Nothing
    lblType.Caption = vbNullString
    txtName.Text = vbNullString
    txtLeftX1.Text = vbNullString
    txtTopY1.Text = vbNullString
    txtWidthX2.Text = vbNullString
    txtHeightY2.Text = vbNullString

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

  Dim Control As VB.Control

    Select Case Button.Key
     Case "TextBox"
        ControlNumber = ControlNumber + 1
        With frmPaper
            Load .TextBox(ControlNumber)
            .TextBox(ControlNumber).Left = 100
            .TextBox(ControlNumber).Top = 100
            .TextBox(ControlNumber).Visible = True
            .TextBox(ControlNumber).ZOrder
            If .BoxesOn Then
                .TextBox(ControlNumber).BorderStyle = 1
             Else
                .TextBox(ControlNumber).BorderStyle = 0
            End If
        End With 'frmPaper
     Case "StraitLine"
        ControlNumber = ControlNumber + 1
        With frmPaper
            Load .sLine(ControlNumber)
            .sLine(ControlNumber).Left = 0
            .sLine(ControlNumber).Top = 0
            frmPaper.sLine(ControlNumber).Width = frmPaper.ScaleWidth
            frmPaper.sLine(ControlNumber).Height = frmPaper.ScaleHeight
            .sLine(ControlNumber).Visible = True
            .sLine(ControlNumber).ZOrder
            If .BoxesOn Then
                .sLine(ControlNumber).IsVisible = True
             Else 'FRMPAPER.BOXESON = FALSE/0
                .sLine(ControlNumber).IsVisible = False
            End If
        End With 'frmPaper
     Case "PictureBox"
        ControlNumber = ControlNumber + 1
        With frmPaper
            Load .PictureBox(ControlNumber)
            .PictureBox(ControlNumber).Visible = True
            .PictureBox(ControlNumber).ZOrder
            If .BoxesOn Then
                .PictureBox(ControlNumber).BorderStyle = 1
             Else
                .PictureBox(ControlNumber).BorderStyle = 0
            End If
        End With 'frmPaper
     Case "Create Code"
        GenerateCode
     Case "Grid"
        frmPaper.Toggle_Grid
     Case "Boxes"
        frmPaper.Toggle_Boxes
     Case "Load"
        On Error Resume Next
        With CD1
            .CancelError = True
            .Filter = "Forms (*.pfm)|*.pfm"
            .FileName = CurrentFileName
            .ShowOpen
            If Err.Number = 0 Then
                CurrentFileName = .FileName
                Load_Form (.FileName)
            End If
        End With
        On Error GoTo 0
     Case "Save"
        On Error Resume Next
        With CD1
            .CancelError = True
            .Filter = "Forms (*.pfm)|*.pfm"
            .FileName = CurrentFileName
            .DefaultExt = "pfm"
            .ShowSave
            If Err.Number = 0 Then
                CurrentFileName = .FileName
                Save_Form (.FileName)
            End If
        End With
        On Error GoTo 0
     Case "Close"
        For Each Control In frmPaper.Controls
            With Control
                If .Name <> "Gripper" Then
                    If .Index <> 0 Then
                        Unload Control
                    End If
                End If
            End With 'Control
        Next Control
        frmPaper.Gripper.Visible = False
        CurrentFileName = "*.pfm"
    End Select

End Sub

Private Sub txtHeightY2_GotFocus()

    timControl.Enabled = False

End Sub

Private Sub txtHeightY2_LostFocus()

    If m_ActiveControl.Name = "sLine" Then
        m_ActiveControl.Y2 = Val(txtHeightY2)
     Else
        m_ActiveControl.Height = Val(txtHeightY2)
    End If
    timControl.Enabled = True

End Sub

Private Sub txtLeftX1_GotFocus()

    timControl.Enabled = False

End Sub

Private Sub txtLeftX1_LostFocus()

    If m_ActiveControl.Name = "sLine" Then
        m_ActiveControl.X1 = Val(txtLeftX1)
     Else
        m_ActiveControl.Left = Val(txtLeftX1)
    End If
    timControl.Enabled = True

End Sub

Private Sub txtName_GotFocus()

    timControl.Enabled = False

End Sub

Private Sub txtName_LostFocus()

    m_ActiveControl.DataField = txtName
    timControl.Enabled = True

End Sub

Private Sub txtPath_GotFocus()

    timControl.Enabled = False

End Sub

Private Sub txtPath_LostFocus()

    If m_ActiveControl.Name = "PictureBox" Then
        m_ActiveControl.DataMember = txtPath.Text
    End If
    timControl.Enabled = True

End Sub

Private Sub txtTopY1_GotFocus()

    timControl.Enabled = False

End Sub

Private Sub txtTopY1_LostFocus()

    If m_ActiveControl.Name = "sLine" Then
        m_ActiveControl.Y1 = Val(txtTopY1)
     Else
        m_ActiveControl.Top = Val(txtTopY1)
    End If
    timControl.Enabled = True

End Sub

Private Sub txtWidthX2_GotFocus()

    timControl.Enabled = False

End Sub

Private Sub txtWidthX2_LostFocus()

    If m_ActiveControl.Name = "sLine" Then
        m_ActiveControl.X2 = Val(txtWidthX2)
     Else
        m_ActiveControl.Width = Val(txtWidthX2)
    End If
    timControl.Enabled = True

End Sub

''Private Sub Text4_Change()
''
''
''
''
''End Sub

':) Roja's VB Code Fixer V1.0.99 (6/18/2003 8:12:36 PM) 10 + 558 = 568 Lines Thanks Ulli for inspiration and lots of code.
