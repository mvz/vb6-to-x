VERSION 5.00
Begin VB.Form RandTextForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chameller Random Text"
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1500
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "RANDTEXT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4245
   ScaleWidth      =   8220
   Begin VB.PictureBox CommonDialog 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7980
      ScaleHeight     =   450
      ScaleWidth      =   1170
      TabIndex        =   20
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Frame f_Analysis 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Anal&ysis"
      ForeColor       =   &H00000000&
      Height          =   4215
      Left            =   4320
      TabIndex        =   14
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton b_Clear 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clea&r"
         Height          =   315
         Left            =   2400
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton b_Save_S 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "S&ave..."
         Height          =   315
         Left            =   1260
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton b_Load_S 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "L&oad..."
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.PictureBox Grid1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   120
         ScaleHeight     =   3465
         ScaleWidth      =   3585
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.Frame f_Proc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Text &Produced"
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   60
      TabIndex        =   9
      Top             =   2520
      Width           =   4215
      Begin VB.HScrollBar scr_Size 
         Height          =   255
         LargeChange     =   500
         Left            =   3000
         Max             =   4000
         SmallChange     =   10
         TabIndex        =   13
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CommandButton b_Create 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Create"
         Height          =   315
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton b_Save_P 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sa&ve..."
         Height          =   315
         Left            =   3000
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox t_Produced 
         Appearance      =   0  'Flat
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lbl_Size 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Size: 0"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3000
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame f_Text 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Text to Analyse"
      ForeColor       =   &H00000000&
      Height          =   2475
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.Frame f_Options 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Options"
         ForeColor       =   &H00000000&
         Height          =   1035
         Left            =   3000
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
         Begin VB.PictureBox c_Add 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   60
            ScaleHeight     =   165
            ScaleWidth      =   945
            TabIndex        =   8
            Top             =   780
            Width           =   975
         End
         Begin VB.PictureBox o_2D 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   60
            ScaleHeight     =   225
            ScaleWidth      =   645
            TabIndex        =   7
            Top             =   480
            Width           =   675
         End
         Begin VB.PictureBox o_1D 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   60
            ScaleHeight     =   225
            ScaleWidth      =   645
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.CommandButton b_Analyse 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Analyse"
         Height          =   315
         Left            =   3000
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton b_Save_A 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Save"
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton b_Load 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Load..."
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox t_Analysis 
         Appearance      =   0  'Flat
         Height          =   2115
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "RandTextForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MatrixInt(27, 27, 27) As Integer
Dim MatrixFrac(27, 27, 27) As Double
Dim RunningTotal2D(27, 27) As Integer
Dim RunningTotal1D(27) As Integer
Const ASC_A_1 = 96
Dim Buffer As String

Private Sub b_Analyse_Click()
    'Voorlopig alleen 3D
    Dim I%, J%, K%

    Dim TempInt As Integer
    Dim Temp1Int As Integer
    Dim Temp2Int As Integer
    Dim Temp As String
    Dim Temp1 As String

    If Len(t_Analysis.Text) = 0 Then Exit Sub
    If c_Add.Value = False Then
        ClearAnalysis
    End If
    Screen.MousePointer = 11

    Temp1Int = 27
    Temp2Int = 27

    For I% = 1 To Len(t_Analysis.Text)
        Temp$ = Mid$(t_Analysis.Text, I%, 1)
        Temp$ = LCase(Temp$)
        TempInt% = Asc(Temp$) - ASC_A_1
        If TempInt% < 1 Or TempInt% > 26 Then TempInt% = 27
        If Not (TempInt = 27 And Temp1Int = 27 And Temp2Int = 27) Then
            MatrixInt(Temp2Int, Temp1Int, TempInt) = MatrixInt(Temp2Int, Temp1Int, TempInt) + 1
            RunningTotal2D(Temp2Int, Temp1Int) = RunningTotal2D(Temp2Int, Temp1Int) + 1
            RunningTotal1D(Temp1Int) = RunningTotal1D(Temp1Int) + 1
            Temp2Int = Temp1Int
            Temp1Int = TempInt
        End If
    Next I%
    For I% = 1 To 27
        If RunningTotal1D(I%) > 0 Then
            Grid1.Row = I%
            For J% = 1 To 27
                If RunningTotal2D(I%, J%) > 0 Then
                    For K% = 1 To 27
                        If MatrixInt(I%, J%, K%) > 0 Then
                            MatrixFrac(I%, J%, K%) = MatrixInt(I%, J%, K%) / RunningTotal2D(I%, J%)
                        End If
                    Next K%
                    Grid1.Col = J%
                    Grid1.Text = Str$(RunningTotal2D(I%, J%))
                End If
            Next J%
        End If
    Next I%
    'For I% = 1 To 27
     '   If RunningTotal(I%) = 0 Then
      '      Grid1.ColWidth(I%) = 1
       '     Grid1.RowHeight(I%) = 1
        'Else
         '   Grid1.ColWidth(I%) = 300
          '  Grid1.RowHeight(I%) = 250
    '    End If
    'Next I%
    Screen.MousePointer = 0
End Sub

Private Sub b_Clear_Click()
    ClearAnalysis
End Sub

Private Sub b_Create_Click()
    Dim I%, J%, K%
    Dim LastChar1%
    Dim LastChar2%
    Dim NextRandom As Double
    Screen.MousePointer = 11
    Do
        LastChar2 = Int(27 * Rnd + 1)
    Loop While RunningTotal1D(LastChar2) = 0
    Do
        LastChar1 = Int(27 * Rnd + 1)
    Loop While RunningTotal2D(LastChar2, LastChar1) = 0

    If LastChar2 = 27 Then
        Buffer = " "
    Else
        Buffer = Chr$(LastChar2 + ASC_A_1)
    End If
    If LastChar1 = 27 Then
        Buffer = Buffer + " "
    Else
        Buffer = Buffer + Chr$(LastChar1 + ASC_A_1)
    End If
    For J = 0 To scr_Size.Value
        NextRandom = Rnd
        For I% = 1 To 27
            NextRandom = NextRandom - MatrixFrac(LastChar2, LastChar1, I)
            If NextRandom < 0 Then
                LastChar2 = LastChar1
                LastChar1 = I%
                Exit For
            End If
        Next I%
        If LastChar1 = 27 Then
            If LastChar2 <> 27 Then Buffer = Buffer + " "
        Else
            Buffer = Buffer + Chr$(LastChar1 + ASC_A_1)
        End If
    Next J
    t_Produced.Text = Buffer
    Screen.MousePointer = 0
End Sub

Private Sub b_Load_Click()
    Dim TempInt%
    CommonDialog.Filter = "Text files|*.TXT|All files|*.*"
    CommonDialog.FilterIndex = 1
    CommonDialog.DefaultExt = "TXT"
    CommonDialog.FileName = ""
    CommonDialog.Action = 1
    If Not (CommonDialog.FileName = "") Then
        If FileLen(CommonDialog.FileName) > 4000 Then
            TempInt = 4000
        Else
            TempInt = FileLen(CommonDialog.FileName)
        End If
        Open CommonDialog.FileName For Input As #1
        t_Analysis.Text = Input$(TempInt, 1)
        Close #1
    End If
End Sub

Private Sub b_Load_S_Click()
    Dim I%, J%, K%, I1%, J1%, K1%, I2%, J2%, K2%
    CommonDialog.Filter = "Text Analysis files|*.TAF|All files|*.*"
    CommonDialog.FilterIndex = 1
    CommonDialog.DefaultExt = "TAF"
    CommonDialog.FileName = ""
    CommonDialog.Action = 1
    If Not (CommonDialog.FileName = "") Then
        Screen.MousePointer = 11
        ClearAnalysis
        Open CommonDialog.FileName For Random As #1 Len = 2
        Get #1, , I
        Get #1, , I2
        If (I <> Asc("A") * 256 + Asc("T")) Or (I2 <> Asc("2") * 256 + Asc("F")) Then
            MsgBox "No Text Analysis File", 48
            Exit Sub
        End If
'        For I% = 1 To 27
 '           Get #1, , RunningTotal1D(I%)
  '          If RunningTotal1D(I%) > 0 Then
   '             Grid1.Row = I%
    '            For J% = 1 To 27
     '               Get #1, , RunningTotal2D(I%, J%)
      '              If RunningTotal2D(I%, J%) > 0 Then
       '                 For K% = 1 To 27
        '                    Get #1, , MatrixInt(I%, J%, K%)
         '                   If MatrixInt(I%, J%, K%) > 0 Then
        '                        MatrixFrac(I%, J%, K%) = MatrixInt(I%, J%, K%) / RunningTotal2D(I%, J%)
       '                     End If
      '                  Next K%
     '                   Grid1.Col = J%
    '                    Grid1.Text = Str$(RunningTotal2D(I%, J%))
   '                 End If
  '              Next J%
 '           End If
'        Next I%
        Get #1, , I1
        For I% = 1 To I1
            Get #1, , I2
            Get #1, , RunningTotal1D(I2)
            Grid1.Row = I2
            Get #1, , J1
            For J% = 1 To J1
                Get #1, , J2
                Get #1, , RunningTotal2D(I2, J2)
                Get #1, , K1
                For K% = 1 To K1
                    Get #1, , K2
                    Get #1, , MatrixInt(I2, J2, K2)
                    MatrixFrac(I2, J2, K2) = MatrixInt(I2, J2, K2) / RunningTotal2D(I2, J2)
                Next K%
                Grid1.Col = J2
                Grid1.Text = Str$(RunningTotal2D(I2, J2))
            Next J%
        Next I%
        Close #1
        Screen.MousePointer = 0
    End If

End Sub

Private Sub b_Save_A_Click()
    Dim CancelVar%
    CommonDialog.Filter = "Text files|*.TXT|All files|*.*"
    CommonDialog.FilterIndex = 1
    CommonDialog.DefaultExt = "TXT"
    CommonDialog.FileName = ""
    CommonDialog.Action = 2
    If Not (CommonDialog.FileName = "") Then
        If Dir$(CommonDialog.FileName) <> "" Then
            CancelVar = MsgBox("File Exists, Overwrite?", 36, "Save As")
            If CancelVar <> 6 Then Exit Sub
        End If
        Open CommonDialog.FileName For Output As #1
        Print #1, t_Analysis.Text
        Close #1
    End If
End Sub

Private Sub b_Save_P_Click()
    Dim CancelVar%
    CommonDialog.Filter = "Text files|*.TXT|All files|*.*"
    CommonDialog.FilterIndex = 1
    CommonDialog.DefaultExt = "TXT"
    CommonDialog.FileName = ""
    CommonDialog.Action = 2
    If Not (CommonDialog.FileName = "") Then
        If Dir$(CommonDialog.FileName) <> "" Then
            CancelVar = MsgBox("File Exists, Overwrite?", 36, "Save As")
            If CancelVar <> 6 Then Exit Sub
        End If
        Open CommonDialog.FileName For Output As #1
        Print #1, t_Produced.Text
        Close #1
    End If
End Sub

Private Sub b_Save_S_Click()
    Dim CancelVar%
    Dim I%, J%, K%, I1%, J1%, K1%
    CommonDialog.Filter = "Text Analysis files|*.TAF|All files|*.*"
    CommonDialog.FilterIndex = 1
    CommonDialog.DefaultExt = "TAF"
    CommonDialog.FileName = ""
    CommonDialog.Action = 2
    If Not (CommonDialog.FileName = "") Then
        If Dir$(CommonDialog.FileName) <> "" Then
            CancelVar = MsgBox("File Exists, Overwrite?", 36, "Save As")
            If CancelVar <> 6 Then Exit Sub
        End If
        Screen.MousePointer = 11
        Open CommonDialog.FileName For Random As #1 Len = 2
        I% = Asc("A") * 256 + Asc("T")
        Put #1, , I
        I% = Asc("2") * 256 + Asc("F")
        Put #1, , I
'        For I% = 1 To 27
'            Put #1, , RunningTotal1D(I%)
'            If RunningTotal1D(I%) > 0 Then
'                For J% = 1 To 27
 '                   Put #1, , RunningTotal2D(I%, J%)
  '                  If RunningTotal2D(I%, J%) > 0 Then
   '                     For K% = 1 To 27
    '                        Put #1, , MatrixInt(I%, J%, K%)
     '                   Next K%
      '              End If
       '         Next J%
        '    End If
        'Next I%
        I1% = 0
        For I% = 1 To 27
            If RunningTotal1D(I%) > 0 Then I1% = I1% + 1
        Next I
        Put #1, , I1
        For I% = 1 To 27
            If RunningTotal1D(I%) > 0 Then
                Put #1, , I
                Put #1, , RunningTotal1D(I%)
                J1% = 0
                For J% = 1 To 27
                    If RunningTotal2D(I%, J%) > 0 Then J1% = J1% + 1
                Next J
                Put #1, , J1
                For J% = 1 To 27
                    If RunningTotal2D(I%, J%) > 0 Then
                        Put #1, , J
                        Put #1, , RunningTotal2D(I%, J%)
                        K1% = 0
                        For K% = 1 To 27
                            If MatrixInt(I%, J%, K%) > 0 Then K1% = K1% + 1
                        Next K
                        Put #1, , K1
                        For K% = 1 To 27
                            If MatrixInt(I%, J%, K%) > 0 Then
                                Put #1, , K
                                Put #1, , MatrixInt(I%, J%, K%)
                            End If
                        Next K%
                    End If
                Next J%
            End If
        Next I%
        Close #1
        Screen.MousePointer = 0
    End If
End Sub

Private Sub ClearAnalysis()
    Dim I%, J%, K%
    For I% = 1 To 27
        Grid1.Row = I%
        For J% = 1 To 27
            Grid1.Col = J%
            Grid1.Text = ""
            For K% = 1 To 27
                MatrixInt(I%, J%, K%) = 0
                MatrixFrac(I%, J%, K%) = 0
            Next K%
            RunningTotal2D(I%, J%) = 0
        Next J%
        RunningTotal1D(I%) = 0
    Next I%
End Sub

Private Sub Form_Load()
    Dim I%
    Grid1.Col = 0
    ClearAnalysis
    Grid1.Col = 0
    For I% = 1 To 26
        Grid1.Row = I%
        Grid1.Text = Chr$(I% + ASC_A_1)
    Next I%
    Grid1.Row = 0
    For I% = 1 To 26
        Grid1.Col = I%
        Grid1.Text = Chr$(I% + ASC_A_1)
        Grid1.ColWidth(I%) = 400
    Next I%
End Sub

Private Sub scr_Size_Change()
    lbl_Size.Caption = "Size: " + Format$(scr_Size.Value)
End Sub

