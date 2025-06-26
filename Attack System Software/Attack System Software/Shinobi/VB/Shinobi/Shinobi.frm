VERSION 5.00
Begin VB.Form From1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shinobi - J [Press J For Help]"
   ClientHeight    =   1560
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7230
   FillColor       =   &H00404040&
   Icon            =   "Shinobi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   840
      Picture         =   "Shinobi.frx":030A
      ScaleHeight     =   315
      ScaleWidth      =   6150
      TabIndex        =   24
      Top             =   2160
      Width           =   6150
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   120
      Picture         =   "Shinobi.frx":685C
      ScaleHeight     =   570
      ScaleWidth      =   735
      TabIndex        =   23
      Top             =   2160
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   120
      Picture         =   "Shinobi.frx":7E96
      ScaleHeight     =   570
      ScaleWidth      =   735
      TabIndex        =   22
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   840
      Picture         =   "Shinobi.frx":94D0
      ScaleHeight     =   315
      ScaleWidth      =   6150
      TabIndex        =   21
      Top             =   120
      Width           =   6150
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Shinobi.frx":FA22
      Left            =   5280
      List            =   "Shinobi.frx":FA24
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   780
      Width           =   1815
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   180
      Left            =   2520
      TabIndex        =   31
      Top             =   3360
      Width           =   90
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   180
      Left            =   1440
      TabIndex        =   30
      Top             =   3360
      Width           =   90
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      Height          =   180
      Left            =   600
      TabIndex        =   29
      Top             =   3360
      Width           =   90
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   180
      Left            =   5760
      TabIndex        =   28
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   180
      Left            =   3360
      TabIndex        =   27
      Top             =   3000
      Width           =   90
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   180
      Left            =   600
      TabIndex        =   26
      Top             =   3000
      Width           =   90
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Single"
      Enabled         =   0   'False
      ForeColor       =   &H80000003&
      Height          =   180
      Left            =   3240
      TabIndex        =   25
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "«∞"
      Height          =   180
      Left            =   6720
      TabIndex        =   8
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "‘⁄"
      Height          =   180
      Left            =   6480
      TabIndex        =   7
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "¡“"
      Height          =   180
      Left            =   6240
      TabIndex        =   6
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Íá"
      Height          =   180
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ω‘"
      Height          =   180
      Left            =   5760
      TabIndex        =   4
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "’ﬂ"
      Height          =   180
      Left            =   5520
      TabIndex        =   3
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ÍL"
      Height          =   180
      Left            =   5280
      TabIndex        =   2
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "±¯"
      Height          =   180
      Left            =   5040
      TabIndex        =   1
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "≈R"
      Height          =   180
      Left            =   4800
      TabIndex        =   0
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retry"
      Height          =   180
      Left            =   5520
      TabIndex        =   20
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attack"
      Enabled         =   0   'False
      Height          =   180
      Left            =   6240
      TabIndex        =   19
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 1-A “°π‚"
      Height          =   180
      Left            =   3840
      TabIndex        =   17
      Top             =   840
      Width           =   1260
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   4350
      TabIndex        =   16
      Top             =   480
      Width           =   90
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double"
      Enabled         =   0   'False
      ForeColor       =   &H80000003&
      Height          =   180
      Left            =   2280
      TabIndex        =   15
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special"
      Enabled         =   0   'False
      ForeColor       =   &H80000003&
      Height          =   180
      Left            =   1200
      TabIndex        =   14
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Defiance"
      Enabled         =   0   'False
      ForeColor       =   &H80000003&
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Front"
      Height          =   180
      Left            =   3240
      TabIndex        =   12
      Top             =   840
      Width           =   450
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Slash"
      Height          =   180
      Left            =   2280
      TabIndex        =   11
      Top             =   840
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hotsuma"
      Height          =   180
      Left            =   1200
      TabIndex        =   10
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Normal"
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   540
   End
End
Attribute VB_Name = "From1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ShinobiData
    Name As String
    Hp As Double
    Shuriken As Double
    Flame As Double
    Side As Double
    Back As Double
    Tate As Double
End Type
Dim Data(16, 7) As ShinobiData
Dim Stage(16) As String
Dim xx As Integer
Dim yy As Integer
Dim A As Double
Dim B As Double
Dim C As Double
Dim D As Double
Dim E As Double
Dim F As Double
Dim G As Double
Dim H As Double
Dim Percent As Double
Sub Hp()
If xx = 15 And yy = 2 Then
    If Label10 = "Easy" Then
    Data(15, 2).Hp = 15
    ElseIf Label10 = "Normal" Then Data(15, 2).Hp = 15
    ElseIf Label10 = "Hard" Then Data(15, 2).Hp = 22.5
    ElseIf Label10 = "Super" Then Data(15, 2).Hp = 30
    End If
End If
End Sub
Sub S0()
If yy = 6 Then
    If xx = 2 Or xx = 6 Or xx = 8 Or xx = 10 Or xx = 12 Or xx = 14 Then Label14.Enabled = True
Else
    Label14.ForeColor = &H80000003
    Label14.Enabled = False
End If

If Label12 = "Fallen" Then
    Label14.ForeColor = &H80000003
    Label14.Enabled = False
End If

End Sub

Sub S1()
If Label12 = "Shuriken" Then
        If Label21 = "Single" Then Label16.ForeColor = &H80000003
        If Label21 = "Burst" Then Label15.ForeColor = &H80000003
    Label16.Enabled = True
    Label21.Enabled = True
    Label21.ForeColor = &H80000012
Else
    Label21.ForeColor = &H80000003
    Label21.Enabled = False
    
    If Label12 = "Slash" Then
        Label16.Enabled = True
    Else
        Label16.ForeColor = &H80000003
        Label16.Enabled = False
    End If
    
End If
Label15.Enabled = True
End Sub
Sub S2()
'//C
Select Case Label10
    Case "Super"
        Select Case Label11
            Case "Hotsuma"
                If Label12 = "Shuriken" Then
                    C = Data(xx, yy).Shuriken
                ElseIf Label12 = "Slash" Then
                    C = 2.5
                End If
            
            Case "Moritsune"
                If Label12 = "Shuriken" Then
                    C = Data(xx, yy).Shuriken
                ElseIf Label12 = "Slash" Then
                    C = 5
                End If
            
            Case "Joe Musashi"
                If Label12 = "Shuriken" Then
                    C = 1
                ElseIf Label12 = "Slash" Then
                    C = 2.5
                End If
        End Select
                
    Case "Hard"
        Select Case Label11
            Case "Hotsuma"
                If Label12 = "Shuriken" Then
                    C = Data(xx, yy).Shuriken
                ElseIf Label12 = "Slash" Then
                    C = 3.74999
                End If
            
            Case "Moritsune"
                If Label12 = "Shuriken" Then
                    C = Data(xx, yy).Shuriken
                ElseIf Label12 = "Slash" Then
                    C = 7.4999
                End If
            
            Case "Joe Musashi"
                If Label12 = "Shuriken" Then
                    C = 1.5
                ElseIf Label12 = "Slash" Then
                    C = 3.74999
                End If
        End Select
    Case "Normal"
        Select Case Label11
            Case "Hotsuma"
                If Label12 = "Shuriken" Then
                    C = Data(xx, yy).Shuriken
                ElseIf Label12 = "Slash" Then
                    C = 5
                End If
            
            Case "Moritsune"
                If Label12 = "Shuriken" Then
                    C = Data(xx, yy).Shuriken
                ElseIf Label12 = "Slash" Then
                    C = 10
                End If
            
            Case "Joe Musashi"
                If Label12 = "Shuriken" Then
                    C = 2
                ElseIf Label12 = "Slash" Then
                    C = 5
                End If
        End Select
    Case "Easy"
        Select Case Label11
            Case "Hotsuma"
                If Label12 = "Shuriken" Then
                    C = Data(xx, yy).Shuriken
                ElseIf Label12 = "Slash" Then
                    C = 10
                End If
            
            Case "Moritsune"
                If Label12 = "Shuriken" Then
                    C = Data(xx, yy).Shuriken
                ElseIf Label12 = "Slash" Then
                    C = 14.999
                End If
            
            Case "Joe Musashi"
                If Label12 = "Shuriken" Then
                    C = 2
                ElseIf Label12 = "Slash" Then
                    C = 10
                End If
        End Select
End Select

If Label12 = "Ka'en" Then C = Data(xx, yy).Flame

If Label12 = "Kamaitachi" Then C = 10

If Label12 = "Missile" Then
    If Label10 = "Super" Then
        C = 7.5
        ElseIf Label10 = "Hard" Then C = 12.5
        ElseIf Label10 = "Normal" Then C = 25
        ElseIf Label10 = "Easy" Then C = 50
    End If
End If

If Label12 = "Fallen" Then
    If Label10 = "Super" Then
        C = 12.5
        ElseIf Label10 = "Hard" Then C = 18.75
        ElseIf Label10 = "Normal" Then C = 37.5
        ElseIf Label10 = "Easy" Then C = 50
    End If
End If

'//D
If Label8.FontBold = True Then
    D = 60
    ElseIf Label7.FontBold = True Then D = 40
    ElseIf Label6.FontBold = True Then D = 30
    ElseIf Label5.FontBold = True Then D = 15
    ElseIf Label4.FontBold = True Then D = 10
    ElseIf Label3.FontBold = True Then D = 4
    ElseIf Label2.FontBold = True Then D = 3
    ElseIf Label1.FontBold = True Then D = 2
    ElseIf Label1.FontBold = False Then D = 1
End If

'//E
Select Case Label13
    Case "Front"
        E = 1
    Case "Side"
        E = Data(xx, yy).Side
    Case "Back"
        E = Data(xx, yy).Back
End Select

'//F
If Label14.ForeColor = &H80000003 Then
    F = 1
    ElseIf Label14.ForeColor = &H80000012 Then F = 3
End If

'//G
If Label15.ForeColor = &H80000003 Then
    G = 1
    ElseIf Label15.ForeColor = &H80000012 Then G = 2
End If

'//H
If Label16.ForeColor = &H80000003 Then
    H = 1
    ElseIf Label16.ForeColor = &H80000012 Then H = 2
End If


S0

End Sub
Sub S3()
Label1.ForeColor = &H80000012
Label2.ForeColor = &H80000012
Label3.ForeColor = &H80000012
Label4.ForeColor = &H80000012
Label5.ForeColor = &H80000012
Label6.ForeColor = &H80000012
Label7.ForeColor = &H80000012
Label8.ForeColor = &H80000012
Label9.ForeColor = &H80000012

Label1.FontBold = False
Label2.FontBold = False
Label3.FontBold = False
Label4.FontBold = False
Label5.FontBold = False
Label6.FontBold = False
Label7.FontBold = False
Label8.FontBold = False
Label9.FontBold = False

End Sub
Sub S4(L As Control, Str1$, Str2$)
Str3$ = Str1 + "  " + Str2
len1% = Len(Str1)
len2% = Len(Str2)
m% = len1 + 2
K% = 1
Y# = 0.1 / (m + 1)
Z = max(len1, len2)
For K = 1 To m + 1
    L = Mid(Str3, K, K + Z - 1)
    DelayLoop Y
Next K
L = Right(L, len2)
End Sub
Sub S5(L As Control, Str1$, Str2$)
Str3$ = Str2 + "  " + Str1
len1% = Len(Str1)
len2% = Len(Str2)
m% = len1 + 2
K% = 1
Y# = 0.1 / (m + 1)
Z = max(len1, len2)
For K = 1 To len1 + 1
    L = Mid(Str3, len1 + 2 - K, len1 + 2 - K + Z - 1)
    DelayLoop Y
Next K
L = Left(L, len2)
End Sub
Sub S6()
If Combo1 <> "" Then
    Label17 = Format(A, "0.00###") + "/" + Format(Data(xx, yy).Hp) + "  " + "0" + "/" + "0.00%"
Else
    Label17 = ""
End If
    
S3

End Sub
Sub DelayLoop(DelayTime#)
Const SecondsInDay = 24& * 60& * 60&
loopfinish = Timer + DelayTime
If loopfinish > SecondsInDay Then
    loopfinish = loopfinish - SecondsInDay
    Do While Timer > loopfinish
    Loop
End If
Do While Timer < loopfinish
Loop
End Sub
Function max(X, Y)
If X >= Y Then max = X Else max = Y
End Function
Sub Stageselect()
If Label18 = "STAGE 1-A “°π‚" Then
    xx = 1
    ElseIf Label18 = "STAGE 1-B ∆∆‹ä" Then xx = 2
    ElseIf Label18 = "STAGE 2-A È_Íñ" Then xx = 3
    ElseIf Label18 = "STAGE 2-B Œ‰«˙" Then xx = 4
    ElseIf Label18 = "STAGE 3-A ”Ò∫‚" Then xx = 5
    ElseIf Label18 = "STAGE 3-B ¡Æÿë" Then xx = 6
    ElseIf Label18 = "STAGE 4-A ÃÏòÿ" Then xx = 7
    ElseIf Label18 = "STAGE 4-B Œƒ«˙" Then xx = 8
    ElseIf Label18 = "STAGE 5-A ÃÏôC" Then xx = 9
    ElseIf Label18 = "STAGE 5-B ¬ª¥Ê" Then xx = 10
    ElseIf Label18 = "STAGE 6-A ÃÏËØ" Then xx = 11
    ElseIf Label18 = "STAGE 6-B æﬁ√≈" Then xx = 12
    ElseIf Label18 = "STAGE 7-A ÿù¿«" Then xx = 13
    ElseIf Label18 = "STAGE 7-B ÃÏò–" Then xx = 14
    ElseIf Label18 = "STAGE 8-A ±±≥Ω" Then xx = 15
    ElseIf Label18 = "STAGE 8-B Ã´“ª" Then xx = 16
    ElseIf Label18 = "STAGE EX" Then xx = 0
End If
For i = 1 To 6
    If Data(xx, i).Name <> "" Then Combo1.AddItem Data(xx, i).Name
Next i
If Data(xx, 0).Name <> "" Then Combo1.AddItem Data(xx, 0).Name
If Label18 = "STAGE EX" Then
    If Label10 = "Easy" Then S4 Label10, "Easy", "Normal"
    If Label10 = "Super" Then S5 Label10, "Super", "Hard"
    If Label10 = "Hard" Then S5 Label10, "Hard", "Normal"
End If
End Sub
Sub Retry()
If Combo1 <> "" Then
    Hp
    A = Data(xx, yy).Hp
    Label19.Enabled = True
    Picture1.Visible = True
    Picture2.Picture = Picture4
    Picture2.Visible = True
    S6
    S2
End If
End Sub
Sub Tate()
Label1.Left = 6960 - Data(xx, yy).Tate * 240

If (Label1.Left + 240) > 6720 Then
    Label2.Visible = False
Else
    Label2.Visible = True
End If
Label2.Left = Label1.Left + 240

If (Label2.Left + 240) > 6720 Then
    Label3.Visible = False
Else
    Label3.Visible = True
End If
Label3.Left = Label2.Left + 240

If (Label3.Left + 240) > 6720 Then
    Label4.Visible = False
Else
    Label4.Visible = True
End If
Label4.Left = Label3.Left + 240

If (Label4.Left + 240) > 6720 Then
    Label5.Visible = False
Else
    Label5.Visible = True
End If
Label5.Left = Label4.Left + 240

If (Label5.Left + 240) > 6720 Then
    Label6.Visible = False
Else
    Label6.Visible = True
End If
Label6.Left = Label5.Left + 240

If (Label6.Left + 240) > 6720 Then
    Label7.Visible = False
Else
    Label7.Visible = True
End If
Label7.Left = Label6.Left + 240

If (Label7.Left + 240) > 6720 Then
    Label8.Visible = False
Else
    Label8.Visible = True
End If
Label8.Left = Label7.Left + 240

If (Label8.Left + 240) > 6720 Then
    Label9.Visible = False
Else
    Label9.Visible = True
End If
Label9.Left = Label8.Left + 240

End Sub
Sub T0()
Label1.Left = 4800
Label2.Left = Label1.Left + 240
Label3.Left = Label2.Left + 240
Label4.Left = Label3.Left + 240
Label5.Left = Label4.Left + 240
Label6.Left = Label5.Left + 240
Label7.Left = Label6.Left + 240
Label8.Left = Label7.Left + 240
Label9.Left = Label8.Left + 240

Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
End Sub
Sub T1()
    Label1.ForeColor = &HFF&
    Label2.ForeColor = &H80000012
    Label3.ForeColor = &H80000012
    Label4.ForeColor = &H80000012
    Label5.ForeColor = &H80000012
    Label6.ForeColor = &H80000012
    Label7.ForeColor = &H80000012
    Label8.ForeColor = &H80000012
    Label9.ForeColor = &H80000012
    
    Label1.FontBold = True
    Label2.FontBold = False
    Label3.FontBold = False
    Label4.FontBold = False
    Label5.FontBold = False
    Label6.FontBold = False
    Label7.FontBold = False
    Label8.FontBold = False
    Label9.FontBold = False
End Sub
Sub T2()
    Label1.ForeColor = &HFF&
    Label2.ForeColor = &HFF&
    Label3.ForeColor = &H80000012
    Label4.ForeColor = &H80000012
    Label5.ForeColor = &H80000012
    Label6.ForeColor = &H80000012
    Label7.ForeColor = &H80000012
    Label8.ForeColor = &H80000012
    Label9.ForeColor = &H80000012
    
    Label1.FontBold = True
    Label2.FontBold = True
    Label3.FontBold = False
    Label4.FontBold = False
    Label5.FontBold = False
    Label6.FontBold = False
    Label7.FontBold = False
    Label8.FontBold = False
    Label9.FontBold = False
End Sub
Sub T3()
    Label1.ForeColor = &HFF&
    Label2.ForeColor = &HFF&
    Label3.ForeColor = &HFF&
    Label4.ForeColor = &H80000012
    Label5.ForeColor = &H80000012
    Label6.ForeColor = &H80000012
    Label7.ForeColor = &H80000012
    Label8.ForeColor = &H80000012
    Label9.ForeColor = &H80000012
    
    Label1.FontBold = True
    Label2.FontBold = True
    Label3.FontBold = True
    Label4.FontBold = False
    Label5.FontBold = False
    Label6.FontBold = False
    Label7.FontBold = False
    Label8.FontBold = False
    Label9.FontBold = False
End Sub
Sub T4()
    Label1.ForeColor = &HFF&
    Label2.ForeColor = &HFF&
    Label3.ForeColor = &HFF&
    Label4.ForeColor = &HFF&
    Label5.ForeColor = &H80000012
    Label6.ForeColor = &H80000012
    Label7.ForeColor = &H80000012
    Label8.ForeColor = &H80000012
    Label9.ForeColor = &H80000012
    
    Label1.FontBold = True
    Label2.FontBold = True
    Label3.FontBold = True
    Label4.FontBold = True
    Label5.FontBold = False
    Label6.FontBold = False
    Label7.FontBold = False
    Label8.FontBold = False
    Label9.FontBold = False
End Sub
Sub T5()
    Label1.ForeColor = &HFF&
    Label2.ForeColor = &HFF&
    Label3.ForeColor = &HFF&
    Label4.ForeColor = &HFF&
    Label5.ForeColor = &HFF&
    Label6.ForeColor = &H80000012
    Label7.ForeColor = &H80000012
    Label8.ForeColor = &H80000012
    Label9.ForeColor = &H80000012
    
    Label1.FontBold = True
    Label2.FontBold = True
    Label3.FontBold = True
    Label4.FontBold = True
    Label5.FontBold = True
    Label6.FontBold = False
    Label7.FontBold = False
    Label8.FontBold = False
    Label9.FontBold = False
End Sub
Sub T6()
    Label1.ForeColor = &HFF&
    Label2.ForeColor = &HFF&
    Label3.ForeColor = &HFF&
    Label4.ForeColor = &HFF&
    Label5.ForeColor = &HFF&
    Label6.ForeColor = &HFF&
    Label7.ForeColor = &H80000012
    Label8.ForeColor = &H80000012
    Label9.ForeColor = &H80000012
    
    Label1.FontBold = True
    Label2.FontBold = True
    Label3.FontBold = True
    Label4.FontBold = True
    Label5.FontBold = True
    Label6.FontBold = True
    Label7.FontBold = False
    Label8.FontBold = False
    Label9.FontBold = False
End Sub
Sub T7()
    Label1.ForeColor = &HFF&
    Label2.ForeColor = &HFF&
    Label3.ForeColor = &HFF&
    Label4.ForeColor = &HFF&
    Label5.ForeColor = &HFF&
    Label6.ForeColor = &HFF&
    Label7.ForeColor = &HFF&
    Label8.ForeColor = &H80000012
    Label9.ForeColor = &H80000012
    
    Label1.FontBold = True
    Label2.FontBold = True
    Label3.FontBold = True
    Label4.FontBold = True
    Label5.FontBold = True
    Label6.FontBold = True
    Label7.FontBold = True
    Label8.FontBold = False
    Label9.FontBold = False
End Sub
Sub T8()
    Label1.ForeColor = &HFF&
    Label2.ForeColor = &HFF&
    Label3.ForeColor = &HFF&
    Label4.ForeColor = &HFF&
    Label5.ForeColor = &HFF&
    Label6.ForeColor = &HFF&
    Label7.ForeColor = &HFF&
    Label8.ForeColor = &HFF&
    Label9.ForeColor = &H80000012
    
    Label1.FontBold = True
    Label2.FontBold = True
    Label3.FontBold = True
    Label4.FontBold = True
    Label5.FontBold = True
    Label6.FontBold = True
    Label7.FontBold = True
    Label8.FontBold = True
    Label9.FontBold = False
End Sub
Sub T9()
    Label1.ForeColor = &HFF&
    Label2.ForeColor = &HFF&
    Label3.ForeColor = &HFF&
    Label4.ForeColor = &HFF&
    Label5.ForeColor = &HFF&
    Label6.ForeColor = &HFF&
    Label7.ForeColor = &HFF&
    Label8.ForeColor = &HFF&
    Label9.ForeColor = &HFF&
    
    Label1.FontBold = True
    Label2.FontBold = True
    Label3.FontBold = True
    Label4.FontBold = True
    Label5.FontBold = True
    Label6.FontBold = True
    Label7.FontBold = True
    Label8.FontBold = True
    Label9.FontBold = True
End Sub
Private Sub Combo1_Click()
For i = 0 To 6
    If Combo1 = Data(xx, i).Name Then
        yy = i
        Exit For
    End If
Next i
Hp
A = Data(xx, yy).Hp
Tate
If yy = 6 And (xx = 1 Or xx = 2 Or xx = 4) Then
    
Else
    If Label12 = "Missile" Or Label12 = "Fallen" Then S5 Label12, Label12, "Kamaitachi"
    If yy = 6 And xx = 15 Then
        If Label13 = "Back" Then S5 Label13, "Back", "Side"
        If Label13 = "Side" Then S5 Label13, "Side", "Front"
    End If
End If
S6
S1
S2
Label19.Enabled = True
Picture1.Visible = True
Picture2.Picture = Picture4
Picture2.Visible = True
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 74 Or KeyAscii = 106 Then SendKeys "{F1}"
End Sub

Private Sub Form_DblClick()
SendKeys "{F1}"
End Sub

Private Sub Form_Load()
Rem 1A
Data(1, 1).Name = "¿∂»Ã"
Data(1, 1).Hp = 10
Data(1, 1).Shuriken = 2
Data(1, 1).Flame = 50
Data(1, 1).Side = 1.5
Data(1, 1).Back = 3
Data(1, 1).Tate = 7

Data(1, 2).Name = "¿«"
Data(1, 2).Hp = 5
Data(1, 2).Shuriken = 3
Data(1, 2).Flame = 75
Data(1, 2).Side = 2
Data(1, 2).Back = 3
Data(1, 2).Tate = 7

Data(1, 3).Name = "ÃπøÀ"
Data(1, 3).Hp = 30
Data(1, 3).Shuriken = 0
Data(1, 3).Flame = 50
Data(1, 3).Side = 1
Data(1, 3).Back = 1
Data(1, 3).Tate = 4

Data(1, 4).Name = "∆˚≥µ£®BOSS£©"
Data(1, 4).Hp = 2.5
Data(1, 4).Shuriken = 2
Data(1, 4).Flame = 50
Data(1, 4).Side = 1
Data(1, 4).Back = 1
Data(1, 4).Tate = 5

Data(1, 5).Name = "¿∂»Ã£®BOSS£©"
Data(1, 5).Hp = 10
Data(1, 5).Shuriken = 2
Data(1, 5).Flame = 50
Data(1, 5).Side = 1.5
Data(1, 5).Back = 3
Data(1, 5).Tate = 4

Data(1, 6).Name = "÷±…˝ª˙£®BOSS£©"
Data(1, 6).Hp = 100
Data(1, 6).Shuriken = 0
Data(1, 6).Flame = 50
Data(1, 6).Side = 1
Data(1, 6).Back = 1
Data(1, 6).Tate = 5

Rem 1B
Data(2, 1).Name = "¿∂»Ã"
Data(2, 1).Hp = 10
Data(2, 1).Shuriken = 2
Data(2, 1).Flame = 50
Data(2, 1).Side = 1.5
Data(2, 1).Back = 3
Data(2, 1).Tate = 9

Data(2, 2).Name = "ª∆∑…Õ∑"
Data(2, 2).Hp = 2.5
Data(2, 2).Shuriken = 2
Data(2, 2).Flame = 50
Data(2, 2).Side = 1.5
Data(2, 2).Back = 3
Data(2, 2).Tate = 9

Data(2, 3).Name = "◊œ≈Æ»Ã"
Data(2, 3).Hp = 7.5
Data(2, 3).Shuriken = 3
Data(2, 3).Flame = 75
Data(2, 3).Side = 2
Data(2, 3).Back = 3
Data(2, 3).Tate = 9

Data(2, 5).Name = "¿∂»Ã£®BOSS£©"
Data(2, 5).Hp = 10
Data(2, 5).Shuriken = 2
Data(2, 5).Flame = 50
Data(2, 5).Side = 1.5
Data(2, 5).Back = 3
Data(2, 5).Tate = 4

Data(2, 6).Name = " ÿ∫„£®BOSS£©"
Data(2, 6).Hp = 187.5
Data(2, 6).Shuriken = 0
Data(2, 6).Flame = 50
Data(2, 6).Side = 1.5
Data(2, 6).Back = 3
Data(2, 6).Tate = 5

Rem 2A
Data(3, 1).Name = "¿∂»Ã"
Data(3, 1).Hp = 10
Data(3, 1).Shuriken = 2
Data(3, 1).Flame = 50
Data(3, 1).Side = 1.5
Data(3, 1).Back = 3
Data(3, 1).Tate = 9

Data(3, 2).Name = "∫¸"
Data(3, 2).Hp = 10
Data(3, 2).Shuriken = 3
Data(3, 2).Flame = 75
Data(3, 2).Side = 2
Data(3, 2).Back = 3
Data(3, 2).Tate = 6

Data(3, 3).Name = "—ªÃÏπ∑"
Data(3, 3).Hp = 60
Data(3, 3).Shuriken = 1.5
Data(3, 3).Flame = 37.5
Data(3, 3).Side = 1.5
Data(3, 3).Back = 3
Data(3, 3).Tate = 9

Data(3, 4).Name = "∫¸£®BOSS£©"
Data(3, 4).Hp = 10
Data(3, 4).Shuriken = 3
Data(3, 4).Flame = 75
Data(3, 4).Side = 2
Data(3, 4).Back = 3
Data(3, 4).Tate = 5

Data(3, 5).Name = "“¯ƒ–£®BOSS£©"
Data(3, 5).Hp = 150
Data(3, 5).Shuriken = 2
Data(3, 5).Flame = 50
Data(3, 5).Side = 1.5
Data(3, 5).Back = 3
Data(3, 5).Tate = 6

Data(3, 6).Name = "Õ≠≈Æ£®BOSS£©"
Data(3, 6).Hp = 112.5
Data(3, 6).Shuriken = 2
Data(3, 6).Flame = 50
Data(3, 6).Side = 2
Data(3, 6).Back = 3
Data(3, 6).Tate = 6

Rem 2B
Data(4, 1).Name = "¬Ã»Ã"
Data(4, 1).Hp = 25
Data(4, 1).Shuriken = 2
Data(4, 1).Flame = 50
Data(4, 1).Side = 1.5
Data(4, 1).Back = 3
Data(4, 1).Tate = 9

Data(4, 2).Name = "ª∆∑…Õ∑"
Data(4, 2).Hp = 2.5
Data(4, 2).Shuriken = 2
Data(4, 2).Flame = 50
Data(4, 2).Side = 1.5
Data(4, 2).Back = 3
Data(4, 2).Tate = 9

Data(4, 3).Name = "—ªÃÏπ∑"
Data(4, 3).Hp = 60
Data(4, 3).Shuriken = 1.5
Data(4, 3).Flame = 37.5
Data(4, 3).Side = 1.5
Data(4, 3).Back = 3
Data(4, 3).Tate = 5

Data(4, 5).Name = "¬Ã»Ã£®BOSS£©"
Data(4, 5).Hp = 25
Data(4, 5).Shuriken = 2
Data(4, 5).Flame = 50
Data(4, 5).Side = 1.5
Data(4, 5).Back = 3
Data(4, 5).Tate = 4

Data(4, 6).Name = "«øªØ÷±…˝ª˙£®BOSS£©"
Data(4, 6).Hp = 200
Data(4, 6).Shuriken = 0
Data(4, 6).Flame = 50
Data(4, 6).Side = 1
Data(4, 6).Back = 1
Data(4, 6).Tate = 5

Rem 3A
Data(5, 1).Name = "¬Ã»Ã"
Data(5, 1).Hp = 25
Data(5, 1).Shuriken = 2
Data(5, 1).Flame = 50
Data(5, 1).Side = 1.5
Data(5, 1).Back = 3
Data(5, 1).Tate = 8

Data(5, 2).Name = "π∑"
Data(5, 2).Hp = 12.5
Data(5, 2).Shuriken = 3
Data(5, 2).Flame = 75
Data(5, 2).Side = 2
Data(5, 2).Back = 3
Data(5, 2).Tate = 7

Data(5, 3).Name = "«øªØÃπøÀ"
Data(5, 3).Hp = 75
Data(5, 3).Shuriken = 0
Data(5, 3).Flame = 50
Data(5, 3).Side = 1
Data(5, 3).Back = 1
Data(5, 3).Tate = 4

Data(5, 5).Name = "π∑£®BOSS£©"
Data(5, 5).Hp = 12.5
Data(5, 5).Shuriken = 3
Data(5, 5).Flame = 75
Data(5, 5).Side = 2
Data(5, 5).Back = 3
Data(5, 5).Tate = 6

Data(5, 6).Name = "≤Æ¿÷£®BOSS£©"
Data(5, 6).Hp = 187.5
Data(5, 6).Shuriken = 2
Data(5, 6).Flame = 50
Data(5, 6).Side = 3
Data(5, 6).Back = 0
Data(5, 6).Tate = 7

Rem 3B
Data(6, 1).Name = "¬Ã»Ã"
Data(6, 1).Hp = 25
Data(6, 1).Shuriken = 2
Data(6, 1).Flame = 50
Data(6, 1).Side = 1.5
Data(6, 1).Back = 3
Data(6, 1).Tate = 7

Data(6, 2).Name = "÷©÷Î"
Data(6, 2).Hp = 6.5
Data(6, 2).Shuriken = 2
Data(6, 2).Flame = 50
Data(6, 2).Side = 2
Data(6, 2).Back = 3
Data(6, 2).Tate = 7

Data(6, 3).Name = "∫Ï≈Æ»Ã"
Data(6, 3).Hp = 22.5
Data(6, 3).Shuriken = 3
Data(6, 3).Flame = 75
Data(6, 3).Side = 2
Data(6, 3).Back = 3
Data(6, 3).Tate = 7

Data(6, 4).Name = "÷©÷Î≥≤"
Data(6, 4).Hp = 20
Data(6, 4).Shuriken = 2
Data(6, 4).Flame = 50
Data(6, 4).Side = 1
Data(6, 4).Back = 1
Data(6, 4).Tate = 7

Data(6, 5).Name = "÷©÷Î£®BOSS£©"
Data(6, 5).Hp = 6.5
Data(6, 5).Shuriken = 2
Data(6, 5).Flame = 50
Data(6, 5).Side = 2
Data(6, 5).Back = 3
Data(6, 5).Tate = 5

Data(6, 6).Name = "∞◊‘∆£®BOSS£©"
Data(6, 6).Hp = 225
Data(6, 6).Shuriken = 0
Data(6, 6).Flame = 37.5
Data(6, 6).Side = 1
Data(6, 6).Back = 1
Data(6, 6).Tate = 6

Data(6, 0).Name = "÷©÷Î≥≤£®BOSS£©"
Data(6, 0).Hp = 20
Data(6, 0).Shuriken = 2
Data(6, 0).Flame = 50
Data(6, 0).Side = 1
Data(6, 0).Back = 1
Data(6, 0).Tate = 6

Rem 4A
Data(7, 1).Name = "∫Ï»Ã"
Data(7, 1).Hp = 30
Data(7, 1).Shuriken = 2
Data(7, 1).Flame = 50
Data(7, 1).Side = 1.5
Data(7, 1).Back = 3
Data(7, 1).Tate = 7

Data(7, 2).Name = "∂Í◊”"
Data(7, 2).Hp = 7.5
Data(7, 2).Shuriken = 3
Data(7, 2).Flame = 75
Data(7, 2).Side = 2
Data(7, 2).Back = 3
Data(7, 2).Tate = 7

Data(7, 3).Name = "«øªØÃπøÀ"
Data(7, 3).Hp = 75
Data(7, 3).Shuriken = 0
Data(7, 3).Flame = 50
Data(7, 3).Side = 1
Data(7, 3).Back = 1
Data(7, 3).Tate = 7

Data(7, 5).Name = "∑…Õ∑£®BOSS£©"
Data(7, 5).Hp = 7.5
Data(7, 5).Shuriken = 3
Data(7, 5).Flame = 75
Data(7, 5).Side = 2
Data(7, 5).Back = 3
Data(7, 5).Tate = 6

Data(7, 6).Name = "—Ê£®BOSS£©"
Data(7, 6).Hp = 225
Data(7, 6).Shuriken = 2
Data(7, 6).Flame = 0
Data(7, 6).Side = 1.5
Data(7, 6).Back = 3
Data(7, 6).Tate = 7

Rem 4B
Data(8, 1).Name = "∫Ï»Ã"
Data(8, 1).Hp = 30
Data(8, 1).Shuriken = 2
Data(8, 1).Flame = 50
Data(8, 1).Side = 1.5
Data(8, 1).Back = 3
Data(8, 1).Tate = 9

Data(8, 2).Name = "∂Í◊”"
Data(8, 2).Hp = 7.5
Data(8, 2).Shuriken = 3
Data(8, 2).Flame = 75
Data(8, 2).Side = 2
Data(8, 2).Back = 3
Data(8, 2).Tate = 9

Data(8, 3).Name = "«øªØÃπøÀ"
Data(8, 3).Hp = 75
Data(8, 3).Shuriken = 0
Data(8, 3).Flame = 50
Data(8, 3).Side = 1
Data(8, 3).Back = 1
Data(8, 3).Tate = 9

Data(8, 5).Name = "∂Í◊”£®BOSS£©"
Data(8, 5).Hp = 7.5
Data(8, 5).Shuriken = 3
Data(8, 5).Flame = 75
Data(8, 5).Side = 2
Data(8, 5).Back = 3
Data(8, 5).Tate = 6

Data(8, 6).Name = "∫ÏÃÏ∂Í£®BOSS£©"
Data(8, 6).Hp = 225
Data(8, 6).Shuriken = 0
Data(8, 6).Flame = 50
Data(8, 6).Side = 1
Data(8, 6).Back = 0
Data(8, 6).Tate = 7

Rem 5A
Data(9, 1).Name = "∫⁄»Ã"
Data(9, 1).Hp = 20
Data(9, 1).Shuriken = 2
Data(9, 1).Flame = 50
Data(9, 1).Side = 1.5
Data(9, 1).Back = 3
Data(9, 1).Tate = 9

Data(9, 2).Name = "∫⁄∑…Õ∑"
Data(9, 2).Hp = 8.5
Data(9, 2).Shuriken = 2
Data(9, 2).Flame = 50
Data(9, 2).Side = 1.5
Data(9, 2).Back = 3
Data(9, 2).Tate = 9

Data(9, 3).Name = "∫Ï≈Æ»Ã"
Data(9, 3).Hp = 22.5
Data(9, 3).Shuriken = 3
Data(9, 3).Flame = 75
Data(9, 3).Side = 2
Data(9, 3).Back = 3
Data(9, 3).Tate = 9

Data(9, 5).Name = "∑…Õ∑£®BOSS£©"
Data(9, 5).Hp = 9
Data(9, 5).Shuriken = 3
Data(9, 5).Flame = 75
Data(9, 5).Side = 2
Data(9, 5).Back = 3
Data(9, 5).Tate = 7

Data(9, 6).Name = "Ω∏’£®BOSS£©"
Data(9, 6).Hp = 375
Data(9, 6).Shuriken = 1.5
Data(9, 6).Flame = 37.5
Data(9, 6).Side = 1.5
Data(9, 6).Back = 3
Data(9, 6).Tate = 8

Rem 5B
Data(10, 5).Name = "…ﬂ£®BOSS£©"
Data(10, 5).Hp = 9
Data(10, 5).Shuriken = 3
Data(10, 5).Flame = 75
Data(10, 5).Side = 1.5
Data(10, 5).Back = 3
Data(10, 5).Tate = 8

Data(10, 6).Name = "–˛æ≈…ﬂ£®BOSS£©"
Data(10, 6).Hp = 300
Data(10, 6).Shuriken = 0
Data(10, 6).Flame = 75
Data(10, 6).Side = 1
Data(10, 6).Back = 1
Data(10, 6).Tate = 9

Rem 6A
Data(11, 1).Name = "∫Ï≈Æ»Ã"
Data(11, 1).Hp = 22.5
Data(11, 1).Shuriken = 3
Data(11, 1).Flame = 75
Data(11, 1).Side = 2
Data(11, 1).Back = 3
Data(11, 1).Tate = 9

Data(11, 2).Name = "¬Ã∑…Õ∑"
Data(11, 2).Hp = 10
Data(11, 2).Shuriken = 4
Data(11, 2).Flame = 100
Data(11, 2).Side = 1.5
Data(11, 2).Back = 3
Data(11, 2).Tate = 9

Data(11, 3).Name = "π˜ Ω…Ò"
Data(11, 3).Hp = 30
Data(11, 3).Shuriken = 1.5
Data(11, 3).Flame = 37.5
Data(11, 3).Side = 1.5
Data(11, 3).Back = 1.5
Data(11, 3).Tate = 8

Data(11, 5).Name = "¬Ã∑…Õ∑£®BOSS£©"
Data(11, 5).Hp = 10
Data(11, 5).Shuriken = 4
Data(11, 5).Flame = 100
Data(11, 5).Side = 1.5
Data(11, 5).Back = 3
Data(11, 5).Tate = 7

Data(11, 6).Name = "øÃ£®BOSS£©"
Data(11, 6).Hp = 450
Data(11, 6).Shuriken = 2
Data(11, 6).Flame = 50
Data(11, 6).Side = 1.5
Data(11, 6).Back = 3
Data(11, 6).Tate = 8

Rem 6B
Data(12, 5).Name = "∑…Õ∑£®BOSS£©"
Data(12, 5).Hp = 10
Data(12, 5).Shuriken = 3
Data(12, 5).Flame = 75
Data(12, 5).Side = 2
Data(12, 5).Back = 3
Data(12, 5).Tate = 8

Data(12, 6).Name = "∞À√ÊÕı£®BOSS£©"
Data(12, 6).Hp = 675
Data(12, 6).Shuriken = 0
Data(12, 6).Flame = 37.5
Data(12, 6).Side = 1
Data(12, 6).Back = 1
Data(12, 6).Tate = 9

Rem 7A
Data(13, 1).Name = "¬Ã≈Æ»Ã"
Data(13, 1).Hp = 37.5
Data(13, 1).Shuriken = 3
Data(13, 1).Flame = 75
Data(13, 1).Side = 2
Data(13, 1).Back = 3
Data(13, 1).Tate = 9

Data(13, 2).Name = "∑…Õ∑"
Data(13, 2).Hp = 15
Data(13, 2).Shuriken = 3
Data(13, 2).Flame = 75
Data(13, 2).Side = 2
Data(13, 2).Back = 3
Data(13, 2).Tate = 9

Data(13, 3).Name = "«øªØ—ªÃÏπ∑"
Data(13, 3).Hp = 150
Data(13, 3).Shuriken = 1.5
Data(13, 3).Flame = 37.5
Data(13, 3).Side = 1.5
Data(13, 3).Back = 3
Data(13, 3).Tate = 8

Data(13, 4).Name = "œ„¬Ø£®BOSS£©"
Data(13, 4).Hp = 50
Data(13, 4).Shuriken = 2
Data(13, 4).Flame = 50
Data(13, 4).Side = 1
Data(13, 4).Back = 1
Data(13, 4).Tate = 9

Data(13, 5).Name = "ª®£®BOSS£©"
Data(13, 5).Hp = 12.5
Data(13, 5).Shuriken = 3
Data(13, 5).Flame = 75
Data(13, 5).Side = 2
Data(13, 5).Back = 3
Data(13, 5).Tate = 8

Data(13, 6).Name = "÷Ï»–£®BOSS£©"
Data(13, 6).Hp = 525
Data(13, 6).Shuriken = 2
Data(13, 6).Flame = 50
Data(13, 6).Side = 2
Data(13, 6).Back = 3
Data(13, 6).Tate = 9

Rem 7B
Data(14, 5).Name = "–°¡˙£®BOSS£©"
Data(14, 5).Hp = 12.5
Data(14, 5).Shuriken = 1
Data(14, 5).Flame = 25
Data(14, 5).Side = 1.5
Data(14, 5).Back = 3
Data(14, 5).Tate = 8

Data(14, 6).Name = "≤‘Ú‘¡˙£®BOSS£©"
Data(14, 6).Hp = 750
Data(14, 6).Shuriken = 0
Data(14, 6).Flame = 50
Data(14, 6).Side = 1.5
Data(14, 6).Back = 3
Data(14, 6).Tate = 9

Rem 8A
Data(15, 1).Name = "∫⁄»Ã"
Data(15, 1).Hp = 20
Data(15, 1).Shuriken = 2
Data(15, 1).Flame = 50
Data(15, 1).Side = 1.5
Data(15, 1).Back = 3
Data(15, 1).Tate = 9

Data(15, 2).Name = "∑…Õ∑"
Data(15, 2).Hp = 15
Data(15, 2).Shuriken = 3
Data(15, 2).Flame = 75
Data(15, 2).Side = 2
Data(15, 2).Back = 3
Data(15, 2).Tate = 9

Data(15, 3).Name = "«øªØπ˜ Ω…Ò"
Data(15, 3).Hp = 45
Data(15, 3).Shuriken = 1.5
Data(15, 3).Flame = 37.5
Data(15, 3).Side = 1.5
Data(15, 3).Back = 1.5
Data(15, 3).Tate = 9

Data(15, 4).Name = "∑…Õ∑£®BOSS£©"
Data(15, 4).Hp = 15
Data(15, 4).Shuriken = 3
Data(15, 4).Flame = 75
Data(15, 4).Side = 2
Data(15, 4).Back = 3
Data(15, 4).Tate = 9

Data(15, 5).Name = "«øªØπ˜ Ω…Ò£®BOSS£©"
Data(15, 5).Hp = 45
Data(15, 5).Shuriken = 1.5
Data(15, 5).Flame = 37.5
Data(15, 5).Side = 1.5
Data(15, 5).Back = 1.5
Data(15, 5).Tate = 9

Data(15, 6).Name = "√≈£®BOSS£©"
Data(15, 6).Hp = 300
Data(15, 6).Shuriken = 0
Data(15, 6).Flame = 50
Data(15, 6).Side = 0
Data(15, 6).Back = 0
Data(15, 6).Tate = 9

Rem 8B
Data(16, 1).Name = "◊œ»Ã"
Data(16, 1).Hp = 60
Data(16, 1).Shuriken = 2
Data(16, 1).Flame = 50
Data(16, 1).Side = 1.5
Data(16, 1).Back = 3
Data(16, 1).Tate = 9

Data(16, 2).Name = "∫Ï∑…Õ∑"
Data(16, 2).Hp = 15
Data(16, 2).Shuriken = 2
Data(16, 2).Flame = 50
Data(16, 2).Side = 1.5
Data(16, 2).Back = 3
Data(16, 2).Tate = 9

Data(16, 3).Name = "«øªØπ˜ Ω…Ò"
Data(16, 3).Hp = 45
Data(16, 3).Shuriken = 1.5
Data(16, 3).Flame = 37.5
Data(16, 3).Side = 1.5
Data(16, 3).Back = 1.5
Data(16, 3).Tate = 9

Data(16, 5).Name = "∑˚£®BOSS£©"
Data(16, 5).Hp = 15
Data(16, 5).Shuriken = 3
Data(16, 5).Flame = 75
Data(16, 5).Side = 2
Data(16, 5).Back = 5
Data(16, 5).Tate = 8

Data(16, 6).Name = "ÆbÕ¡ÚŒªß£®BOSS£©"
Data(16, 6).Hp = 900
Data(16, 6).Shuriken = 0
Data(16, 6).Flame = 50
Data(16, 6).Side = 1.5
Data(16, 6).Back = 3
Data(16, 6).Tate = 9

Rem EX
Data(0, 0).Name = "Ú⁄Úº"
Data(0, 0).Hp = 105
Data(0, 0).Shuriken = 1.5
Data(0, 0).Flame = 37.5
Data(0, 0).Side = 2
Data(0, 0).Back = 3
Data(0, 0).Tate = 7

Stageselect
S1
S2
Label17 = ""
Label15.Enabled = False
End Sub

Private Sub Label1_Click()
If ((Label1.ForeColor = &HFF&) And (Label2.ForeColor = &H80000012)) Or Label2.Visible = False Then
   S3
Else
    T1
End If

S2
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label18 <> "STAGE EX" Then
    If Button = 1 Then
        If Label10 = "Easy" Then
        S4 Label10, "Easy", "Normal"
        ElseIf Label10 = "Normal" Then S4 Label10, "Normal", "Hard"
        ElseIf Label10 = "Hard" Then S4 Label10, "Hard", "Super"
        End If
    End If

    If Button = 2 Then
        If Label10 = "Super" Then
        S5 Label10, "Super", "Hard"
        ElseIf Label10 = "Hard" Then S5 Label10, "Hard", "Normal"
        ElseIf Label10 = "Normal" Then S5 Label10, "Normal", "Easy"
        End If
    End If
End If
Picture1.Visible = True
Picture2.Picture = Picture4
Picture2.Visible = True
Hp
A = Data(xx, yy).Hp
S6
S2
If Combo1 <> "" Then Label19.Enabled = True
End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If Label11 = "Hotsuma" Then
    S4 Label11, "Hotsuma", "Moritsune"
    ElseIf Label11 = "Moritsune" Then S4 Label11, "Moritsune", "Joe Musashi"
    End If
End If

If Button = 2 Then
    If Label11 = "Joe Musashi" Then
    S5 Label11, "Joe Musashi", "Moritsune"
    ElseIf Label11 = "Moritsune" Then S5 Label11, "Moritsune", "Hotsuma"
    End If
End If
Picture1.Visible = True
Picture2.Picture = Picture4
Picture2.Visible = True
A = Data(xx, yy).Hp

S6
S1
S2
If Combo1 <> "" Then Label19.Enabled = True
End Sub

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If Label12 = "Slash" Then
        S4 Label12, "Slash", "Shuriken"
        ElseIf Label12 = "Shuriken" Then S4 Label12, "Shuriken", "Ka'en"
        ElseIf Label12 = "Ka'en" Then S4 Label12, "Ka'en", "Kamaitachi"
        ElseIf Label12 = "Kamaitachi" Then
            If yy = 6 And (xx = 1 Or xx = 4) Then S4 Label12, "Kamaitachi", "Missile"
            If yy = 6 And xx = 2 Then S4 Label12, "Kamaitachi", "Fallen"
    End If
End If

If Button = 2 Then
    If Label12 = "Kamaitachi" Then
        S5 Label12, "Kamaitachi", "Ka'en"
        ElseIf Label12 = "Ka'en" Then S5 Label12, "Ka'en", "Shuriken"
        ElseIf Label12 = "Shuriken" Then S5 Label12, "Shuriken", "Slash"
        ElseIf Label12 = "Missile" Then S5 Label12, "Missile", "Kamaitachi"
        ElseIf Label12 = "Fallen" Then S5 Label12, "Fallen", "Kamaitachi"
    End If
End If

S1
S2
End Sub

Private Sub Label13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Data(xx, yy).Name <> Data(15, 6).Name Then
    If Button = 1 Then
        If Label13 = "Front" Then
        S4 Label13, "Front", "Side"
        ElseIf Label13 = "Side" Then S4 Label13, "Side", "Back"
        End If
    End If
    
    If Button = 2 Then
        If Label13 = "Back" Then
        S5 Label13, "Back", "Side"
        ElseIf Label13 = "Side" Then S5 Label13, "Side", "Front"
        End If
    End If
End If
S2
End Sub

Private Sub Label14_Click()
If Label14.ForeColor = &H80000003 Then
    Label14.ForeColor = &H80000012
    Label16.ForeColor = &H80000003
Else
    Label14.ForeColor = &H80000003
End If

S2
End Sub

Private Sub Label15_Click()
If Label15.ForeColor = &H80000003 Then
    Label15.ForeColor = &H80000012
    If Label12 = "Shuriken" And Label21 = "Burst" Then
        S5 Label21, "Burst", "Single"
        Label16.ForeColor = &H80000003
    End If
Else: Label15.ForeColor = &H80000003
End If

S2
End Sub
Private Sub Label16_Click()
If Label16.ForeColor = &H80000003 Then
    Label16.ForeColor = &H80000012
    Label14.ForeColor = &H80000003
    If Label12 = "Shuriken" And Label21 = "Single" Then
        S4 Label21, "Single", "Burst"
        Label15.ForeColor = &H80000003
    End If
Else: Label16.ForeColor = &H80000003
End If

S2
End Sub

Private Sub Label18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If Label18 = "STAGE 1-A “°π‚" Then
    S4 Label18, "STAGE 1-A “°π‚", "STAGE 1-B ∆∆‹ä"
    ElseIf Label18 = "STAGE 1-B ∆∆‹ä" Then S4 Label18, "STAGE 1-B ∆∆‹ä", "STAGE 2-A È_Íñ"
    ElseIf Label18 = "STAGE 2-A È_Íñ" Then S4 Label18, "STAGE 2-A È_Íñ", "STAGE 2-B Œ‰«˙"
    ElseIf Label18 = "STAGE 2-B Œ‰«˙" Then S4 Label18, "STAGE 2-B Œ‰«˙", "STAGE 3-A ”Ò∫‚"
    ElseIf Label18 = "STAGE 3-A ”Ò∫‚" Then S4 Label18, "STAGE 3-A ”Ò∫‚", "STAGE 3-B ¡Æÿë"
    ElseIf Label18 = "STAGE 3-B ¡Æÿë" Then S4 Label18, "STAGE 3-B ¡Æÿë", "STAGE 4-A ÃÏòÿ"
    ElseIf Label18 = "STAGE 4-A ÃÏòÿ" Then S4 Label18, "STAGE 4-A ÃÏòÿ", "STAGE 4-B Œƒ«˙"
    ElseIf Label18 = "STAGE 4-B Œƒ«˙" Then S4 Label18, "STAGE 4-B Œƒ«˙", "STAGE 5-A ÃÏôC"
    ElseIf Label18 = "STAGE 5-A ÃÏôC" Then S4 Label18, "STAGE 5-A ÃÏôC", "STAGE 5-B ¬ª¥Ê"
    ElseIf Label18 = "STAGE 5-B ¬ª¥Ê" Then S4 Label18, "STAGE 5-B ¬ª¥Ê", "STAGE 6-A ÃÏËØ"
    ElseIf Label18 = "STAGE 6-A ÃÏËØ" Then S4 Label18, "STAGE 6-A ÃÏËØ", "STAGE 6-B æﬁ√≈"
    ElseIf Label18 = "STAGE 6-B æﬁ√≈" Then S4 Label18, "STAGE 6-B æﬁ√≈", "STAGE 7-A ÿù¿«"
    ElseIf Label18 = "STAGE 7-A ÿù¿«" Then S4 Label18, "STAGE 7-A ÿù¿«", "STAGE 7-B ÃÏò–"
    ElseIf Label18 = "STAGE 7-B ÃÏò–" Then S4 Label18, "STAGE 7-B ÃÏò–", "STAGE 8-A ±±≥Ω"
    ElseIf Label18 = "STAGE 8-A ±±≥Ω" Then S4 Label18, "STAGE 8-A ±±≥Ω", "STAGE 8-B Ã´“ª"
    ElseIf Label18 = "STAGE 8-B Ã´“ª" Then S4 Label18, "STAGE 8-B Ã´“ª", "STAGE EX"
    End If
End If

If Button = 2 Then
    If Label18 = "STAGE EX" Then
    S5 Label18, "STAGE EX", "STAGE 8-B Ã´“ª"
    ElseIf Label18 = "STAGE 8-B Ã´“ª" Then S5 Label18, "STAGE 8-B Ã´“ª", "STAGE 8-A ±±≥Ω"
    ElseIf Label18 = "STAGE 8-A ±±≥Ω" Then S5 Label18, "STAGE 8-A ±±≥Ω", "STAGE 7-B ÃÏò–"
    ElseIf Label18 = "STAGE 7-B ÃÏò–" Then S5 Label18, "STAGE 7-B ÃÏò–", "STAGE 7-A ÿù¿«"
    ElseIf Label18 = "STAGE 7-A ÿù¿«" Then S5 Label18, "STAGE 7-A ÿù¿«", "STAGE 6-B æﬁ√≈"
    ElseIf Label18 = "STAGE 6-B æﬁ√≈" Then S5 Label18, "STAGE 6-B æﬁ√≈", "STAGE 6-A ÃÏËØ"
    ElseIf Label18 = "STAGE 6-A ÃÏËØ" Then S5 Label18, "STAGE 6-A ÃÏËØ", "STAGE 5-B ¬ª¥Ê"
    ElseIf Label18 = "STAGE 5-B ¬ª¥Ê" Then S5 Label18, "STAGE 5-B ¬ª¥Ê", "STAGE 5-A ÃÏôC"
    ElseIf Label18 = "STAGE 5-A ÃÏôC" Then S5 Label18, "STAGE 5-A ÃÏôC", "STAGE 4-B Œƒ«˙"
    ElseIf Label18 = "STAGE 4-B Œƒ«˙" Then S5 Label18, "STAGE 4-B Œƒ«˙", "STAGE 4-A ÃÏòÿ"
    ElseIf Label18 = "STAGE 4-A ÃÏòÿ" Then S5 Label18, "STAGE 4-A ÃÏòÿ", "STAGE 3-B ¡Æÿë"
    ElseIf Label18 = "STAGE 3-B ¡Æÿë" Then S5 Label18, "STAGE 3-B ¡Æÿë", "STAGE 3-A ”Ò∫‚"
    ElseIf Label18 = "STAGE 3-A ”Ò∫‚" Then S5 Label18, "STAGE 3-A ”Ò∫‚", "STAGE 2-B Œ‰«˙"
    ElseIf Label18 = "STAGE 2-B Œ‰«˙" Then S5 Label18, "STAGE 2-B Œ‰«˙", "STAGE 2-A È_Íñ"
    ElseIf Label18 = "STAGE 2-A È_Íñ" Then S5 Label18, "STAGE 2-A È_Íñ", "STAGE 1-B ∆∆‹ä"
    ElseIf Label18 = "STAGE 1-B ∆∆‹ä" Then S5 Label18, "STAGE 1-B ∆∆‹ä", "STAGE 1-A “°π‚"
    End If
End If
Label19.Enabled = False
Combo1.Clear
Stageselect
If Label12 = "Missile" Or Label12 = "Fallen" Then Label12 = "Shuriken"
S6
S2
Label14.ForeColor = &H80000003
Label14.Enabled = False
Label15.ForeColor = &H80000003
Label15.Enabled = False
Label16.ForeColor = &H80000003
Label16.Enabled = False
Picture1.Visible = True
Picture2.Picture = Picture4
Picture2.Visible = True
T0
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Hp

If Label12 = "Slash" Or Label12 = "Shuriken" Then
    B = C * D * E * F * G * H
    If Label12 = "Shuriken" And Label21 = "Single" Then B = C * E * F
    If (yy = 3 And (xx = 11 Or xx = 15 Or xx = 16)) And Label12 = "Shuriken" And Label21 = "Single" Then B = C
    If Label11 = "Hotsuma" And (Label18 = "STAGE 1-A “°π‚" Or Label18 = "STAGE 1-B ∆∆‹ä") Then B = C * E * F * G * H
Else
    If Label12 = "Ka'en" Or Label12 = "Kamaitachi" Then B = C * E * F
    If Label12 = "Missile" Or Label12 = "Fallen" Then B = C
    If yy = 6 And xx = 16 And Label10 = "Easy" And Label12 = "Ka'en" Then B = 2 * C * E * F
    
    If Label12 = "Missile" Then
        If Label10 = "Super" And xx = 1 And yy = 6 And A - B < 2.5 Then
            If A <= 2.5 Then
                B = 0
            Else
                B = A - 2.5
            End If
        End If

        If Label10 = "Hard" And xx = 1 And yy = 6 And A - B < 3.75 Then
            If A <= 3.75 Then
                B = 0
            Else
                B = A - 3.75
            End If
        End If

        If (Label10 = "Normal" Or Label10 = "Easy") And xx = 1 And yy = 6 And A - B < 5 Then
            If A <= 5 Then
                B = 0
            Else
                B = A - 5
            End If
        End If
        
        
               
        If Label10 = "Super" And xx = 4 And yy = 6 And A - B < 25 Then
            If A <= 25 Then
                B = 0
            Else
                B = A - 25
            End If
        End If

        If Label10 = "Hard" And xx = 4 And yy = 6 And A - B < 37.5 Then
            If A <= 37.5 Then
                B = 0
            Else
                B = A - 37.5
            End If
        End If

        If (Label10 = "Normal" Or Label10 = "Easy") And xx = 4 And yy = 6 And A - B < 50 Then
            If A <= 50 Then
                B = 0
            Else
                B = A - 50
            End If
        End If
    End If

    
    
    
        
    
    If Label12 = "Fallen" Then
        If Label10 = "Super" And xx = 2 And yy = 6 And A - B < 7.5 Then
            If A <= 7.5 Then
                B = 0
            Else
                B = A - 7.5
            End If
        End If

        If Label10 = "Hard" And xx = 2 And yy = 6 And A - B < 11.25 Then
            If A <= 11.25 Then
                B = 0
            Else
                B = A - 11.25
            End If
        End If

        If (Label10 = "Normal" Or Label10 = "Easy") And xx = 2 And yy = 6 And A - B < 15 Then
            If A <= 15 Then
                B = 0
            Else
                B = A - 15
            End If
        End If
    End If
End If


A = A - B
Percent = B / Data(xx, yy).Hp

If B > 0 Then
    If H = 1 Then
        Picture2.Line (A * 6000 / Data(xx, yy).Hp, 150)-((A + B) * 6000 / Data(xx, yy).Hp, 200), &HFFFF00, BF
    Else
        Picture2.Line ((A + B / 2) * 6000 / Data(xx, yy).Hp, 150)-((A + B) * 6000 / Data(xx, yy).Hp, 200), &HC0C000, BF
        Picture2.Line (A * 6000 / Data(xx, yy).Hp, 150)-((A + B / 2) * 6000 / Data(xx, yy).Hp, 200), &HFFFF00, BF
    End If
    
    If A <= 0 Then
        A = 0
        If (xx = 1 And yy = 4) Or (xx = 6 And yy = 4) Or (xx = 6 And yy = 0) Or (xx = 13 And yy = 4) Then
        
        Else
            If Label12 = "Slash" Then
                If Label8.ForeColor = &HFF& Then
                    T9
                ElseIf Label7.ForeColor = &HFF& Then
                    T8
                ElseIf Label6.ForeColor = &HFF& Then
                    T7
                ElseIf Label5.ForeColor = &HFF& Then
                    T6
                ElseIf Label4.ForeColor = &HFF& Then
                    T5
                ElseIf Label3.ForeColor = &HFF& Then
                    T4
                ElseIf Label2.ForeColor = &HFF& Then
                    T3
                ElseIf Label1.ForeColor = &HFF& Then
                    T2
                Else
                    T1
                End If
            End If
        End If
    End If
End If
Label17 = Format(A, "0.00###") + "/" + Format(Data(xx, yy).Hp) + "  " + Format(B) + "/" + Format(Percent, "0.00###%")

'From1.Cls
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""
'From1.Print ""

'From1.Print "A ="; A
'From1.Print "B ="; B
'From1.Print "C ="; C
'From1.Print "D ="; D
'From1.Print "E ="; E
'From1.Print "F ="; F
'From1.Print "G ="; G
'From1.Print "H ="; H
'From1.Print "xx ="; xx
'From1.Print "yy ="; yy
'From1.Print "A * 6000 / Data(xx, yy).Hp ="; A * 6000 / Data(xx, yy).Hp
'From1.Print "(A + B) * 6000 / Data(xx, yy).Hp ="; (A + B) * 6000 / Data(xx, yy).Hp

End Sub

Private Sub Label19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If B > 0 Then
    DelayLoop 0.3
    Picture2.Line (A * 6000 / Data(xx, yy).Hp, 150)-(6000, 200), &H404040, BF
End If

If A = 0 Then
    S3
    DelayLoop 0.2
    Picture1.Visible = False
    Picture2.Visible = False
    Label19.Enabled = False
End If

End Sub

Private Sub Label2_Click()
If ((Label2.ForeColor = &HFF&) And (Label3.ForeColor = &H80000012)) Or Label3.Visible = False Then
   S3
Else
T2
End If

S2
End Sub

Private Sub Label20_Click()
Retry
End Sub


Private Sub Label21_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If Label21 = "Single" Then S4 Label21, "Single", "Burst"
    Label15.ForeColor = &H80000003
End If
    
If Button = 2 Then
    If Label21 = "Burst" Then S5 Label21, "Burst", "Single"
    Label16.ForeColor = &H80000003
End If
End Sub

Private Sub Label3_Click()
If ((Label3.ForeColor = &HFF&) And (Label4.ForeColor = &H80000012)) Or Label4.Visible = False Then
   S3
Else
T3
End If

S2
End Sub

Private Sub Label4_Click()
If ((Label4.ForeColor = &HFF&) And (Label5.ForeColor = &H80000012)) Or Label5.Visible = False Then
   S3
Else
T4
End If

S2
End Sub

Private Sub Label5_Click()
If ((Label5.ForeColor = &HFF&) And (Label6.ForeColor = &H80000012)) Or Label6.Visible = False Then
   S3
Else
T5
End If

S2
End Sub

Private Sub Label6_Click()
If ((Label6.ForeColor = &HFF&) And (Label7.ForeColor = &H80000012)) Or Label7.Visible = False Then
   S3
Else
T6
End If

S2
End Sub

Private Sub Label7_Click()
If ((Label7.ForeColor = &HFF&) And (Label8.ForeColor = &H80000012)) Or Label8.Visible = False Then
   S3
Else
T7
End If

S2
End Sub

Private Sub Label8_Click()
If ((Label8.ForeColor = &HFF&) And (Label9.ForeColor = &H80000012)) Or Label9.Visible = False Then
   S3
Else
T8
End If

S2
End Sub

Private Sub Label9_Click()
S3
S2
End Sub

