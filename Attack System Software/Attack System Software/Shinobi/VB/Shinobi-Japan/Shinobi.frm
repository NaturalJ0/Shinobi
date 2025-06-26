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
      Caption         =   "前"
      Height          =   180
      Left            =   6720
      TabIndex        =   8
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "在"
      Height          =   180
      Left            =   6480
      TabIndex        =   7
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "烈"
      Height          =   180
      Left            =   6240
      TabIndex        =   6
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "陣"
      Height          =   180
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "皆"
      Height          =   180
      Left            =   5760
      TabIndex        =   4
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "者"
      Height          =   180
      Left            =   5520
      TabIndex        =   3
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "闘"
      Height          =   180
      Left            =   5280
      TabIndex        =   2
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "兵"
      Height          =   180
      Left            =   5040
      TabIndex        =   1
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "臨"
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
      Caption         =   "STAGE 1-A 摇光"
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
Sub S4(L As Control, Str1 As String, Str2 As String)
str3$ = Str1 + "  "
str3 = str3 + Str2
len1% = Len(Str1)
len2% = Len(Str2)
m% = len1 + 2
K% = 1
y# = 0.1 / (m + 1)
Z = max(len1, len2)
For K = 1 To m + 1
    L = Mid(str3, K, K + Z - 1)
    DelayLoop y
Next K
L = Right(L, len2)
End Sub
Sub S5(L As Control, Str1 As String, Str2 As String)
str3$ = Str2 + "  "
str3 = str3 + Str1
len1% = Len(Str1)
len2% = Len(Str2)
m% = len1 + 2
K% = 1
y# = 0.1 / (m + 1)
Z = max(len1, len2)
For K = 1 To len1 + 1
    L = Mid(str3, len1 + 2 - K, len1 + 2 - K + Z - 1)
    DelayLoop y
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
Function max(x, y)
If x >= y Then max = x Else max = y
End Function
Sub Stageselect()
If Label18 = Stage(1) Then
    xx = 1
    ElseIf Label18 = Stage(2) Then xx = 2
    ElseIf Label18 = Stage(3) Then xx = 3
    ElseIf Label18 = Stage(4) Then xx = 4
    ElseIf Label18 = Stage(5) Then xx = 5
    ElseIf Label18 = Stage(6) Then xx = 6
    ElseIf Label18 = Stage(7) Then xx = 7
    ElseIf Label18 = Stage(8) Then xx = 8
    ElseIf Label18 = Stage(9) Then xx = 9
    ElseIf Label18 = Stage(10) Then xx = 10
    ElseIf Label18 = Stage(11) Then xx = 11
    ElseIf Label18 = Stage(12) Then xx = 12
    ElseIf Label18 = Stage(13) Then xx = 13
    ElseIf Label18 = Stage(14) Then xx = 14
    ElseIf Label18 = Stage(15) Then xx = 15
    ElseIf Label18 = Stage(16) Then xx = 16
    ElseIf Label18 = Stage(17) Then xx = 0
End If
For i = 1 To 6
    If Data(xx, i).Name <> "" Then Combo1.AddItem Data(xx, i).Name
Next i
If Data(xx, 0).Name <> "" Then Combo1.AddItem Data(xx, 0).Name
If Label18 = Stage(17) Then
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
pre
Label1 = Tatename(1)
Label2 = Tatename(2)
Label3 = Tatename(3)
Label4 = Tatename(4)
Label5 = Tatename(5)
Label6 = Tatename(6)
Label7 = Tatename(7)
Label8 = Tatename(8)
Label9 = Tatename(9)
Label18 = Stage(1)
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

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Label18 <> Stage(17) Then
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

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub Label13_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub Label18_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    If Label18 = Stage(1) Then
    S4 Label18, Stage(1), Stage(2)
    ElseIf Label18 = Stage(2) Then S4 Label18, Stage(2), Stage(3)
    ElseIf Label18 = Stage(3) Then S4 Label18, Stage(3), Stage(4)
    ElseIf Label18 = Stage(4) Then S4 Label18, Stage(4), Stage(5)
    ElseIf Label18 = Stage(5) Then S4 Label18, Stage(5), Stage(6)
    ElseIf Label18 = Stage(6) Then S4 Label18, Stage(6), Stage(7)
    ElseIf Label18 = Stage(7) Then S4 Label18, Stage(7), Stage(8)
    ElseIf Label18 = Stage(8) Then S4 Label18, Stage(8), Stage(9)
    ElseIf Label18 = Stage(9) Then S4 Label18, Stage(9), Stage(10)
    ElseIf Label18 = Stage(10) Then S4 Label18, Stage(10), Stage(11)
    ElseIf Label18 = Stage(11) Then S4 Label18, Stage(11), Stage(12)
    ElseIf Label18 = Stage(12) Then S4 Label18, Stage(12), Stage(13)
    ElseIf Label18 = Stage(13) Then S4 Label18, Stage(13), Stage(14)
    ElseIf Label18 = Stage(14) Then S4 Label18, Stage(14), Stage(15)
    ElseIf Label18 = Stage(15) Then S4 Label18, Stage(15), Stage(16)
    ElseIf Label18 = Stage(16) Then S4 Label18, Stage(16), Stage(17)
    End If
End If

If Button = 2 Then
    If Label18 = Stage(17) Then
    S5 Label18, Stage(17), Stage(16)
    ElseIf Label18 = Stage(16) Then S5 Label18, Stage(16), Stage(15)
    ElseIf Label18 = Stage(15) Then S5 Label18, Stage(15), Stage(14)
    ElseIf Label18 = Stage(14) Then S5 Label18, Stage(14), Stage(13)
    ElseIf Label18 = Stage(13) Then S5 Label18, Stage(13), Stage(12)
    ElseIf Label18 = Stage(12) Then S5 Label18, Stage(12), Stage(11)
    ElseIf Label18 = Stage(11) Then S5 Label18, Stage(11), Stage(10)
    ElseIf Label18 = Stage(10) Then S5 Label18, Stage(10), Stage(9)
    ElseIf Label18 = Stage(9) Then S5 Label18, Stage(9), Stage(8)
    ElseIf Label18 = Stage(8) Then S5 Label18, Stage(8), Stage(7)
    ElseIf Label18 = Stage(7) Then S5 Label18, Stage(7), Stage(6)
    ElseIf Label18 = Stage(6) Then S5 Label18, Stage(6), Stage(5)
    ElseIf Label18 = Stage(5) Then S5 Label18, Stage(5), Stage(4)
    ElseIf Label18 = Stage(4) Then S5 Label18, Stage(4), Stage(3)
    ElseIf Label18 = Stage(3) Then S5 Label18, Stage(3), Stage(2)
    ElseIf Label18 = Stage(2) Then S5 Label18, Stage(2), Stage(1)
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

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Hp

If Label12 = "Slash" Or Label12 = "Shuriken" Then
    B = C * D * E * F * G * H
    If Label12 = "Shuriken" And Label21 = "Single" Then B = C * E * F
    If (yy = 3 And (xx = 11 Or xx = 15 Or xx = 16)) And Label12 = "Shuriken" And Label21 = "Single" Then B = C
    If Label11 = "Hotsuma" And (Label18 = Stage(1) Or Label18 = Stage(2)) Then B = C * E * F * G * H
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

Private Sub Label19_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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


Private Sub Label21_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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

