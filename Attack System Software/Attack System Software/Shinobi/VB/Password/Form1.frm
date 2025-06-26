VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   "Password"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   4230
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text26 
      Height          =   270
      Left            =   1298
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   5280
      Width           =   1635
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1253
      TabIndex        =   2
      Top             =   952
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1253
      TabIndex        =   0
      Top             =   585
      Width           =   615
   End
   Begin VB.TextBox Text24 
      Height          =   270
      Left            =   3293
      TabIndex        =   23
      Top             =   4545
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Password"
      Height          =   255
      Left            =   2340
      TabIndex        =   50
      Top             =   4898
      Width           =   1575
   End
   Begin VB.TextBox Text25 
      Height          =   270
      Left            =   1253
      TabIndex        =   25
      Top             =   4890
      Width           =   615
   End
   Begin VB.TextBox Text23 
      Height          =   270
      Left            =   1253
      TabIndex        =   22
      Top             =   4545
      Width           =   615
   End
   Begin VB.TextBox Text22 
      Height          =   270
      Left            =   3293
      TabIndex        =   21
      Top             =   4185
      Width           =   615
   End
   Begin VB.TextBox Text21 
      Height          =   270
      Left            =   1253
      TabIndex        =   20
      Top             =   4185
      Width           =   615
   End
   Begin VB.TextBox Text20 
      Height          =   270
      Left            =   3293
      TabIndex        =   19
      Top             =   3825
      Width           =   615
   End
   Begin VB.TextBox Text19 
      Height          =   270
      Left            =   1253
      TabIndex        =   18
      Top             =   3825
      Width           =   615
   End
   Begin VB.TextBox Text18 
      Height          =   270
      Left            =   3293
      TabIndex        =   17
      Top             =   3465
      Width           =   615
   End
   Begin VB.TextBox Text17 
      Height          =   270
      Left            =   1253
      TabIndex        =   16
      Top             =   3465
      Width           =   615
   End
   Begin VB.TextBox Text13 
      Height          =   270
      Left            =   1253
      TabIndex        =   12
      Top             =   2715
      Width           =   615
   End
   Begin VB.TextBox Text14 
      Height          =   270
      Left            =   3293
      TabIndex        =   13
      Top             =   2715
      Width           =   615
   End
   Begin VB.TextBox Text15 
      Height          =   270
      Left            =   1253
      TabIndex        =   14
      Top             =   3075
      Width           =   615
   End
   Begin VB.TextBox Text16 
      Height          =   270
      Left            =   3293
      TabIndex        =   15
      Top             =   3075
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Left            =   3293
      TabIndex        =   11
      Top             =   2355
      Width           =   615
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   1253
      TabIndex        =   10
      Top             =   2355
      Width           =   615
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   3293
      TabIndex        =   9
      Top             =   1995
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1253
      TabIndex        =   8
      Top             =   1995
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   3293
      TabIndex        =   7
      Top             =   1635
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1253
      TabIndex        =   6
      Top             =   1635
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   3293
      TabIndex        =   5
      Top             =   1275
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1253
      TabIndex        =   4
      Top             =   1275
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   3293
      TabIndex        =   3
      Top             =   945
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   3293
      TabIndex        =   1
      Top             =   592
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 2A"
      Height          =   255
      Left            =   300
      TabIndex        =   51
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入下列各关中封印石的数量"
      Height          =   180
      Left            =   855
      TabIndex        =   49
      Top             =   195
      Width           =   2520
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EX FINAL"
      Height          =   180
      Left            =   300
      TabIndex        =   48
      Top             =   4935
      Width           =   720
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EX 08"
      Height          =   180
      Left            =   2340
      TabIndex        =   47
      Top             =   4560
      Width           =   450
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EX 07"
      Height          =   180
      Left            =   300
      TabIndex        =   46
      Top             =   4560
      Width           =   450
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EX 06"
      Height          =   180
      Left            =   2340
      TabIndex        =   45
      Top             =   4200
      Width           =   450
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EX 05"
      Height          =   180
      Left            =   300
      TabIndex        =   44
      Top             =   4200
      Width           =   450
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EX 04"
      Height          =   180
      Left            =   2340
      TabIndex        =   43
      Top             =   3840
      Width           =   450
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EX 03"
      Height          =   180
      Left            =   300
      TabIndex        =   42
      Top             =   3840
      Width           =   450
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EX 02"
      Height          =   180
      Left            =   2340
      TabIndex        =   41
      Top             =   3480
      Width           =   450
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EX 01"
      Height          =   180
      Left            =   300
      TabIndex        =   40
      Top             =   3480
      Width           =   450
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 7A"
      Height          =   180
      Left            =   300
      TabIndex        =   39
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 7B"
      Height          =   180
      Left            =   2340
      TabIndex        =   38
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 8A"
      Height          =   180
      Left            =   300
      TabIndex        =   37
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 8B"
      Height          =   180
      Left            =   2340
      TabIndex        =   36
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 1A"
      Height          =   255
      Left            =   300
      TabIndex        =   35
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 1B"
      Height          =   255
      Left            =   2340
      TabIndex        =   34
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 6B"
      Height          =   180
      Left            =   2340
      TabIndex        =   33
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 6A"
      Height          =   180
      Left            =   300
      TabIndex        =   32
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 5B"
      Height          =   180
      Left            =   2340
      TabIndex        =   31
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 5A"
      Height          =   180
      Left            =   300
      TabIndex        =   30
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 4B"
      Height          =   180
      Left            =   2340
      TabIndex        =   29
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 4A"
      Height          =   180
      Left            =   300
      TabIndex        =   28
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 3B"
      Height          =   180
      Left            =   2340
      TabIndex        =   27
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 3A"
      Height          =   180
      Left            =   300
      TabIndex        =   26
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAGE 2B"
      Height          =   255
      Left            =   2340
      TabIndex        =   24
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
total = 0
If Text1 = "0" Then total = total + 1
If Text2 = "6" Then total = total + 1
If Text3 = "0" Then total = total + 1
If Text4 = "18" Then total = total + 1
If Text5 = "0" Then total = total + 1
If Text6 = "10" Then total = total + 1
If Text7 = "0" Then total = total + 1
If Text8 = "18" Then total = total + 1
If Text9 = "14" Then total = total + 1
If Text10 = "0" Then total = total + 1
If Text11 = "19" Then total = total + 1
If Text12 = "0" Then total = total + 1
If Text13 = "0" Then total = total + 1
If Text14 = "0" Then total = total + 1
If Text15 = "12" Then total = total + 1
If Text16 = "20" Then total = total + 1
If Text17 = "0" Then total = total + 1
If Text18 = "6" Then total = total + 1
If Text19 = "3" Then total = total + 1
If Text20 = "4" Then total = total + 1
If Text21 = "0" Then total = total + 1
If Text22 = "3" Then total = total + 1
If Text23 = "0" Then total = total + 1
If Text24 = "3" Then total = total + 1
If Text25 = "0" Then total = total + 1
If total = 25 Then Text26 = "~!@#$%^&*()_+|" Else Text26 = "      错误"
End Sub

