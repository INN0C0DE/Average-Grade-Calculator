VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   Caption         =   "Average Grade Calculator"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdcompute 
      BackColor       =   &H008080FF&
      Caption         =   "COMPUTE"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   4695
   End
   Begin VB.TextBox Txtscience 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Top             =   3480
      Width           =   3735
   End
   Begin VB.TextBox Txtmath 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Top             =   2520
      Width           =   3735
   End
   Begin VB.TextBox Txtenglish 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   7920
      X2              =   7920
      Y1              =   120
      Y2              =   8400
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   120
      X2              =   7920
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   120
      X2              =   7920
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   8400
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Made by: Raphael Arnaldo Cruz"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   7800
      Width           =   3135
   End
   Begin VB.Label lblaverage 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   9
      Top             =   4440
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "AVERAGE:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Science Grade:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Math Grade:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "English Grade:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Average Grade Calculator"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Cmdcompute_Click()
Dim average As Integer
Dim units As Integer
Const math = 4
Const science = 3
Const english = 3
units = math + science + english
average = ((Val(Txtenglish) * english + Val(Txtscience) * science + Val(Txtmath) * math)) / units
lblaverage.Caption = average

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_Change()

End Sub

