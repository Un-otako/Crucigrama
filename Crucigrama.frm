VERSION 5.00
Begin VB.Form Crucigrama 
   Caption         =   "Nivel 1"
   ClientHeight    =   8310
   ClientLeft      =   4275
   ClientTop       =   600
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   Picture         =   "Crucigrama.frx":0000
   ScaleHeight     =   8310
   ScaleWidth      =   11415
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   48
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Cmdsuiguiente 
      Caption         =   "Siguiente"
      Height          =   615
      Left            =   7680
      TabIndex        =   46
      Top             =   7320
      Width           =   3375
   End
   Begin VB.TextBox Text32 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      MaxLength       =   1
      TabIndex        =   45
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text31 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      MaxLength       =   1
      TabIndex        =   44
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text26 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   43
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox Text30 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      MaxLength       =   1
      TabIndex        =   42
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text29 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   41
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text28 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   40
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text27 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   39
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   38
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox Text25 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   37
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text24 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   36
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text23 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   34
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Resultado"
      Height          =   615
      Left            =   360
      TabIndex        =   33
      Top             =   7320
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   615
      Left            =   4080
      TabIndex        =   32
      Top             =   7320
      Width           =   3375
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   24
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   23
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   22
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   21
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   20
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   19
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   18
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   17
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   16
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   15
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   14
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Horizontal"
      Height          =   1695
      Left            =   7440
      TabIndex        =   9
      Top             =   2160
      Width           =   3495
      Begin VB.Label Label10 
         Caption         =   "5-Servicio que se la da a una computadora para que funcione correctamente."
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label9 
         Caption         =   "4-Conexion en que los componentes el�ctricos se conectan uno a continuaci�n del otro."
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vertical"
      Height          =   2175
      Left            =   7440
      TabIndex        =   8
      Top             =   0
      Width           =   3495
      Begin VB.Label Label3 
         Caption         =   "3-Es un firmware de sistema b�sico de entrada y salida"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "2-Componente para conectar dispositivos externos."
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "1-Palabra utilizada para referise a objetos que usan electrisidad."
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaxLength       =   1
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "Puntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   47
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   35
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   29
      Top             =   4080
      Width           =   120
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   1680
      Width           =   120
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4440
      TabIndex        =   27
      Top             =   3240
      Width           =   120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   26
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2160
      TabIndex        =   25
      Top             =   720
      Width           =   120
   End
End
Attribute VB_Name = "Crucigrama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'Juego creado por Mateo Uriel Villafañe Barboza el 14|8|21 y terminado el 27|8|21
'
'Alumno de 6° año de la escuela N°20 Antonio Berni Argentina Villamercedes, San Luis.
'
'Puedes editar este juego pero tienes que mencionarme.
'------------------------------------------------------------------------------------

Option Explicit

Dim res1 As Integer

Dim res2 As Integer

Dim res3 As Integer

Dim res4 As Integer

Dim res5 As Integer

Dim resultado As Integer
Private Sub Command1_Click()

End

End Sub
Private Sub Command2_Click()

Call contenido

resultado = res1 + res2 + res3 + res4 + res5

Label11.Caption = resultado

End Sub
Private Sub Cmdsuiguiente_Click()

Crucigrama2.Show

End Sub
Private Sub Form_Load()

Label11.Caption = 0

End Sub
Public Sub contenido()

If Text1.Text = "u" And Text2.Text = "s" And Text3.Text = "b" Then

 res2 = "20"

ElseIf Text1.Text = "U" And Text2.Text = "S" And Text3.Text = "B" Then

 res2 = "20"
 
Else
 
 res2 = "0"

End If

If Text4.Text = "e" And Text5.Text = "r" And Text6.Text = "i" And Text7.Text = "e" Then

 res4 = "20"
 
ElseIf Text4.Text = "E" And Text5.Text = "R" And Text6.Text = "I" And Text7.Text = "E" Then

 res4 = "20"
 
Else
 
 res4 = "0"

End If

If Text10.Text = "e" And Text9.Text = "l" And Text16.Text = "c" And Text17.Text = "t" And Text18.Text = "r" And Text19.Text = "o" And Text22.Text = "n" And Text26.Text = "i" And Text8.Text = "c" And Text27.Text = "o" Then

 res1 = "20"

ElseIf Text10.Text = "E" And Text9.Text = "R" And Text16.Text = "I" And Text17.Text = "E" And Text18.Text = "R" And Text19.Text = "O" And Text22.Text = "O" And Text26.Text = "I" And Text8.Text = "C" And Text27.Text = "O" Then

 res1 = "20"
 
Else
 
 res1 = "0"

End If

If Text13.Text = "b" And Text11.Text = "i" And Text14.Text = "o" And Text15.Text = "s" Then

 res3 = "20"
 
ElseIf Text13.Text = "B" And Text11.Text = "I" And Text14.Text = "O" And Text15.Text = "S" Then

 res3 = "20"
 
Else
 
 res3 = "0"

End If

If Text20.Text = "m" And Text21.Text = "a" And Text22.Text = "n" And Text23.Text = "t" And Text24.Text = "e" And Text25.Text = "n" And Text11.Text = "i" And Text12.Text = "m" And Text28.Text = "i" And Text29.Text = "e" And Text30.Text = "n" And Text31.Text = "t" And Text32.Text = "o" Then

 res5 = "20"
 
ElseIf Text20.Text = "M" And Text21.Text = "A" And Text22.Text = "N" And Text23.Text = "T" And Text24.Text = "E" And Text25.Text = "N" And Text11.Text = "I" And Text12.Text = "M" And Text28.Text = "I" And Text29.Text = "E" And Text30.Text = "N" And Text31.Text = "T" And Text32.Text = "O" Then

 res5 = "20"
 
Else
 
 res5 = "0"

End If

End Sub
