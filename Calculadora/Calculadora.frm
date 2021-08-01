VERSION 5.00
Begin VB.Form frmCalculadora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculadora"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnClearUltimo 
      Caption         =   "<-"
      Height          =   615
      Left            =   1800
      TabIndex        =   22
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton btnNegativar 
      Caption         =   "+/-"
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton btn0 
      Caption         =   "0"
      Height          =   615
      Left            =   960
      TabIndex        =   20
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox lblOperador 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton btnDivisao 
      Caption         =   "%"
      Height          =   615
      Left            =   1800
      TabIndex        =   18
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton btnMultiplicacao 
      Caption         =   "*"
      Height          =   615
      Left            =   2640
      TabIndex        =   17
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton btnClearPrincipal 
      Caption         =   "CE"
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton btn9 
      Caption         =   "9"
      Height          =   615
      Left            =   1800
      TabIndex        =   15
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton btn8 
      Caption         =   "8"
      Height          =   615
      Left            =   960
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton btn6 
      Caption         =   "6"
      Height          =   615
      Left            =   1800
      TabIndex        =   13
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton btn5 
      Caption         =   "5"
      Height          =   615
      Left            =   960
      TabIndex        =   12
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton btn4 
      Caption         =   "4"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton btn7 
      Caption         =   "7"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "C"
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox lblResultado 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Text            =   "0"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton btnSubtracao 
      Caption         =   "-"
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton btnSoma 
      Caption         =   "+"
      Height          =   615
      Left            =   2640
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton btnIgual 
      Caption         =   "="
      Height          =   1335
      Left            =   2640
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton btn3 
      Caption         =   "3"
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton btn2 
      Caption         =   "2"
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton btn1 
      Caption         =   "1"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox lblPrincipal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox lblConta 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmCalculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adicionaValor(intNumero As Integer)
   
   Dim intDisplay As Variant
   
   intDisplay = Val(lblPrincipal.Text)
   
   If intDisplay = 0 Then
      lblPrincipal.Text = intNumero
   Else
      lblPrincipal.Text = CStr(intDisplay) + "" + CStr(intNumero)
   End If

End Sub

Private Sub adicionaOperador(strOperador As String)

   Dim varConta As Variant
   Dim intSubtotal As Integer
   
   varConta = lblConta.Text
   
   If varConta = "" Then
      lblConta.Text = CStr(lblPrincipal.Text) + CStr(strOperador)
   Else
      lblConta.Text = CStr(lblConta) + CStr(lblPrincipal.Text) + CStr(strOperador)
   End If

   If lblResultado.Text <> 0 Then
      Select Case strOperador
         Case "+"
            intSubtotal = Val(lblResultado.Text) + Val(lblPrincipal.Text)
            lblOperador.Text = "+"
         Case "-"
            intSubtotal = Val(lblResultado.Text) - Val(lblPrincipal.Text)
            lblOperador.Text = "-"
         Case "*"
            intSubtotal = Val(lblResultado.Text) * Val(lblPrincipal.Text)
            lblOperador.Text = "*"
         Case "/"
            intSubtotal = Val(lblResultado.Text) / Val(lblPrincipal.Text)
            lblOperador.Text = "/"
      End Select
   Else
      intSubtotal = Val(lblPrincipal.Text)
      lblOperador.Text = strOperador
   End If

   lblResultado.Text = intSubtotal
   lblPrincipal.Text = 0
   
End Sub

Private Sub calculaEquacao()
   
   Select Case lblOperador.Text
      Case "+"
         lblPrincipal.Text = Val(lblResultado.Text) + Val(lblPrincipal.Text)
      Case "-"
         lblPrincipal.Text = Val(lblResultado.Text) - Val(lblPrincipal.Text)
      Case "*"
         lblPrincipal.Text = Val(lblResultado.Text) * Val(lblPrincipal.Text)
      Case "/"
         lblPrincipal.Text = Val(lblResultado.Text) / Val(lblPrincipal.Text)
   End Select
   
   lblConta.Text = ""
   lblResultado.Text = 0
   lblOperador.Text = ""
   
End Sub

Private Sub btn0_Click()
   adicionaValor (0)
End Sub

Private Sub btn1_Click()
   adicionaValor (1)
End Sub

Private Sub btn2_Click()
   adicionaValor (2)
End Sub

Private Sub btn3_Click()
   adicionaValor (3)
End Sub

Private Sub btn4_Click()
   adicionaValor (4)
End Sub

Private Sub btn5_Click()
   adicionaValor (5)
End Sub

Private Sub btn6_Click()
   adicionaValor (6)
End Sub

Private Sub btn7_Click()
   adicionaValor (7)
End Sub

Private Sub btn8_Click()
   adicionaValor (8)
End Sub

Private Sub btn9_Click()
   adicionaValor (9)
End Sub

Private Sub btnClear_Click()
   lblResultado.Text = 0
   lblPrincipal.Text = 0
   lblConta.Text = ""
   lblOperador.Text = ""
End Sub

Private Sub btnClearPrincipal_Click()
   lblPrincipal.Text = 0
End Sub

Private Sub btnClearUltimo_Click()
   
   If lblPrincipal.Text = "" Then
      lblPrincipal.Text = 0
   Else
      If lblPrincipal.Text <> 0 Then
         lblPrincipal.Text = Left(lblPrincipal.Text, Len(lblPrincipal.Text) - 1)
         
         If lblPrincipal.Text = "" Then
            lblPrincipal.Text = 0
         End If
         
      End If
   End If
   
End Sub

Private Sub btnDivisao_Click()
   adicionaOperador ("/")
End Sub

Private Sub btnIgual_Click()
   calculaEquacao
End Sub

Private Sub btnMultiplicacao_Click()
   adicionaOperador ("*")
End Sub

Private Sub btnNegativar_Click()

      lblPrincipal.Text = Val(lblPrincipal.Text) * (-1)
   
End Sub

Private Sub btnSoma_Click()
   adicionaOperador ("+")
End Sub

Private Sub btnSubtracao_Click()
   adicionaOperador ("-")
End Sub
