VERSION 5.00
Begin VB.Form frmJuros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculadora de Juros"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnClear 
      Caption         =   "Limpar Campos"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   3855
   End
   Begin VB.TextBox txtResultado 
      Height          =   405
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox txtJuros 
      Height          =   405
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox txtParcela 
      Height          =   405
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Juros"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton btnCalcular 
         Caption         =   "Calcular Juros ao Mês"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   3855
      End
      Begin VB.TextBox txtValor 
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Valor"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Juros ao mês (Em %)"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Duração da Aplicação (Em meses)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Valor da Compra"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmJuros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCalcular_Click()

   Dim dblCompra As Double
   Dim intParcelas As Integer
   Dim dblJuros As Double
   Dim dblTotalJuros As Double
   Dim dblPorcentagemJuros As Double
   
   dblCompra = Val(txtValor.Text)
   intParcela = Val(txtParcela.Text)
   dblJuros = Val(txtJuros.Text)
   
   If dblCompra = 0 Then
      MsgBox ("Ooops! Você precisa informar um valor para calculo de Juros."), vbCritical
      Exit Sub
   End If
   
   If intParcela = 0 Then
      MsgBox ("Ooops! Você precisa informar a quantidade de parcelas."), vbCritical
      Exit Sub
   End If
   
   If dblJuros = 0 Then
      MsgBox ("Ooops! Você precisa informar a porcentagem de juros ao mês."), vbCritical
      Exit Sub
   End If
   
   dblTotalJuros = dblCompra * ((1 + (dblJuros / 100)) ^ intParcela)
   txtResultado.Text = "R$ " & Format(dblTotalJuros, "###0.00")
   
End Sub

Private Sub btnClear_Click()

   txtValor.Text = ""
   txtParcela.Text = ""
   txtJuros.Text = ""
   txtResultado.Text = ""
   
End Sub
