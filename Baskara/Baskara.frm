VERSION 5.00
Begin VB.Form Calcular 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calcular Baskara"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox frmX2 
      Height          =   405
      Left            =   1440
      TabIndex        =   12
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox frmX1 
      Height          =   405
      Left            =   1440
      TabIndex        =   10
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox frmDelta 
      Height          =   405
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton frmCalcular 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox frmC 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox frmB 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox frmA 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "X2"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "X1"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Delta"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "C"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "B"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "A"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Calcular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub frmCalcular_Click()

   Dim intA As Single
   Dim intB As Single
   Dim intC As Single
   Dim intDelta As Single
   Dim intX1 As Single
   Dim intX2 As Single
   
   intA = Val(frmA.Text)
   intB = Val(frmB.Text)
   intC = Val(frmC.Text)
   
   If intA = 0 Then
      MsgBox ("O Coeficiente de A precisa ser maior que 0"), vbCritical
      Exit Sub
   End If
   
   intDelta = (intB ^ 2 - 4 * intA * intC)
   
   If intDelta < 0 Then
     frmDelta.Text = intDelta
     frmX1.Text = 0
     frmX2.Text = 0
     
     MsgBox ("Esta equação não possui raizes reais!"), vbCritical
     Exit Sub
   End If
   
   If intDelta = 0 Then
      MsgBox ("Esta equação possui duas raizes iguais!"), vbInformation
   End If
   
   intX1 = (-intB + Sqr(intDelta) / (2 * intA))
   intX2 = (-intB - Sqr(intDelta) / (2 * intA))
   
   frmDelta.Text = intDelta
   frmX1.Text = intX1
   frmX2.Text = intX2

End Sub
