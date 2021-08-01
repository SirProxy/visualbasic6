VERSION 5.00
Begin VB.Form ecoa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ecoa Texto"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox lblEco 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Frame frmEcoa 
      Caption         =   "Formulário Ecoa"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton btnLimpar 
         Caption         =   "Limpar"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   4215
      End
      Begin VB.TextBox lblCopy 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "Texto digitado"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Digite o texto"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Digite o texto"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
   End
End
Attribute VB_Name = "ecoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLimpar_Click()
   
   lblEco.Text = ""
   lblCopy.Text = ""

End Sub

Private Sub lblCopy_Change()

   lblEco.Text = lblCopy.Text

End Sub
