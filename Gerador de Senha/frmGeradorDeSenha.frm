VERSION 5.00
Begin VB.Form frmGeradorDeSenha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador de Senhas"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox checkLetrasMaiusculas 
      Caption         =   "Letras Maiúsculas"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   3975
   End
   Begin VB.CheckBox checkNumeros 
      Caption         =   "Incluir Números"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Frame frameGerador 
      Caption         =   "Gerador de Senhas"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtQuantidadeCaracteres 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "8"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton btnGerar 
         Caption         =   "Gerar Senha"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   4095
      End
      Begin VB.CheckBox checkSimbolos 
         Caption         =   "Incluir Simbolos"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox txtSenha 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Quantidade de caracteres"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label lblSenha 
         Caption         =   "Senha Gerada"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmGeradorDeSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGerar_Click()
    
    Dim intQuantidadeCaracteres As Integer
    Dim intFor As Integer
    Dim strSenha As String
    Dim strAlfabeto As String
    Dim arrAlfabeto() As String
    Dim intTamanhoAlfabeto As Integer
    
    intQuantidadeCaracteres = Val(txtQuantidadeCaracteres.Text)
    strAlfabeto = "a;b;c;d;e;f;g;h;i;j;k;l;m;n;o;p;q;r;s;t;u;v;w;x;y;z;"
    
    If checkLetrasMaiusculas.Value = 1 Then
        strAlfabeto = strAlfabeto + "A;B;C;D;E;F;G;H;I;J;K;L;M;N;O;P;Q;R;S;T;U;V;W;X;Y;Z;"
    End If
    
    If checkNumeros.Value = 1 Then
        strAlfabeto = strAlfabeto + "1;2;3;4;5;6;7;8;9;0;"
    End If
    
    If checkSimbolos.Value = 1 Then
        strAlfabeto = strAlfabeto + ";@;#;$;%;*;!;.;,;];[;|;{;};(;);"
    End If
    
    arrAlfabeto = Split(strAlfabeto, ";")
    
    For intFor = 1 To intQuantidadeCaracteres
        strSenha = strSenha + arrAlfabeto(Val(Rnd(50) * UBound(arrAlfabeto)))
    Next intFor

    txtSenha.Text = strSenha
    
End Sub
