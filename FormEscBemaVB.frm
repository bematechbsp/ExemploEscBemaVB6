VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FormCmdDiretoVB 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6240
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton CommandMostraCOM 
      Caption         =   "Mostra COM"
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   240
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FormEscBemaVB.frx":0000
      Left            =   4800
      List            =   "FormEscBemaVB.frx":0002
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton CommandQrCode 
      Caption         =   "QRCode"
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Text            =   "33181236206274000189650010000010259000010258"
      Top             =   3600
      Width           =   4455
   End
   Begin VB.CommandButton CommandCodBarras 
      Caption         =   "Código de Barras"
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Text            =   "1234567890"
      Top             =   3120
      Width           =   4455
   End
   Begin VB.CommandButton CommandImprime 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2520
      Width           =   2775
   End
   Begin VB.OptionButton OptionExpandido 
      Caption         =   "Expandido"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.OptionButton OptionNegrito 
      Caption         =   "Negrito"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.OptionButton OptionItalico 
      Caption         =   "Itálico"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.OptionButton OptionCondensado 
      Caption         =   "Condensado"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox TextImpressao 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "Texto a ser impresso"
      Top             =   1320
      Width           =   6255
   End
   Begin VB.TextBox TextPorta 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "COM8"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Texto a ser impresso"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Porta Impressora:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FormCmdDiretoVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmd As String

Private Sub CommandImprime_Click()
porta = Replace(Combo1.Text, "COM", "")
MSComm1.CommPort = porta


MSComm1.PortOpen = True

If OptionCondensado.LBound Then
    MSComm1.Output = Chr(27) + Chr(40) + TextImpressao.Text
    MSComm1.PortOpen = False
    
End If

End Sub


Private Sub Form_Load()
Dim i As Integer
Combo1.Clear
For i = 1 To 16
   If DetectaPortaCOM(i) <> 0 Then
       Combo1.AddItem "COM" & i
   End If
Next
Combo1.ListIndex = 0
End Sub
