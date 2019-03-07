VERSION 5.00
Begin VB.Form FormAviso 
   Caption         =   "Aguarde.."
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3885
   LinkTopic       =   "Form1 "
   ScaleHeight     =   1725
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Reiniciando impressora, por favor, aguarde..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "FormAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
