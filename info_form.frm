VERSION 5.00
Begin VB.Form info_form 
   Caption         =   "Information"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label_info 
      Height          =   5895
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   5175
   End
End
Attribute VB_Name = "Info_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Info_form.Hide
 Unload Info_form
End Sub

