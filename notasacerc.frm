VERSION 5.00
Begin VB.Form Acercade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerdca de..."
   ClientHeight    =   495
   ClientLeft      =   5280
   ClientTop       =   3030
   ClientWidth     =   4335
   Icon            =   "notasacerc.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VB.Label lbl 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para cualquier duda o sugerencia, contacten con migo XDD                     gabumoneselmejor@gmail.com"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "Acercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
consola.Visible = True
End Sub

Private Sub lbl_Click()

End Sub
