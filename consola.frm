VERSION 5.00
Begin VB.MDIForm consola 
   BackColor       =   &H8000000C&
   Caption         =   "Notedit"
   ClientHeight    =   7515
   ClientLeft      =   2235
   ClientTop       =   2295
   ClientWidth     =   9255
   Icon            =   "consola.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu arch 
      Caption         =   "&Archivo"
      Begin VB.Menu nuevo 
         Caption         =   "&Nuevo"
         Begin VB.Menu notas 
            Caption         =   "Documento de &notas"
         End
         Begin VB.Menu encript 
            Caption         =   "Notas &encriptadas"
            Shortcut        =   ^E
         End
      End
      Begin VB.Menu salir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "consola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub encript_Click()
inteligencia.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim segur As String
segur = MsgBox("¿Está se guro de que desea salir?", vbYesNo + vbQuestion, "Saliendo...")
If segur = vbNo Then
   Cancel = -1
Else
End If
End Sub

Private Sub notas_Click()
notedit.Show
End Sub
