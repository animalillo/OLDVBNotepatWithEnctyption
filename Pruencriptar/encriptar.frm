VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8115
   ClientLeft      =   1545
   ClientTop       =   1575
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   6585
   Begin VB.OptionButton dobl 
      Caption         =   "doble"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   6240
      Width           =   1575
   End
   Begin VB.OptionButton segur 
      Caption         =   "segura"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   5760
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton decript 
      Caption         =   "desencriptar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton encript 
      Caption         =   "encriptar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox encriptado 
      Height          =   1935
      Left            =   1200
      TabIndex        =   1
      Top             =   3360
      Width           =   4095
   End
   Begin VB.TextBox normal 
      Height          =   2175
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   4095
   End
   Begin VB.Frame opciones 
      Caption         =   "metodo"
      Height          =   1215
      Left            =   1680
      TabIndex        =   6
      Top             =   5520
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clave As String
Dim caracter As String
Dim p As Integer

Private Sub decript_Click()
On Error Resume Next
If segur.Value = True Then
    clave = 50
ElseIf dobl.Value = True Then
    clave = 100
End If

'''limpiar campo Normalfield
normal.Text = ""
'''''''''''''Avance = 100 / Len(.ActiveForm.EncryptedField.Text)
'''''''''''''Avanzado = Avance
    For p = 1 To Len(encriptado.Text)
        '''''Progress.Value = Avance
        '''''Avance = Avance + Avanzado
        '''''LblStatus.Caption = "Status :" & Round(Avance, 0) & "% recuperado..."
        caracter = Chr((Asc(Mid(encriptado.Text, p, 1)) + (256 - clave)) Mod 256)
        normal.Text = normal.Text & caracter
        DoEvents
    Next
'''reinicio de variables
                'Avance = 0
                'Avanzado = 0
                'Progress.Value = 0
                caracter = 0
'''ocultar controles
'LblStatus.Visible = False
'Progress.Visible = False
End Sub

Private Sub encript_Click()
On Error Resume Next
If segur.Value = True Then
    clave = 50
ElseIf dobl.Value = True Then
    clave = 100
End If

'''limpiar campo encryptedfield
encriptado.Text = ""
'Avance = .Progress.Max / Len(.ActiveForm.NormalField.Text)
'Avanzado = Avance
    For p = 1 To Len(normal.Text)
        'Avance = Avance + Avanzado
        'LblStatus.Caption = "Status :" & Round(Avance, 0) & "% encriptado..."
        'Progress.Value = Avance
        caracter = Chr((Asc(Mid(normal.Text, p, 1)) + clave) Mod 256)
        encriptado.Text = encriptado.Text & caracter
    DoEvents
    Next
'''reinicio de variables
                'Avance = 0
                'Avanzado = 0
                'Progress.Value = 0
                caracter = 0
'''ocultar controles
'LblStatus.Visible = False
'Progress.Visible = False
End Sub
