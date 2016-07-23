Attribute VB_Name = "acciones"
Option Explicit

Dim clave, caracter, avance, avanzado As String
Public texto As String
Dim p As Integer
Dim a, b As String

Sub decript()
On Error Resume Next
a = Val(InputBox("Inserte la clave"))
If a >= 250 Or a = 0 Then
    
Else
clave = a
With inteligencia
.progreso.Visible = True
'limpiar texto Normal
.normal.Text = ""
avance = 100 / Len(.encriptado.Text)
avanzado = avance
    For p = 1 To Len(.encriptado.Text)
        .progreso.Value = avance
        avance = avance + avanzado
        caracter = Chr((Asc(Mid(.encriptado.Text, p, 1)) + (50970 - clave)) Mod 50970)
        .normal.Text = .normal.Text & caracter
        DoEvents
    Next
'reinicio de variables
                avance = 0
                avanzado = 0
                .progreso.Value = 0
                caracter = 0
                p = 0
                clave = ""
'ocultar controles
.progreso.Visible = False
End With
End If
End Sub

Sub encript()
On Error Resume Next
a = Val(InputBox("inserte una clave alfanumerica inferior a 50970", "Insertar clave..."))
If a >= 250 Or a = 0 Or a = "" Then
 MsgBox "Clave incorrecta", vbCritical + vbOKOnly, "¡ERROR!"
Else
    b = Val(InputBox("Confirme la clave", "Confirmar clave..."))
        If a = b Then
            clave = a
        Else
            MsgBox "Las claves introducidas no coinciden", vbExclamation + vbOKOnly, "¡ERROR!"
        End If
With inteligencia
.progreso.Visible = True
'limpiar campo encryptado
.encriptado.Text = ""
'inteligencia.encriptado.Text = ""
avance = .progreso.Max / Len(.normal.Text)
avanzado = avance
    For p = 1 To Len(.normal.Text)
        avance = avance + avanzado
        .progreso.Value = avance
        caracter = Chr((Asc(Mid(.normal.Text, p, 1)) + clave) Mod 50970)
        .encriptado.Text = .encriptado.Text & caracter
    DoEvents
    Next
'''reinicio de variables
                avance = 0
                avanzado = 0
                .progreso.Value = 0
                caracter = 0
                p = 0
                clave = ""
'''ocultar controles
.progreso.Visible = False
End With
End If
End Sub
