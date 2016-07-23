VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{FC07EBD4-FE92-11D0-A199-A0077383D901}#5.5#0"; "ccrpprg.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor de notas"
   ClientHeight    =   6375
   ClientLeft      =   3915
   ClientTop       =   3915
   ClientWidth     =   5535
   Icon            =   "notasfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin CCRProgressBar.ccrpProgressBar progreso 
      Height          =   375
      Left            =   3120
      ToolTipText     =   "Abierto..."
      Top             =   5640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BackColor       =   8421504
      BorderStyle     =   1
      Caption         =   "Abriendo....."
      FillColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      Min             =   1
      Style           =   1
   End
   Begin MSComDlg.CommonDialog d2 
      Left            =   240
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "dll"
      FileName        =   "C:\Archivos de programa\Notedit\archabr.dll"
      Filter          =   "dllfiles (*.dll)"
      InitDir         =   "C:\Archivos de programa\Notedit\"
   End
   Begin MSComDlg.CommonDialog d1 
      Left            =   240
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "not"
      DialogTitle     =   """Archivos de notas"""
      FileName        =   "*.not; *.txt"
      Filter          =   "Notas (*.not); Texto (*.txt)"
      FontItalic      =   -1  'True
      InitDir         =   "C:\Documents and Settings\Usuario\Escritorio"
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000013&
      DragIcon        =   "notasfrm.frx":0442
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label lbl 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu nuev 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu z 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "&Abrir"
         Shortcut        =   ^A
      End
      Begin VB.Menu l 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "&Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGuardarcomo 
         Caption         =   "Guardar como"
      End
      Begin VB.Menu ll 
         Caption         =   "-"
      End
      Begin VB.Menu print 
         Caption         =   "Im&primir"
         Shortcut        =   ^P
      End
      Begin VB.Menu u2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu edic 
      Caption         =   "&Edición"
      Begin VB.Menu block 
         Caption         =   "&Bloquear"
         Shortcut        =   ^B
      End
      Begin VB.Menu desb 
         Caption         =   "&Desbloquear"
         Shortcut        =   ^D
      End
      Begin VB.Menu lll 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuInsKlav 
         Caption         =   "Insertar clave al archivo"
      End
   End
   Begin VB.Menu acerca 
      Caption         =   "&Acerca de ..."
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sal As Integer
Dim modi As Integer
Dim error As Integer
Dim key As String
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub acerca_Click()
Form2.Visible = True
Form1.Visible = False
End Sub

Private Sub block_Click()
Text1.Enabled = False
desb.Enabled = True
block.Enabled = False
nuev.Enabled = False
mnuAbrir.Enabled = False
mnuGuardar.Enabled = False
mnuSalir.Enabled = False
sal = 1
End Sub


Private Sub desb_Click()
Dim clave As String
clave = "aquaventur"
Dim proc As String
proc = InputBox("Intruduzca la clave porfavor", "Desbloqueando...")
If proc = clave Then
    Text1.Enabled = True
    block.Enabled = True
    desb.Enabled = False
    nuev.Enabled = True
    mnuAbrir.Enabled = True
    mnuGuardar.Enabled = True
    mnuSalir.Enabled = True
    sal = 0
Else
    If error > 0 Then
        MsgBox "La clave introducida es incorrecta, porfavor intente nuevamente. " & " Le quedan: " & error & " intentos", vbCritical, "ERROOORRRRR"
        error = error - 1
        Call desb_Click
    Else
        MsgBox "Ha superado el Nº máximo de intentos.", vbCritical
        error = 3
    End If
End If
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
desb.Enabled = False
lbl.Caption = ""
Dim abr As String
Open d2.FileName For Input As #1
Text1.Text = ""
Do While Not EOF(1)
Line Input #1, abr
lbl.Caption = lbl.Caption & abr & (Chr$(13) + Chr$(10))
Loop
Close #1
d1.InitDir = lbl.Caption
modi = 0
error = 3
key = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim gu As String
If sal = 1 Then
    MsgBox "Eta bloqueado, porfavor desbloquee el programa 1º", vbCritical, "Bloqueo"
    Cancel = -1
Else
    Dim segur As String
    segur = MsgBox("¿Está se guro de que desea salir?", vbYesNo + vbQuestion, "Saliendo...")
    If segur = vbNo Then
        Cancel = -1
    Else
        If Text1.Text = "" Or modi = 0 Then

        Else
            gu = MsgBox("¿Desea guardar los cambios efectuados?", vbQuestion + vbYesNoCancel)
    
            If gu = vbYes Then
                Call mnuGuardar_Click
            ElseIf gu = vbNo Then
        
            ElseIf gu = vbCancel Then
                Cancel = -1
            End If
        End If
    End If
End If
End Sub

Private Sub mnuAbrir_Click()
On Error GoTo anomalia

Dim linea As String
d1.InitDir = lbl.Caption
d1.FileName = "*.not; *.txt"
d1.ShowOpen
Open d1.FileName For Input As #1
Text1.Text = ""
Do While Not EOF(1)
Line Input #1, linea
Text1.Text = Text1.Text & linea & (Chr$(13) + Chr$(10))
Loop
Close #1
Form1.Caption = d1.FileName
lbl.Caption = d1.FileName
Dim arch As Integer
d2.FileName = "C:\Archivos de programa\Notedit\archabr.dll"
arch = FreeFile(0)
Open d2.FileName For Output As arch
Print #1, lbl.Caption
Close arch
modi = 0
anomalia:
    If Err.Number = 75 Then
    MsgBox "No se a podido abrir el archivo solicitado", vbCritical, "Error"
        Else
    End If
End Sub

Private Sub mnuGuardar_Click()
Dim f As Integer
If d1.FileName = "*.not; *.txt" Then
   Call mnuGuardarcomo_Click

    Else
        
        f = FreeFile(0)
        Open d1.FileName For Output As f
            Print #1, Text1.Text & "cl5540p3" & key
        Close (f)
    
        Me.Caption = "Notas - " & d1.FileName
        modi = 0
    End If
End Sub

Private Sub mnuInsKlav_Click()
Dim comp1 As String
Dim comp2 As String
comp1 = InputBox("Introduzca la calve que desea poner al archivo", "Clave.....")
comp2 = InputBox("Confirme la clave", "Clave.....")
If comp1 = comp2 Then
    key = com2
Else
    MsgBox "Las claves introducidas no coinciden", vbExclamation, "Error 335"
End If
End Sub

Private Sub mnuSalir_Click()
Unload Me
End Sub

Private Sub mnuGuardarcomo_Click()
On Error GoTo anomalia
Dim f As Integer
If d1.FileName = "*.not; *.txt" Then
    d1.DialogTitle = "Guardar como..."
    d1.InitDir = lbl.Caption
    d1.ShowSave
    f = FreeFile(0)
    Open d1.FileName For Output As f
        Print #1, Text1.Text & "cl5540p3" & key
    Close (f)
End If
Me.Caption = "Notas - " & d1.FileName
modi = 0
anomalia:
    If Err.Number = 52 Then
        MsgBox "no se ha podido guardar el archivo", vbCritical, "ERRORRRR"
    End If
d1.FileName = "*.not; *.txt"
lbl.Caption = d1.FileName
Dim arch As Integer
d2.FileName = "C:\Archivos de programa\Notedit\archabr.dll"
arch = FreeFile(0)
Open d2.FileName For Output As arch
Print #1, lbl.Caption
Close arch
End Sub

Private Sub nuev_Click()
Text1.Text = ""
d1.FileName = "*.not; *.txt"
Form1.Caption = "Editor de notas"
modi = 0
End Sub

Private Sub Text1_Change()
modi = 1
End Sub
