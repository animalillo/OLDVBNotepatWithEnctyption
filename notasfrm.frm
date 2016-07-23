VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form notedit 
   Caption         =   "Editor de notas"
   ClientHeight    =   5055
   ClientLeft      =   4050
   ClientTop       =   4260
   ClientWidth     =   5805
   Icon            =   "notasfrm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   5805
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog d2 
      Left            =   3720
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "dll"
      FileName        =   "app.path & archabr.dll"
      Filter          =   "dllfiles (*.dll)"
      InitDir         =   "app.path"
   End
   Begin MSComDlg.CommonDialog d1 
      Left            =   3720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "not"
      DialogTitle     =   """Archivos de notas"""
      FileName        =   "*.not; *.txt"
      Filter          =   "Notas (*.not); Texto (*.txt)"
      FontItalic      =   -1  'True
      InitDir         =   "C:\Documents and Settings\Usuario\Escritorio"
      MaxFileSize     =   16959
   End
   Begin VB.Label lbl 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2400
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
         Caption         =   "Encriptar archivo"
      End
   End
   Begin VB.Menu acerca 
      Caption         =   "&Acerca de ..."
   End
End
Attribute VB_Name = "notedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sal As Integer
Dim modi As Integer
Dim error As Integer
Dim key As String
Dim initdir As String
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub acerca_Click()
Acercade.Show
consola.Visible = False
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
clave = "piraguas"
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
initdir = App.Path
d2.initdir = initdir
d2.FileName = initdir & "\archabr.dll"
Dim abr As String
Open d2.FileName For Input As #1
lbl.Caption = ""
Text1.Text = ""
Do While Not EOF(1)
Line Input #1, abr
lbl.Caption = lbl.Caption & abr & (Chr$(13) + Chr$(10))
Loop
Close #1
d1.initdir = lbl.Caption
modi = 0
error = 3
key = 0
End Sub

Private Sub Form_Resize()
On Error GoTo anomalia
    Text1.Width = Me.Width - 120
    Text1.Height = Me.Height - 750
anomalia:
    If Err.Number = 380 Then
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim gu As String
If sal = 1 Then
    MsgBox "Eta bloqueado, porfavor desbloquee el programa 1º", vbCritical, "Bloqueo"
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
End Sub

Private Sub mnuAbrir_Click()
On Error GoTo anomalia
Dim linea As String
Dim prog As String

prog = 0
d1.initdir = lbl.Caption
d1.FileName = "*.not; *.txt"
d1.ShowOpen
Open d1.FileName For Input As #1
Text1.Text = ""
Do While Not EOF(1)
Line Input #1, linea
    Text1.Text = Text1.Text & linea & (Chr$(13) + Chr$(10))
Loop
Close #1
modi = 0
notedit.Caption = d1.FileName
lbl.Caption = d1.FileName
Dim arch As Integer
d2.FileName = "C:\Archivos de programa\Notedit\archabr.dll"
arch = FreeFile(0)
Open d2.FileName For Output As arch
Print #1, lbl.Caption
Close arch
anomalia:
    If Err.Number = 75 Then
    MsgBox "No se a podido abrir el archivo solicitado", vbCritical, "Error"
        Else
    End If
End Sub

Private Sub mnuGuardar_Click()
Dim f As Integer
If modi = 0 Then

Else

    If d1.FileName = "*.not; *.txt" Then
        Call mnuGuardarcomo_Click

    Else
        
        f = FreeFile(0)
        If key <> "" Then
            Open d1.FileName For Output As f
                Print #1, Text1.Text & key
            Close (f)
        Else
            Open d1.FileName For Output As f
                Print #1, Text1.Text
            Close (f)
        End If
    
        Me.Caption = "Notas - " & d1.FileName
        modi = 0
    End If
End If
End Sub

Private Sub mnuInsKlav_Click()
texto = Text1.Text
inteligencia.Show
End Sub

Private Sub mnuSalir_Click()

Unload Me
End Sub

Private Sub mnuGuardarcomo_Click()
On Error GoTo anomalia
Dim f As Integer
If d1.FileName = "*.not; *.txt" Then
    d1.DialogTitle = "Guardar como..."
    d1.initdir = lbl.Caption
    d1.ShowSave
    f = FreeFile(0)
    If key <> "" Then
        Open d1.FileName For Output As f
            Print #1, Text1.Text & "33plo" & key & "33plo"
        Close (f)
    Else
        Open d1.FileName For Output As f
            Print #1, Text1.Text
        Close (f)
    End If
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
notedit.Caption = "Editor de notas"
modi = 0
End Sub

Private Sub Text1_Change()
modi = 1
End Sub
