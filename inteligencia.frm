VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form inteligencia 
   BackColor       =   &H8000000C&
   Caption         =   "Encriptador"
   ClientHeight    =   11355
   ClientLeft      =   795
   ClientTop       =   3030
   ClientWidth     =   13965
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "inteligencia.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   11355
   ScaleWidth      =   13965
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar progreso 
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   9960
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox encriptado 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   3480
      Width           =   15015
   End
   Begin VB.TextBox normal 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   15015
   End
   Begin VB.Label guaNormal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Guardar archivo normal"
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      ToolTipText     =   "Guarda el archivo de texto sin encriptar."
      Top             =   9120
      Width           =   3975
   End
   Begin VB.Label guaEncriptado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Guardar archivo encriptado"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      ToolTipText     =   "Guarda el archivo encriptado operativo"
      Top             =   8520
      Width           =   3975
   End
   Begin VB.Shape Shape 
      Height          =   1335
      Index           =   1
      Left            =   6480
      Top             =   8400
      Width           =   4215
   End
   Begin VB.Label abrEncriptado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Abrir archivo encriptado"
      Height          =   495
      Left            =   720
      MousePointer    =   3  'I-Beam
      TabIndex        =   5
      ToolTipText     =   "Abre un archivo encriptado."
      Top             =   9240
      Width           =   3255
   End
   Begin VB.Label abrNormal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Abrir archivo normal"
      Height          =   495
      Left            =   720
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      ToolTipText     =   "Abre un archivo no encriptado"
      Top             =   8760
      Width           =   3255
   End
   Begin VB.Shape Shape 
      Height          =   1215
      Index           =   0
      Left            =   600
      Top             =   8640
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   4200
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label encriptar 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Encriptar"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label desencriptar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Desencriptar"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Menu archivo 
      Caption         =   "&Archivo"
      Begin VB.Menu nuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu abrir 
         Caption         =   "&Abrir"
         Shortcut        =   ^A
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu guardar 
         Caption         =   "&Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu guardarcomo 
         Caption         =   "Guardar &como"
      End
      Begin VB.Menu c 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu acerca 
      Caption         =   "&Acerca de ..."
   End
End
Attribute VB_Name = "inteligencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub acerca_Click()
Acercade.Show
consola.Visible = False
End Sub

Private Sub desencriptar_Click()
Call decript
End Sub

Private Sub desencriptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
desencriptar.ForeColor = &HFF00&
End Sub

Private Sub encriptar_Click()
Call encript
End Sub

Private Sub encriptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
encriptar.ForeColor = &HFF00&
End Sub

Private Sub Form_Load()
If texto <> "" Then
    normal.Text = texto
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
encriptar.ForeColor = &H80000012
desencriptar.ForeColor = &H80000012
End Sub

