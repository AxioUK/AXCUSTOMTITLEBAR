VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   Caption         =   "Proyecto1 - Form1 (Form)"
   ClientHeight    =   5085
   ClientLeft      =   1470
   ClientTop       =   2250
   ClientWidth     =   7365
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin UCAxCustomTitleBar.axCustomTitleBar axCustomTitleBar 
      Left            =   4530
      Top             =   795
      _ExtentX        =   4551
      _ExtentY        =   926
      WhiteIcons      =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3390
      Top             =   2070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   855
      Left            =   4890
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "utilizando el espacio de la baarra de titulo"
      Height          =   495
      Left            =   4875
      TabIndex        =   7
      Top             =   2925
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Color de la barra de titulo sin foco"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fondo con imagen"
      Height          =   495
      Left            =   4875
      TabIndex        =   5
      Top             =   2370
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Color de la barra de titulo"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FF9999&
      Caption         =   "Mostrar barra de titulo (SI=movible / NO=modal)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   3675
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FF9999&
      Caption         =   "Mostrar Icono"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF9999&
      Caption         =   "Utilizar iconos y titulo Blanco."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Color del Formulario"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1215
      Left            =   720
      TabIndex        =   8
      Top             =   3600
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    axCustomTitleBar.WhiteIcons = Check1.Value
End Sub

Private Sub Check2_Click()
    axCustomTitleBar.ShowIcon = Check2.Value
End Sub

Private Sub Check3_Click()
    axCustomTitleBar.ShowTitlebar = Check3.Value
End Sub

Private Sub Command1_Click()
    Dim Ctl As Object
    CommonDialog1.ShowColor
    Me.BackColor = CommonDialog1.Color
    On Error Resume Next
    For Each Ctl In Me.Controls
        Ctl.BackColor = CommonDialog1.Color
    Next
End Sub

Private Sub Command2_Click()
CommonDialog1.ShowColor
axCustomTitleBar.TitleBarBackColor = CommonDialog1.Color
axCustomTitleBar.ShowTitlebar = True
End Sub


Private Sub Command4_Click()
    CommonDialog1.ShowColor
    axCustomTitleBar.TitleBarBackColorDesactivate = CommonDialog1.Color
    axCustomTitleBar.ShowTitlebar = True
End Sub

Private Sub Command5_Click()
    Dim Cadena As String
    
    Cadena = "cadena de texto" & vbCrLf
    
    If Right$(Cadena, 1) = vbLf Then
        MsgBox "Retorno de Carro detectado"
    Else
        MsgBox "Cadena Limpia"
    End If
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'        If Y < axCustomTitleBar.ControlBoxHeight * Screen.TwipsPerPixelY Then
'            axCustomTitleBar.DragForm
'        End If
'    End If
End Sub


