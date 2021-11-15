VERSION 5.00
Begin VB.Form FrmBots 
   Caption         =   "Form2"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5595
   LinkTopic       =   "Form2"
   ScaleHeight     =   6480
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox obpkbots 
      Caption         =   "Es PK"
      Height          =   255
      Left            =   3360
      TabIndex        =   16
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   15
      Text            =   "Seleciones el Bots que quiere borrar"
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eliminar Ultimo Bots Agregado"
      Height          =   495
      Left            =   3840
      TabIndex        =   14
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtclanbots 
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Top             =   2370
      Width           =   1935
   End
   Begin VB.CommandButton cmdCrearbots 
      Caption         =   "Crear"
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox TxtNombrebots 
      Height          =   285
      Left            =   3360
      TabIndex        =   10
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtybots 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Txtxbots 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Textmapbots 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox clasebots 
      Height          =   315
      ItemData        =   "FrmBots.frx":0000
      Left            =   3360
      List            =   "FrmBots.frx":000D
      TabIndex        =   0
      Text            =   "Seleccionar"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblclanbots 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el Clan"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Line Line 
      X1              =   0
      X2              =   5520
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblNombrebots 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el Nombre"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblybots 
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label LblXbots 
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblCordenadabots 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese las Cordenadas"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lbMapaBots 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el mapa"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Lbclasebots 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el Tipo de bots"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "FrmBots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim clase As Long
Dim CantidadBots As Long
Dim NombreBost(15) As String
Private Sub clasebots_Click()
    Select Case clasebots.Text
    Case "Mago"
        clase = 1
    Case "Clero"
        clase = 2
    Case "cazador"
    clase = 3
    End Select
End Sub

Private Sub clasebots_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdCrearbots_Click()
   Dim Clan As String
    
   If clase = 0 Then Exit Sub
   If TxtNombrebots.Text = "" Then Exit Sub
  If Textmapbots.Text = "" Then Exit Sub
  If Txtxbots.Text = "" Then Exit Sub
    If txtybots.Text = "" Then Exit Sub
    Possbots.map = val(Textmapbots.Text)
    Possbots.X = val(Txtxbots.Text)
    Possbots.Y = val(txtybots.Text)
  If txtclanbots.Text = "" Then
  Clan = ""
  Else
  Clan = "<" & txtclanbots.Text & ">"
  End If
  


    Call ia_Spawn(clase, Possbots, TxtNombrebots.Text & Clan, False, obpkbots.Value, 0)
    CantidadBots = CantidadBots + 1
    'ReDim NombreBost(CantidadBots) As String
    NombreBost(CantidadBots) = TxtNombrebots.Text
    Combo1.AddItem TxtNombrebots.Text
    
End Sub

Private Sub Combo1_Click()



If Combo1.ListIndex <> -1 Then
For X = 1 To 15
If NombreBost(X) = Combo1.Text Then EsteES = X
Next X

Call ia_EraseChar(EsteES, False)
Combo1.RemoveItem Combo1.ListIndex
CantidadBots = CantidadBots - 1
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()

If CantidadBots = 0 Then Exit Sub
Call ia_EraseChar(CantidadBots, False)
CantidadBots = CantidadBots - 1
End Sub


