VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'Kein
   Caption         =   "cGFX Demo"
   ClientHeight    =   6000
   ClientLeft      =   1650
   ClientTop       =   1755
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   4515
   Begin VB.CommandButton cmdFadeIn 
      Caption         =   "Fade In"
      Height          =   330
      Left            =   2340
      TabIndex        =   5
      Top             =   2460
      Width           =   1890
   End
   Begin VB.PictureBox picBack 
      Height          =   6045
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   5985
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      Begin VB.CommandButton cmdFadeOut 
         Caption         =   "Fade Out"
         Height          =   330
         Left            =   300
         TabIndex        =   4
         Top             =   2430
         Width           =   1890
      End
      Begin VB.CommandButton cmdSlide 
         Caption         =   "Slide to Fade Rate"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2340
         TabIndex        =   3
         Top             =   3000
         Width           =   1860
      End
      Begin VB.CheckBox chkOnChange 
         BackColor       =   &H00C8C8C8&
         Caption         =   "On Change Fade"
         Height          =   285
         Left            =   330
         TabIndex        =   2
         Top             =   3030
         Value           =   1  'Aktiviert
         Width           =   1860
      End
      Begin MSComctlLib.Slider slFade 
         Height          =   315
         Left            =   285
         TabIndex        =   1
         Top             =   3435
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Min             =   1
         Max             =   255
         SelStart        =   255
         Value           =   255
      End
      Begin VB.Label lblClose 
         BackStyle       =   0  'Transparent
         Height          =   120
         Left            =   4260
         TabIndex        =   6
         Top             =   135
         Width           =   105
      End
      Begin VB.Shape Shape1 
         Height          =   945
         Left            =   225
         Top             =   2925
         Width           =   4050
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'cGFX Demo
'by tHa_imaX
'Binary Crew & Digital Death Crew

'FOR NEAR INFORMATION TAKE A LOOK TO THE CLASS.

'FIRST STEP: Declare an Instance of the Class
Private gX As New cGfX

Private Sub Form_Load()
'SECOND STEP: Cut your Window if needed.
    gX.PolyTrans picBack, Me
'THIRD STEP: If you're going to fade set your current fade rate to 255
    gX.SetCurrentRate 255

'OPTIONAL STEP: Let your Form Fadein :)
    gX.FadeIn Me.HwnD
End Sub


Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'FOURTH STEP: Make your form movable
    gX.MoveForm Me
End Sub

' NOW YOURE READY TO PLAY ;)


Private Sub slFade_Click()
    If chkOnChange.Value Then
        gX.DoTrans Me.HwnD, slFade.Value
    End If
End Sub

Private Sub chkOnChange_Click()
    If chkOnChange Then
        cmdSlide.Enabled = False
    Else
        cmdSlide.Enabled = True
    End If
End Sub

Private Sub cmdFadeIn_Click()
    gX.FadeIn Me.HwnD
End Sub

Private Sub cmdFadeOut_Click()
    gX.FadeOut Me.HwnD
    MsgBox "Fade out complete. Click to restore."
    gX.DoTrans Me.HwnD, 255
End Sub

Private Sub cmdSlide_Click()
    gX.FadeTo Me.HwnD, slFade.Value
End Sub

Private Sub lblClose_Click()
    gX.FadeOut Me.HwnD
    Unload Me
    End
End Sub

