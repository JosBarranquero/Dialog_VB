VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Marco digital"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOpen 
      Caption         =   "Cargar imagen"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog diaOpen 
      Left            =   840
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccionar imagen"
      Filter          =   "Imágenes JPG|*.jpg|Imágenes GIF|*.gif|Imágenes BMP|*.bmp"
      InitDir         =   "App.Path"
   End
   Begin VB.Image imgRect 
      BorderStyle     =   1  'Fixed Single
      Height          =   4335
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Author:  José Antonio Barranquero Fernández
'Version: 0.1
'Date:    03/04/2017
'Remark:  Simple photo frame, needs COMDLG32.OCX installed in C:\Windows\System (Win9x), C:\Windows\System32 (WinNT x86) or C:\Windows\SysWow64 (WinNT x64) to run

Dim imageUri As String  'String which stores the image location

Const MIN_FORM_HEIGHT = 1290
Const IMAGE_HORI_POSITION = 345
Const IMAGE_VERT_POSITION = 1185
Const BUTTON_VERT_POSITION = 960

Private Sub btnOpen_Click()
On Error GoTo errorHandler
    imageUri = ""           'Variable deleted so it's empty every time before choosing an image
    diaOpen.FileName = ""   'Preventing the FileName cache

    diaOpen.ShowOpen
    imageUri = diaOpen.FileName
    
    If imageUri <> "" Then
        MousePointer = vbHourglass              'Set mouse pointer to a hourglass
        imgRect.Picture = LoadPicture(imageUri) 'Load image selected into imgRect
    End If
    
errorHandler:
    'Error 53: Not found
    If (Err.Number = 53) Then MsgBox "El archivo especificado no se ha encontrado", vbCritical, "Error de archivo"
    'Error 481: Invalid format
    If (Err.Number = 481) Then MsgBox "El formato del archivo no es compatible", vbCritical, "Error de formato"
    
    MousePointer = vbDefault   'Set mouse pointer back to default, even if an error occurs
End Sub

Private Sub Form_Resize()
On Error GoTo errorHandler
    imgRect.Width = Me.Width - IMAGE_HORI_POSITION          'Set imgRect to properly resize
    imgRect.Height = Me.Height - IMAGE_VERT_POSITION
    
    btnOpen.Top = Me.Height - BUTTON_VERT_POSITION           'Reposition the button to properly resize
    btnOpen.Left = (Me.Width / 2) - (btnOpen.Width / 2)
    
errorHandler:
    'Error 380: Invalid property value, height is negative
    If (Err.Number = 380) Then Me.Height = MIN_FORM_HEIGHT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (MsgBox("¿Quieres salir?", vbYesNo + vbQuestion, "Salir") = vbNo) Then Cancel = 1
End Sub
