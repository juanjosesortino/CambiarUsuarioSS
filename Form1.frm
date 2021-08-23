VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Usuario SourceSafe"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   3360
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   6
      ToolTipText     =   "C:\VSS Carpetas de Trabajo - Ant"
      Top             =   1230
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   2160
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   5
      ToolTipText     =   "C:\VSS Carpetas de Trabajo Comp"
      Top             =   1230
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   960
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   4
      ToolTipText     =   "C:\VSS Carpetas de Trabajo"
      Top             =   1230
      Width           =   1215
   End
   Begin VB.PictureBox PicImagen 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   5310
      Picture         =   "Form1.frx":1CCA
      ScaleHeight     =   1500
      ScaleWidth      =   6900
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   6900
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuario SourceSafe"
      Height          =   1065
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox Combo 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2865
      End
      Begin VB.Image Image 
         Height          =   720
         Index           =   0
         Left            =   60
         Picture         =   "Form1.frx":57E4
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   2355
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8061
            MinWidth        =   6526
         EndProperty
      EndProperty
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "VB6 Recent"
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   1050
      Width           =   945
   End
   Begin VB.Line Line 
      X1              =   30
      X2              =   4830
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Image Image 
      Height          =   780
      Index           =   5
      Left            =   30
      Picture         =   "Form1.frx":730A
      Stretch         =   -1  'True
      Top             =   1380
      Width           =   960
   End
   Begin VB.Image Image 
      Height          =   780
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":A34C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private bForm_Load As Boolean
Private strUsuario As String
Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long

Dim MouseForm As Boolean
Dim MousePicture As Boolean
Dim bPicturedown As Boolean
Dim strActualRecent As String

Private Sub Combo_Click()

   Dim Env_Var_Name As String
   Dim Env_Value As String
   Dim ix As Integer
        
   If bForm_Load Then Exit Sub
   
   StatusBar.Panels(1).Text = "Modificando Variables de Entorno..."
   Combo.Enabled = False
   
   Env_Var_Name = "SSUSER"
   Env_Value = Combo.Text
   DoEvents
   SetEnvironmentVar Env_Var_Name, Env_Value
   DoEvents
   
   StatusBar.Panels(1).Text = "Modificando Registro..."
   For ix = 1 To 50
      DeleteRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles", ix
   Next
   
   If strUsuario = Combo.Text Then
      SetRegistry "C:\VSS Carpetas de Trabajo\"
   Else
      SetRegistry "C:\VSS Carpetas de Trabajo Comp\"
   End If
   
   Combo.Enabled = True
   StatusBar.Panels(1).Text = ""
   
   ActualRecent
End Sub

Private Sub Form_Load()
   Dim Nombre As String, ret As Long, strClave As String
   
   bForm_Load = True
   bPicturedown = False
   
   Nombre = Space$(250)
   ret = Len(Nombre)
   If GetUserName(Nombre, ret) = 0 Then
      strUsuario = vbNullString
   Else
      strUsuario = Left$(Nombre, ret - 1)
   End If
   
   Combo.AddItem strUsuario
   Combo.AddItem strUsuario & ".comp"

   strClave = GetRegistryValue(HKEY_CURRENT_USER, "Environment", "SSUSER")
   If Len(strClave) = 0 Then
      Combo.ListIndex = 0
   Else
      If strUsuario = GetRegistryValue(HKEY_CURRENT_USER, "Environment", "SSUSER") Then
         Combo.ListIndex = 0
      Else
         Combo.ListIndex = 1
      End If
   End If
   
   bForm_Load = False
   
   ActualRecent
End Sub
Private Sub ActualRecent()

   strActualRecent = GetRegistryValue(HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "1")
   If InStr(strActualRecent, "C:\VSS Carpetas de Trabajo\") Then
      Activate 1
      Picture1.Visible = True
      Picture2.Visible = False
      Picture3.Visible = True
   End If
   If InStr(strActualRecent, "C:\VSS Carpetas de Trabajo Comp\") Then
      Activate 2
      Picture1.Visible = False
      Picture2.Visible = True
      Picture3.Visible = False
   End If
   If InStr(strActualRecent, "C:\VSS Carpetas de Trabajo - Ant\") Then
      Activate 3
      Picture1.Visible = True
      Picture2.Visible = False
      Picture3.Visible = True
   End If

End Sub
Private Sub SetRegistry(ByVal strCarpeta As String)

   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "1", strCarpeta & "Inicio\Inicio.vbp"
   
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "2", strCarpeta & "Cereales\Cereales.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "3", strCarpeta & "GestionComercial\GestionComercial.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "4", strCarpeta & "Contabilidad\Contabilidad.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "5", strCarpeta & "Fiscal\Fiscal.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "6", strCarpeta & "AdministradorGeneral\AdministradorGeneral.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "7", strCarpeta & "Produccion\Produccion.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "8", strCarpeta & "Seguridad\Seguridad.vbp"
   
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "9", strCarpeta & "ReportsCereales\ReportsCereales.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "10", strCarpeta & "ReportsCereales2\ReportsCereales2.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "11", strCarpeta & "ReportsGescom\ReportsGescom.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "12", strCarpeta & "ReportsGescom2\ReportsGescom2.vbp"
   
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "13", strCarpeta & "COM DLLs\BOCereales\BOCereales.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "14", strCarpeta & "COM DLLs\BOContabilidad\BOContabilidad.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "15", strCarpeta & "COM DLLs\BOGesCom\BOGesCom.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "16", strCarpeta & "COM DLLs\BOFiscal\BOFiscal.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "17", strCarpeta & "COM DLLs\BOGeneral\BOGeneral.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "18", strCarpeta & "COM DLLs\BOSeguridad\BOSeguridad.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "19", strCarpeta & "COM DLLs\BOProduccion\BOProduccion.vbp"
   
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "20", strCarpeta & "COM DLLs\DSCereales\DSCereales.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "21", strCarpeta & "COM DLLs\DSContabilidad\DSContabilidad.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "22", strCarpeta & "COM DLLs\DSGesCom\DSGesCom.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "23", strCarpeta & "COM DLLs\DSFiscal\DSFiscal.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "24", strCarpeta & "COM DLLs\DSGeneral\DSGeneral.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "25", strCarpeta & "COM DLLs\DSSeguridad\DSSeguridad.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "26", strCarpeta & "COM DLLs\DSProduccion\DSProduccion.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "27", strCarpeta & "COM DLLs\SPCereales\SPCereales.vbp"
   
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "28", strCarpeta & "COM DLLs\DataShare\DataShare.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "29", strCarpeta & "COM DLLs\DataAccess\DataAccess.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "30", strCarpeta & "AlgStart\AlgStart.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "31", strCarpeta & "COM DLLs\AlgStdFunc\AlgStdFunc.vbp"
   
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "32", strCarpeta & "Mobile\AlgMobile\AlgMobile.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "33", strCarpeta & "Mobile\AlgInterop\AlgInterop.vbp"
   SetRegistryValue HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles\", "34", strCarpeta & "Mobile\AlgInterop\AlgInteropTest\AlgInteropTest.vbp"
End Sub

Private Sub Picture1_Click()
   SetRegistry "C:\VSS Carpetas de Trabajo\"
   StatusBar.Panels(1).Text = "VB6 Recent Modificado con VSS Carpetas de Trabajo"
End Sub
Private Sub Picture2_Click()
   SetRegistry "C:\VSS Carpetas de Trabajo Comp\"
   StatusBar.Panels(1).Text = "VB6 Recent Modificado con VSS Carpetas de Trabajo Comp"
End Sub
Private Sub Picture3_Click()
   SetRegistry "C:\VSS Carpetas de Trabajo - Ant\"
   StatusBar.Panels(1).Text = "VB6 Recent Modificado con VSS Carpetas de Trabajo - Ant"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Activate 1
End Sub
Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Activate 2
End Sub
Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Activate 3
End Sub
Private Sub Activate(Button As Integer)
   Select Case Button
      Case 1
         BitBlt Picture1.hDC, 0, 0, 100, 100, PicImagen.hDC, 370, 15, vbSrcCopy
         BitBlt Picture2.hDC, 0, 0, 100, 100, PicImagen.hDC, 193, 15, vbSrcCopy
         BitBlt Picture3.hDC, 0, 0, 100, 100, PicImagen.hDC, 193, 15, vbSrcCopy
         StatusBar.Panels(1).Text = "VB6 Recent VSS Carpetas de Trabajo"
      Case 2
         BitBlt Picture1.hDC, 0, 0, 100, 100, PicImagen.hDC, 193, 15, vbSrcCopy
         BitBlt Picture2.hDC, 0, 0, 100, 100, PicImagen.hDC, 370, 15, vbSrcCopy
         BitBlt Picture3.hDC, 0, 0, 100, 100, PicImagen.hDC, 193, 15, vbSrcCopy
         StatusBar.Panels(1).Text = "VB6 Recent VSS Carpetas de Trabajo Comp"
      Case 3
         BitBlt Picture1.hDC, 0, 0, 100, 100, PicImagen.hDC, 193, 15, vbSrcCopy
         BitBlt Picture2.hDC, 0, 0, 100, 100, PicImagen.hDC, 193, 15, vbSrcCopy
         BitBlt Picture3.hDC, 0, 0, 100, 100, PicImagen.hDC, 370, 15, vbSrcCopy
         StatusBar.Panels(1).Text = "VB6 Recent VSS Carpetas de Trabajo - Ant"
   End Select

   If bPicturedown Then
      bPicturedown = False
   Else
      bPicturedown = True
   End If
   refrescar
End Sub
Private Sub refrescar()
    Picture1.Refresh
    Picture2.Refresh
    Picture3.Refresh
End Sub

