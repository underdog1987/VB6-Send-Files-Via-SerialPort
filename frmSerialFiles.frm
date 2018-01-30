VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSerialFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envío y recepción de archivos vía puerto serie"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   Icon            =   "frmSerialFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Recibir"
      Height          =   3855
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   3375
      Begin VB.CommandButton cmdSaveAs 
         Caption         =   "Guardar buffer"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   1560
         Top             =   3360
      End
      Begin VB.CommandButton cmdRecibe 
         Caption         =   "Recibir"
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txBytesRecibidos 
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Puertos serie"
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   3375
      Begin VB.ComboBox cmbSerial 
         Height          =   315
         ItemData        =   "frmSerialFiles.frx":0442
         Left            =   120
         List            =   "frmSerialFiles.frx":0444
         TabIndex        =   11
         Top             =   360
         Width           =   2535
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   2760
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InBufferSize    =   1
         OutBufferSize   =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enviar"
      Height          =   5415
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      Begin MSComctlLib.ProgressBar progresoEnvio 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4800
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdEnviar 
         Caption         =   "Enviar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txSendFile 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3840
         Width           =   4095
      End
      Begin VB.DirListBox dirCarpetas 
         Height          =   990
         Left            =   1080
         TabIndex        =   4
         Top             =   960
         Width           =   3135
      End
      Begin VB.FileListBox flArchivos 
         Height          =   1650
         Left            =   1080
         TabIndex        =   3
         Top             =   2040
         Width           =   3135
      End
      Begin VB.DriveListBox drvUnidades 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Archivo:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Directorio:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Unidad:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSerialFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private curPort As String
Private buf() As Byte
Private bytesRecibidos As Long
Private Sub cmbSerial_Change()
    cmbSerial_Click
End Sub

Private Sub cmbSerial_Click()
    Dim errMsg As String
    On Local Error GoTo er
    Dim iPort As Integer, a As Integer
    If MSComm1.PortOpen Then
        MSComm1.PortOpen = False
        For a = 1 To 1024: DoEvents: DoEvents: Next
    End If
    iPort = CInt(Right$(cmbSerial.Text, Len(cmbSerial.Text) - 3))
    
    MSComm1.CommPort = iPort
    MSComm1.PortOpen = True
    curPort = "COM" & iPort
er:
    If Err.Number <> 0 Then
        errMsg = IIf(Err.Number = 8002, "No se puede conectar al puerto indicado.", Err.Description)
        MsgBox errMsg, vbCritical + vbOKOnly + vbSystemModal, "Envío de archivos por puerto serie"
        cmbSerial.Text = curPort
        Exit Sub
    End If
End Sub

Private Sub cmdEnviar_Click()
    Dim nFile As Integer
    Dim bytes() As Byte
    Dim dataa As Byte
    Dim fileLenght As Long, i As Long
    Dim datas As String
   If MsgBox("Está seguro que desea enviar el archivo " & txSendFile.Text & " por el puerto serie?", vbYesNo + vbQuestion, "Enviar archivos por puerto serie") = vbYes Then
        nFile = FreeFile
        Open txSendFile.Text For Binary As #nFile
        
        fileLenght = LOF(nFile)
        ReDim bytes(fileLenght)
        i = 0
        Do While Not EOF(nFile)
            Get nFile, , dataa
            bytes(i) = dataa
            DoEvents: DoEvents: DoEvents
            progresoEnvio.Value = (100 / fileLenght) * i
            i = i + 1
            Debug.Print (Chr$(dataa))
            MSComm1.Output = Chr$(dataa)
            Sleep (50)
        Loop
        Close nFile
        
        datas = StrConv(bytes, vbUnicode)
        MsgBox ("Envío completado" & vbCrLf & i & " bytes enviados."), vbInformation + vbOKOnly, "Envío de archivos por puerto serie"
   
    End If
End Sub

Private Sub cmdRecibe_Click()

Timer1.Enabled = Not Timer1.Enabled
If Timer1.Enabled Then
    cmdRecibe.Caption = "Escuchando"

Else ' borrar todo
    If MsgBox("¿Guardar bufer actual?", vbQuestion + vbYesNo, "Enviar y recibir archivos por el puerto serie") = vbYes Then
        cmdSaveAs_Click
    End If
    cmdRecibe.Caption = "Recibir"
    bytesRecibidos = 0
    ReDim buf(0)
    txBytesRecibidos.Text = ""
End If
End Sub

Private Sub cmdSaveAs_Click()
    Dim datas As String
    Dim filito As Integer
    
    datas = StrConv(buf, vbUnicode)
    'MsgBox datas
    filito = FreeFile
    Open "\bufer.txt" For Append As #filito
    Print #filito, datas
    Close #filito
End Sub

Private Sub dirCarpetas_Change()
    txSendFile.Text = ""
    flArchivos.Path = dirCarpetas.Path
End Sub

Private Sub drvUnidades_Change()
    dirCarpetas.Path = Left$(drvUnidades.Drive, 2) & "\"
End Sub

Private Sub flArchivos_Click()
    Dim slash As String
    slash = IIf(Right$(dirCarpetas.Path, 1) = "\", "", "\")
    txSendFile.Text = dirCarpetas.Path & slash & flArchivos.List(flArchivos.ListIndex)
    cmdEnviar.Enabled = MSComm1.PortOpen
    
End Sub

Private Sub Form_Load()
    Dim c As Integer
    For c = 1 To 254
        cmbSerial.AddItem ("COM" & c)
    Next
    cmbSerial.Text = cmbSerial.List(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim a As Integer
    If MSComm1.PortOpen Then
        MSComm1.PortOpen = False
        For a = 1 To 1024: DoEvents: DoEvents: Next
    End If
End Sub

Private Sub Timer1_Timer()
    If MSComm1.InBufferCount > 0 Then
        ReDim Preserve buf(bytesRecibidos)
        buf(bytesRecibidos) = Asc(MSComm1.Input)
        txBytesRecibidos.Text = txBytesRecibidos.Text & Chr$(buf(bytesRecibidos))
        bytesRecibidos = bytesRecibidos + 1
        'Debug.Print (MSComm1.Input)
        txBytesRecibidos.SelStart = Len(txBytesRecibidos.Text)
        
    End If
End Sub
