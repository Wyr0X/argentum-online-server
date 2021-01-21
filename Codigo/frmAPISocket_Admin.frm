VERSION 5.00
Begin VB.Form frmAPISocket 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visor de API"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Conectar"
      Height          =   360
      Left            =   6120
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdShutdown 
      Caption         =   "Cerrar Socket"
      Height          =   360
      Left            =   7800
      TabIndex        =   5
      Top             =   210
      Width           =   1500
   End
   Begin VB.TextBox txtResponse 
      Height          =   4245
      Left            =   570
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4770
      Width           =   8685
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   600
      Left            =   1740
      TabIndex        =   1
      Top             =   3360
      Width           =   6150
   End
   Begin VB.TextBox txtSend 
      Height          =   2325
      Left            =   540
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   8685
   End
   Begin VB.Label lblOutput 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Respuesta"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3720
      TabIndex        =   4
      Top             =   4170
      Width           =   2055
   End
   Begin VB.Label lblInput 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Peticion"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3720
      TabIndex        =   3
      Top             =   150
      Width           =   1815
   End
End
Attribute VB_Name = "frmAPISocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Socket As clsSocket
Attribute Socket.VB_VarHelpID = -1

Public Sub Connect()
    '*********************************************************************
    'Author: Jopi
    'Conexion a la API.
    '*********************************************************************
    
    ' Me fijo que estamos usando la API.
    If Not API_Enabled Then Exit Sub
    
    If Socket Is Nothing Then
        Set Socket = New clsSocket
    End If
    
    ' Si estamos en VB, tenemos que llamar a la API directamente para cerrar el socket
    ' Sino tenemos que re-abrir el VB para que el hilo se cierre
    ' Por lo que primero lo cerramos, para evitar errores.
    If RunningInVB() Or Socket.State <> (sckClosed Or sckConnecting) Then
        Call Socket.CloseSck
        DoEvents
    End If

    'Usamos la API de Windows
    Call Socket.Connect(API_HostName, API_Port)

End Sub

Private Sub cmdConnect_Click()
    Call frmAPISocket.Connect
End Sub

Private Sub cmdShutdown_Click()
    Call Socket.CloseSck
End Sub

Private Sub cmdEnviar_Click()

    On Error Resume Next

    If Len(txtSend.Text) = 0 Then Exit Sub
    
    Call SendDataAPI(txtSend.Text)

End Sub

Private Sub Socket_BeforeSend(ByRef data As String)
    'Agregamos el separador de paquetes al string que vamos a enviar
    data = data & ";"
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    '*********************************************************************
    'Author: Jopi
    'Que hacemos con los datos que recibimos de la API.
    '*********************************************************************
    Dim recievedData As String
    
    Call Socket.GetData(recievedData, vbString, bytesTotal)
    
    ' Si no llego nada, nos vamos alv.
    If Len(recievedData) = 0 Then Exit Sub

    'Process the data we recieved from the API
    Call API.HandleIncomingAPIData(recievedData)
    
    Debug.Print vbNewLine
    Debug.Print "Tamaño: " & bytesTotal
    Debug.Print recievedData
    Debug.Print vbNewLine
    
End Sub

Private Sub Socket_Connect()
    '*********************************************************************
    'Author: Jopi
    'Que hacemos apenas nos conectamos a la API.
    '*********************************************************************
    
    Debug.Print "Conectado a la API"
    
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    Debug.Print vbNewLine
    Debug.Print "Error al conectarse a la API..."
    Debug.Print "Error: " & Number
    Debug.Print "Description: " & Description
    Debug.Print vbNewLine
    
End Sub

