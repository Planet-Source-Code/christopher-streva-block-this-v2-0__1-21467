VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBlockThis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Block This!"
   ClientHeight    =   4050
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBlockThis2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBlockThis2.frx":030A
   ScaleHeight     =   4050
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txt32BitOctetNum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1680
      Width           =   4455
   End
   Begin VB.TextBox txtBinaryNum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2160
      Width           =   4455
   End
   Begin VB.TextBox txtLongBinaryNum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2640
      Width           =   4455
   End
   Begin VB.TextBox txtDecimalConversionNum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3120
      Width           =   4455
   End
   Begin VB.CommandButton cmdGetIP 
      Caption         =   "Get IP!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtFinalURL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3600
      Width           =   4455
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtIP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblBlockThis 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Block This!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label lblIP 
      BackStyle       =   0  'Transparent
      Caption         =   " This Is The IP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "This Is The Final URL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label lblDecimalConversion 
      BackStyle       =   0  'Transparent
      Caption         =   "This Is The Decimal Conversion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label lblLongBinary 
      BackStyle       =   0  'Transparent
      Caption         =   "This Is The Long Binary:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lblBinary 
      BackStyle       =   0  'Transparent
      Caption         =   " This Is The Binary Conversion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label lbl32BitOctet 
      BackStyle       =   0  'Transparent
      Caption         =   " This Is The 32-Bit Octet:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblChooseURL 
      BackStyle       =   0  'Transparent
      Caption         =   "Put The URL Here (No WWW):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "frmBlockThis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConvert_Click()

  Dim intIPDec(1 To 4) As Integer  'Contains each section of IP in decimal
  Dim strIPBin(1 To 4) As String  'Contains each section of the IP in binary
  Dim strLongBinaryNum As String 'Contains long binary number
  Dim lngLongDecimalNum As String 'Contains long decimal number
  Call SeparateIPIntoSections(intIPDec())
  Call ConvertDecimalToLongBinary(intIPDec(), strIPBin(), strLongBinaryNum)
  Call ConvertLBinaryToLDecimal(strLongBinaryNum, lngLongDecimalNum)

End Sub

Private Sub SeparateIPIntoSections(ByRef intIPDec() As Integer)

Dim strIP As String 'Original IP
strIP = txtIP
Dim intLocation As Integer 'Location of character in original IP
Dim intPeriodLoc(1 To 3) As Integer 'Location of period in original IP
Dim intPeriodNumber As Integer 'Counter for intPeriodLoc
intPeriodNumber = 1

For intLocation = 1 To Len(strIP)
  If Mid(strIP, intLocation, 1) = "." Then
    intPeriodLoc(intPeriodNumber) = intLocation
    intPeriodNumber = intPeriodNumber + 1
  End If
Next intLocation

intIPDec(1) = Mid(strIP, 1, intPeriodLoc(1) - 1)
intIPDec(2) = Mid(strIP, intPeriodLoc(1) + 1, intPeriodLoc(2) - intPeriodLoc(1))
intIPDec(3) = Mid(strIP, intPeriodLoc(2) + 1, intPeriodLoc(3) - intPeriodLoc(2))
intIPDec(4) = Mid(strIP, intPeriodLoc(3) + 1, Len(strIP) - intPeriodLoc(3))

txt32BitOctetNum = intIPDec(1) & "  " & intIPDec(2) & "  " & _
                   intIPDec(3) & "  " & intIPDec(4)

End Sub

Private Sub ConvertDecimalToLongBinary(ByRef intIPDec() As Integer, _
  ByRef strIPBin() As String, strLongBinaryNum)

Dim intSection As Integer 'Section of IP
Dim intLeftOver As Integer 'Used in conversion calculations
Dim intCounter As Integer 'Used to add zeros to binary sections

For intSection = 4 To 1 Step -1
  intLeftOver = intIPDec(intSection)
  Do While intLeftOver <> 0
    strIPBin(intSection) = (intLeftOver Mod 2) & strIPBin(intSection)
    intLeftOver = intLeftOver \ 2
  Loop
  For intCounter = 1 To (8 - Len(strIPBin(intSection)))
    strIPBin(intSection) = "0" & strIPBin(intSection)
  Next intCounter
Next intSection

txtBinaryNum.Text = strIPBin(1) & "  " & strIPBin(2) & _
                        "  " & strIPBin(3) & "  " & strIPBin(4)
strLongBinaryNum = strIPBin(1) & strIPBin(2) & _
                   strIPBin(3) & strIPBin(4)
txtLongBinaryNum.Text = strLongBinaryNum

End Sub

Private Sub ConvertLBinaryToLDecimal(ByVal strLongBinaryNum As String, _
    ByVal lngLongDecimalNum As String)

  Dim intLocation As Integer
  lngLongDecimalNum = Left(strLongBinaryNum, 1)
  
  For intLocation = 2 To Len(strLongBinaryNum)
    If Mid(strLongBinaryNum, intLocation, 1) = "1" Then
      lngLongDecimalNum = lngLongDecimalNum * 2 + 1
    Else
      lngLongDecimalNum = lngLongDecimalNum * 2
    End If
  Next intLocation

  txtDecimalConversionNum = lngLongDecimalNum
  txtFinalURL = "http://" & lngLongDecimalNum
End Sub

Private Sub cmdGetIP_Click()
  
  cmdGetIP.Enabled = False
  sckClient.Connect txtURL.Text, 80

End Sub

Private Sub sckClient_Connect()

txtIP.Text = sckClient.RemoteHostIP
sckClient.Close

End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "You screwed up somewhere along the line!", vbCritical, "You Messed Up!"
sckClient.Close

End Sub

Private Sub txtIP_Change()

If txtIP = "" Then
    cmdConvert.Enabled = False
Else
    cmdConvert.Enabled = True
End If

End Sub

Private Sub txtURL_Change()

txt32BitOctetNum = ""
txtBinaryNum = ""
txtLongBinaryNum = ""
txtDecimalConversionNum = ""
txtFinalURL = ""
txtIP = ""
If txtURL = "" Then
    cmdGetIP.Enabled = False
Else
    cmdGetIP.Enabled = True
End If

End Sub
