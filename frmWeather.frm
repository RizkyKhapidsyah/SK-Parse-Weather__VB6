VERSION 5.00
Begin VB.Form frmWeather 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Get Weather"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4110
   Icon            =   "frmWeather.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   38
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   36
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame fraWeather 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3975
      Begin VB.Label lblZip 
         BackStyle       =   0  'Transparent
         Caption         =   "99999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   8415
      End
      Begin VB.Image imgDay 
         Height          =   495
         Index           =   4
         Left            =   720
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   495
      End
      Begin VB.Image imgDay 
         Height          =   495
         Index           =   3
         Left            =   720
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   495
      End
      Begin VB.Image imgDay 
         Height          =   495
         Index           =   2
         Left            =   720
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   495
      End
      Begin VB.Image imgDay 
         Height          =   495
         Index           =   1
         Left            =   720
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   495
      End
      Begin VB.Image imgDay 
         Height          =   495
         Index           =   0
         Left            =   720
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblLow 
         BackStyle       =   0  'Transparent
         Caption         =   "Low 5"
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   34
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblLow 
         BackStyle       =   0  'Transparent
         Caption         =   "Low 4"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   33
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label lblLow 
         BackStyle       =   0  'Transparent
         Caption         =   "Low 3"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   32
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblLow 
         BackStyle       =   0  'Transparent
         Caption         =   "Low 2"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   31
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblLow 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Low 1"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   30
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "High 5"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   29
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "High 4"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   28
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "High 3"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   27
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblHigh 
         BackStyle       =   0  'Transparent
         Caption         =   "High 2"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   26
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblHigh 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "High 1"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   25
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label lblWeather 
         BackStyle       =   0  'Transparent
         Caption         =   "Weather 5"
         Height          =   375
         Index           =   4
         Left            =   1320
         TabIndex        =   24
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label lblWeather 
         BackStyle       =   0  'Transparent
         Caption         =   "Weather 4"
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   23
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label lblWeather 
         BackStyle       =   0  'Transparent
         Caption         =   "Weather 3"
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   22
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lblWeather 
         BackStyle       =   0  'Transparent
         Caption         =   "Weather 2"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   21
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblWeather 
         BackStyle       =   0  'Transparent
         Caption         =   "Weather 1"
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   20
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date 5"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   19
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date 4"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   18
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date 3"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   17
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date 2"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   16
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date 1"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Day 5"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   14
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Day 4"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   13
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Day 3"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   12
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Day 2"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Day 1"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image imgCurrent 
         Height          =   630
         Left            =   120
         Top             =   480
         Width           =   750
      End
      Begin VB.Label lblCurrentW 
         BackStyle       =   0  'Transparent
         Caption         =   "Weather"
         Height          =   495
         Left            =   960
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblCurrentH 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "High"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Humidity:"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblHumid 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dew Point: "
         Height          =   255
         Left            =   2220
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblDew 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3120
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "High"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblCurrent 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Weather"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Zip Code"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgNoPic 
      Height          =   780
      Left            =   2640
      Picture         =   "frmWeather.frx":000C
      Top             =   5880
      Visible         =   0   'False
      Width           =   780
   End
End
Attribute VB_Name = "frmWeather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function GetWeather(strZip As String) As Boolean

' Get Weather
Dim vntSource As Variant
Dim strWeather As String
Dim intPos As Long
Dim intNum As Integer
Dim lngTmp As Long
Dim blnPic As Boolean


  
  On Error GoTo errWeather
  GetWeather = True
  vntSource = GetUrlSource("http://www.weather.com/weather/local/" & strZip)
  ' Get City and Zip
  intPos = InStr(vntSource, "bodyText")
  intPos = InStr(intPos, vntSource, "<B>") + 3
  strWeather = Mid(vntSource, intPos, InStr(vntSource, "<!-- Insert City Name and Zip Code -->") - intPos)
  lblZip.Caption = LTrim(strWeather)
  
  ' Get current conditions
  intPos = InStr(vntSource, "<!-- insert forecast text -->") + 29
  strWeather = Mid(vntSource, intPos, InStr(intPos, vntSource, "</td>") - intPos)
  lblCurrentW.Caption = LTrim(strWeather)
  If FileExists(App.Path & "\Images\Weather\" & LTrim(strWeather) & ".gif") Then
    imgCurrent.Picture = LoadPicture(App.Path & "\Images\Weather\" & LTrim(strWeather) & ".gif")
  Else
    imgCurrent.Picture = imgNoPic.Picture
  End If
  intPos = InStr(vntSource, "<!-- insert current temp -->") + 28
  strWeather = Mid(vntSource, intPos, InStr(intPos, vntSource, "&") - intPos)
  lblCurrentH.Caption = strWeather & Chr(176) & " F"
  intPos = InStr(vntSource, "<!-- insert dew point -->") + 25
  strWeather = Mid(vntSource, intPos, InStr(intPos, vntSource, "&") - intPos)
  lblDew.Caption = strWeather & Chr(176) & " F"
  intPos = InStr(vntSource, "<!-- insert humidity -->") + 24
  strWeather = Mid(vntSource, intPos, InStr(intPos, vntSource, "</td>") - intPos)
  lblHumid.Caption = strWeather
  
  vntSource = GetUrlSource("http://www.weather.com/weather/tenday/" & strZip)
  intPos = InStr(intPos, vntSource, "<!-- insert load change day/date here -->") + 41
  
  For intNum = 0 To 4
    blnPic = False
    ' Get the day of the week
    If intPos = 0 Or intNum = 0 Or intNum = 1 Or (intNum = 2 And lblDay(0).Caption = "Tonight") Then
      If intPos = 0 Then
        intPos = 1
      End If
      intPos = InStr(intPos, vntSource, "<!-- insert load change day/date here -->") + 41
    End If
    If intPos = 0 Or intNum > 2 Or (intNum = 2 And lblDay(0).Caption <> "Tonight") Then
      If intPos = 0 Then
        intPos = 1
      End If
      intPos = InStr(intPos, vntSource, "<!-- insert no link day name here -->") + 37
    End If
    If intNum < 2 Then
      intPos = InStr(intPos + 1, vntSource, Chr(10)) + 1
      strWeather = Mid(vntSource, intPos, InStr(intPos + 1, vntSource, "</A>") - intPos)
    Else
      If lblDay(0).Caption = "Tonight" And intNum = 2 Then
        intPos = InStr(intPos + 1, vntSource, Chr(10)) + 1
        strWeather = Mid(vntSource, intPos, InStr(intPos + 1, vntSource, "</A>") - intPos)
      Else
        intPos = InStr(intPos, vntSource, Chr(10)) + 1
        strWeather = Mid(vntSource, intPos, InStr(intPos + 1, vntSource, "<BR>") - intPos)
      End If
    End If
    lblDay(intNum).Caption = LTrim(strWeather)
        
    ' Get Month and Day
    intPos = InStr(intPos, vntSource, "<BR>") + 4
    strWeather = Mid(vntSource, intPos, InStr(intPos, vntSource, Chr(10)) - intPos)
    lblDate(intNum).Caption = strWeather
      
    ' Get Weather Condition
    intPos = InStr(intPos, vntSource, "dataText") + 10
    intPos = InStr(intPos, vntSource, Chr(10)) + 1
    strWeather = Mid(vntSource, intPos, InStr(intPos + 1, vntSource, Chr(10)) - intPos)
    lblWeather(intNum).Caption = LTrim(strWeather)
    If FileExists(App.Path & "\Images\Weather\" & LTrim(strWeather) & ".gif") Then
      imgDay(intNum).Picture = LoadPicture(App.Path & "\Images\Weather\" & LTrim(strWeather) & ".gif")
    Else
      If InStr(strWeather, "Fog") > 0 Then
        imgDay(intNum).Picture = LoadPicture(App.Path & "\Images\Weather\Fog.gif")
        blnPic = True
      End If
      If InStr(strWeather, "Showers") > 0 Then
        imgDay(intNum).Picture = LoadPicture(App.Path & "\Images\Weather\Showers.gif")
        blnPic = True
      End If
      If Not blnPic Then
        imgDay(intNum).Picture = imgNoPic.Picture
      End If
    End If
  
    ' Get High Temp
    intPos = InStr(intPos, vntSource, "dataText") + 10
    intPos = InStr(intPos, vntSource, "&nbsp;") + 8
    intPos = InStr(intPos, vntSource, Chr(10)) + 1
    strWeather = Mid(vntSource, intPos, InStr(intPos + 1, vntSource, "&") - intPos)
    If Val(LTrim(strWeather)) = 0 Then
      lblHigh(intNum).Caption = ""
    Else
      lblHigh(intNum).Caption = Val(LTrim(strWeather)) & Chr(176)
    End If
    
    ' Get Low Temp
    intPos = InStr(intPos, vntSource, "dataText") + 10
    intPos = InStr(intPos, vntSource, "&nbsp;") + 8
    strWeather = Mid(vntSource, intPos, InStr(intPos + 1, vntSource, "&") - intPos)
    lblLow(intNum).Caption = LTrim(strWeather) & Chr(176)
  Next
  lblCurrent.Caption = "Current Weather (" & Now & ")"
  Exit Function
  
errWeather:
  If Err.Number = 5 Then
    MsgBox "Zip Code does not exist", vbExclamation, "Invalid Zip"
  Else
    If Err.Number = 0 Then
      Resume
    Else
      MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
    End If
  End If
  GetWeather = False
  
End Function


Private Sub cmdRefresh_Click()
  Screen.MousePointer = vbHourglass
  If GetWeather(txtZip.Text) Then
    SetValue App.Path & "\Weather.ini", "Settings", "Zip", txtZip.Text
  Else
    txtZip.SelStart = 0
    txtZip.SelLength = Len(txtZip.Text)
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim strZip As String

  strZip = GetValue(App.Path & "\Weather.ini", "Settings", "Zip", "")
  txtZip.Text = strZip
  GetWeather strZip
  
End Sub


