VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Get Data"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   615
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   120
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "init.frx":0000
      Top             =   840
      Width           =   6855
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1920
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Tags"
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Ending Text"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Starting Text"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Site URL"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public flag1
Public flag2
Dim dflag, pos
Function GetData(ByVal URL As String)
Dim Data() As Byte
Dim i As Integer
Dim st As Long
Dim s, d, g

st = 1
Init
Data = Inet1.OpenURL(URL)
On Error GoTo err
Text1.Text = ""
    
s = InStr(st, Data, Text3.Text, 1)
While s > 0 And dflag = 0
d = InStr(Len(Text3.Text) + s, Data, Text4.Text, 1)
g = Mid(Data, s, d - s + Len(Text4.Text)) 'Len(Text3.Text) +
DoEvents
Text1.Text = Text1.Text & vbCrLf & _
"=================================================" & _
"=======================" & vbCrLf & g
st = d + 1
s = InStr(st, Data, Text3.Text, 1)
Wend

err:
If err Then
Text1.Text = "No Contents Found between selected Text..."
End If
End Function

Public Function Init()
flag1 = 0
flag2 = 0
End Function

Private Sub Command3_Click()
dflag = 0
Text1.Text = "Executing The Request......"
Command3.Enabled = False
Command1.Enabled = False
Me.GetData (Text2.Text)
Command3.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
dflag = 1
Command3.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Command1_Click()
dflag = 0
Text1.Text = "Executing The Request......"
Command3.Enabled = False
Command1.Enabled = False
Me.GetTags (Text2.Text)
Command3.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Form_Load()
dflag = 0
Form1.Caption = "Capture Data from any Website"
Text1.Text = "Result Window"
Text2.Text = "www.envy.nu/prashant"
Text3.Text = "<a"
Text4.Text = "</a>"
End Sub
Function GetTags(ByVal URL As String)
Dim Data() As Byte
Dim str As String
Dim i As Long

On Error Resume Next
Init
Data = Inet1.OpenURL(URL)
Text1.Text = ""
For i = 1 To UBound(Data)
    If dflag = 1 Then
    Exit For
    End If
        
    If Chr(Data(i)) = "<" Then flag1 = 1

 If Data(i) > 0 And flag1 = 1 Then
    str = str + Chr(Data(i))
    If Chr(Data(i)) = ">" Then
        flag2 = 1
        Text1.Text = Text1.Text & vbCrLf & str
        str = ""
        Init
    End If
 End If
DoEvents
Next i
err:
If err Then
Text1.Text = "This site has internal Server Errors....."
End If
End Function

