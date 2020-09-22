VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Better PlanetSourceCode Search - by Dinosurfer"
   ClientHeight    =   720
   ClientLeft      =   3180
   ClientTop       =   2175
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboWorld 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmSearch.frx":0000
      Left            =   2040
      List            =   "frmSearch.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton optInternal 
      Caption         =   "Internal Browser"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton optIE 
      Caption         =   "Popup Internet Explorer"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Enter Your Search Here"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   300
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picSearching 
      BackColor       =   &H00800000&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   4995
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Searching...Please Wait"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1665
         TabIndex        =   6
         Top             =   240
         Width           =   1725
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Display Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect
Public SearchString As String

Private Sub cmdSearch_Click()
If cboWorld.Text = "Select World" Or cboWorld.Text = "" Then MsgBox "Please Select A World To Search In.", , "Better PSC Search": Exit Sub
If txtSearch.Text = "" Or txtSearch.Text = "Enter Your Search Here" Then MsgBox "Please Enter Something To Search For.", , "Better PSC Search": Exit Sub
Me.Caption = txtSearch.Text + " - Better PlanetSourceCode Search - by Dinosurfer"
picSearching.Visible = True
picSearching.ZOrder 0
lngWorld = worldID(cboWorld.Text)
SearchString = ConvertSpaces(txtSearch.Text)
Select Case optIE.Value
Case True
OpenIE SearchURL & SearchString & URLWorld & lngWorld
picSearching.Visible = False
Case False
frmBrowser.Caption = "Search Results For " & txtSearch.Text
frmBrowser.SearchInternal SearchString, lngWorld
End Select
End Sub

Private Sub txtSearch_Click()
If txtSearch.Text = "Enter Your Search Here" Then txtSearch.Text = ""
End Sub

Function worldID(wName As String) As Long
Select Case wName
Case "ASP/VBScript"
worldID = 4
Case "Visual Basic"
worldID = 1
Case "C++/C"
worldID = 3
Case "Javascript/Java"
worldID = 2
Case "Perl"
worldID = 6
Case "Delphi"
worldID = 7
Case "PHP"
worldID = 8
Case "SQL"
worldID = 5
End Select
End Function

Function ConvertSpaces(strData As String) As String
'Special Character Conversions Coming Soon.
For i = 1 To Len(strData)
ch = Mid(strData, i, 1)
If ch = " " Then
sQ = sQ & "+"
Else
sQ = sQ & ch
End If
Next
ConvertSpaces = sQ
End Function
