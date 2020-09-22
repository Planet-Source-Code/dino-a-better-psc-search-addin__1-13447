VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Results For"
   ClientHeight    =   4800
   ClientLeft      =   2070
   ClientTop       =   3120
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      ExtentX         =   13785
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   0
      Width           =   6015
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "X"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   ">>"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SearchURL = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&optSort=Alphabetical&txtCriteria="
Const URLWorld = "&lngWId="

Sub SearchInternal(Query As String, worldID)
Me.Show , frmSearch
web1.Navigate SearchURL & Query & URLWorld & worldID
txtAddress.Text = web1.LocationURL
Me.SetFocus
frmSearch.picSearching.Visible = False
End Sub

Private Sub cmdBack_Click()
web1.GoBack
End Sub

Private Sub cmdForward_Click()
web1.GoForward
End Sub

Private Sub cmdRefresh_Click()
web1.Refresh
End Sub

Private Sub cmdStop_Click()
web1.Stop
End Sub

Private Sub Form_Load()
web1.Width = Me.ScaleWidth
web1.Height = Me.ScaleHeight - 255
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
web1.Navigate txtAddress.Text
End If
End Sub

Private Sub web1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
txtAddress.Text = web1.LocationURL
End Sub
