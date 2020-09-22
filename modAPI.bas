Attribute VB_Name = "modAPI"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global Const SearchURL = "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&optSort=Alphabetical&txtCriteria="
Global Const URLWorld = "&lngWId="
Const SW_NORMAL = 1
Const SW_SHOW = 5

Sub OpenIE(URL)
ShellExecute frmSearch.hWnd, "open", URL, "", "", SW_NORMAL Or SW_SHOW
End Sub

Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub
