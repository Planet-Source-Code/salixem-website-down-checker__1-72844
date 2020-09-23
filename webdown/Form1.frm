VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Website down checker"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   1380
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   450
      TabIndex        =   0
      Text            =   "google.com"
      Top             =   45
      Width           =   3345
   End
   Begin SHDocVwCtl.WebBrowser W 
      Height          =   150
      Left            =   450
      TabIndex        =   2
      Top             =   3825
      Width           =   150
      ExtentX         =   265
      ExtentY         =   265
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   555
      Left            =   45
      TabIndex        =   1
      Top             =   405
      Width           =   3750
   End
   Begin VB.Label Label2 
      Caption         =   "IDLE"
      Height          =   375
      Left            =   45
      TabIndex        =   4
      Top             =   990
      Width           =   3750
   End
   Begin VB.Label Label1 
      Caption         =   "URL:"
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   90
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple website down checker
'Coded by Salar Zeynali
'Salixem@Gmail.Com
Private Sub Command1_Click()
W.Navigate "http://downforeveryoneorjustme.com/" & Text1.Text
End Sub
Private Function WebPageContains(ByVal s As String) As Boolean
Dim i As Long, HTMLElement
For i = 1 To W.Document.All.length
Set HTMLElement = _
W.Document.All.Item(i)
If Not (HTMLElement Is Nothing) Then
If InStr(1, HTMLElement.innerHTML, _
s, vbTextCompare) > 0 Then
WebPageContains = True
Exit Function
End If
End If
Next i
End Function
Private Sub W_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If WebPageContains("It's not just you!") = True Then
Label2.Caption = Text1 & " is down from here."
Else
End If

If WebPageContains("It's just you.") = True Then
Label2.Caption = Text1 & " isn't down from here."
End If
End Sub
