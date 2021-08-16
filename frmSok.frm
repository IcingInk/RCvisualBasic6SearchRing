VERSION 5.00
Begin VB.Form frmSok 
   Caption         =   "Ringnummer - Sök!"
   ClientHeight    =   8505
   ClientLeft      =   75
   ClientTop       =   720
   ClientWidth     =   16545
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   16545
   Begin VB.TextBox RingTillRC 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   37
      Top             =   6120
      Width           =   11295
      Visible         =   0   'False
   End
   Begin VB.TextBox lblMaerkFangst 
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   240
      TabIndex        =   36
      Top             =   3540
      Width           =   11295
      Visible         =   0   'False
   End
   Begin VB.TextBox txtColText 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5040
      TabIndex        =   27
      Top             =   1665
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   6045
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   315
      Width           =   3705
   End
   Begin VB.TextBox txtCeName 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox RingTillMaerkare 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5760
      Width           =   11295
   End
   Begin VB.TextBox LeveransUppg 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5280
      Width           =   11295
   End
   Begin VB.Timer Tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9690
      Top             =   120
   End
   Begin VB.PictureBox txtRing1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      ScaleHeight     =   315
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Skriv märkdata"
      Height          =   495
      Left            =   10050
      TabIndex        =   4
      Top             =   2535
      Width           =   1500
   End
   Begin VB.CommandButton cmdGranska 
      Caption         =   "Granska allt"
      Height          =   495
      Left            =   10050
      TabIndex        =   3
      Top             =   1680
      Width           =   1500
   End
   Begin VB.CommandButton cmdSok 
      Caption         =   "Sök"
      Default         =   -1  'True
      Height          =   495
      Left            =   10050
      TabIndex        =   2
      Top             =   600
      Width           =   1500
   End
   Begin VB.PictureBox txtCentral 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      ScaleHeight     =   375
      ScaleWidth      =   615
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   615
   End
   Begin VB.Line L3 
      BorderColor     =   &H00FF00FF&
      Index           =   2
      X1              =   11700
      X2              =   11700
      Y1              =   360
      Y2              =   6600
   End
   Begin VB.Line L2 
      BorderColor     =   &H00FF00FF&
      Index           =   0
      X1              =   11700
      X2              =   15540
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lblVaenta 
      BackColor       =   &H8000000E&
      Caption         =   "VÄNTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   10290
      TabIndex        =   42
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label ObserveraSkylt 
      Caption         =   "OBSERVERA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   12360
      TabIndex        =   41
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label VarningsSkylt 
      BackColor       =   &H8000000E&
      Caption         =   "VARNINGAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   12480
      TabIndex        =   40
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblVarningsLabel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   11880
      TabIndex        =   39
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label lblF1 
      Caption         =   "tryck F1 för ny sökning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10050
      TabIndex        =   38
      Top             =   2265
      Width           =   1695
   End
   Begin VB.Label lblEkoAvskr 
      BackColor       =   &H80000005&
      Caption         =   "Avskriven Fagel2-kontr finns med detta ringnummer - granska"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   11880
      TabIndex        =   35
      Top             =   2640
      Width           =   3495
      Visible         =   0   'False
   End
   Begin VB.Label lblUrspMarkRing 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2880
      TabIndex        =   34
      Top             =   2130
      Width           =   6855
      WordWrap        =   -1  'True
   End
   Begin VB.Line L1 
      BorderColor     =   &H00FF00FF&
      Index           =   1
      X1              =   120
      X2              =   15240
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label lblAvskr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Avskrivet fynd finns på denna ring - granska!"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   11880
      TabIndex        =   33
      Top             =   2040
      Width           =   3135
      Visible         =   0   'False
   End
   Begin VB.Label lblFRV 
      BackColor       =   &H80000005&
      Caption         =   "Kontroll av ringmärkare (annan än märkaren) finns - granska"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   11880
      TabIndex        =   32
      Top             =   5160
      Width           =   3015
      Visible         =   0   'False
   End
   Begin VB.Label lblVarning 
      Height          =   735
      Left            =   0
      TabIndex        =   31
      Top             =   2400
      Width           =   4815
      Visible         =   0   'False
   End
   Begin VB.Label lblTreffRub2 
      Caption         =   "  Centr   Ringnummer  "
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
      Left            =   6060
      TabIndex        =   30
      Top             =   90
      Width           =   3690
   End
   Begin VB.Label lblTreffRub1 
      Caption         =   "  Centr   Ringnummer      Färgkod"
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
      Left            =   6060
      TabIndex        =   29
      Top             =   90
      Width           =   3690
   End
   Begin VB.Label lblAvbrytSpec 
      BackColor       =   &H00C0C0FF&
      Caption         =   " Avbryt specialsök"
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
      Left            =   2880
      TabIndex        =   28
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblInstr5 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   120
      TabIndex        =   26
      Top             =   4440
      Width           =   10935
   End
   Begin VB.Label lblInstr4 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   25
      Top             =   5340
      Width           =   10935
   End
   Begin VB.Label lblInstr3 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   3960
      Width           =   10935
   End
   Begin VB.Label lblInstr2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   10935
   End
   Begin VB.Label lblInstr 
      Caption         =   "Instruktioner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   10170
      TabIndex        =   22
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblFringTreff 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   5040
      TabIndex        =   21
      Top             =   360
      Width           =   960
   End
   Begin VB.Label lblColMark 
      BackColor       =   &H00C000C0&
      Caption         =   " Sökning färgringtext  >"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   1665
      Width           =   2055
      Visible         =   0   'False
   End
   Begin VB.Label lblUtanCentr 
      BackColor       =   &H00FF0000&
      Caption         =   " Sökning utan central"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   840
      Width           =   2055
      Visible         =   0   'False
   End
   Begin VB.Label lblSTOCKHOLM 
      BackColor       =   &H0080FF80&
      Caption         =   " STOCKHOLM"
      ForeColor       =   &H80000010&
      Height          =   315
      Left            =   2880
      TabIndex        =   17
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblTranspHeld 
      BackColor       =   &H80000005&
      Caption         =   "   Märkdata:   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   15
      Top             =   3150
      Width           =   11295
      Visible         =   0   'False
   End
   Begin VB.Label lblEkoFinns 
      BackColor       =   &H80000005&
      Caption         =   "Egen (el associerad märkares) kontroll finns "
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   11880
      TabIndex        =   14
      Top             =   5760
      Width           =   3255
      Visible         =   0   'False
   End
   Begin VB.Label lblFyndFinns 
      BackColor       =   &H80000005&
      Caption         =   "Fynd finns !!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   345
      Left            =   11880
      TabIndex        =   13
      Top             =   4680
      Width           =   2175
      Visible         =   0   'False
   End
   Begin VB.Label MarkTillclash 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OBS! Uspr märkaren ej samme som fått ursprungliga ringen!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   570
      Left            =   11880
      TabIndex        =   12
      Top             =   3360
      Width           =   2175
      Visible         =   0   'False
   End
   Begin VB.Label lblOrt 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   4290
      Width           =   11295
   End
   Begin VB.Image Image1 
      Height          =   2070
      Left            =   240
      Picture         =   "frmSok.frx":0000
      Top             =   120
      Width           =   2550
   End
   Begin VB.Label lblMaerkare 
      Height          =   300
      Left            =   2880
      TabIndex        =   8
      Top             =   1785
      Width           =   6855
      Visible         =   0   'False
   End
   Begin VB.Label lblEfter 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   7320
      Width           =   11295
   End
   Begin VB.Label lblFoere 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   6840
      Width           =   11295
   End
   Begin VB.Label lblMaerkData 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   5
      Top             =   3885
      Width           =   11295
   End
End
Attribute VB_Name = "frmSok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Vareja As String
Private Ring As clsRingRc, OldRtr As Integer
Private MaerkDataHittad As Boolean, KontrHittad As Boolean, FyndHittat As Boolean
Private bSoekSVS As Boolean, ListVal As Boolean, Testar As Boolean
Private bSoekAnnanSvensk As Boolean, OlSoektring As String, OlFring As String
Private bSoekUtland As Boolean, Acv As String, Burknamn As String
Private RM1 As String, RM2 As String, RM3 As String, RM4 As String, RM5 As String

Private Sub cmdGranska_Click()   'Inte riktigt klar
md
On Error Resume Next

'Set Ac = New Access.Application
'Ac.OpenCurrentDatabase "C:\program files\sokringrc\soekring.mdb"
' If DE.rscmdSoekRing.State = 1 Then DE.rscmdSoekRing.Close
' If DE.rscmdSoekRing.State = 0 Then DE.rscmdSoekRing.Open
' rptRing.Show vbModal, Me

'Ac.DoCmd.RunMacro "rappvisa"
'Ac.Visible = True

End Sub

Private Sub cmdGranska_GotFocus()
  Me.Tmr.Enabled = False
End Sub

Private Sub cmdGranska_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
  txtRing1 = ""
  txtCentral = "SVS"
  txtCeName.Visible = False
  lblSTOCKHOLM.Visible = True
  txtRing1.SetFocus
End If

End Sub

Private Sub cmdPrint_Click()
'ListVal = False
'rptRing.PrintReport False, rptRangeAllPages
If DE.rscmdMärkdata.State = 1 Then DE.rscmdMärkdata.Close
If DE.rscmdMärkdata.State = 0 Then DE.rscmdMärkdata.Open
rptMärkdata.PrintReport

'Ac.DoCmd.OpenReport "rptData", acViewNormal
'Ac.DoCmd.OpenReport "rptData", acNormal
'Ac.DoCmd.Close
'Ac.quit
'Set Ac = Nothing

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
  txtRing1 = ""
  txtCentral = "SVS"
  txtCeName.Visible = False
  lblSTOCKHOLM.Visible = True
  txtRing1.SetFocus
End If

End Sub

Private Sub cmdSok_Click()
Dim Ringar As clsRingarRc
' Dim Ringar As clInkiRingar
Dim Col As Collection
Dim sArtkod As String
Dim C As Integer
Dim bMnr As Boolean
Dim sMnr As String
Dim Rnr As String
Dim uRing As String, SoekAringR As String, SoekAringC As String
Dim uCentr As String, BoxRad As String
Dim bOmmarkt As Boolean
Dim visRing As String
Dim visCentr As String, Vistr As String
Dim OlCe As String, Olce2 As String, Rir As String, Kontir As String, Fyri As String
Dim Uri As String, Ufri As String

Dim bOmmarktFinns As Boolean
Dim bOmmarktSaknas As Boolean
Dim bDataHittad As Boolean
'Dim rs As ADODB.Recordset
On Error GoTo Err_Handler

Conn.Execute ("DELETE FROM SökRing")

Art_Skyddad = False


If Trim(txtRing1) = "" Then
  Screen.MousePointer = vbDefault
  txtRing1.SetFocus
  Exit Sub
End If

lblFRV.Visible = False
lblAvskr.Visible = False
lblEkoAvskr.Visible = False
lblMaerkFangst.Visible = False
lblMaerkare.Visible = False
VarningsSkylt.Visible = False
lblVarningsLabel.Visible = False
ObserveraSkylt.Visible = False
If Trim(txtCentral = "") Then lblUtanCentr.Visible = True
List1.Visible = False
List1.Clear
lblAvbrytSpec.Visible = False
lblTreffRub1.Visible = False
lblTreffRub2.Visible = False
lblFringTreff.Visible = False
Refresh

Utland = False
UtlandSVS = False
FyndFinns = False
KontrHittad = False
MaerkDataHittad = False
Sparas = False
Basdataut = False
VarFrom = ""

'Rensar tabellerna i den lokala databasen SoekRing
Vareja = "inget"
'************************************
' Gällande Connectionstring skall vara: med lite mer ändringar
'som Providern vilken typ av Access det är. 2000 eller 97
Vareja = "A"
Vareja = "efterA"

'Set rs = New ADODB.Recordset
'Set rs = Conn.Execute("DELETE * FROM SökRing")
'Set rs = Nothing

'testar först att txtRing1 har data i sig

If txtRing1 = "" Then
  txtRing1.SetFocus
  Exit Sub 'före connection
End If

bMnr = False
sMnr = ""
bOmmarkt = False
bOmmarktFinns = False
bOmmarktSaknas = False
bFRVHittad = False

Screen.MousePointer = vbHourglass
Set Ringar = New clsRingarRc

'rensa formuläret
lblFyndFinns.Visible = False
lblEkoFinns.Visible = False
lblMaerkData = ""
lblFoere = ""
lblEfter = ""
lblMaerkare = ""
lblOrt = ""
RingTillMaerkare = ""
lblVarningsLabel.Caption = ""
LeveransUppg = ""
OlCe = ""
Conn.Execute ("DELETE FROM SökRingMtext") 'formuläret är rensat


If ListVal = True Then GoTo BoerjaSoek 'Hoppar förbi förberedande sökningar när man redan varit i List1

If txtCentral = "F" Then   'sökning på färgmärkes text
  
  txtColText = UCase(txtRing1)
  OlFring = txtRing1
  'Set rs = New ADODB.Recordset
  MaerkeClash = False
  Rir = ""
  Kontir = ""
  Fyri = ""
  Uri = ""
  Ufri = ""
  
  If rs.State = 1 Then rs.Close
  rs.Open "SELECT centr,text1,text2,ring,colorkod1,colorkod2 FROM ringon WHERE ((text1= '" & RTrim(txtRing1) & "') OR (text2= '" & RTrim(txtRing1) & "')) and (tr='1' or tr = 'M')", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If RM1 <> "" Then
        'Då är det en andra märkpost med samma märketext och detta bör annonseras
        lblVarningsLabel.Caption = "Det finns flera märkposter med samma färgmärke-text!"
        lblVarningsLabel.Visible = True
        VarningsSkylt.Visible = True
        Refresh
      End If

      If IsNull(rs!text1) = False And RTrim(rs!text1) = txtRing1 Then
        RM1 = rs!Centr & "   " & rs!Ring & "       " & rs!colorkod1
      Else
        RM1 = rs!Centr & "   " & rs!Ring & "       " & rs!colorkod2
      End If
      
      Conn.Execute ("INSERT INTO SökRingMtext VALUES ('" & txtRing1 & "','','U','" & RM1 & "')")

      List1.AddItem RM1
      rs.MoveNext
    Loop
  End If
  rs.Close
  
  If rs.State = 1 Then rs.Close
  rs.Open "SELECT centr,text1,text2,ring,colorkod1,colorkod2 FROM ringon WHERE ((text1= '" & RTrim(txtRing1) & "') OR (text2= '" & RTrim(txtRing1) & "')) and (tr='7'  or tr = '4' or tr = 'N'  or tr = 'L')", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      
      If IsNull(rs!text1) = False And RTrim(rs!text1) = txtRing1 Then
        RM1 = rs!Centr & "   " & rs!Ring & "       " & rs!colorkod1
        MT = rs!text1
      Else
        RM1 = rs!Centr & "   " & rs!Ring & "       " & rs!colorkod2
        MT = rs!text2
      End If
      
      If RsMt.State = 1 Then RsMt.Close
      RsMt.Open "SELECT * FROM SökRingMtext WHERE Ceringcol= '" & RM1 & "'", Conn, adOpenStatic, adLockReadOnly
      If RsMt.EOF = True And RsMt.BOF = True Then
        Conn.Execute ("INSERT INTO SökRingMtext VALUES ('" & txtRing1 & "','','R','" & RM1 & "')")
        List1.AddItem RM1
      End If
      
      rs.MoveNext
    Loop
  End If
  rs.Close
  
  rs.Open "SELECT centr,ring,rappcent,rappring,text1,text2,colorkod1,colorkod2 FROM f2kontr WHERE ((text1= '" & RTrim(txtRing1) & "') OR (text2= '" & RTrim(txtRing1) & "'))", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!Centr) = False And Trim(rs!Centr) > "" And IsNull(rs!Ring) = False And Trim(rs!Ring) > "" Then
        If IsNull(rs!text1) = False And RTrim(rs!text1) = txtRing1 Then
          RM1 = rs!Centr & "   " & rs!Ring & "       " & rs!colorkod1
          MT = rs!text1
        Else
          RM1 = rs!Centr & "   " & rs!Ring & "       " & rs!colorkod2
          MT = rs!text2
        End If
      Else
        If IsNull(rs!text1) = False And RTrim(rs!text1) = txtRing1 Then
          RM1 = rs!RappCent & "   " & rs!RappRing & "       " & rs!colorkod1
          MT = rs!text1
        Else
          RM1 = rs!RappCent & "   " & rs!RappRing & "       " & rs!colorkod2
          MT = rs!text2
        End If
      End If
      
      If RsMt.State = 1 Then RsMt.Close
      RsMt.Open "SELECT * FROM SökRingMtext WHERE Ceringcol= '" & RM1 & "'", Conn, adOpenStatic, adLockReadOnly
      If RsMt.EOF = True And RsMt.BOF = True Then
        Conn.Execute ("INSERT INTO SökRingMtext VALUES ('" & txtRing1 & "','','K','" & RM1 & "')")
        List1.AddItem RM1
      End If
      
      rs.MoveNext
    Loop
  End If
  rs.Close
  
  rs.Open "SELECT centr,ring,rappcent,rappring,text1,text2,colorkod1,colorkod2 FROM fynd WHERE ((text1= '" & RTrim(txtRing1) & "') OR (text2= '" & RTrim(txtRing1) & "'))", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!Centr) = False And Trim(rs!Centr) > "" And IsNull(rs!Ring) = False And Trim(rs!Ring) > "" Then
        If IsNull(rs!text1) = False And RTrim(rs!text1) = txtRing1 Then
          RM1 = rs!Centr & "   " & rs!Ring & "       " & rs!colorkod1
          MT = rs!text1
        Else
          RM1 = rs!Centr & "   " & rs!Ring & "       " & rs!colorkod2
          MT = rs!text2
        End If
      Else
        If IsNull(rs!text1) = False And RTrim(rs!text1) = txtRing1 Then
          RM1 = rs!RappCent & "   " & rs!RappRing & "       " & rs!colorkod1
          MT = rs!text1
        Else
          RM1 = rs!RappCent & "   " & rs!RappRing & "       " & rs!colorkod2
          MT = rs!text2
        End If
      End If
      
      If RsMt.State = 1 Then RsMt.Close
      RsMt.Open "SELECT * FROM SökRingMtext WHERE Ceringcol= '" & RM1 & "'", Conn, adOpenStatic, adLockReadOnly
      If RsMt.EOF = True And RsMt.BOF = True Then
        Conn.Execute ("INSERT INTO SökRingMtext VALUES ('" & txtRing1 & "','','F','" & RM1 & "')")
        List1.AddItem RM1
      End If
      rs.MoveNext
    Loop
  End If
  rs.Close
  
  rs.Open "SELECT Uring.centr,Uring.ring,text1,text2,colorkod1,colorkod2 FROM uring left join Ceadress on Uring.centr=Ceadress.centr WHERE ((text1= '" & RTrim(txtRing1) & "') OR (text2= '" & RTrim(txtRing1) & "'))", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    RM1 = ""
    rs.MoveFirst
    Do Until rs.EOF
      If RM1 <> "" Then
        'Då är det en andra märkpost med samma märketext och detta bör annonseras
        lblVarningsLabel.Caption = "Det finns flera märkposter med samma färgmärke-text!"
        lblVarningsLabel.Visible = True
        VarningsSkylt.Visible = True
        Refresh
      End If
      
      If IsNull(rs!text1) = False And RTrim(rs!text1) = txtRing1 Then
        RM1 = rs!Centr & "   " & rs!Ring & "       " & rs!colorkod1
        MT = rs!text1
      Else
        RM1 = rs!Centr & "   " & rs!Ring & "       " & rs!colorkod2
        MT = rs!text2
      End If
      
      If RsMt.State = 1 Then RsMt.Close
      RsMt.Open "SELECT * FROM SökRingMtext WHERE Ceringcol= '" & RM1 & "'", Conn, adOpenStatic, adLockReadOnly
      If RsMt.EOF = True And RsMt.BOF = True Then
        Conn.Execute ("INSERT INTO SökRingMtext VALUES ('" & txtRing1 & "','','U','" & RM1 & "')")
        List1.AddItem RM1
      End If
      rs.MoveNext
    Loop
  End If
  rs.Close
  
  rs.Open "SELECT centr,ring,rappcent,rappring,text1,text2,colorkod1,colorkod2 FROM Ufynd WHERE ((text1= '" & RTrim(txtRing1) & "') OR (text2= '" & RTrim(txtRing1) & "')) ", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!Centr) = False And Trim(rs!Centr) > "" And IsNull(rs!Ring) = False And Trim(rs!Ring) > "" Then
        If IsNull(rs!text1) = False And RTrim(rs!text1) = txtRing1 Then
          RM1 = rs!Centr & "   " & rs!Ring & "       " & rs!colorkod1
          MT = rs!text1
        Else
          RM1 = rs!Centr & "   " & rs!Ring & "       " & rs!colorkod2
          MT = rs!text2
        End If
      Else
        If IsNull(rs!text1) = False And RTrim(rs!text1) = txtRing1 Then
          RM1 = rs!RappCent & "   " & rs!RappRing & "       " & rs!colorkod1
          MT = rs!text1
        Else
          RM1 = rs!RappCent & "   " & rs!RappRing & "       " & rs!colorkod2
          MT = rs!text2
        End If
      End If
      
      If RsMt.State = 1 Then RsMt.Close
      RsMt.Open "SELECT * FROM SökRingMtext WHERE Ceringcol= '" & RM1 & "'", Conn, adOpenStatic, adLockReadOnly
      If RsMt.EOF = True And RsMt.BOF = True Then
        Conn.Execute ("INSERT INTO SökRingMtext VALUES ('" & txtRing1 & "','','F','" & RM1 & "')")
        List1.AddItem RM1
      End If
      rs.MoveNext
    Loop
  End If
  rs.Close
 
  If List1.ListCount > 0 Then List1.Visible = True
  If List1.ListCount > 1 Then
    lblFringTreff.Visible = True
    lblFringTreff.Caption = "Mer än ett alternativ finns - välj med klick - tryck sedan SÖK >"
    List1.SetFocus
  End If
  If List1.ListCount = 1 Then
   lblFringTreff.Visible = True
   lblFringTreff.Caption = "Finns bara en - söks nu direkt"
   List1.Selected(0) = True
   GoTo AdrFixaRing
  End If
   
  If List1.ListCount = 0 Then
   lblMaerkData = "Det finns ingen uppgift om färgmärke med texten " & txtColText
   lblVaenta.Visible = False
   Screen.MousePointer = Default
   Exit Sub
  End If
  
  List1.Visible = True
  lblAvbrytSpec.Visible = True
  lblTreffRub1.Visible = True
  Refresh

  Screen.MousePointer = Default
  lblVaenta.Visible = False
  Exit Sub
Else 'Gör förberedande sökning på enbart ringnummer för att få lista på förekommande centraler
  If Right(txtCentral, 1) <> "X" Then
    Rnr = Ringar.FixaRingnr(txtRing1.Text) 'fixar ringnumret
    Soektring = Rnr
    txtRing1.Text = Rnr
  Else
    Rnr = txtRing1.Text
  End If
  If rs.State = 1 Then rs.Close 'Ringon - sök Ring
  rs.Open "SELECT DISTINCT Centr, Ring FROM Ringon WHERE Ring= '" & Soektring & "'", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!Ring) = False And Trim(rs!Ring) > "" Then List1.AddItem rs!Centr & "   " & rs!Ring & "       Ringon - Ring"
      rs.MoveNext
    Loop
  End If
  If rs.State = 1 Then rs.Close 'Ringon - sök GRing
  rs.Open "SELECT DISTINCT GCentr, GRing FROM Ringon WHERE GRing= '" & Soektring & "'", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!GRing) = False And Trim(rs!GRing) > "" Then List1.AddItem rs!GCentr & "   " & rs!GRing & "       Ringon - GRing"
      rs.MoveNext
    Loop
  End If
  If rs.State = 1 Then rs.Close 'URing - sök Ring
  rs.Open "SELECT DISTINCT Centr, Ring FROM URing WHERE Ring= '" & Soektring & "'", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!Ring) = False And Trim(rs!Ring) > "" Then List1.AddItem rs!Centr & "   " & rs!Ring & "       URing - Ring"
      rs.MoveNext
    Loop
  End If
  If rs.State = 1 Then rs.Close 'URing - sök GRing
  rs.Open "SELECT DISTINCT GCentr, GRing FROM URing WHERE GRing= '" & Soektring & "'", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!GRing) = False And Trim(rs!GRing) > "" Then List1.AddItem rs!GCentr & "   " & rs!GRing & "       URing - GRing"
      rs.MoveNext
    Loop
  End If
    If rs.State = 1 Then rs.Close 'UFynd - sök Rappring
  rs.Open "SELECT DISTINCT Rappcent, Rappring FROM UFynd WHERE Rappring= '" & Soektring & "'", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
    rs.Close 'UFynd - sök färgringtext i Rappring
    rs.Open "SELECT DISTINCT Rappcent, Rappring FROM UFynd WHERE Rappring= '" & Replace(Soektring, " ", "") & "' AND Right(RappCent,1)='X'", Conn, adOpenStatic, adLockReadOnly
    If rs.EOF = True And rs.BOF = True Then
    Else
      rs.MoveFirst
      Do Until rs.EOF
        If IsNull(rs!RappRing) = False Then List1.AddItem rs!RappCent & "   " & rs!RappRing & "       UFynd - Rappring"
        rs.MoveNext
      Loop
    End If
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!RappRing) = False Then List1.AddItem rs!RappCent & "   " & rs!RappRing & "       UFynd - Rappring"
      rs.MoveNext
    Loop
  End If
  If rs.State = 1 Then rs.Close 'UFynd - sök Ring
  rs.Open "SELECT DISTINCT Centr, Ring FROM UFynd WHERE Ring= '" & Soektring & "'", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!Ring) = False And Trim(rs!Ring) > "" Then List1.AddItem rs!Centr & "   " & rs!Ring & "       UFynd - Ring"
      rs.MoveNext
    Loop
  End If
  If rs.State = 1 Then rs.Close 'Fynd - sök Rappring
  rs.Open "SELECT DISTINCT Rappcent, Rappring FROM Fynd WHERE Rappring= '" & Soektring & "'", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!RappRing) = False Then List1.AddItem rs!RappCent & "   " & rs!RappRing & "       Fynd - Rappring"
      rs.MoveNext
    Loop
  End If
  If rs.State = 1 Then rs.Close 'Fynd - sök Ring
  rs.Open "SELECT DISTINCT Centr, Ring FROM Fynd WHERE Ring= '" & Soektring & "'", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!Ring) = False And Trim(rs!Ring) > "" Then List1.AddItem rs!Centr & "   " & rs!Ring & "       Fynd - Ring"
      rs.MoveNext
    Loop
  End If
  If rs.State = 1 Then rs.Close 'F2Kontr - sök Rappring
  rs.Open "SELECT DISTINCT Rappcent, Rappring FROM F2Kontr WHERE Rappring= '" & Soektring & "'", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!RappRing) = False Then List1.AddItem rs!RappCent & "   " & rs!RappRing & "       F2Kontr - Rappring"
      rs.MoveNext
    Loop
  End If
  If rs.State = 1 Then rs.Close 'F2Kontr - sök Ring
  rs.Open "SELECT DISTINCT Centr, Ring FROM F2Kontr WHERE Ring= '" & Soektring & "'", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True And rs.BOF = True Then
  Else
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!Ring) = False And Trim(rs!Ring) > "" Then List1.AddItem rs!Centr & "   " & rs!Ring & "       F2Kontr - Ring"
      rs.MoveNext
    Loop
  End If
  rs.Close
  
  If List1.ListCount > 1 Then
    List1.Visible = True
    lblTreffRub2.Visible = True
    Testar = True
    C = 0
    List1.Selected(C) = True
    visCentr = Left(List1.Text, 3)
    C = 1
    Do While C < List1.ListCount
      List1.Selected(C) = True
      If visCentr <> Left(List1.Text, 3) Then
        Screen.MousePointer = vbDefault
        lblFringTreff.Visible = True
        lblFringTreff.Caption = "Mer än ett alternativ finns - välj med klick - tryck sedan SÖK >"
        Testar = False
        Exit Sub
      End If
      C = C + 1
    Loop
    Testar = False
    List1_Click
    GoTo BoerjaSoek
  End If
  If List1.ListCount = 1 Then
    List1.Selected(0) = True
    'Exit Sub
    List1_Click
    GoTo BoerjaSoek
  End If
  If List1.ListCount = 0 Then
    lblMaerkData.Visible = True
    lblMaerkData = "Det finns inga uppgifter registrerade om ringen " & txtRing1 & " - fortsätter sökning efter andra uppgifter."
    Refresh
  End If

End If

AdrFixaRing:

If Right(txtCentral, 1) <> "X" Then
  Rnr = Ringar.FixaRingnr(txtRing1.Text) 'fixar ringnumret
  Soektring = Rnr
  txtRing1.Text = Rnr
Else
  Rnr = txtRing1.Text
End If
If txtRing1 <> OlSoektring And OlSoektring <> "" Then ListVal = False
OlSoektring = txtRing1

uRing = Rnr
visRing = Rnr
Refresh
'Set rs = New ADODB.Recordset
If ListVal = True Then GoTo BoerjaSoek


'SÖKNING UTAN CENTRAL eller om central= SVS
If txtCentral = "" Or txtCentral.Text = "SVS" Then

  If rs.State = 1 Then rs.Close
  rs.Open "SELECT mnr,tr, ringon.centr,ring, gcentr, gring, ceadress.cekortad, Gcek=Gcead.cekortad FROM Ringon LEFT JOIN Ceadress ON ringon.centr = Ceadress.centr  LEFT JOIN Ceadress as Gcead ON ringon.centr = Gcead.centr WHERE Ring= '" & Soektring & "'", Conn, adOpenStatic, adLockReadOnly
  'Conn.Search_Ringon txtRing1, rs
  If rs.EOF = True And rs.BOF = True Then
    If txtCentral.Text = "SVS" Then 'Sök även i Uring om inget hittas - finns några SVS där också
      rs.Close
      rs.Open "SELECT mnr,tr, uring.centr,ring, gcentr, gring, ceadress.cekortad, Gcek=Gcead.cekortad FROM URing LEFT JOIN Ceadress ON Uring.centr = Ceadress.centr  LEFT JOIN Ceadress as Gcead ON uring.centr = Gcead.centr WHERE Ring= '" & Soektring & "'", Conn, adOpenStatic, adLockReadOnly
      If rs.EOF = True And rs.BOF = True Then
      Else
        rs.MoveFirst
        Do Until rs.EOF
          If IsNull(rs!Ring) = False And Trim(rs!Ring) > "" And rs!Centr = "SVS" Then
            List1.AddItem rs!Centr & "   " & rs!Ring & "       " & rs!cekortad
            UtlandSVS = True
          End If
          OlCe = rs!Centr
          If IsNull(rs!GRing) = False And Trim(rs!GRing) > "" And rs!GCentr <> Olce2 Then
            List1.AddItem rs!GCentr & "   " & rs!GRing & "       " & rs!cekortad
            Olce2 = rs!GCentr
            List1.AddItem rs!GCentr & "   " & rs!GRing & "       " & rs!gcek
          End If
          rs.MoveNext
        Loop
      End If
    End If
  Else
    OlCe = ""
    Olce2 = ""
    rs.MoveFirst
    Do Until rs.EOF
      If IsNull(rs!Ring) = False And Trim(rs!Ring) > "" And rs!Centr <> OlCe Then List1.AddItem rs!Centr & "   " & rs!Ring & "       " & rs!cekortad
      OlCe = rs!Centr
      If rs!Centr <> txtCentral Then Olce2 = rs!Centr
      rs.MoveNext
    Loop
  End If
  rs.Close

  If txtCentral = "" Then
    Conn.Search_URing txtRing1, rs
    If rs.State = 0 Then
    Else
      If rs.EOF = False Then
        'rs.MoveFirst
        Do Until rs.EOF
          If IsNull(rs!Ring) = False And Trim(rs!Ring) > "" And rs!Centr <> OlCe Then List1.AddItem rs!Centr & "   " & rs!Ring & "       " & rs!cekortad
          OlCe = rs("centr")
          If IsNull(rs!GRing) = False And Trim(rs!GRing) > "" And rs!GCentr <> Olce2 Then
            List1.AddItem rs!GCentr & "   " & rs!GRing & "       " & rs!cekortad
            Olce2 = rs!GCentr
            List1.AddItem rs!GCentr & "   " & rs!GRing & "       " & rs!gcek
          End If
          rs.MoveNext
        Loop
      End If
      rs.Close
    End If
  End If

  Screen.MousePointer = Default
 
  If List1.ListCount > 1 Then
    List1.Visible = True
    'If RTrim(txtCentral) = "" Then lblAvbrytSpec.Visible = True
    lblTreffRub2.Visible = True
    lblFringTreff.Visible = True
    lblFringTreff.Caption = "Mer än ett alternativ finns - välj med klick - tryck sedan SÖK >"
    lblTreffRub2.Visible = True
    Exit Sub
  End If
  If List1.ListCount = 1 Then
    If Olce2 <> "" And txtColText = "" Then
      MsgBox "OBS!" & Chr(13) & Chr(13) & "Sökning sker nu på central " & Olce2 & " som inte är den som först angavs i urvalet (" & txtCentral & ").", vbOKOnly, "Annan central!"
      txtCentral = Olce2
      lblTreffRub2.Visible = False
      List1.ListIndex = 0
      txtCeName.Visible = True
      txtCeName = Mid(List1.Text, 22)
      If Trim(txtCeName) = "RIKSMUSEET" Then txtCeName = txtCeName & " gamla"
      Refresh
    End If
    List1.Selected(0) = True
    GoTo BoerjaSoek
  End If
  If List1.ListCount = 0 Then
    lblMaerkData.Visible = True
    lblMaerkData = "Det finns ingen svensk märkuppgift om ringen " & txtRing1 & " - fortsätter sökning efter andra uppgifter."
    Refresh
  End If

  lblVaenta.Visible = False
  'Exit Sub
End If


BoerjaSoek:

'Kolla först om sökta ringen är usrpr eller en ny ring
SoektTR = ""
GammalRing = ""
GammalCentr = ""
ListVal = False
lblVaenta.Visible = True
Refresh

If rs.State = 1 Then rs.Close
rs.Open "SELECT mnr,tr, centr,ring,gcentr, gring FROM Ringon WHERE Ring= '" & Soektring & "' and centr='" & txtCentral & "' ORDER BY tr DESC", Conn, adOpenStatic, adLockReadOnly
If rs.EOF = True And rs.BOF = True Then
Else
  SoektTR = rs!Tr
  If IsNull(rs!GCentr) = False And Trim(rs!GCentr) > "" Then GammalCentr = rs!GCentr
  If IsNull(rs!GRing) = False And Trim(rs!GRing) > "" Then GammalRing = rs!GRing
  If rs!Tr = "1" Or rs!Tr = "M" Then
    UrspMnr = rs!Mnr
'    txtCentral = rs!Centr
'    txtRing1 = rs!Ring
    UrspCentr = rs!Centr
    Urspring = rs!Ring
  End If
End If
rs.Close

Conn.Execute ("DELETE FROM SökRingMtext")
'Kolla upp detta
Central = txtCentral.Text
SoektCentr = txtCentral.Text
Soektring = txtRing1.Text
uCentr = Central
visCentr = Central
SoekAringR = Rnr
SoekAringC = Central
lblMaerkData = "Söker..."
Vareja = "Conn.open conSql = " + conSql
x = Right(conSql, 20)
'Conn.Open conSql

bSoekUtland = False
If Left(Central, 2) <> "SV" Or UtlandSVS = True Then bSoekUtland = True

' 2021-08-16, ingi smäller på txtRing1 ? provar att ersätta med txtRing1 med txtRing1.Text
'Set Col = Ringar.LetMaerkData(txtRing1, Central, bSoekUtland)
Set Col = Ringar.LetMaerkData(txtRing1.Text, Central, bSoekUtland)


If Col.Count > 4 Then MaerkDataHittad = True
For C = 1 To Col.Count 'gå igenom collectionen med märkdata
  Set Ring = Col(C)
  FyllRapptab
Next C

Rnr = Urspring
Set Col = Ringar.LetKontrData(bSoekUtland)
If Col.Count > 0 Then KontrHittad = True
For C = 1 To Col.Count 'gå igenom collectionen med kontrolldata
  Set Ring = Col(C)
  FyllRapptab
Next C

Fyndets_Centr = ""
Fyndets_Ring = ""
Set Col = Ringar.LetFyndData(bSoekUtland)
If Col.Count > 0 Then FyndFinns = True
For C = 1 To Col.Count 'gå igenom collectionen med fynddata
  Set Ring = Col(C)
  FyllRapptab
Next C

If Fyndets_Centr <> "" Then
  Dim bla As Boolean
  bla = False
  If Left(Fyndets_Centr, 2) <> "SV" Then bla = True
  Set Col = Ringar.LetMaerkData(Fyndets_Ring, Fyndets_Centr, bla)
  If Col.Count > 4 Then MaerkDataHittad = True
  For C = 1 To Col.Count 'gå igenom collectionen med märkdata
    Set Ring = Col(C)
    FyllRapptab
  Next C

End If
'Exit Sub

   
'Nu finns all data i rapptab (utom ev F2-post - se nedan)
'kör en refresh på rapporten för att öka hastigheten
'rptRing.Refresh
'On Error Resume Next
'On Error GoTo Err_Handler
'DE.rscmdSoekRing.Open
'Refresh
'DE.rscmdSoekRing.Requery
'Refresh
'DE.rscmdSoekRing.Close
'Refresh

If MaerkDataHittad = True Then
  Set Col = Nothing '--behövs nog inte
  Tmr.Enabled = True
Else
  If KontrHittad = False Then
    lblMaerkData = "Det finns ingen uppgift om ring " & txtCentral & " " & txtRing1
    If txtCentral > "" Then lblMaerkData = lblMaerkData & " OBSERVERA att du sökt på central " & txtCentral
  Else
    lblMaerkData = "Märkdata saknas, men kontrolldata finns"
  End If
  
End If

KollaLev:

'Hoppa ur om det är en utländsk ring
If Left(txtCentral, 2) = "SV" Then
Else
  Screen.MousePointer = vbDefault
  GoTo SoekSlut
End If

'om det inte blev någon match så skall vi
' 1. Hämta leverans mnr och ring -- fixa leverans kan nog vara sub och
   'hitta märkarnummer för närliggande märkdata

' 2021-08-16, ingimar -'txtRing' blir -> 'txtRing1.Text'
If GammalRing = "" Then
  If Left(GammalCentr, 2) <> "SV" Then Set Col = Ringar.Leverans(txtRing1.Text)
Else
  If GammalCentr = "SVS" Then
    Set Col = Ringar.Leverans(GammalRing)
  Else
    Set Col = Ringar.Leverans(txtRing1.Text)
  End If
End If


If Col.Count > 0 Then Set Ring = Col(1)

If Ring.Datafinns = MTyp.rcFinns Then
  'Det finns leveransdata
  bMnr = True
  sMnr = Ring.Mnr
  If lblMaerkData.Caption = "Söker..." Then lblMaerkData = ""
Else
  If GammalRing = "" Then
   If MaerkDataHittad = False And KontrHittad = False And FyndFinns = False Then
     lblMaerkData = "  Finns ingen uppgift alls om ringen " & Central & " " & txtRing1
     If txtCentral <> "SVS" Then lblMaerkData = lblMaerkData & " OBSERVERA att du sökt på central " & txtCentral
   End If
   If MaerkDataHittad = False And (KontrHittad = True Or FyndFinns = True) Then lblMaerkData = "  Märkdata saknas för denna ring " & Central & " " & txtRing1
  Else
   If MaerkDataHittad = False And KontrHittad = False And FyndFinns = False Then lblMaerkData = "  Sökt ring har ersatt " & GammalCentr & " " & GammalRing & " Finns ingen uppgift alls om ringen"
'   If KontrHittad = False And FyndFinns = False Then lblMaerkData = "  Sökt ring har ersatt " & GammalCentr & " " & GammalRing & " Finns ingen uppgift alls om ringen"
   If MaerkDataHittad = False And (KontrHittad = True Or FyndFinns = True) Then
     If SoektTR = "7" Or SoektTR = "N" Then lblMaerkData = " Sökt ring har ersatt " & GammalCentr & " " & GammalRing & " Märkdata saknas för denna ring"
     If SoektTR = "4" Or SoektTR = "L" Then lblMaerkData = " Sökt ring har tillagts " & GammalCentr & " " & GammalRing & " Märkdata saknas för denna ring"
   End If
  End If
End If

If MaerkDataHittad = True Then GoTo SoekSlut

Conn.Execute "INSERT INTO SökRing (Rdatum, Rtr, Rubrik, Märkdata) VALUES('980000',1, '" & frmSok.RingTillMaerkare & "', 'J')"
Conn.Execute "INSERT INTO SökRing (Rdatum, Rtr, Rubrik, Märkdata) VALUES('980000',2, '" & frmSok.RingTillRC & "', 'J')"
Conn.Execute "INSERT INTO SökRing (Rdatum, Rtr, Rubrik, Märkdata) VALUES('980000',3, '" & frmSok.LeveransUppg & "', 'J')"
Conn.Execute "INSERT INTO SökRing (Rdatum, Rtr, Rubrik, Märkdata) VALUES('980000',4, '======================================================================================', 'J')"

' 2. Hämta närliggande fynd dvs närliggande ringon
'    sätt variabler på bmnr och mnr
' 2021-08-16, Ingimar, txtRing1 till txtRing1.Text
If GammalRing = "" Then Set Col = Ringar.NarliggandeNummer(txtRing1.Text, bMnr, sMnr)
If GammalRing <> "" And GammalCentr = "SVS" Then Set Col = Ringar.NarliggandeNummer(GammalRing, bMnr, sMnr)

For C = 1 To Col.Count
  Set Ring = Col(C)
  If Ring.Datafinns = MTyp.rcSaknas Then
  Else
    If Ring.Artkod = "Fore" Then
      lblFoere.Visible = True
      If Ring.Centr = txtCentral Then
        lblFoere = "  Närmast lägre ringnummer:  " & Ring.Ring & " användes " & Ring.Datum & " av " & Ring.MarkarNamn & "  " & Ring.Mnr
      Else
        lblFoere = "  Närmast lägre ringnummer (OBS! annan central = " & Ring.Centr & "): " & Ring.Ring & " användes " & Ring.Datum & " av " & Ring.MarkarNamn & " " & Ring.Mnr
      End If
    End If
    
    If Ring.Artkod = "Efter" Then
      lblEfter.Visible = True
      If Ring.Centr = txtCentral Then
        lblEfter = "  Närmast högre ringnummer:  " & Ring.Ring & " användes " & Ring.Datum & " av " & Ring.MarkarNamn & "  " & Ring.Mnr
      Else
        lblEfter = "  Närmast högre ringnummer (OBS! annan central = " & Ring.Centr & "): " & Ring.Ring & " användes " & Ring.Datum & " av " & Ring.MarkarNamn & " " & Ring.Mnr
      End If
    End If
  End If
Next C

SoekSlut:

'Kolla om det finns märken registrerade och i så fall om med olika text, dvs flera poster i tabellen SökRingMtext
' - då är det illa
If RsMt.State = 1 Then RsMt.Close
RsMt.Open "SELECT ant=count(Mtext1) FROM SökRingMtext", Conn, adOpenStatic, adLockReadOnly
'RsMt.Open "SELECT * FROM SökRingMtext ORDER BY MTEXT,Posttyp", Conn, adOpenStatic, adLockReadOnly
If RsMt.EOF = True And RsMt.BOF = True Then
Else
  If RsMt!ant > 1 Then
    lblVarningsLabel.Caption = "Det finns olika färgmärke-texter till en och samma fågel - kolla granskning!"
    lblVarningsLabel.Visible = True
    VarningsSkylt.Visible = True
    L2(0).Visible = True
    L3(2).Visible = True
    Refresh
  End If
End If

Set Ringar = Nothing
Set Ring = Nothing
Set Col = Nothing

'Conn.Close

Screen.MousePointer = vbDefault
Tmr.Enabled = True
lblVaenta.Visible = False
Exit Sub

Resume
Err_Handler:
Screen.MousePointer = vbDefault
Select Case Err.Number
  Case 3705
    Resume Next
  Case 3704
    Resume Next
  Case -2147217896
    Resume Next
  '  If Conn.State = adStateClosed Then
  '    Conn.Open conSql
  '    Resume
  '  End If
  Case -2147467259
    MsgBox ("Mappnamnet är felaktigt, det går inte att hitta den begärda filen. Vareja=" & Vareja)
  Case 3706
    MsgBox "Du försöker leta i en fil som inte är Access97 - format."
  Case Else
    MsgBox Err.Number & Err.Description & " i sub CmdSok.Click / Vi går vidare!"
    Resume Next
End Select
Set Ringar = Nothing
End Sub

Private Sub cmdSok_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
  txtRing1 = ""
  txtCentral = "SVS"
  txtCeName.Visible = False
  lblSTOCKHOLM.Visible = True
  txtRing1.SetFocus
End If

End Sub

Private Sub lblavbrytspec_Click()
 lblFringTreff.Visible = False
 lblUtanCentr.Visible = False
 lblColMark.Visible = False
 txtColText.Visible = False
 lblAvbrytSpec.Visible = False
 lblTreffRub1.Visible = False
 lblTreffRub2.Visible = False
 List1.Visible = False
 List1.Clear
 txtCentral = "SVS"
 lblSTOCKHOLM.Visible = True
 txtRing1 = ""
 
 txtRing1.SetFocus
 
End Sub

Private Sub lblInstr_Click()
If lblInstr2.Visible = True Then
lblInstr2.Visible = False
lblInstr3.Visible = False
lblInstr4.Visible = False
lblInstr5.Visible = False
Exit Sub
End If
lblInstr2.Visible = True
lblInstr3.Visible = True
lblInstr4.Visible = True
lblInstr5.Visible = True

End Sub

Private Sub List1_Click()
If Testar = True Then Exit Sub

ListVal = True
lblVaenta.Visible = True
txtCentral = Left(List1.Text, 3)
txtRing1 = Mid(List1.Text, 7, 8)
lblUtanCentr.Visible = False
If txtCentral <> "SVS" Then txtCeName.Visible = True
If txtColText <> "" Then
  txtColText.Visible = True
  lblColMark.Caption = "Sökning färgringtext >"
Else
  txtCeName = Mid(List1.Text, 16)
  If Trim(txtCeName) = "RIKSMUSEET" Then txtCeName = txtCeName & " gamla"
End If
Refresh

Soektring = txtRing1
SoektCentr = txtCentral
'cmdSok_Click
End Sub

Private Sub Form_Load()
'Sätter connection string till servern som används genom hela programmet
Set Conn = New Connection
SQLUser = Environ("username")
If Conn.State = 0 Then
  'Öppna kommunikationen (vilken metod funkar bäst? ODBC används i Fyndin)
  'conSql = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Ring;Data Source=ringdb.nrm.se"
  ' conSql = "ODBC;Description=Ring;DRIVER=SQL Server;SERVER=ringdb.nrm.se;DATABASE=Ring;Trusted_Connection=Yes"
  conSql = "ODBC;Description=Ring;DRIVER=SQL Server;SERVER=nrmmssql05.nrm.se;DATABASE=Ring;Trusted_Connection=Yes"
  Conn.ConnectionString = conSql
  Conn.Open
End If

'Ställer in så att formuläret nu blir lagom stort - tidigare fyllde hela skärmen
Me.Top = 0
Me.Left = 0
Me.Height = 10000
Me.Width = 15600
'Me.Height = Screen.Height
'Me.Width = Screen.Width

'Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Me.BackColor = RGB(100, 200, 150)
lblMaerkare.BackColor = RGB(100, 200, 150)
Me.txtCentral.BackColor = RGB(100, 200, 150)
Me.lblSTOCKHOLM.BackColor = RGB(100, 200, 150)
Me.lblInstr.BackColor = RGB(100, 200, 150)
Me.lblF1.BackColor = RGB(100, 200, 150)
bSoekSVS = True
txtCentral = "SVS"
OlSoektring = ""
txtCeName.Visible = False
lblSTOCKHOLM.Visible = True
lblFringTreff.Visible = False
lblFRV.Visible = False
txtColText.Visible = False
List1.Visible = False
lblInstr2.Visible = False
lblInstr3.Visible = False
lblInstr4.Visible = False
lblInstr5.Visible = False
lblAvbrytSpec.Visible = False
lblTreffRub1.Visible = False
lblTreffRub2.Visible = False
lblVarningsLabel.Visible = False
VarningsSkylt.Visible = False
ObserveraSkylt.Visible = False
lblVaenta.Visible = False

Refresh
End Sub

Private Sub FyllRapptab()
Dim Rsr As ADODB.Recordset
On Error GoTo RappErr
'Dessa skall hämtas från Access.ini
Vareja = "B"
Set Rsr = New ADODB.Recordset
If Sparas = True Then VarFrom = "S"

'If IsNull(Ring.Basdata) = False And RTrim(Ring.Basdata) <> "" Then Basdataut = True
   
'Rsr.Open "SELECT * FROM SökRing WHERE Rdatum= '" & Ring.Rdatum & "' AND rtr= " & Ring.RTr & " AND dnr= '" & Ring.Dnr & "' AND basdata= '" & Ring.Basdata & "'", cnAc, adOpenStatic, adLockReadOnly
'If Rsr.State = 1 And Rsr.EOF = False Then
'Else
'  Conn.Execute "INSERT INTO SökRing (RDatum, RTr, Basdata, Dnr, Koordinat, Prov, Ort, Lokal, Tr, Datum, TNog, Sex, Age, Status, Varn, Pullus, Reserv, Vinge, Vikt, Fett, Pjm, Fynddetalj, Extradata, Rubrik, Varfrom)" & _
'    " VALUES (" & Ring.RDatum & ", " & Ring.RTr & ", '" & Ring.Basdata & "', '" & Ring.Dnr & "', '" & Ring.Koordinat & "', '" & Ring.Prov & "', '" & Ring.Ort & "', '" & Ring.Lokal & "', '" & Ring.Tr & "', '" & Ring.Datum & "', '" & Ring.TNog & "', '" & Ring.Sex & _
'    "', '" & Ring.Age & "', '" & Ring.Status & "', '" & Ring.Varn & "', '" & Ring.Pullus & "', '" & Ring.Reserv & "', '" & Ring.Vinge & "', '" & Ring.Vikt & "', '" & Ring.Fett & "', '" & Ring.Pjm & "', '" & Ring.Fynddetalj & "', '" & Ring.Extradata & "', '" & Ring.Rubrik & "', '" & VarFrom & "')"
'  If Ring.Dnr = "Ommärkt" And (Ring.Tr = "7" Or Ring.Tr = "N" Or Ring.Tr = "L") Then VarFrom = "S"

  Conn.Execute "INSERT INTO SökRing (RDatum, RTr, Basdata, Dnr, Koordinat, Prov, Ort, Lokal, Tr, Datum, TNog, Signatur, Sex, Age, Status, Varn, Pullus, Reserv, Vinge, Vikt, Fett, Pjm, Fynddetalj, Extradata, Rubrik, Varfrom, Märkdata)" & _
    " VALUES ('" & Ring.Rdatum & "', " & Ring.RTr & ", '" & Ring.Basdata & "', '" & Ring.Dnr & "', '" & Ring.Koordinat & "', '" & Ring.Prov & "', '" & Ring.Ort & "', '" & Ring.Lokal & "', '" & Ring.Tr & "', '" & Ring.Datum & "', '" & Ring.Tnog & "', '" & Ring.Signatur & "', '" & Ring.Sex & "', '" & Ring.Age & _
    "', '" & Ring.Status & "', '" & Ring.Varn & "', '" & Ring.Pullus & "', '" & Ring.Reserv & "', '" & Ring.Vinge & "', '" & Ring.Vikt & "', '" & Ring.Fett & "', '" & Ring.Pjm & "', '" & Ring.Fynddetalj & "', '" & Ring.Extradata & "', '" & Ring.Rubrik & "', '" & VarFrom & "', '" & Ring.Märkdata & "')"
  
  Exit Sub
'  If IsNull(Ring.Ort2) = False And RTrim(Ring.Ort2) <> "" Then

'  Conn.Execute "INSERT INTO SökRing (RDatum, RTr, Ort2, Varfrom)" & _
'    " VALUES ('" & Ring.Rdatum & "', " & (Ring.RTr + 1) & ", '" & Ring.Ort2 & "', '" & VarFrom & "')"
'End If

'If Ring.RTr = 20 Or Ring.RTr = 30 Or Ring.RTr = 40 Or Ring.RTr = 50 And OldRtr > 19 Then 'lägg in tom rad före fynd
  If (Right(Left((Ring.Dnr), 3), 1) = "/" Or Right(Left((Ring.Dnr), 3), 1) = ":" Or Right(Left((Ring.Dnr), 3), 1) = "Y" Or Right(Left((Ring.Dnr), 3), 1) = "F" Or Right(Left((Ring.Dnr), 3), 1) = "G" Or Right(Left((Ring.Dnr), 3), 1) = "L") And Ring.Tr <> "F" Then
'    Conn.Execute "INSERT INTO SökRing (RDatum, RTr,rubrik) VALUES (" & Ring.RDatum & ", " & 49 & ")"
    Conn.Execute "INSERT INTO SökRing (RDatum, RTr,rubrik,varfrom) VALUES ('" & Ring.Rdatum & "', " & 74 & ", " & "'--------------------------------------------------------------------------------------'" & ",'" & "S" & "')"
  End If

'End If

OldRtr = Ring.RTr

If Rsr.State = 1 Then Rsr.Close
Set Rsr = Nothing
Exit Sub

Resume
RappErr:
MsgBox Err.Number & " " & Err.Description
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
'stänger connection till SQL
If DE.rscmdSoekRing.State = 1 Then DE.rscmdSoekRing.Close
If DE.rscmdMärkdata.State = 1 Then DE.rscmdMärkdata.Close
If Conn.State = 1 Then Conn.Close
Set Conn = Nothing
End Sub

Private Sub Tmr_Timer()
cmdGranska.Enabled = True
cmdPrint.Enabled = True
cmdGranska.SetFocus
End Sub

Private Sub txtCentral_gotFocus()
txtCentral = ""
txtRing1 = ""
End Sub

Private Sub txtCentral_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRing1.SetFocus
If (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 228 Or KeyAscii = 229 Or KeyAscii = 246 Then KeyAscii = KeyAscii - 32
If Len(txtCentral) = 2 Then
  txtCentral = txtCentral + Chr(KeyAscii)
  txtRing1.SetFocus
End If
End Sub

Private Sub txtCentral_LostFocus()
lblColMark.Visible = False
txtColText.Visible = False
lblAvbrytSpec.Visible = False
lblFringTreff.Visible = False
txtColText = ""
List1.Clear
List1.Visible = False
lblTreffRub1.Visible = False
lblTreffRub2.Visible = False

txtCentral = UCase(txtCentral)

If txtCentral = "" Then Exit Sub

Dim Rsc As New ADODB.Recordset
'Set Rsc = New ADODB.Recordset
'If Conn.State = 0 Then Conn.Open conSql

If txtCentral = "SUM" Then txtCentral = "RUM" 'Ryssland ny kod

If txtCentral = "F" Then
  OlFring = ""
  lblColMark.Visible = True
  lblColMark.Caption = "Sökning färgringtext"
  'txtColtext.Visible = True
  Exit Sub
End If

Conn.sel_central txtCentral, Rsc
Screen.MousePointer = vbDefault
If Rsc.EOF = False Then
  Me.txtCeName = " " & Rsc("cekortad")
  If Trim(txtCeName) = "RIKSMUSEET" Then txtCeName = txtCeName & " gamla"
Else
  Me.txtCeName = "Ogiltig centralkod"
End If
Rsc.Close
'Set Rsc = Nothing

End Sub

Private Sub txtRing_Change()
If txtRing1 <> Soektring Then
  ListVal = False
End If
If ListVal = True Then Exit Sub

'Dim rs As ADODB.Recordset
'**************
'Rensa labels och delete i SoekRing tabellen,
'eventuellt refresh på rptring
'*****************
'rensa formuläret
lblFRV.Visible = False
frmSok.MarkTillclash.Visible = False
frmSok.lblTranspHeld.Visible = False
frmSok.lblMaerkFangst.Visible = False
If List1.Visible = False Then
  lblTreffRub2.Visible = False
  frmSok.lblMaerkare.Visible = False
End If
frmSok.lblFyndFinns.Visible = False
frmSok.lblEkoFinns.Visible = False
frmSok.lblVarningsLabel.Visible = False

LeveransUppg = ""
RingTillMaerkare = ""
lblMaerkData = ""
lblFoere = ""
lblEfter = ""
lblMaerkare = ""
lblOrt = ""
Urspring = ""
cmdGranska.Enabled = False
'formuläret är rensat
End Sub

Private Sub txtRing_GotFocus()
If ListVal = True Then Exit Sub

lblFRV.Visible = False
lblFyndFinns.Visible = False
lblFoere.Visible = False
lblEfter.Visible = False
lblMaerkData.Visible = False
lblUrspMarkRing.Visible = False
lblOrt.Visible = False
lblAvskr.Visible = False
lblEkoAvskr.Visible = False
LeveransUppg.Visible = False
RingTillMaerkare.Visible = False
RingTillRC.Visible = False
VarningsSkylt.Visible = False
ObserveraSkylt.Visible = False
cmdPrint.Enabled = False
cmdGranska.Enabled = False
L2(0).Visible = False
L3(2).Visible = False

Urspring = ""

If lblColMark.Visible = True And ListVal = False Then
  List1.Clear
  txtColText = ""
  txtRing1 = OlFring
  txtCentral = "F"
End If


If txtCentral = "SVS" Then
  lblSTOCKHOLM.Visible = True
  txtCeName.Visible = False
  lblUtanCentr.Visible = False
  GoTo Trogon
End If
If txtCentral = "F" Then
  lblSTOCKHOLM.Visible = False
  txtCeName.Visible = False
  lblUtanCentr.Visible = False
  GoTo Trogon
End If
If Trim(txtCentral) = "" Then
  lblSTOCKHOLM.Visible = False
  txtCeName.Visible = False
  lblUtanCentr.Visible = True
  GoTo Trogon
End If

lblUtanCentr.Visible = False
txtCeName.Visible = True
lblSTOCKHOLM.Visible = False

Trogon:
If Conn.State = 1 Then
  'Set rs = Conn.Execute("DELETE * FROM SökRing")
  'Set rs = Nothing
  Conn.Execute ("DELETE FROM SökRing")
  'Conn.Close
  'Set cnAc = Nothing
End If

'Varning! Programmet måste ställas in för rätt Access. Det är nu inställt
'för Access2000. Ändring görs i koden där Accessdatabasen öppnas -
'där finns två rader dne med 4.0 är för Access2000- den andra för Access97.
'Så måste DE justeras - där högerklickar man på cnAc och väljer egenskaper
'och där ska man välja rätt komponenter samt ange fullständig sökväg.
'I den burk där versionen för Access2000  körs måste också Accessdatabasen
'Soekring.mdb vara konverterad till Access2000.

'cnAc.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Program Files\sokringrc\SoekRing.mdb"
'cnAc.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=C:\Program Files\sokringrc\SoekRing.mdb"
'Set rs = cnAc.Execute("DELETE * FROM SökRing")
'Set rs = Nothing
'cnAc.Execute ("DELETE * FROM SökRing")

'Me.txtRing = ""
End Sub

Private Sub txtRing_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then txtCentral.SetFocus
End Sub

Private Sub txtRing_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub

If (KeyAscii >= 45 And KeyAscii <= 58) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 228 Or KeyAscii = 229 Or KeyAscii = 246 Or KeyAscii = 43 Or KeyAscii = 63 Or KeyAscii = 196 Or KeyAscii = 197 Or KeyAscii = 214 Then
Else
  KeyAscii = 0
End If

If (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 228 Or KeyAscii = 229 Or KeyAscii = 246 Then KeyAscii = KeyAscii - 32
If Len(txtRing1.Text) > 7 And txtRing1.SelLength = 0 Then
  KeyAscii = 0
  Exit Sub
End If
End Sub

















' 2021-08-16, ingi testar - behövs inte
Private Function InkimarLetMaerkData(Ringnr As String, Central As String, Utland As Boolean) As Collection
Dim Svar As String
'Dim Tr As String
Dim omUtland As Boolean 'Används vid sökning av ommärkt fågel
Dim C As Integer 'Används vid For loop vid ommärkning
Dim GRing As String, sTr As String, Ringensid As String
Dim GCentr As String, Soekid As Long
Dim Rsr As New ADODB.Recordset
Dim rsG As New ADODB.Recordset, rsFy As New ADODB.Recordset
Dim FyndRapp As Boolean, MaerkeRad As String

Dim Nyring As String

Set LetMaerkData = New Collection
On Error GoTo LetMaerkErr
bRubrik = False
omUtland = Utland 'ställs som indatavariabeln

'Set GetRingnr = New Collection ' sätter collectionen

Set rs = New ADODB.Recordset
Set rsA = New ADODB.Recordset

FyndRapp = False 'Fyndrapp sätts om det finns ett fynd där rappring är den sökta ringen

'Urspring = Ringnr
'UrspCentr = Central
LetRing1 = ""
LetRing2 = ""
LetRing3 = ""
LetRing4 = ""
SoektMnr = ""
UrspMnr = ""
En_hittad_ring = ""
En_hittad_Centr = ""

MT1 = ""
MT2 = ""

'Sök först reda på märkarnummer till soekta ringen om post finns
If Utland = False Then
  rs.Open "SELECT mnr,tr, gcentr, gring FROM Ringon WHERE Ring= '" & Ringnr & "' AND Centr= '" & Central & "' AND (TR = '1' OR TR ='M' or TR = 'N' or TR ='L')", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = False Then
    SoektMnr = rs!Mnr
    If rs!Tr = "1" Or rs!Tr = "M" Then UrspMnr = rs!Mnr
  End If
  rs.Close
End If
'Här skall det skiljas på om sökningen sker i URing eller i Ring
'Gör en identisk vy för utlandsmärkningen
Soekigen:


If rs.State = 1 Then rs.Close
If Utland = False Then
  rs.Open "SELECT * FROM Ringon_View WHERE Ring= '" & Ringnr & "' AND Centr= '" & Central & "' ORDER BY Datum,Tr DESC", Conn, adOpenStatic, adLockReadOnly
Else
  rs.Open "SELECT * FROM URing_View WHERE Ring= '" & Ringnr & "' AND Centr = '" & Central & "' ORDER BY Datum, Tr DESC", Conn, adOpenStatic, adLockReadOnly
End If

If rs.BOF = False And rs.EOF = False Then
  'Spara värden på märketext om sådant finns
  'Kolla nu om märketext finns och om denna misstämmer mot märkdata
  Reg_Märketext "U"

  'Om gammal ring finns då är det ej ursprungliga ringen - denna ska sökas
  If IsNull(rs!GRing) = False And Trim(rs!GRing) > "" Then
    GammalRing = rs!GRing
    GammalCentr = rs!GCentr
    En_hittad_ring = rs!GRing
    En_hittad_Centr = rs!GCentr
    


    ' gammal ring finns - den sökta ringen är alltså inte den ursprungliga
    ' sök märkpost för den ursprungliga
    Utland = False
    Select Case rs!GCentr
      Case "SVS", "SVR", "SVO", "SVJ", "SVG", "SVL", "SVN"
      Case Else
        Utland = True
      End Select
    Ringnr = rs!GRing
    'Central = rs!GCentr 'Avmarkerad 2009-04-14 TWR
'     If rs!GRing <> Soektring And rs!Centr = "SVS" Then
    If rs!GRing <> Soektring Then
      If LetRing1 = "" Then  'spara nu ringnumret för ev sökning i F2Kontr längre fram
        LetRing1 = rs!GRing
        GoTo Soekigen
      End If
      If LetRing2 = "" Then
        LetRing2 = rs!GRing
        GoTo Soekigen
      End If
      If LetRing3 = "" Then
        LetRing3 = rs!GRing
        GoTo Soekigen
      End If
      If LetRing4 = "" Then
        LetRing4 = rs!GRing
        GoTo Soekigen
      End If
    End If
    
    'här ska ev in en registrering i rapport-tabellen så att om ingen urspr märkdata hittas
    'ska ändå synas att och var och när den ommärkts

    GoTo Soekigen

  End If
  
  If rs!Ring <> Soektring And (rs!Tr = "1" Or rs!Tr = "M") Then
    If Utland = False Then UrspMnr = rs!Mnr
    frmSok.lblUrspMarkRing.Visible = True
    If Utland = False Then frmSok.lblUrspMarkRing.Caption = "Urspr  " & rs!Centr & " " & rs!Ring & "  (" & rs!Mnr & ") "
    If Utland = True Then frmSok.lblUrspMarkRing.Caption = "Urspr  " & rs!Centr & " " & rs!Ring
    If Utland = False Then If IsNull(rs!enamn) = False Then frmSok.lblUrspMarkRing.Caption = frmSok.lblUrspMarkRing.Caption & RTrim(rs!enamn)
    If Utland = False Then If IsNull(rs!fnamn) = False Then frmSok.lblUrspMarkRing.Caption = frmSok.lblUrspMarkRing.Caption & ", " & RTrim(rs!fnamn)
  End If

  
  'cr.Datafinns = True
  If Utland = False Then
    Ringensid = rs!Ringid
    SoektaRingensId = rs!Ringid
  Else
    Ringensid = rs!Uringid
    SoektaRingensId = rs!Uringid
  End If
  TidLokSort = FixaSort(rs!Datum, rs!Tr, rs!Tnog, rs!Lokal) 'Fixa sorteringsnyckel för denna rad i rapport-tabellen
  
Else 'kolla då om ringen är rappring i fynd
  If FyndRapp = True Then GoTo FixaMaerkData ' då ingen idé söka i fynd igen
  If Utland = False Then
    rsFy.Open "SELECT * FROM fynd_View WHERE Nyring = '" & Soektring & "' AND Nycentr= '" & Central & "' ORDER BY Datum,Tr DESC", Conn, adOpenStatic, adLockReadOnly
  Else
    rsFy.Open "SELECT * FROM ufynd_View WHERE Nyring= '" & Soektring & "' AND Nycentr = '" & Central & "' ORDER BY Datum, Tr DESC", Conn, adOpenStatic, adLockReadOnly
  End If

  If rsFy.EOF = False Then
    FyndRapp = True ' KOLLA HÄR FYNDFINNS
    Ringnr = rsFy!Ring
    Central = rsFy!Centr
    rsFy.Close
    GoTo Soekigen
  End If
End If

FixaMaerkData:

frmSok.lblMaerkData.Visible = True
   
Set cr = New clsRingRc
cr.Rdatum = 0
cr.RTr = 0
cr.Märkdata = "J"

If rs.EOF = False Then 'Märkdata hittad
  If rs!Tr = "1" Or rs!Tr = "M" Then
    Urspring = rs!Ring
    UrspCentr = rs!Centr
  End If
  
  If Utland = False Then
  MarkarNamn = "(" & rs!Mnr & ") "
    If IsNull(rs!enamn) = False Then MarkarNamn = MarkarNamn + RTrim(rs!enamn)
    If IsNull(rs!fnamn) = False Then MarkarNamn = MarkarNamn + ", " + RTrim(rs!fnamn)
    frmSok.lblMaerkare = MarkarNamn
  End If
'  If Soektring = Urspring Then frmSok.lblMaerkare.Caption = "Märkare = " & frmSok.lblMaerkare.Caption
'  If Trim(frmSok.lblMaerkare.Caption) <> "" Then frmSok.lblMaerkare.Visible = True
  
  cr.Basdata = rs!Centr & " " & rs!Ring
  If rs!Ring <> Soektring Then
    frmSok.lblMaerkare.Visible = True
    cr.Basdata = cr.Basdata & " (ursprunglig)"
  End If
  cr.Basdata = cr.Basdata & "   " & rs!Artkod
  If Utland = False Then cr.Basdata = cr.Basdata & "   mnr = " & rs!Mnr & " " & Mid(MarkarNamn, 7)
  LetMaerkData.Add cr
     
  If rs!Ring <> Soektring Then
    Set cr = New clsRingRc
    cr.Rdatum = 0
    cr.RTr = 1
    cr.Märkdata = "J"
    cr.Basdata = "Sökt ring " & Soektring
    If Fyndets_Centr = "" And Central <> "SVS" Then cr.Basdata = cr.Basdata + " central " & Central
    LetMaerkData.Add cr
  End If
  
End If

If rs.EOF = True Then 'Märkdata ej hittad
  Set cr = New clsRingRc
  cr.Rdatum = 0
  cr.RTr = 1
  cr.Märkdata = "J"
  If GammalRing = "" Then
    cr.Basdata = "Ursprunglig märkdata saknas till sökta ringen = " & Central & " " & Soektring
  Else
    If Soektring <> GammalRing Then
      If SoektTR = "7" Or SoektTR = "N" Then cr.Basdata = "Sökt ring = " & Central & " " & Soektring & " har ersatt " & GammalCentr & " " & GammalRing & " märkdata saknas"
      If SoektTR = "4" Or SoektTR = "L" Then cr.Basdata = "Sökt ring = " & Central & " " & Soektring & " har tillagts  " & GammalCentr & " " & GammalRing & " märkdata saknas"
    End If
  End If
  LetMaerkData.Add cr
  Set cr = New clsRingRc
  cr.Rdatum = 0
  cr.RTr = 2
  cr.Märkdata = "J"
  cr.Rubrik = "======================================================================================"
  LetMaerkData.Add cr
  rs.Close
  Set rs = Nothing
  Exit Function
End If

frmSok.lblOrt.Visible = True

Set cr = New clsRingRc
cr.Rdatum = 0
cr.RTr = 2
cr.Märkdata = "J"
cr.Rubrik = "======================================================================================"
LetMaerkData.Add cr

Set cr = New clsRingRc
cr.Rdatum = 0
cr.RTr = 3
cr.Märkdata = "J"
cr.Rubrik = "        tr datum      dn  k åld st va hp r  vinge vikt fett pjm sign fkd"
LetMaerkData.Add cr

Set cr = New clsRingRc
cr.Rdatum = 0
cr.RTr = 4
cr.Märkdata = "J"
cr.Rubrik = "--------------------------------------------------------------------------------------"
LetMaerkData.Add cr

'Gå igenom recordsetet med märkdata - kan finnas separata fångst- och släpprapporter
rs.MoveFirst
Do Until rs.EOF
  
  If rs!Tr = "1" Or rs!Tr = "M" Then
    Urspring = rs!Ring
    UrspCentr = rs!Centr
  End If
  SortStr = FixaSort(rs!Datum, rs!Tr, rs!Tnog, rs!Lokal) 'Fixa sorteringsnyckel för denna rad i rapport-tabellen

  Set cr = New clsRingRc
  Dataraden Utland
  If rs!Tr = "M" Then
    cr.RTr = 10
  Else
    cr.RTr = 0
  End If
  cr.Dnr = "Märkdata"
  If rs!Tr = "0" Then cr.Dnr = "Fångad"
  If rs!Tr = "M" Or rs!Tr = "N" Or rs!Tr = "L" Then cr.Dnr = "Släppt"
  cr.Rdatum = SortStr
  cr.Märkdata = "J"
  LetMaerkData.Add cr

  Set cr = New clsRingRc
  Platsen (rs!Tr)  'Fixa hela strängen med ortsangivelsen
  If rs!Tr = "M" Then
    cr.RTr = 11
  Else
    cr.RTr = 1
  End If
  cr.Rdatum = SortStr
  cr.Märkdata = "J"
  LetMaerkData.Add cr
  
  If RTrim(Ort2) <> "" Then
    Set cr = New clsRingRc
    If rs!Tr = "M" Then
      cr.RTr = 12
    Else
      cr.RTr = 2
    End If
    cr.Ort = Ort2
    cr.Rdatum = SortStr
    cr.Märkdata = "J"
    LetMaerkData.Add cr
  End If
  
  'Finns Kulldata ?
  If Utland = False Then
    If IsNull(rs!Idkull) = False Then
      Set cr = New clsRingRc
      Kulldata
      If rs!Tr = "M" Then
        cr.RTr = 13
      Else
        cr.RTr = 4
      End If
      cr.Rdatum = SortStr
      cr.Märkdata = "J"
      LetMaerkData.Add cr
    End If
  End If
  
  'eventuell extradata till kulldata
  If Utland = False Then
    If IsNull(rs("kKondition")) = False Or IsNull(rs("kTextrad")) = False Or IsNull(rs("kRingtext")) = False Then
      Set cr = New clsRingRc
      If rs!Tr = "M" Then
        cr.RTr = 13
      Else
        cr.RTr = 4
      End If
      KullXData
      cr.Rdatum = SortStr
      cr.Märkdata = "J"
      LetMaerkData.Add cr
    End If
  End If
  
  'Finns extradata till märkposten ?
  If IsNull(rs!kondition) = False Or IsNull(rs!textrad) = False Or IsNull(rs!ringtext) = False Then
    Set cr = New clsRingRc
    XData
    If rs!Tr = "M" Then
      cr.RTr = 14
    Else
      cr.RTr = 5
    End If
    cr.Rdatum = SortStr
    cr.Märkdata = "J"
    LetMaerkData.Add cr
  End If
  
  'Finns mfext-post till märkposten (MFEXT-post är MEXT från fyndbasen i DATAFLEX)
  Svar = rs!Centr + rs!Ring
  If rsA.State = 1 Then rsA.Close
  rsA.Open "SELECT text FROM Mfext WHERE Mfxnyckel= '" & Svar & "'", Conn, adOpenStatic, adLockReadOnly
  If rsA.EOF = False Then
    Set cr = New clsRingRc
    If rs!Tr = "M" Then
      cr.RTr = 15
    Else
      cr.RTr = 6
    End If
    cr.Extradata = "Mext: " & rsA!Text
    cr.Rdatum = SortStr
    cr.Märkdata = "J"
    LetMaerkData.Add cr
  End If
  rsA.Close
  
  'Finns annat märke
  If IsNull(rs!typ1) = False Or IsNull(rs!typ2) = False Then
    Set cr = New clsRingRc
    AndraMaerken
    If rs!Tr = "M" Then
      cr.RTr = 16
    Else
      cr.RTr = 7
    End If
    cr.Rdatum = SortStr
    cr.Märkdata = "J"
    LetMaerkData.Add cr
  End If
  
  'Visa  märkdata på skärmen
  
  Timme = "" 'fixa uttryck för timangivelse om sådan finns
  If IsNull(rs!Tnog) = False And RTrim(rs!Tnog) <> "" Then Timme = "kl " & rs!Tnog

  If rs!Tr = "0" Then
    frmSok.lblTranspHeld.Visible = True
    frmSok.lblTranspHeld = "  Märkdata:"
    frmSok.lblMaerkFangst.Visible = True
    If Soektring <> Urspring Then frmSok.lblTranspHeld = "  Ursprunglig märkdata:  " & GammalCentr & "  " & GammalRing
    frmSok.lblMaerkFangst = "   fångad:  Tr 0  " & rs!Datum & " " & Timme
    If IsNull(rs!Varn) = False And rs!Varn <> "" Then frmSok.lblMaerkFangst = frmSok.lblMaerkFangst & "   varn " & rs!Varn
    frmSok.lblMaerkFangst = frmSok.lblMaerkFangst & "  lokal " & RTrim(Helort)
    If RTrim(Ort2) <> "" Then frmSok.lblMaerkFangst = frmSok.lblMaerkFangst & ", " & Ort2
  Else
    If rs!Tr = "1" Or rs!Tr = "7" Or rs!Tr = "4" Then
      frmSok.lblMaerkData = "  Märkdata:  "
      If Soektring <> Urspring Then frmSok.lblMaerkData = "Ursprunglig märkdata: "
    Else
      frmSok.lblMaerkData = "  släppt  :  "
    End If
    frmSok.lblMaerkData = frmSok.lblMaerkData & UrspCentr & "  " & Urspring & "  Tr " & rs!Tr & "  " & rs!Datum & " " & Timme & "  " & rs!Artkod & "  " & rs!Sex & "  " & rs!Age
    If IsNull(rs!Varn) = False And rs!Varn <> "" Then frmSok.lblMaerkData = frmSok.lblMaerkData & "   varn " & rs!Varn

    If rs!Varn = "4" Or rs!Varn = "7" Or rs!Varn = "9" Then
      MaerkeRad = ""
      If IsNull(rs!text1) = False Then MaerkeRad = RTrim(rs!text1)
      If IsNull(rs!colorkod1) = False Then MaerkeRad = MaerkeRad & " " & RTrim(rs!colorkod1)
      If IsNull(rs!text2) = False Then MaerkeRad = MaerkeRad & " " & RTrim(rs!text2)
      If IsNull(rs!colorkod2) = False Then MaerkeRad = MaerkeRad & " " & RTrim(rs!colorkod2)
      If MaerkeRad <> "" Then frmSok.lblMaerkData = frmSok.lblMaerkData & " =" & " " & MaerkeRad
    End If
      
    frmSok.lblOrt = "  " & VisOrt
  End If
  rs.MoveNext
Loop

Set cr = New clsRingRc
cr.Rdatum = SortStr
cr.RTr = 20
cr.Märkdata = "J"
cr.Rubrik = "======================================================================================"
LetMaerkData.Add cr
Set rsA = Nothing

Exit Function

Resume
LetMaerkErr:
MsgBox Err.Number & Err.Description & " Felet i function LetMaerkData"
Resume Next
End Function

Public Function FixaRingnr(Ringnr As String) As String

Dim ringbeg As String, ringteck As String, ringlab As String
Dim rlen As Byte, lastbok As Byte, ringblank As String, siffend As String
Dim ringbit As String, bitpos As Byte

Ringnr = Replace(Ringnr, " ", "", 1)
Ringnr = UCase(Ringnr)

'On Error GoTo Err_1
bitpos = 1

lastbok = 0

ringbeg = Trim(Ringnr)
rlen = Len(ringbeg)

Do Until bitpos = rlen + 1
  If IsNull(Mid(ringbeg, bitpos, 1)) = False Or Mid(ringbeg, bitpos, 1) > " " Then ringlab = ringlab + Mid(ringbeg, bitpos, 1)
  bitpos = bitpos + 1
Loop

rlen = Len(ringlab)

If rlen = 8 Then
  FixaRingnr = Ringnr
  Exit Function
End If
ringteck = Mid(ringbeg, 1, 1)
If ringteck >= "A" And ringteck <= "Z" Then lastbok = 1
If rlen = 1 Then GoTo fixaring

ringteck = Mid(ringbeg, 2, 1)
If ringteck >= "A" And ringteck <= "Z" Then lastbok = 2
If rlen = 2 Then GoTo fixaring

ringteck = Mid(ringbeg, 3, 1)
If ringteck >= "A" And ringteck <= "Z" Then lastbok = 3

fixaring:
siffend = Right(ringlab, rlen - lastbok)
bitpos = 1

Do Until bitpos = (rlen - lastbok) + 1
  If Mid(siffend, bitpos, 1) >= "0" And Mid(siffend, bitpos, 1) <= "9" Then
    bitpos = bitpos + 1
  Else
    Exit Do
    Ringnr = ""
    Exit Function
  End If
Loop

If lastbok = 0 Then
  Do Until rlen = 8
    ringblank = ringblank + " "
    rlen = rlen + 1
  Loop
ringlab = ringblank + ringlab
End If

If lastbok >= 1 Then
  ringblank = Mid(ringlab, 1, lastbok)
  Do Until rlen = 8
    ringblank = ringblank + " "
    rlen = rlen + 1
  Loop
ringlab = ringblank + siffend
End If
FixaRingnr = ringlab
Exit Function

Resume
Err_1:
Resume Next
End Function
