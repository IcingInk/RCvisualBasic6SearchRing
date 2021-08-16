Attribute VB_Name = "basSok"
Option Explicit
Public conSql As String, SQLUser As String
Public Conn As New ADODB.Connection
Public FyndFinns As Boolean, Sparas As Boolean, UtlandSVS As Boolean
Public VarFrom As String, Dnrut As String
Public rs As New ADODB.Recordset
Public rsA As New ADODB.Recordset
Public Rsr As New ADODB.Recordset
Public RsMt As New ADODB.Recordset
Public SoektaRingensId As Long
Public TidLokSort As String, SoektMnr As String
Public MT1 As String, MT2 As String
Public Soektring As String, SoektCentr As String, SoektTR As String
Public KontrFynd_TR As String
Public En_hittad_ring As String, En_hittad_Centr As String
Public Central As String, UrspCentr As String, Urspring As String
Public Art_Skyddad As Boolean, VisOrt As String
Public GammalRing As String, GammalCentr As String, MaerkeClash As Boolean
Public Fyndets_Ring As String, Fyndets_Centr As String

Enum MTyp
  rcSaknas = 0
  rcFinns = 1
  rcSaknasOmmärkt = 2
  rcFinnsOmmärkt = 3
End Enum

Enum Sorder
  Datarad = 0
  LokalochLinje1 = 1
  KulldataochRubrik = 2
  KullXdataochLinje2 = 3
  NyochRappring = 4
  Fynddetalj = 5
  XData = 6
  Marken = 7
  Rapportoer = 8
  FyndAdmin = 9
  Markdata = 10
  EgfyndTr2 = 20
  EgfyndTransp = 30
  EgfyndTr3 = 40
  Fynd = 50
End Enum

Public Function FixaSort(Dat, Tr, Tnog, Lokal)

FixaSort = CStr(Dat)
If IsNull(Tnog) = False Then
  FixaSort = FixaSort + Tnog
Else
  FixaSort = FixaSort + "00"
End If
FixaSort = FixaSort + Tr + Lokal

End Function

