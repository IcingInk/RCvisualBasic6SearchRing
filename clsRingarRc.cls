VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRingarRc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private cr As clsRingRc
Private rs As ADODB.Recordset
Private rsA As ADODB.Recordset
Private Rx As ADODB.Recordset
Private MarkarNamn As String, UrspMnr As String, FirstRapp As Boolean
Private LetRing1 As String, LetRing2 As String, LetRing3 As String, LetRing4 As String
Private bRubrik As Boolean, BasRad As String, SortStr As String, Posttyp As String
Private Helort As String, Ort As String, Ort2 As String, Timme As String
Private MFXirad As Integer, MFXtr As String, MFXrub As String, MFXtext As String, OlSort As String
Public MFXnyckel As String, FdetHela As String, Fdet1 As String, Fdet2 As String, Fdet3 As String, Fdet4 As String
Private Rapp1 As String, Rapp2 As String, Rapp3 As String, Rapp4 As String, Epost As String
Dim Rapp5 As String, Rapp6 As String, RappRest As String

Public Function LetKontrData(Utland As Boolean) As Collection
Dim Svar As String
Dim Tr As String
Dim omUtland As Boolean 'Används vid sökning av ommärkt fågel
Dim C As Integer 'Används vid For loop vid ommärkning
Dim GRing As String, sTr As String, Ringensid As String
Dim GCentr As String, Soekid As Long
Dim Rapp As String, EkoAvskriven As Boolean

Dim Nyring As String

On Error GoTo LetKontrErr
bRubrik = False
EkoAvskriven = False
omUtland = Utland 'ställs som indatavariabeln

Set LetKontrData = New Collection ' sätter collectionen

Set rs = New ADODB.Recordset
Set rsA = New ADODB.Recordset

If rs.State = 1 Then rs.Close

KontrFynd_TR = ""

'Fältet Dfynd = D = finns som färdigbehandlat fynd och tas därifrån
'MnrFlagga = Y är avskriven kontroll och ska inte tas med
'Sök nu på såväl sökta ringnumret som ursprungliga som oftast är = sökta men som kan skilja om fågeln ommärkts
'samt på ev andra nummer (Letring1-4) som kan förekomma som Ring i Fagel2-poster i Ringon
rs.Open "SELECT * FROM F2kontr_View WHERE ((Ring= '" & Soektring & "' AND Centr= '" & SoektCentr & "') OR (Ring= '" & Urspring & "' AND Centr= '" & UrspCentr & "') OR (Centr = 'SVS' and (Ring='" & LetRing1 & "' OR Ring='" & LetRing2 & "' OR Ring='" & LetRing3 & "' OR Ring='" & LetRing4 & "') OR (Ring= '" & En_hittad_ring & "' AND Centr= '" & En_hittad_Centr & "'))) AND (Dfynd is null OR (Dfynd is not null and Dfynd <> 'D' and dfynd <> 'Z')) ORDER BY Datum, Tr DESC", Conn, adOpenStatic, adLockReadOnly

If rs.EOF = True Then  'Då sök på Rappring i F2kontr
  rs.Close
  rs.Open "SELECT * FROM F2kontr_View WHERE ((RappRing= '" & Soektring & "' AND RappCent= '" & SoektCentr & "') OR (RappRing= '" & Urspring & "' AND RappCent= '" & UrspCentr & "') OR (RappCent = 'SVS' and (RappRing='" & LetRing1 & "' OR RappRing='" & LetRing2 & "' OR RappRing='" & LetRing3 & "' OR RappRing='" & LetRing4 & "') OR (RappRing= '" & En_hittad_ring & "' AND RappCent= '" & En_hittad_Centr & "'))) AND (Dfynd is null OR (Dfynd is not null and Dfynd <> 'D' and dfynd <> 'Z')) ORDER BY Datum, Tr DESC", Conn, adOpenStatic, adLockReadOnly
  If rs.EOF = True Then Exit Function
End If

If rs!mnrflagga = "F" Then
  frmSok.lblFRV.Visible = True
  frmSok.ObserveraSkylt.Visible = True
End If
If rs!mnrflagga = "E" Or rs!mnrflagga = "A" Then
  frmSok.lblEkoFinns.Visible = True
  frmSok.ObserveraSkylt.Visible = True
End If
 
rs.MoveFirst
 
Do Until rs.EOF

 'Kolla nu om märketext finns och om denna misstämmer mot märkdata
  Reg_Märketext "K"

'Fixa först rapportör-rad om ej egen kontroll
Rapp = ""
If rs!mnrflagga = "F" Or rs!mnrflagga = "A" Then
  Rapp = "Rapportör: "
  If rs!mnrflagga = "A" Then Rapp = Rapp & "(associerad märkare) "
  Rapp = Rapp & " " & rs!Mnr
  If IsNull(rs!enamn) = False Then Rapp = Rapp & " " & RTrim(rs!enamn)
  If IsNull(rs!fnamn) = False Then Rapp = Rapp & ", " & RTrim(rs!fnamn)
End If

  SortStr = FixaSort(rs!Datum, rs!Tr, rs!Tnog, rs!Lokal) 'Fixa sorteringsnyckel för denna rad i rapport-tabellen
  
  If rs!mnrflagga = "Y" Or rs!dfynd = "Y" Or rs!Tr = "Y" Then 'Avskriven kontroll
    EkoAvskriven = True
    frmSok.lblEkoAvskr.Visible = True
    frmSok.VarningsSkylt.Visible = True
    Set cr = New clsRingRc
    cr.Extradata = "OBS!  Avskriven Fagel2-kontroll  OBS !"
    cr.Dnr = "F2-post"
    cr.Rdatum = SortStr
    cr.RTr = 0
    LetKontrData.Add cr
  End If
  
  If rs!mnrflagga = "F" And (IsNull(rs!dfynd) = True Or Trim(rs!dfynd) = "") Then 'Obehandlad främmande kontroll
    Set cr = New clsRingRc
    cr.Extradata = "Ring " & rs!Centr & " " & rs!Ring & " Obehandlad främmande Fagel2-post"
    cr.Dnr = "F2-post"
    cr.Rdatum = SortStr
    cr.RTr = 0
    LetKontrData.Add cr
  End If
  
  If rs!mnrflagga <> "F" And rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Or rs!Tr = "7" Or rs!Tr = "4" Then '
    Set cr = New clsRingRc
    If rs!mnrflagga = "E" Then cr.Dnr = "Egkontr"
    If rs!mnrflagga = "A" Then cr.Dnr = "Assoc."
    cr.Rdatum = SortStr
    cr.RTr = 9
    LetKontrData.Add cr
  End If
  
  Set cr = New clsRingRc
  Dataraden Utland
  If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
    cr.RTr = 10
  Else
    cr.RTr = 1
  End If
  If rs!mnrflagga = "E" Then cr.Dnr = "Egkontr"
  If rs!mnrflagga = "A" Then cr.Dnr = "Assoc."
  If rs!Tr = "F" Then cr.Dnr = "Fångad"
  If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then cr.Dnr = "Släppt"
  cr.Rdatum = SortStr
  cr.Fynddetalj = rs("KND") & rs("DT") & rs("DTA")
  cr.Lokal = "    " & rs!Artkod
  LetKontrData.Add cr

  Set cr = New clsRingRc
  Platsen (rs!Tr)  'Fixa hela strängen med ortsangivelsen
  If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
    cr.RTr = 11
  Else
    cr.RTr = 2
  End If
  cr.Rdatum = SortStr
  LetKontrData.Add cr
  
  If RTrim(Ort2) <> "" Then
    Set cr = New clsRingRc
    If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
      cr.Dnr = "Släppt"
      cr.RTr = 12
    Else
      cr.RTr = 3
    End If
    cr.Ort = Ort2
    cr.Rdatum = SortStr
    LetKontrData.Add cr
  End If
  
  'Är denna kontroll också en ommärkning, dvs det har satts på en ny ring
  If rs!Tr = "7" Or rs!Tr = "4" Or rs!Tr = "N" Or rs!Tr = "L" Then
    Set cr = New clsRingRc
    If rsA.State = 1 Then rsA.Close
    'If IsNull(rs!Centr) = False Then rsA.Open "SELECT Centr, Ring FROM Ringon WHERE Gcentr='" & rs!Centr & "' AND Gring= '" & rs!Ring & "'", Conn, adOpenStatic, adLockReadOnly
    'If IsNull(rs!Centr) = True Then rsA.Open "SELECT Centr, Ring FROM Ringon WHERE Gcentr='" & rs!RappCent & "' AND Gring= '" & rs!rappring & "'", Conn, adOpenStatic, adLockReadOnly
    'Feltänkt! För att bli rätt vid upprepade ommärkningar måste alltid NyRing hämtas från märkdata för RappRing
    rsA.Open "SELECT Centr, Ring FROM Ringon WHERE Gcentr='" & rs!RappCent & "' AND Gring= '" & rs!RappRing & "'", Conn, adOpenStatic, adLockReadOnly
    If rsA.EOF = True And InStr(rs!RappRing, "?") > 0 Then
      rsA.Close
      rsA.Open "SELECT Centr, Ring FROM Ringon WHERE Gcentr='" & rs!Centr & "' AND Gring= '" & rs!Ring & "'", Conn, adOpenStatic, adLockReadOnly
    End If
    If rs!Tr = "7" Or rs!Tr = "N" Then cr.Dnr = "Ommärkt"
    If rs!Tr = "4" Or rs!Tr = "L" Then cr.Dnr = "Tillagd"
    
    If rsA.EOF = False Then cr.Extradata = "Ny ring: " & rsA!Centr & " " & rsA!Ring
    If rsA.EOF = True Then cr.Extradata = "nya ringen kan inte hittas av programmet!!"
    cr.RTr = 13
    cr.Rdatum = SortStr
    LetKontrData.Add cr
  End If
 
  'Finns extradata till kontrollposten ?
  If IsNull(rs!kondition) = False Or IsNull(rs!textrad) = False Or IsNull(rs!ringtext) = False Then
    Set cr = New clsRingRc
    XData
    If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
      cr.RTr = 14
    Else
      cr.RTr = 5
    End If
    cr.Rdatum = SortStr
    LetKontrData.Add cr
  End If
  
  'Finns annat märke
  If IsNull(rs!typ1) = False Or IsNull(rs!typ2) = False Then
    Set cr = New clsRingRc
    AndraMaerken
    If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
      cr.RTr = 16
    Else
      cr.RTr = 7
    End If
    cr.Rdatum = SortStr
    LetKontrData.Add cr
  End If
  
  'Finns mfext-post till kontrollposten (MFEXT-post är MEXT från fyndbasen i DATAFLEX)
  Svar = rs!Ring & CStr(rs!Datum)
  If rsA.State = 1 Then rsA.Close
  rsA.Open "SELECT text FROM Mfext WHERE Mfxnyckel= '" & Svar & "'", Conn, adOpenStatic, adLockReadOnly
  If rsA.EOF = False Then
    Set cr = New clsRingRc
    If rs!Tr = "F" Then
      cr.RTr = 15
    Else
      cr.RTr = 6
    End If
    cr.Extradata = "Fext: " & rsA!Text
    cr.Rdatum = SortStr
    LetKontrData.Add cr
  End If
  rsA.Close
  

  If Urspring <> "" And rs!Ring <> Urspring And rs!Ring = Soektring Then
    Set cr = New clsRingRc
    cr.Rdatum = SortStr
    cr.RTr = 20
    cr.Extradata = "Kontrollerade ring " & rs!Centr & " " & rs!Ring & " är ej den ursprungliga"
    LetKontrData.Add cr
  End If

  If Rapp <> "" Then
    Set cr = New clsRingRc
    cr.Rdatum = SortStr
    cr.RTr = 21
    cr.Extradata = Rapp
    LetKontrData.Add cr
  End If

  Set cr = New clsRingRc
  cr.Rdatum = SortStr
  cr.RTr = 22
  cr.Rubrik = "--------------------------------------------------------------------------------------"
  LetKontrData.Add cr

  rs.MoveNext
Loop

If EkoAvskriven = True Then
  Set cr = New clsRingRc
  cr.Rdatum = SortStr
  cr.RTr = 98
  cr.Rubrik = "OBS!!  Det finns en Fagel2-kontrollpost med denna ring som avskrivits - se ovan !!"
  frmSok.VarningsSkylt.Visible = True
  LetKontrData.Add cr

  Set cr = New clsRingRc
  cr.Rdatum = SortStr
  cr.RTr = 99
  cr.Rubrik = "--------------------------------------------------------------------------------------"
  LetKontrData.Add cr
End If


Set rs = Nothing
Set rsA = Nothing

Exit Function

Resume
LetKontrErr:
If Err.Number = 3265 Then Resume Next
MsgBox Err.Number & Err.Description & " Felet i function LetKontrData"
If Left(Right(Err.Description, 35), 19) = "Invalid column name" Then Exit Function
Resume Next
End Function
Private Sub Reg_Märketext(Posttyp As String)
MT1 = ""
MT2 = ""
If IsNull(rs!text1) = False And Trim(rs!text1) <> "" Then MT1 = rs!text1
If IsNull(rs!text2) = False And Trim(rs!text2) <> "" Then MT2 = rs!text2
If MT1 <> "" Or MT2 <> "" Then
  If RsMt.State = 1 Then RsMt.Close
  RsMt.Open "SELECT * FROM SökRingMtext WHERE mtext1= '" & MT1 & "' or mtext2= '" & MT1 & "' or mtext1= '" & MT2 & "' or mtext2= '" & MT2 & "'", Conn, adOpenStatic, adLockReadOnly
  If RsMt.EOF = True And RsMt.BOF = True Then
    Conn.Execute ("INSERT INTO SökRingMtext VALUES ('" & MT1 & "','" & MT2 & "','" & Posttyp & "','')")
    Conn.Execute ("UPDATE SökRingMtext SET Mtext1='--' WHERE mtext1 is null or mtext1 = ''")
    Conn.Execute ("UPDATE SökRingMtext SET Mtext2='--' WHERE mtext2 is null or mtext2 = ''")
  End If
End If

End Sub

Public Function LetFyndData(Utland As Boolean) As Collection
Dim rsA As New ADODB.Recordset
Dim OldRappnyckel As String, OldRapptyp As String, OldMom As String, OldBdat As String
Dim Olddnr As String, OldDeldnr As String
On Error GoTo LetFyndErr
Set LetFyndData = New Collection ' sätter collectionen

Set rs = New ADODB.Recordset
Set rsA = New ADODB.Recordset

'Här skall det skiljas på om sökningen sker i Fynd eller i Ufynd

If rs.State = 1 Then rs.Close
If Utland = False Then
  If Urspring = "" Then
    Conn.Fynd_Soek SoektCentr, Soektring, rs
  Else
    Conn.Fynd_Soek UrspCentr, Urspring, rs
  End If
Else
  If Urspring = "" Then
    Conn.UFynd_Soek SoektCentr, Soektring, rs
  Else
    Conn.UFynd_Soek UrspCentr, Urspring, rs
  End If
End If

If rs.EOF = False Then GoTo FyndaVidare 'Fynd hittat

If rs.State = 1 Then rs.Close

'Kolla nu om ringen finns som nyring i FYND eller UFYND
If Urspring = "" Then
  Conn.Fynd_Soek_Nyring SoektCentr, Soektring, rs
Else
  Conn.Fynd_Soek_Nyring UrspCentr, Urspring, rs
End If
If rs.EOF = False Then
  Fyndets_Ring = rs!Ring
  Fyndets_Centr = rs!Centr
  GoTo FyndaVidare 'Fynd hittat
End If

If rs.State = 1 Then rs.Close
If Urspring = "" Then
  Conn.UFynd_Soek_Nyring SoektCentr, Soektring, rs
Else
  Conn.UFynd_Soek_Nyring UrspCentr, Urspring, rs
End If
If rs.EOF = False Then
  Fyndets_Ring = rs!Ring
  Fyndets_Centr = rs!Centr
  GoTo FyndaVidare 'Fynd hittat
End If

If rs.State = 1 Then rs.Close
Set rs = Nothing
Set rsA = Nothing
Exit Function

FyndaVidare:
frmSok.lblFyndFinns.Caption = "Fynd finns!"
frmSok.lblFyndFinns.Visible = True
frmSok.ObserveraSkylt.Visible = True

Do Until rs.EOF
  'Kolla nu om märketext finns och om denna misstämmer mot märkdata
  Reg_Märketext "F"
  
  If rs.EOF Then
    GoTo FySlut
  End If
  
  'om nu detta inte är första fyndposten ska kollas om den har samma dnr-början som föregående
  'om det inte är samma rapport då ska rapportör, ev brevdataum och slutrad skrivas ut
  If Left(rs!Dnr, 9) <> Olddnr And Olddnr <> "" Then Reporter OldRappnyckel, OldRapptyp, OldMom, OldBdat, OldDeldnr
  
  SortStr = FixaSort(rs!Datum, rs!Tr, rs!Tnog, rs!Lokal) 'Fixa sorteringsnyckel för denna rad i rapport-tabellen
  
  ' Tr=Y eller Artkod=AVSKR betyder avskrivet fynd
  If rs!Artkod = "AVSKR" Or rs!Tr = "Y" Then
    Set cr = New clsRingRc
    cr.Rdatum = 0
    cr.RTr = 0
    cr.Basdata = "OBS! Detta fynd " & rs!Dnr & " " & " har avskrivits !!"
    LetFyndData.Add cr
    frmSok.lblAvskr.Visible = True
    frmSok.VarningsSkylt.Visible = True
  End If
  
  'Tänd label som berättar att slutfynd finns - antingen är TR = 3 eller så är TR =2 och samtidigt VR = 1 (ringen insänd)
  If rs!Tr = "3" Or (rs!Tr = "2" And rs!vr = "1") Then
    frmSok.lblFyndFinns.Caption = "Slutfynd finns!"
    frmSok.ObserveraSkylt.Visible = True
  End If
  
  If rs!Tr = "F" Then   'en extrarad i början med Dnr eftersom ordte Fångad sk in i dnr-fältet på dataraden
    Set cr = New clsRingRc
    cr.Dnr = Right(Left(rs!Dnr, 9), 7)
    cr.RTr = 0
    cr.Rdatum = SortStr
    LetFyndData.Add cr
  End If
  
  If Right(rs!Dnr, 2) <> "01" Then  'då ska en extrarad som avdelare in med diarienumrets sista två tecken
    Set cr = New clsRingRc
    cr.Dnr = "-" & Right(rs!Dnr, 2) & "-"
    cr.RTr = 0
    cr.Rdatum = SortStr
    LetFyndData.Add cr
  End If
  
  
  Set cr = New clsRingRc   'Dataraden
  Dataraden Utland         'subr Dataraden fixar allt som ska stå där
  If rs!Tr = "N" Or rs!Tr = "L" Then
    cr.RTr = 12
  Else
    cr.RTr = 2
  End If
  cr.Dnr = Right(Left(rs!Dnr, 9), 7)
  If rs!Tr = "F" Then cr.Dnr = "Fångad"
  cr.Fynddetalj = rs("KND") & rs("DT") & rs("DTA")
  cr.Rdatum = SortStr
  If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then cr.Dnr = "Släppt"
  If rs!Tr = "F" Then cr.Dnr = "Fångad"
  cr.Lokal = "    " & rs!RappArt
  LetFyndData.Add cr
  
  Set cr = New clsRingRc  'Raden med ortsangivelsen
  Platsen (rs!Tr) 'Subr Platsen fixar hela strängen med ortsangivelsen
  If rs!Tr = "N" Or rs!Tr = "L" Then
    cr.RTr = 13
  Else
      cr.RTr = 3
  End If
  cr.Rdatum = SortStr
  LetFyndData.Add cr
  
  If RTrim(Ort2) <> "" Then
    Set cr = New clsRingRc  'en till rad behövs för att få rum med ortsangivlesen
    If rs!Tr = "N" Or rs!Tr = "L" Then
      cr.RTr = 14
    Else
      cr.RTr = 4
    End If
    cr.Ort = Ort2
    cr.Rdatum = SortStr
    LetFyndData.Add cr
  End If
  
  If rs!RappRing <> Urspring Then  'då ska rad ut med uppgift om vilken rapporterade ringen är
    Set cr = New clsRingRc
    If rs!Tr = "N" Or rs!Tr = "L" Then
      cr.RTr = 15
    Else
      cr.RTr = 5
    End If
    cr.Rdatum = SortStr
    cr.Extradata = "Rapporterad ring= " & rs!RappCent & " " & rs!RappRing
    LetFyndData.Add cr
  End If
  
  'Finns extradata till kontrollposten ?
  If IsNull(rs!kondition) = False Or IsNull(rs!textrad) = False Or IsNull(rs!ringtext) = False Then
    Set cr = New clsRingRc
    XData                 'subr Xdata fixar raden med uppgifter ur tabellen Xdata
    If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
      cr.RTr = 16
    Else
      cr.RTr = 6
    End If
    cr.Rdatum = SortStr
    LetFyndData.Add cr
  End If
  
  'Finns mfext-post till märkposten (MFEXT-post är FEXT från fyndbasen i DATAFLEX)
  If rsA.State = 1 Then rsA.Close
  rsA.Open "SELECT text FROM Mfext WHERE Mfxnyckel= '" & rs!Dnr & "'", Conn, adOpenStatic, adLockReadOnly
  If rsA.EOF = False Then
    Set cr = New clsRingRc
    If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
      cr.RTr = 17
    Else
      cr.RTr = 7
    End If
    cr.Extradata = "Fext: " & rsA!Text
    cr.Rdatum = SortStr
    LetFyndData.Add cr
  End If
  rsA.Close
  
  
  'Är detta fynd också en ommärkning, dvs det har satts på en ny ring
  If rs!Tr = "7" Or rs!Tr = "4" Or rs!Tr = "N" Or rs!Tr = "L" Or (rs!Tr = "F" And IsNull(rs!Nyring) = False) Then
     Dim TR47 As String
     TR47 = rs!Tr
     Set cr = New clsRingRc
    'Om tr = F då kan det finnas en separat släpp-post där rätt TR-kod finns
    'MEN! det kan likaväl vara så betr fynd importerade från DATA FLEX att det inte finns en
    'separat släpp-post utan släppet anges i en mfext-post. Då vet nu inte programmet vilken
    'TR-kod det ska vara
    If rs!Tr = "F" Then
      TR47 = "F"

      Set Rx = New ADODB.Recordset
      Rx.Open "SELECT tr FROM Fynd WHERE left(dnr,9)= '" & Left(rs!Dnr, 9) & "' and dnr <> '" & rs!Dnr & "' and (tr = 'N' or tr = 'L') and datum > '" & rs!Datum & "'", Conn, adOpenStatic, adLockReadOnly
      If Rx.EOF = True And Rx.BOF = True Then
      Else
        'det finns en sådan post - då kommer uppgifterna om ny ring ut ändå med denna post
        Rx.Close
        Set Rx = Nothing
        GoTo NyraKlar
      End If
      If Rx.State = 1 Then Rx.Close
      Set Rx = Nothing
    End If

    If TR47 = "7" Or TR47 = "N" Then cr.Dnr = "Ommärkt"
    If TR47 = "4" Or TR47 = "L" Then cr.Dnr = "Tillagd"
    If TR47 = "F" Then
       cr.Dnr = "Om/Till?"
    End If
    If IsNull(rs!nycentr) = False Then
      cr.Extradata = "Ny ring: " & rs!nycentr & " " & rs!Nyring
      If TR47 = "F" Then cr.Extradata = cr.Extradata & " (progr kan inte avgöra om byte eller tillägg!)"
    Else
      cr.Extradata = "Ny ring ska ha påsatts enl TR-koden men uppgift om nummer saknas"
    End If
    cr.RTr = 13
    cr.Rdatum = SortStr
    LetFyndData.Add cr
  End If
 
NyraKlar:
    
  'Finns annat märke
  If IsNull(rs!typ1) = False Or IsNull(rs!typ2) = False Then
    Set cr = New clsRingRc
    AndraMaerken
    If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
      cr.RTr = 18
    Else
      cr.RTr = 8
    End If
    cr.Rdatum = SortStr
    LetFyndData.Add cr
  End If

    
  'Finns fynddetaljer i klartext i tabellen Fynddet?
  If rsA.State = 1 Then rsA.Close
  rsA.Open "SELECT fdet, fdetutl FROM fynddet WHERE dnr= '" & rs!Dnr & "'", Conn, adOpenStatic, adLockReadOnly
  If rsA.EOF = False Then
    FdetHela = ""
    If IsNull(rsA!fdet) = False And RTrim(rsA!fdet) <> "" Then FdetHela = rsA!fdet
    If IsNull(rsA!fdetutl) = False And RTrim(rsA!fdetutl) <> "" Then
      If FdetHela = "" Then
        FdetHela = rsA!fdetutl
      Else
        FdetHela = FdetHela & "$" & rsA!fdetutl
      End If
    End If
    
    Fixa_Fdet   'subr som fixar innehållet till ett antal variabler beroende
                'på hur mycket som står i rsa!Fdet
  
    If Fdet1 <> "" Then
      Set cr = New clsRingRc
      If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
        cr.RTr = 19
      Else
        cr.RTr = 9
      End If
      cr.Extradata = "Fd:" & Fdet1
      cr.Rdatum = SortStr
      LetFyndData.Add cr
    End If
    If Fdet2 <> "" Then
      Set cr = New clsRingRc
      If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
        cr.RTr = 20
      Else
        cr.RTr = 10
      End If
      cr.Extradata = "   " & Fdet2
      cr.Rdatum = SortStr
      LetFyndData.Add cr
    End If
    If Fdet3 <> "" Then
      Set cr = New clsRingRc
      If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
        cr.RTr = 21
      Else
        cr.RTr = 11
      End If
      cr.Extradata = "   " & Fdet3
      cr.Rdatum = SortStr
      LetFyndData.Add cr
    End If
    If Fdet4 <> "" Then
      Set cr = New clsRingRc
      If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
        cr.RTr = 22
      Else
        cr.RTr = 12
      End If
      cr.Extradata = "   " & Fdet4
      cr.Rdatum = SortStr
      LetFyndData.Add cr
    End If
    
  End If
  rsA.Close
  
      
  'Finns Accnr eller Cflag
  If IsNull(rs!Accnr) = False And RTrim(rs!Accnr) <> "" Then
    Set cr = New clsRingRc
    cr.Extradata = ""
    If IsNull(rs!Accnr) = False Then cr.Extradata = "Accnr:  " & rs!Accnr
    'If IsNull(rs!Cflag) = False Then cr.Extradata = cr.Extradata & "  Cflag = " & rs!Cflag
    If rs!Tr = "K" Or rs!Tr = "N" Or rs!Tr = "L" Then
      cr.RTr = 23
    Else
      cr.RTr = 13
    End If
    cr.Rdatum = SortStr
    LetFyndData.Add cr
  End If


  'Sätt oldvariablerna för att kunna skriva ut rapportör ev brevdataum och slutrad för dnr:et

  OlSort = SortStr
  Olddnr = Left(rs!Dnr, 9)
  OldDeldnr = Left(rs!Dnr, 9)
  OldRapptyp = rs!Rapptyp
  OldRappnyckel = rs!Rappnyckel
  OldMom = ""
  If IsNull(rs!Mom) = False Then OldMom = RTrim(rs!Mom)
  OldBdat = ""
  If IsNull(rs!Bdat) = False Then OldBdat = RTrim(CStr(rs!Bdat))
 
  rs.MoveNext
  
Loop

FySlut:
'skriv ut rapportör, ev brevdatum och slutlinje för sista fyndposten
Reporter OldRappnyckel, OldRapptyp, OldMom, OldBdat, OldDeldnr


Set rs = Nothing
Set rsA = Nothing

Exit Function

Resume
LetFyndErr:
MsgBox Err.Number & Err.Description & " Felet i function LetFyndData"
Resume Next
End Function

Public Function LetMaerkData(Ringnr As String, Central As String, Utland As Boolean) As Collection
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

Private Sub Dataraden(Utland As Boolean)

cr.Tr = rs("Tr")
cr.Datum = CStr(rs("Datum"))
If IsNull(rs("TNog")) = False Then cr.Tnog = rs("TNog")
If IsNull(rs("Sex")) = False Then cr.Sex = rs("Sex") Else cr.Sex = "-"
If IsNull(rs("Age")) = False Then cr.Age = rs("Age") Else cr.Age = "--"
If IsNull(rs("Status")) = False Then cr.Status = rs("Status") Else cr.Status = "-"
If IsNull(rs("Varn")) = False Then cr.Varn = rs("Varn") Else cr.Varn = "-"
If Utland = False Then
  If IsNull(rs("Pullus")) = False Then cr.Pullus = rs("Pullus") Else cr.Pullus = "-"
  If IsNull(rs("Reserv")) = False Then cr.Reserv = rs("Reserv") Else cr.Reserv = "-"
  If IsNull(rs("Signatur")) = False Then cr.Signatur = rs("Signatur") Else cr.Signatur = "-"
End If
If rs("Vinge") > 0 Then cr.Vinge = CStr(rs("Vinge")) Else cr.Vinge = "-"
If rs("Vikt") > 0 Then cr.Vikt = CStr(rs("Vikt")) Else cr.Vikt = "-"
'If CStr(rs("Vikt")) = "0.0" Then cR.Vikt = "-"
If rs("Fett") < 10 Then cr.Fett = CStr(rs("Fett")) Else cr.Fett = "-"
If rs("Pjm") < 10 Then cr.Pjm = CStr(rs("Pjm")) Else cr.Pjm = "-"

End Sub

Private Sub Platsen(Tr As String)
On Error GoTo PlatsErr

Helort = ""
Ort = ""
Ort2 = ""

If Tr = "1" Or Tr = "0" Or Tr = "M" Or Tr = "N" Or Tr = "L" Or Tr = "4" Or Tr = "7" Then
  'Kolla först om arten är skyddvärd - i så fall får koordinat och ortsnamn ej visas
  Set rsA = New Recordset

  rsA.Open "SELECT Skyddad FROM Artlista WHERE Artkod= '" & rs!Artkod & "'", Conn, adOpenStatic, adLockReadOnly
  If rsA.EOF = False Then
    If IsNull(rsA!Skyddad) = False Then
      VisOrt = "Skyddad art endast provinskoden visas: " & rs!Prov
    End If
  End If
  rsA.Close
  'Exit Sub
End If

If IsNull(rs("Greenwich")) = False And IsNull(rs("ruta")) = True Then
  Helort = rs("Greenwich")
End If

If IsNull(rs("Ruta")) = False Then
  Helort = rs("Ruta")
  If IsNull(rs("Greenwich")) = False Then Helort = Helort + "  " + rs("greenwich")
End If
 
If IsNull(rs("knog")) = False Then Helort = Helort + "  " + rs("knog")
If IsNull(rs("Prov")) = False Then Helort = Helort + "  " + rs("Prov")

If IsNull(rs("Storort")) = False And IsNull(rs("Litenort")) = True Then
  Ort = Ort + " " + rs("Storort")
ElseIf IsNull(rs("Litenort")) = False And IsNull(rs("Storort")) = True Then
  Ort = Ort + " " + rs("Litenort")
Else
  Ort = Ort + " " + rs("Storort") & ", " & rs("Litenort")
End If

If IsNull(rs("Kommentar")) = False Then Ort = Trim(Ort) & " #" & Trim(rs("kommentar")) & "#"

VisOrt = Helort + " " + Ort
If Len(Helort + Ort) > 65 Then
  Ort2 = Ort
Else
  Helort = Helort + " " + Ort
End If

cr.Ort = Helort
   
cr.Lokal = "-"
If rs("Secure") = "R" Then cr.Lokal = "+"
'If Left(rs("lokal"), 1) = "R" Then
  cr.Lokal = cr.Lokal & rs("lokal")
'Else
'  cr.Lokal = cr.Lokal & Left(rs("lokal"), 4) & " " & Mid(rs("Lokal"), 5, 4)
'End If

Exit Sub

Resume
PlatsErr:
MsgBox "Fel: " & Err.Number & " " & Err.Description & " felet i subr Platsen (modul cslRingar", vbOKOnly
Resume Next
End Sub

Private Sub XData()

cr.Extradata = "Extradata: "
If IsNull(rs("Kondition")) = False Then cr.Extradata = cr.Extradata & rs("Kondition") & " "
If IsNull(rs("Textrad")) = False Then cr.Extradata = cr.Extradata & rs("Textrad") & " "
If IsNull(rs("Ringtext")) = False Then cr.Extradata = cr.Extradata & rs("Ringtext")

End Sub

Private Sub AndraMaerken()

cr.Extradata = "Märken: "
If IsNull(rs("Typ1")) = False Then cr.Extradata = cr.Extradata & rs("Typ1") & " "
If IsNull(rs("Text1")) = False Then cr.Extradata = cr.Extradata & rs("Text1") & " "
If IsNull(rs("Colorkod1")) = False Then cr.Extradata = cr.Extradata & rs("Colorkod1") & " "
If IsNull(rs("Typ2")) = False Then cr.Extradata = cr.Extradata & " / " & rs("Typ2") & " "
If IsNull(rs("Text2")) = False Then cr.Extradata = cr.Extradata & rs("Text2") & " "
If IsNull(rs("Colorkod2")) = False Then cr.Extradata = cr.Extradata & rs("Colorkod2") & " "
      
End Sub

Private Sub Kulldata()
'kulldata endast på nymärkningar Tr =1

cr.RTr = Sorder.Markdata + Sorder.KulldataochRubrik
If rs("age") = "00" Then
  cr.Extradata = "I kull: "
Else
  cr.Extradata = "Har kull: "
End If
If IsNull(rs("alive")) = False Then
  cr.Extradata = cr.Extradata & rs("Alive") & " levande"
  If IsNull(rs("kPullus")) = False Then cr.Extradata = cr.Extradata & ", utv. " & rs("kpullus")
  cr.Extradata = cr.Extradata & " (döda " & rs("Dead") & ", ägg " & rs("Eggs") & "), id=" & Right(rs("idkull"), 7)
  cr.Idkull = rs("Idkull")
  Exit Sub
End If
If Left(rs("idkull"), 2) = "RC" Then cr.Extradata = cr.Extradata & Right(RTrim(rs("idkull")), 2) & " levande " & " (döda  --, ägg --)"

End Sub

Private Sub KullXData()
cr.RTr = Sorder.Markdata + Sorder.KullXdataochLinje2
If IsNull(rs("kKondition")) = False Then cr.Extradata = cr.Extradata & rs("kKondition") & " "
If IsNull(rs("kTextrad")) = False Then cr.Extradata = cr.Extradata & rs("kTextrad") & " "
If IsNull(rs("kRingtext")) = False Then cr.Extradata = cr.Extradata & rs("kRingtext")
End Sub
Private Sub URingData(iTr As Integer)
       
'cR.RDatum = CLng(rs("Datum"))
TidLokSort = FixaSort(rs!Datum, rs!Tr, rs!Tnog, rs!Lokal) 'Fixa sorteringsnyckel för denna rad i rapport-tabellen
            
cr.Rdatum = TidLokSort

Select Case iTr
  Case 0
    cr.RTr = Sorder.FyndAdmin
  Case 1
    cr.RTr = Sorder.Markdata + Sorder.FyndAdmin
End Select

If IsNull(rs("IDat")) = False Then cr.Extradata = "Inmatad den " & CStr(rs("IDat")) & " - "
If IsNull(rs("UDat")) = False Then cr.Extradata = cr.Extradata & "ändrad den " & CStr(rs("UDat"))
End Sub

Function RCLok(Kod As String, Mnr As String) As Boolean 'Används inte längre

'gör om så att den tar lokalkoden i skicket mnr & lokal
If Kod = "FBO" And Mnr = "002" Then RCLok = True
If Kod = "FYR" And Mnr = "002" Then RCLok = True
If Kod = "FLO" And Mnr = "002" Then RCLok = True
If Kod = "NAB" And Mnr = "002" Then RCLok = True
If Kod = "MÅK" And Mnr = "002" Then RCLok = True
If Kod = "SKA" And Mnr = "002" Then RCLok = True
If Kod = "GET" And Mnr = "005" Then RCLok = True
If Kod = "NID" And Mnr = "016" Then RCLok = True
If Kod = "KVI" And Mnr = "006" Then RCLok = True
If Kod = "ÄSS" And Mnr = "006" Then RCLok = True
If Kod = "UTK" And Mnr = "007" Then RCLok = True
If Kod = "TOR" And Mnr = "007" Then RCLok = True
If Kod = "HEN" And Mnr = "009" Then RCLok = True
If Kod = "EGG" And Mnr = "018" Then RCLok = True
If Kod = "FJÄ" And Mnr = "020" Then RCLok = True
If Kod = "LAN" And Mnr = "223" Then RCLok = True
If Kod = "PIL" And Mnr = "050" Then RCLok = True
If Kod = "HAM" And Mnr = "061" Then RCLok = True
If Kod = "SLÄ" And Mnr = "284" Then RCLok = True
If Kod = "SKO" And Mnr = "349" Then RCLok = True
If Kod = "HOB" And Mnr = "349" Then RCLok = True
If Kod = "HSR" And Mnr = "420" Then RCLok = True
If Kod = "SVK" And Mnr = "284" Then RCLok = True
If Kod = "ALM" And Mnr = "017" Then RCLok = True
If Kod = "FLI" And Mnr = "017" Then RCLok = True
If Kod = "STAV" And Mnr = "245" Then RCLok = True
End Function

Public Function Leverans(Ringnr As String) As Collection 'Collection = samling av variabeluppsättningar

Dim bLag As Boolean
Set Leverans = New Collection ' sätter collectionen
Set cr = New clsRingRc
Set rs = New ADODB.Recordset

If rs.State = 1 Then rs.Close
rs.Open "SELECT TOP 5 * FROM dbo.Ringmnr_View WHERE TNum >= '" & Ringnr & "' ORDER BY TNum", Conn, adOpenForwardOnly, adLockReadOnly
'om delar av serien levererats, dvs om det finns mer än en post i rs.
  
If rs.BOF = True And rs.EOF = True Then
  'om det inte finns leveransuppgifter
  GoTo SlutLev
  
  'sätt Saknas som standard, om det inte blir någon match
  'i loopen så behålls variabeln annars byts den ut.
Else
  rs.MoveFirst
  Do Until rs.EOF
    cr.Datafinns = MTyp.rcSaknas
    'testar matchning
    If Ringnr >= rs("FNum") Then 'det är rätt serie
      Select Case rs("Mnr")
        Case "LEV"
          frmSok.RingTillRC.Visible = True
          frmSok.RingTillRC = "  Till Rc: " & rs("FNum") & " - " & rs("TNum")
          If IsNull(rs("dat")) = False Then frmSok.RingTillRC = frmSok.RingTillRC & " " & CStr(rs("Dat"))
          cr.Datafinns = MTyp.rcFinns
        Case "VAK"
          frmSok.LeveransUppg = "  Sökta ringen är registrerad som vakant " & rs("FNum") & " - " & rs("TNum")
          If IsNull(rs("dat")) = False Then frmSok.LeveransUppg = frmSok.LeveransUppg & " " & CStr(rs("Dat"))
          cr.Datafinns = MTyp.rcFinns
          rs.Close
          Leverans.Add cr
          Exit Function
        Case "LAG"
          frmSok.LeveransUppg = "  Sökta ringen finns kvar i lagret " & rs("FNum") & " - " & rs("TNum")
          If IsNull(rs("dat")) = False Then frmSok.LeveransUppg = frmSok.LeveransUppg & " " & CStr(rs("Dat"))
          If rs("mat") = "S" Then
            frmSok.LeveransUppg = frmSok.LeveransUppg & " stål"
          Else
            frmSok.LeveransUppg = frmSok.LeveransUppg & " alum"
          End If
            bLag = True
            cr.Datafinns = MTyp.rcFinns
        Case Else
          If IsNull(rs("Ukod")) = False And rs("Ukod") <> "I" Then
            Select Case rs("Ukod")
              Case "R"
                frmSok.RingTillMaerkare.Visible = True
                If IsNull(rs("FNamn")) = False Then
                  frmSok.RingTillMaerkare = "  I returnerad serie: " & rs("Fnum") & " - " & rs("Tnum") & " från " & rs("Mnr") & " " & Trim(rs("FNamn")) & " " & Trim(rs("ENamn"))
                Else
                  frmSok.RingTillMaerkare = "  I returnerad serie: " & rs("Fnum") & " - " & rs("Tnum") & " från " & rs("Mnr") & " " & rs("ENamn")
                End If
              Case "X"
                frmSok.RingTillMaerkare.Visible = True
                If IsNull(rs("FNamn")) = False Then
                  frmSok.RingTillMaerkare = "  Anmäld förlorad " & rs("dat") & "  - ringarna: " & rs("Fnum") & " - " & rs("Tnum") & " av " & rs("Mnr") & " " & Trim(rs("FNamn")) & " " & Trim(rs("ENamn"))
                Else
                  frmSok.RingTillMaerkare = "  Anmäld förlorad " & rs("dat") & "  - ringarna: " & rs("Fnum") & " - " & rs("Tnum") & " av " & rs("Mnr") & " " & rs("ENamn")
                End If
            End Select
            
            cr.Datafinns = MTyp.rcFinns
            If bLag = True Then
              If rs.State = 1 Then rs.Close
              Leverans.Add cr
              Exit Function
            End If
          Else
            cr.Datafinns = MTyp.rcFinns
            cr.Mnr = rs("Mnr")
            
            If RTrim(UrspMnr) <> "" And UrspMnr <> rs!Mnr And rs!ukod <> "R" Then
              frmSok.MarkTillclash.Visible = True
              frmSok.MarkTillclash.Caption = "OBS! Urspr. märkaren " & SoektMnr & " är ej samme som ringen registrerats på >  " & rs!Mnr & " - utred detta!"
              frmSok.VarningsSkylt.Visible = True
            End If
            frmSok.lblMaerkare = "Sökt ring  ("
            If Soektring = Urspring Then frmSok.lblMaerkare = "  ("
            frmSok.LeveransUppg = ""
            If Ringnr = GammalRing Then frmSok.lblMaerkare = "Utbytt ring   " & GammalRing & "  ("
            If Ringnr = GammalRing Then frmSok.LeveransUppg = GammalRing
            If IsNull(rs("FNamn")) = False Then
              frmSok.lblMaerkare = frmSok.lblMaerkare & RTrim(rs!Mnr) & ") " & RTrim(rs!enamn) & ", " & RTrim(rs!fnamn)
              frmSok.LeveransUppg = frmSok.LeveransUppg & " Sänd till " & rs("MNR") & "  " & Trim(rs("FNamn")) & " " & Trim(rs("ENamn")) & " i serien  " & rs("Fnum") & " - " & rs("TNUM")
              If IsNull(rs("dat")) = False Then frmSok.LeveransUppg = frmSok.LeveransUppg & "  " & CStr(rs("Dat"))
            Else
              frmSok.lblMaerkare = frmSok.lblMaerkare & RTrim(rs!Mnr) & ") " & RTrim(rs!enamn)
              frmSok.LeveransUppg = frmSok.LeveransUppg & " Sänd till " & rs("MNR") & "  " & Trim(rs("ENamn")) & " i serien  " & rs("Fnum") & " - " & rs("TNUM")
              If IsNull(rs("dat")) = False Then frmSok.LeveransUppg = frmSok.LeveransUppg & "  " & CStr(rs("Dat"))
            End If
            If IsNull(rs("lagkod")) = False And rs("lagkod") = "N" Then frmSok.LeveransUppg = frmSok.LeveransUppg & "  Nya"
            If IsNull(rs("lagkod")) = False And rs("lagkod") = "G" Then frmSok.LeveransUppg = frmSok.LeveransUppg & "  Gamla"
            If IsNull(rs("Fagel2")) = False And rs("Fagel2") = "J" Then frmSok.lblMaerkare = frmSok.lblMaerkare & "  (Fagel2)"
            Leverans.Add cr
          End If
      End Select
    End If
    rs.MoveNext
  Loop
End If

SlutLev:

Leverans.Add cr

If RTrim(frmSok.LeveransUppg) = "" Then
  If RTrim(frmSok.lblMaerkData) = "" Then
    frmSok.LeveransUppg = "  Uppgift om denna ring saknas i ring-registret"
  Else
    If frmSok.lblMaerkare <> "" Then
      frmSok.LeveransUppg = " Märkare: " & RTrim(frmSok.lblMaerkare) & " /  Ringen saknas i ringregistret"
    End If
  End If
End If
If frmSok.LeveransUppg <> "" Then frmSok.LeveransUppg.Visible = True
If rs.State = 1 Then rs.Close
Set rs = Nothing
End Function

Public Function NarliggandeNummer(Ringnr As String, bMnr As Boolean, Optional Mnr As String = "") As Collection
Dim bSaknas As Boolean

Set rs = New ADODB.Recordset

Set NarliggandeNummer = New Collection

If bMnr = False Then 'om märkaren är okänd
  'rs.Open "SELECT Ring, Datum, Mnr, ENamn, FNamn FROM Ring_View WHERE SUBSTRING(Ring,1,5)= '" & Mid(Ringnr, 1, 5) & "' ORDER BY Ring, Tr", Conn, adOpenStatic, adLockReadOnly
  rs.Open "SELECT R.Centr,R.Ring,R.Tr,R.Datum,R.Lokal,R.Artkod,R.Sex,R.Age,R.Status,R.Varn,R.Mnr,M.ENamn,M.FNamn,L.Ruta,L.Greenwich,L.Prov,L.Storort,L.Litenort,L.Kommentar FROM Ringon R INNER JOIN Maerkare M ON R.Mnr=M.Mnr Left JOIN Lokaler L ON R.Lokal=L.Lokal WHERE SUBSTRING(R.Ring,1,5) = '" & Mid(Ringnr, 1, 5) & "' ORDER BY R.Ring, R.Tr", Conn, adOpenStatic, adLockReadOnly
Else 'märkaren är känd
  'rs.Open "SELECT Ring, Datum, Mnr, ENamn, FNamn FROM Ring_View WHERE SUBSTRING(Ring,1,5)= '" & Mid(Ringnr, 1, 5) & "' AND Mnr= '" & Mnr & "'ORDER BY Ring, Tr", Conn, adOpenStatic, adLockReadOnly
  rs.Open "SELECT R.Centr,R.Ring,R.Tr,R.Datum,R.Artkod,R.Sex,R.Age,R.Status,R.Varn,R.Mnr,M.ENamn,M.FNamn FROM Ringon R INNER JOIN Maerkare M ON R.Mnr=M.Mnr WHERE SUBSTRING(R.Ring,1,5) = '" & Mid(Ringnr, 1, 5) & "' AND R.Mnr= '" & Mnr & "' ORDER BY R.Ring, R.Tr", Conn, adOpenStatic, adLockReadOnly
End If

If rs.BOF = True And rs.EOF = True Then ' om det inte finns närliggande nummer
  Set cr = New clsRingRc
  cr.Datafinns = MTyp.rcSaknas
  NarliggandeNummer.Add cr
  
Else ' om det finns närliggande nummer
  rs.MoveLast
  'Här kan man omedelbart testa om den sista
  'ringen är mindre än sökt ring.
 
  '************************************************
  'Ändra några variabler
  '**************************************************
  If rs("Ring") < Ringnr Then 'det finns bara före
    If Mid(rs("Ring"), 1, 3) = Mid(Ringnr, 1, 3) Then
      Set cr = New clsRingRc
      cr.Datafinns = MTyp.rcFinns
      cr.Artkod = "Fore"
      cr.Centr = rs("centr")
      cr.Ring = rs("Ring")
      cr.Datum = CStr(rs("Datum"))
      cr.Mnr = rs("Mnr")
      If IsNull(rs("FNamn")) = True Then
        cr.MarkarNamn = rs("ENamn")
      Else
        cr.MarkarNamn = Trim(rs("ENamn")) & ", " & Trim(rs("FNamn"))
      End If
      NarliggandeNummer.Add cr
    Else
      Set cr = New clsRingRc
      cr.Datafinns = MTyp.rcSaknas
      NarliggandeNummer.Add cr
    End If
  Else 'om data finns och den sista ringen är > Ringnr
    rs.MoveFirst
    If rs("Ring") > Ringnr Then
      Set cr = New clsRingRc
      cr.Datafinns = MTyp.rcFinns
      cr.Centr = rs("centr")
      cr.Ring = rs("Ring")
      cr.Artkod = "Efter"
      cr.Datum = CStr(rs("Datum"))
      cr.Mnr = rs("Mnr")
      If IsNull(rs("FNamn")) = True Then
        cr.MarkarNamn = rs("ENamn")
      Else
        cr.MarkarNamn = Trim(rs("ENamn")) & ", " & Trim(rs("FNamn"))
      End If
      NarliggandeNummer.Add cr
    Else
      Do Until rs.EOF
        If rs("Ring") > Ringnr Then
          rs.MovePrevious 'ringen före
          If Mid(rs("Ring"), 1, 3) = Mid(Ringnr, 1, 3) Then
            Set cr = New clsRingRc
            cr.Datafinns = MTyp.rcFinns
            cr.Centr = rs("centr")
            cr.Ring = rs("Ring")
            cr.Artkod = "Fore"
            cr.Datum = CStr(rs("Datum"))
            cr.Mnr = rs("Mnr")
            If IsNull(rs("FNamn")) = True Then
              cr.MarkarNamn = rs("ENamn")
            Else
              cr.MarkarNamn = Trim(rs("ENamn")) & ", " & Trim(rs("FNamn"))
            End If
            NarliggandeNummer.Add cr
            bSaknas = True
          End If
          rs.MoveNext
                                  
          If Mid(rs("Ring"), 1, 3) = Mid(Ringnr, 1, 3) Then
            Set cr = New clsRingRc
            cr.Datafinns = MTyp.rcFinns
            cr.Centr = rs("centr")
            cr.Ring = rs("Ring")
            cr.Artkod = "Efter"
            cr.Datum = CStr(rs("Datum"))
            cr.Mnr = rs("Mnr")
            If IsNull(rs("FNamn")) = True Then
              cr.MarkarNamn = rs("ENamn")
            Else
              cr.MarkarNamn = Trim(rs("ENamn")) & ", " & Trim(rs("FNamn"))
            End If
            NarliggandeNummer.Add cr
            bSaknas = False
          Else
            bSaknas = True
          End If
          
          If bSaknas = True Then
            Set cr = New clsRingRc
            cr.Datafinns = MTyp.rcSaknas
            NarliggandeNummer.Add cr
          End If
          Exit Do
        End If
        rs.MoveNext
      Loop

    End If
  End If
End If
'NarliggandeNummer.Add cR
If rs.State = 1 Then rs.Close
Set rs = Nothing

End Function

Private Sub Basdata(iTr As Integer, Utland As Boolean)
  cr.Rdatum = 0
  cr.RTr = 0
  cr.Ring = rs("Ring")
  cr.Centr = rs("Centr")
  cr.Artkod = rs("Artkod")
  If Utland = False Then
    cr.Mnr = rs("Mnr")
    If IsNull(rs("FNamn")) = False Then
      cr.MarkarNamn = Trim(rs("ENamn")) & ", " & Trim(rs("FNamn"))
    Else
      cr.MarkarNamn = rs("Enamn")
    End If
  End If
  Select Case iTr
    Case 1
      If rs("ring") = frmSok.txtRing Then
      cr.Basdata = cr.Centr & "    " & cr.Ring & "    " & cr.Artkod & "    mnr = " & cr.Mnr & "  " & cr.MarkarNamn
      Else
      cr.Basdata = cr.Centr & "    " & cr.Ring & " (sökt  " & frmSok.txtRing & ")   " & cr.Artkod & "    mnr = " & cr.Mnr & "  " & cr.MarkarNamn
      End If
    Case 7
      cr.Basdata = rs("GCentr") & "    " & rs("GRing") & " (sökt  " & cr.Centr & "    " & cr.Ring & ") " & cr.Artkod & "  " & cr.Mnr & "  " & cr.MarkarNamn
  End Select
End Sub

Public Function FixaSort(Dat, Tr, Tnog, Lokal)

FixaSort = CStr(Dat)
If IsNull(Tnog) = False Then
  FixaSort = FixaSort + Tnog
Else
  FixaSort = FixaSort + "00"
End If
FixaSort = FixaSort + Tr + Lokal

End Function
Private Sub Fixa_Fdet()
Dim Nti As Integer, Laengd As Integer, Rest As Integer
Fdet1 = ""
Fdet2 = ""
Fdet3 = ""
Fdet4 = ""

Nti = 0
Laengd = Len(FdetHela)
Do Until Nti = Laengd
  Nti = Nti + 1
  If Mid(FdetHela, Nti, 1) = "$" Then
    'GoTo Dollarklar
    FdetHela = Fdet1
    Exit Do
  End If
  Fdet1 = Fdet1 + Mid(FdetHela, Nti, 1)
Loop

Fdet1 = Left(FdetHela, 69)
Fdet2 = Mid(FdetHela, 70, 69)
Fdet3 = Mid(FdetHela, 139, 69)
Fdet4 = Mid(FdetHela, 208, 69)
Exit Sub

Dollarklar:
Laengd = Len(Trim(Fdet1))
'Rad ett
If Laengd < 71 Then
  Fdet2 = Mid(FdetHela, (Nti + 1), 70)
Else
  Fdet1 = Left(FdetHela, 70)
  Rest = Laengd - 70
  Fdet2 = Mid(FdetHela, (Nti - Rest - 1), 70)
End If
  
End Sub

Private Sub Fixa_Rapp()
If Len(Rapp1) + Len(Rapp2) > 70 Then GoTo Rappa2
If Rapp2 <> "" Then
  Rapp1 = Rapp1 + "," + Rapp2
  Rapp2 = ""
End If
If Len(Rapp1) + Len(Rapp3) > 70 Then GoTo Rappa3
If Rapp3 <> "" Then
  Rapp1 = Rapp1 + "," + Rapp3
  Rapp3 = ""
End If
If Len(Rapp1) + Len(Rapp4) > 70 Then GoTo Rappa4
If Rapp4 <> "" Then
  Rapp1 = Rapp1 + "," + Rapp4
  Rapp4 = ""
End If
If Len(Rapp1) + Len(Rapp5) > 70 Then GoTo Rappa5
If Rapp5 <> "" Then
  Rapp1 = Rapp1 + "," + Rapp5
  Rapp5 = ""
End If
Exit Sub

Rappa2:
If Len(Rapp2) + Len(Rapp3) > 70 Then GoTo Rappa3
If Rapp3 <> "" Then
  Rapp2 = Rapp2 + "," + Rapp3
  Rapp3 = ""
End If
If Len(Rapp2) + Len(Rapp4) > 70 Then GoTo Rappa4
If Rapp4 <> "" Then
  Rapp2 = Rapp2 + "," + Rapp4
  Rapp4 = ""
End If
If Len(Rapp2) + Len(Rapp5) > 70 Then GoTo Rappa5
If Rapp5 <> "" Then
  Rapp2 = Rapp2 + "," + Rapp5
  Rapp5 = ""
End If
Rapp2 = Replace(Rapp2, "'", "´")
Exit Sub

Rappa3:
If Len(Rapp3) + Len(Rapp4) > 70 Then GoTo Rappa4
If Rapp4 <> "" Then
  Rapp3 = Rapp3 + "," + Rapp4
  Rapp4 = ""
End If
If Len(Rapp2) + Len(Rapp5) > 70 Then GoTo Rappa5
If Rapp5 <> "" Then
  Rapp3 = Rapp3 + "," + Rapp5
  Rapp5 = ""
End If
Rapp3 = Replace(Rapp3, "'", "´")
Exit Sub

Rappa4:
If Len(Rapp4) + Len(Rapp5) > 70 Then GoTo Rappa5
If Rapp5 <> "" Then
  Rapp4 = Rapp4 + "," + Rapp5
  Rapp5 = ""
End If
Rapp4 = Replace(Rapp4, "'", "´")
Exit Sub

Rappa5:
If Len(Rapp5) + Len(Rapp6) > 70 Then GoTo Rappa6
If Rapp6 <> "" Then
  Rapp5 = Rapp5 + "," + Rapp6
  Rapp6 = ""
End If
Rapp5 = Replace(Rapp5, "'", "´")
Exit Sub

Rappa6:
Rapp6 = Replace(Rapp6, "'", "´")

End Sub

Private Sub Fixa_Eposten()

If Epost = "" Then Exit Sub

If Rapp5 <> "" Then
  Rapp5 = Rapp5 + " / " + Epost
  Exit Sub
End If
If Rapp4 <> "" Then
  If Len(Rapp4 + Epost) < 68 Then
    Rapp4 = Rapp4 + " / " + Epost
  Else
    Rapp5 = Epost
  End If
  Exit Sub
End If
If Rapp3 <> "" Then
  If Len(Rapp3 + Epost) < 68 Then
    Rapp3 = Rapp3 + " / " + Epost
  Else
    Rapp4 = Epost
  End If
  Exit Sub
End If
If Rapp2 <> "" Then
  If Len(Rapp2 + Epost) < 68 Then
    Rapp2 = Rapp2 + " / " + Epost
  Else
    Rapp3 = Epost
  End If
  Exit Sub
End If

If Len(Rapp1 + Epost) < 68 Then
  Rapp1 = Rapp1 + " / " + Epost
Else
  Rapp2 = Epost
End If

End Sub

Private Sub Reporter(Rappnyckel As String, Rapptyp As String, Mom As String, Bdat As String, Deldnr As String)
  
'Rad med uppgifter om fyndets status i fyndhanteringen
Dim Svardat As String, SvarDatRapp1 As Boolean, KopAdr As String
Dim rsRapp As ADODB.Recordset

On Error GoTo RepErr

'If RTrim(rs!Rappnyckel) <> "-ejreg" Then GoTo Slutrad

Set rsRapp = New ADODB.Recordset

FirstRapp = False
Rappnyckel = RTrim(Rappnyckel)

Svardat = ""
SvarDatRapp1 = False
If IsNull(Mom) = False And RTrim(Mom) <> "" Then
  If Mom = "K" Or Mom = "I" Then
    Svardat = "/Korrektur - ej brev ännu/"
  Else
    If IsNull(Bdat) = False Then Svardat = " /svarsbrev " & Bdat & "/"
  End If
End If
  
'Sök nu rapportör
Rapp1 = ""
Rapp2 = ""
Rapp3 = ""
Rapp4 = ""
Rapp5 = ""
Rapp6 = ""
Epost = ""

If Rapptyp = "C" Then 'Rapportör är central
  rsRapp.Open "SELECT Cekortad FROM ceadress WHERE centr= '" & Rappnyckel & "'", Conn, adOpenStatic, adLockReadOnly
  If rsRapp.EOF = True Then
    Rapp1 = "Rapportör: " & RTrim(Rappnyckel) & " Kan ej hittas i Central-registret!"
  Else
    Rapp1 = "Rapportör: (" & RTrim(Rappnyckel) & ") " & rsRapp!cekortad
  End If
End If

If Rapptyp = "M" Then
  rsRapp.Open "SELECT fnamn,enamn FROM maerkare WHERE mnr= '" & Rappnyckel & "'", Conn, adOpenStatic, adLockReadOnly
  If rsRapp.EOF = True Then
    Rapp1 = "Rapportör: " & RTrim(Rappnyckel) & " Kan ej hittas i märkar-registret!"
  Else
    Rapp1 = "Rapportör : (" & RTrim(Rappnyckel) & ") "
    If IsNull(rsRapp!enamn) = False Then Rapp1 = Rapp1 + RTrim(rsRapp!enamn)
    If IsNull(rsRapp!fnamn) = False And RTrim(rsRapp!fnamn) <> "" Then Rapp1 = Rapp1 + ", " + RTrim(rsRapp!fnamn)
  End If
End If

If Rapptyp = "K" Then
  rsRapp.Open "SELECT Namn FROM RcRapp WHERE Rappnyckel = '" & Rappnyckel & "'", Conn, adOpenStatic, adLockReadOnly
  If rsRapp.EOF = True Then
    'kolla då om kodade rapp är ett märkarnummer, dvs rapptyp K är fel ska egentligen vara M
    rsRapp.Close
    rsRapp.Open "SELECT fnamn,enamn FROM maerkare WHERE mnr= '" & Rappnyckel & "'", Conn, adOpenStatic, adLockReadOnly
    If rsRapp.EOF = True Then
      Rapp1 = "Rapportör: " & RTrim(Rappnyckel) & " Kan ej hittas i märkar-registret!"
    Else
      Rapp1 = "Rapportör : (" & RTrim(Rappnyckel) & ") "
      If IsNull(rsRapp!enamn) = False Then Rapp1 = Rapp1 + RTrim(rsRapp!enamn)
      If IsNull(rsRapp!fnamn) = False And RTrim(rsRapp!fnamn) <> "" Then Rapp1 = Rapp1 + ", " + RTrim(rsRapp!fnamn)
    End If
  Else
    Rapp1 = "Rapportör : (" & RTrim(Rappnyckel) & ") " & rsRapp!namn
  End If
End If

If Rapptyp = "A" Then
  rsRapp.Open "SELECT * FROM RcRapp WHERE Rappnyckel= '" & Rappnyckel & "'", Conn, adOpenStatic, adLockReadOnly
  If rsRapp.EOF = True Then
    Rapp1 = "Uppgift om rapportör saknas"
    If Right(RTrim(Rappnyckel), 6) = "anonym" Then Rapp1 = "Rapportör anonym"
  End If
  If rsRapp.EOF = False Then
    Rapp1 = "Rapportör: " & RTrim(rsRapp!namn)
    If IsNull(rsRapp!adr1) = False Then Rapp2 = RTrim(rsRapp!adr1)
    If IsNull(rsRapp!adr2) = False Then Rapp3 = RTrim(rsRapp!adr2)
    If IsNull(rsRapp!adr3) = False Then Rapp4 = RTrim(rsRapp!adr3)
    If IsNull(rsRapp!adr4) = False Then Rapp5 = RTrim(rsRapp!adr4)
    If IsNull(rsRapp!adr5) = False Then Rapp6 = RTrim(rsRapp!adr5)
    If IsNull(rsRapp!email) = False Then Epost = RTrim(rsRapp!email)
    Fixa_Rapp

  End If
End If

If rsRapp.State = 1 Then rsRapp.Close

Fixa_Eposten

If Rapptyp = "C" Or Rapptyp = "M" Or (Rapptyp = "A" And Rapp2 = "") Then
  If Svardat <> "" And Len(RTrim(Rapp1)) < 34 Then
    Rapp1 = Rapp1
    Do Until Len(Rapp1) = 33
      Rapp1 = Rapp1 + " "
    Loop
    Rapp1 = Rapp1 + Svardat
    SvarDatRapp1 = True
  End If
End If
  
Conn.Execute "INSERT INTO SökRing (RDatum, RTr, Extradata)" & " VALUES ('" & OlSort & "', " & 25 & ", '" & Rapp1 & "')"
If Rapp2 <> "" Then Conn.Execute "INSERT INTO SökRing (RDatum, RTr, Extradata)" & " VALUES ('" & OlSort & "', " & 26 & ", '" & Rapp2 & "')"
If Rapp3 <> "" Then Conn.Execute "INSERT INTO SökRing (RDatum, RTr, Extradata)" & " VALUES ('" & OlSort & "', " & 27 & ", '" & Rapp3 & "')"
If Rapp4 <> "" Then Conn.Execute "INSERT INTO SökRing (RDatum, RTr, Extradata)" & " VALUES ('" & OlSort & "', " & 28 & ", '" & Rapp4 & "')"
If Rapp5 <> "" Then Conn.Execute "INSERT INTO SökRing (RDatum, RTr, Extradata)" & " VALUES ('" & OlSort & "', " & 29 & ", '" & Rapp5 & "')"
If Rapp6 <> "" Then Conn.Execute "INSERT INTO SökRing (RDatum, RTr, Extradata)" & " VALUES ('" & OlSort & "', " & 29 & ", '" & Rapp6 & "')"

'Rad med ev kopieadressater
rsRapp.Open "SELECT Rckopad.rappnyckel,Rcrapp.namn,maerkare.fnamn,maerkare.enamn, Ceadress.cekortad FROM Rckopad LEFT JOIN Rcrapp ON Rckopad.rappnyckel=rcrapp.rappnyckel LEFT JOIN Maerkare ON Rckopad.Rappnyckel=Maerkare.Mnr LEFT JOIN Ceadress ON Rckopad.rappnyckel=Ceadress.centr WHERE left(dnr,9)= '" & Deldnr & "'", Conn, adOpenStatic, adLockReadOnly
If rsRapp.EOF = False Then
  KopAdr = "Svarskopia till: "
  rsRapp.MoveFirst
  Do Until rsRapp.EOF
    If IsNull(rsRapp!namn) = False Then
      KopAdr = KopAdr + RTrim(rsRapp!namn) + " / "
    End If
    If IsNull(rsRapp!cekortad) = False Then
      KopAdr = KopAdr + RTrim(rsRapp!cekortad) + " / "
    End If
    If IsNull(rsRapp!enamn) = False Then
      If IsNull(rsRapp!fnamn) = False Then
        KopAdr = KopAdr + RTrim(rsRapp!fnamn) + " " + RTrim(rsRapp!enamn) + " / "
      Else
        KopAdr = KopAdr + RTrim(rsRapp!enamn) + " / "
      End If
    End If
    rsRapp.MoveNext
  Loop
  Conn.Execute "INSERT INTO SökRing (RDatum, RTr, Extradata)" & " VALUES ('" & OlSort & "', " & 30 & ", '" & KopAdr & "')"
End If
  
If SvarDatRapp1 = False And Svardat <> "" Then
  KopAdr = "                                 " & Svardat
  Conn.Execute "INSERT INTO SökRing (RDatum, RTr, Extradata)" & " VALUES ('" & OlSort & "', " & 31 & ", '" & KopAdr & "')"
End If

Slutrad:
'Slutrad för posten
Conn.Execute "INSERT INTO SökRing (RDatum, RTr, rubrik)" & " VALUES ('" & OlSort & "', " & 32 & ", '--------------------------------------------------------------------------------------')"

If rsRapp.State = 1 Then rsRapp.Close
Set rsRapp = Nothing
Exit Sub

Resume
RepErr:
MsgBox Err.Number & Err.Description
Resume Next
End Sub
