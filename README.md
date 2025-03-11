# Inventory-Dashboard

VBA project; focus on SAP <> Excel interactions

NAME: Inventory 'Wingman' Dashboard

---> Please SEARCH/FIND VBA code HIGHLIGHTS with  " __________ " character set

---> Script highlights are more visible and searchable in the TXT version of readme file !


### Removing header of SAP-import (done through text clipboard) with eg LIKE
### Application.Inputbox with : data entry, empty entry (fast re-use of previous data) and aborting
### Item search modifiers: User gets prompted, next put in use
### Looped searching via SAP t-code in case of errors
### Flipping through SAP tables (3, document-flow check equivalanet, now missing in SAP) to pass on data to main Subroutine
### Dictionary-based comparison of 2 lists
### Finding most frequent value in variable based range
### Calling other Subs
### Powerful SWTICH-featuring Formula to check text lengths and start creating other logics
### MB1B (invetory transfer) mass transacting
### Subroutine to Read iDocs and pass its contenst to anoother sub
### Array formula looking for set of digits at 1 time, and returning their positions

+ extensive cross-tab, 2-SAP-t-codes automation of Material, price, value formula determination
+ as in screenshots (sheet example + dashboard looks in general)

Details below:


__________  ### Removing header of SAP-import through text clipboard (with eg LIKE):

For pPos = 1 To 15                                                                              'highlight lines with rolls
    If Range("C" & pPos) Like "| ??#??????? *" Then
        Range("B" & pPos) = "here"
    End If
Next pPos
...                                                                                              'count the highlighted lines, deduct 1
Range("B1").Select
Range(Selection, Selection.End(xlDown)).Select
selcount = Selection.count
selcount = selcount - 1
...                                                                                              'erase all lines from 1-thru-(count-1) to clear out the whole header of export from SAP
Range("B1:B" & selcount).EntireRow.Delete
Range("B:B").ClearContents


__________ ### Application.Inputbox with : data entry, empty entry (fast re-use of previous data) and aborting
On Error Resume Next                                                                            'if error, go below
Set batches = Application.InputBox("HIGHLIGHT  ROLLS : ", " About to SHOW  :", Left:=100, Top:=75, Type:=8)
...
If batches Is Nothing Then   'if none highlighted, stop
  GoTo ErrHan
End If
...
If batches <> 0 Then                                                                            'if new ones, clear below ranges, which hold previous batches; go next block
    Workbooks("macro.xlsm").Worksheets("Show").Range("A:A, B:B, C:C, F:F, G1:G3").ClearContents
    GoTo EditCount
End If
...
If Workbooks("macro.xlsm").Worksheets("Show").Range("F1").Value = "" Then                      'stop in case previous SHOW trial indicated they were not in cons.)
    MsgBox "Macro was NOT SHOWing these batches the last time it tried, highlight others please", vbOKOnly, "Not these ones"
    Exit Sub
End If
...
...                                                                                               'take previous batches (F:F -> A:A) and go to Fast script block
Workbooks("macro.xlsm").Worksheets("Show").Range("A1:A2736").Value = & _
Workbooks("macro.xlsm").Worksheets("Show").Range("F1:F2736").Value
GoTo Fast

EditCount:                                                                                        'play with colors, copy RLs
batches.Interior.Color = vbBlack
etc................
...
On Error GoTo ShowNext:                                                                            'run next SHOW-subroutine part anyway
  session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell 1, ""
  session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").clearSelection
Exit Sub


__________ ### Item search modifiers: User gets prompted, next put in use:
"Missing items in Dashboard -> PREFIX 1-ST BATCH WITH below mo and rerun :" + vbNewLine + vbNewLine + _
      "#b        searches for Batches (GCF-relevant)" + vbNewLine + _
      "#s         toggles search:  1 VS ALL Location SBPs" + vbNewLine + _
      "##        does both" + vbNewLine + vbNewLine + _
      "#z | zz   generates 1 | ALL SBPs snap"..
...
ElseIf iDocWhse = "" Then                                                                         'in case process modifiers were used, calculate them 
    Range("E2").Formula = "=IF(LEN(A1)=12,RIGHT(A1,10),(IF(LEN(A1)=11,LEFT(A1,10),A1)))"          'exlude 1-st 2 marks from RL, if 12-char long
    Range("E2").Value = Range("E2").Value
    Range("E1").Formula = "=IF(or(LEFT(a1,2)=""#s"",LEFT(a1,2)=""zz"",LEFT(a1,2)=""##""),1,0)"    'Many VS 1SBP  & per-BATCHES !
    Range("E1").Value = Range("E1").Value
    Range("E3").Formula = "=IF(OR(LEFT(a1,2)=""#b"",LEFT(a1,2)=""##""),1,0)"                      'per BATCH
    Range("E3").Value = Range("E3").Value
    Range("E4").Formula = "=IF(OR(LEFT(a1,2)=""#z"",LEFT(a1,2)=""zz""),1,0)"                      'snapshots
End If
...
ifbatches = Workbooks("macro.xlsm").Worksheets("Show").Range("E3").Value
''ifbatches = CBool(Range("E3").Value)
ifsnap = Workbooks("macro.xlsm").Worksheets("Show").Range("E4").Value
''ifsnap = CBool(Range("E4").Value)
...
If ifsnap = True Then                                                                               'mark Inventory Report (SNAPSHOT), if ruled so by Modifier'if snapshot being run
   Application.DisplayAlerts = False                                                                'for cases of longer snapshot -disable xls alerts'
   session.findById("wnd[0]/usr/radP_INVRP").Select
   session.findById("wnd[0]/usr/ctxtSO_HUNIT-LOW").Text = ""
   session.findById("wnd[0]/tbar[1]/btn[8]").press
    Exit Sub
Else                                                                                                 'if batches field to be selected
      If ifbatches = False Then
        session.findById("wnd[0]/usr/btn%_SO_HUNIT_%_APP_%-VALU_PUSH").press
      ElseIf ifbatches = True Then                                                                   'if otherwise
        session.findById("wnd[0]/usr/btn%_SO_CHARG_%_APP_%-VALU_PUSH").press
      End If
End If


__________ ### Looped searching via SAP t-code in case of errors
ErrHan3:
If iDocWhse <> "" Then                                                                                'IDOC -> if outside this Inventory Type, run MB51
    GoTo ErrHan1
ElseIf iDocWhse = "" Then                                                                             'non-idoc -> next Loop section
    GoTo ErrHan2
End If
ErrHan2:                                                                                              'increase ShowTrial no. by 1
TryCou = TryCou + 1
If TryCou = 1 Then
    If OneSbp = False Then                                                                            'switch to 1 ship-to
        OneSbp = True
        GoTo again
    ElseIf OneSbp = True Then                                                                         'if 1 ship-to, switch to perBatch
        If ifbatches = False Then
            ifbatches = True
            GoTo again
        ElseIf ifbatches = True Then                                                                  'in case per sub-unit is on, run MB51
            GoTo ErrHan1
        End If
    End If
ElseIf TryCou = 2 Then                                                                                'in case switch to 1 ship-to itself isn't enough,
        If ifbatches = False Then                                                                     'switch to per sub-unit too
            ifbatches = True
            GoTo again
        End If
ElseIf TryCou = 3 Then                                                                                  'unknown fail...
    MsgBox "unknown fail    ??? ", vbOKOnly
    GoTo ErrHan1
Exit Sub
again:                                                                                                  'resumes LOOP
    Resume HHHHOutoffMB51
Exit Sub
ErrHan1:                                                                                                'runs MB51 as last resort
   Call Errorhandler(batch1, wbname, pRange, PriceVar, ilow, fc)
   Exit Sub
End Sub


__________ ### Dictionary-based comparison of 2 lists
   With CreateObject("scripting.dictionary")
   .comparemode = 1
      For i = 1 To UBound(ar, 1)
      .Item(ar(i, 2)) = Empty
      Next
         For i = 1 To UBound(ar, 1)
         If Not .exists(ar(i, 1)) Then
         n = n + 1
         var(n, 1) = ar(i, 1)
         End If
         Next
   End With
   On Error Resume Next                                                                                   'feeds [E16] array with elements missing in [D16] vs [C16]
   [E16].Resize(n).Value = var

__________ ### Flipping through SAP tables (3, document-flow check equivalanet, now missing in SAP) to pass on data to main Subroutine

If IntShip <> "" Then                                        'in case Internal Shipm, plug it in, jump to exporting
    session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "VBFA"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtI3-LOW").Text = IntShip
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    cont = vbNullString
    GoTo exporting
End If

recogn:

If item.Value Like "9000######" Then                        'if TD-Shipment used, plug it in
    cfosShipm = item.Value
    session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "VBFA"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtI3-LOW").Text = cfosShipm
    session.findById("wnd[0]/tbar[1]/btn[8]").press
ElseIf item.Value Like "900#######" Then                    'INVOICE
    inv = item.Value
    session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "VBFA"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtI3-LOW").Text = inv
    session.findById("wnd[0]/tbar[1]/btn[8]").press
                                                            'sorting for ZKEs mixed up with Dlvr's
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1, "VBELV"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "VBELV"
    session.findById("wnd[0]/tbar[1]/btn[28]").press
ElseIf item.Value Like "80########" Then                   'DELIVERY
        Delivery = item.Value
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "VBFA"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtI3-LOW").Text = Delivery
        session.findById("wnd[0]/tbar[1]/btn[8]").press
ElseIf item.Value Like "???[A-Z]#######" Then               'if Container
    cont = item.Value
    session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "VTTK"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/txtI21-LOW").Text = cont
    session.findById("wnd[0]/tbar[1]/btn[8]").press

ElseIf item.Value Like "10#######" Then                      'if order
    ZKE = item.Value
    GoTo ZKE
Else                                                        'else item = Vehicle ID (delivery Note ref.)
    VehID = UCase(VehID)
    If Left(VehID, 3) = "GP2" Then
        VehID = Replace(VehID, "GP2", "")
    End If
    GoTo ZKE
End If
        
exporting:                                                  'exporting from VBFA
 On Error GoTo ErrHan
 session.findById("wnd[0]/tbar[1]/btn[45]").press
 session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
 session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
 session.findById("wnd[1]/tbar[0]/btn[0]").press


__________ ### Finding most frequent value in variable based range
wb.Worksheets("Compare").Range("H2").Formula = "=INDEX(A10:A" & count & ",MODE(MATCH(A10:A" & count & ",A10:A" & count & ",0)))"
mode = wb.Worksheets("Compare").Range("H2").Value                                                        'mode (most frequent) plant


__________ ### Mass range check with LIKE inside a Loop:
pPos = 53
     For cPos = 1 To 33
        If Range("AZ" & cPos) Like "#__#" Then                                                         'if indeed a Port used in consignment, add it to MD, so that MB51 includes it/them)
           Range("AZ" & cPos).Copy
           Range("i" & pPos).Select
           Selection.PasteSpecial Paste:=xlPasteValues
           pPos = pPos + 1
        End If
     Next cPos


__________ ### Calling other Subs
If fc > 0 Then                                                                                            'if FOUNDS are there, run MB51 to get BATCH+PLANT KEEY, weights etc...
   Call Errorhandler(batch1, wbname, pRange, PriceVar, ilow, fc)
   Exit Sub
ElseIf fc = 0 Then                                                                                        'if no FOUNDS, add 1 col. into Prices-TAB col. B (which mimicks range shifting, so the relevnat col.
   Workbooks("macro.xlsm").Worksheets("Prices").Activate                                                'stays where it should be, as if the MB51 was ran)
   Columns("B:B").Select
   Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
End If
...
Call prices2(wbname, PriceVar, pRange, ilow, fc, SapMissCou)                                              'run MM60 directly


__________ ### Powerful SWTICH-featuring Formula to check text lengths and start creating other logics
Range("G" & pfi).Formula = "=IFNA(TRIM(SWITCH(LEN(h" & pfi & "),148,MID(h" & pfi & "" & Chr(38) & """ "",121,6)," & _
    "145,MID(h" & pfi & "" & Chr(38) & """ "",118,6),150,MID(h" & pfi & "" & Chr(38) & """ "",121,8),155," & _
    "MID(h" & pfi & "" & Chr(38) & """ "",125,9),149,MID(h" & pfi & "" & Chr(38) & """ "",122,7))),"""")"


### Multiple XLookUps and variables inside formula, with Variables derived from earlier measuring of xls Ranges
For Each cell In Range("E" & istartmi & ":E" & mic)
  Range("E" & istartmi & ":E" & mic).Formula = "=INDEX(N:N,MATCH((XLOOKUP(C" & istartmi & ",Compare!$E$" & istartmi & ":$E$70000," & _
  "Compare!$A$" & istartmi & ":$A$70000)&(XLOOKUP(C" & istartmi & ",Compare!$E$" & istartmi & ":$E$70000,Compare!$C$" & istartmi & ":$C$70000))),M:M,0))"
Next cell


__________ ### Subroutine to Read iDocs and pass its contenst to anoother sub
...                                                                                                        'preparation formula agaoinst iDoc items          
=IF(LEFT(A27,6)="CREDAT",CONCAT("_"&(RIGHT(A27,8))),IF(LEFT(A27,16)="IDocNumber:00000",RIGHT(A27,10),IF(LEFT(A27,8)="NameDesc", & _
CONCAT("="),IF(LEFT(A27,5)="ABTNR",CONCAT("SCAC_"&(RIGHT(A27,4))),IF(LEFT(A27,3)="MAT",RIGHT(LEFT(A27,15),10),IF(LEFT(A27,4)="MEST", & _
CONCAT("_"&(RIGHT(A27,3))),IF(LEFT(A27,29)="SNDPRNSenderPartnerNumber0000",CONCAT("__"&(RIGHT(A27,6))),A27)))))))
...
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "MESTYP"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "SNDPRN"
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "CREDAT"
...
    Workbooks("macro.xlsm").Worksheets("iDoc-SAP").Range("a1").Select
    ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, _
    DisplayAsIcon:=False, NoHTMLFormatting:=True
    Selection.RemoveDuplicates Columns:=1, Header:=xlNo
    Selection.Replace What:=Chr(124), Replacement:=Chr(0), LookAt:=xlPart
    Selection.Replace What:=Chr(32), Replacement:=Chr(0), LookAt:=xlPart
...
poPos = poPos + 1
   If poPos < 25 Then
      For POck = poPos To 25
         If Range("D" & poPos) Like "BELNR*" Then
            idocPO2 = Range("D" & poPos).Text
            idocPO2 = Replace(idocPO2, "BELNR", "PO/Veh_")
            Range("D" & poPos).Value = idocPO2
            Range("D" & poPos).Interior.ColorIndex = 6
         End If
            If idocPO2 <> "" Then
            Exit For
            End If
      Next POck
   End If
With Workbooks("macro.xlsm").Worksheets("MAP").Range("D2:D299")
   Set c = .Find(iDocWhse, LookIn:=xlValues, LookAt:=xlPart)                                            ' Setting renge reference and SEARCHing
    If Not c Is Nothing Then
        WhName = c.Offset(0, 1).Value
        WhName = UCase(WhName)
        Workbooks("macro.xlsm").Worksheets("iDoc-SAP").Range("D" & wPos).Value = WhName
        Workbooks("macro.xlsm").Worksheets("iDoc-SAP").Range("D" & wPos).Font.Color = vbRed
    Else
        MsgBox "This Location ship-to is missing in macro MAP tab, please update !"
   End If
End With
...
 If Workbooks("macro.xlsm").Worksheets("iDoc-SAP").Range("D1").Value <> "" Then                         'iDoc Delete                  
     yesNo = MsgBox("delete iDoc# " & iDocNo & " from Inbox?" & (Chr(10)) & (Chr(10)) & _
     "[   YES  ]" & (Chr(10)) & (Chr(10)) & _
     "[   NO   ]", vbYesNo + vbCritical, "Delete iDoc now ?")
        Select Case yesNo
          Case vbYes
          session.findById("wnd[0]/tbar[0]/btn[3]").press
          session.findById("wnd[0]/tbar[0]/btn[3]").press
          session.findById("wnd[0]/tbar[1]/btn[14]").press
          session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
          Application.Wait Now + TimeValue("00:00:03")
          session.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton "EREF"
          session.findById("wnd[0]/tbar[0]/btn[15]").press
        Case vbNo
          session.findById("wnd[0]/tbar[0]/btn[3]").press
          session.findById("wnd[0]/tbar[0]/btn[3]").press
          End Select   
 ElseIf Workbooks("macro.xlsm").Worksheets("iDoc-SAP").Range("D1").Value = "" Then
     GoTo ErrHan
 End If
...
 Call showpart2(icount, iDocWhse, SCAC, wbname, iDocNo)
 Exit Sub
...
    Case vbNo                                                                                               'iDoc forward
         session.findById("wnd[0]/tbar[0]/btn[15]").press
         session.findById("wnd[0]/tbar[0]/btn[15]").press
         session.findById("wnd[0]/tbar[0]/btn[15]").press
         session.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").pressToolbarButton "SEND"
         session.findById("wnd[1]/usr/ctxtPCHDY-SEARK").Text = destination
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         Exit Sub
    End Select
...
ErrHan:
If iPos = 0 Then
   MsgBox "InBox appears EMPTY / SAP session NOT backed to home view / ... ?", vbOKOnly
ElseIf iPos = 1 And bPos < 25 Then
   MsgBox "iDoc exporting error / Connection issue ?", vbOKOnly, vbCritical
ElseIf iPos = 1 And bPos >= 25 Then
    Workbooks("macro.xlsm").Worksheets("iDoc-SAP").Range("d" & iPos).Copy
    MsgBox "iDoc seems EMPTY and is kept in INbox" & (Chr(10)) & (Chr(10)) & _
    "--> check in Inbox or t-code: WE02." & (Chr(10)) & (Chr(10)) & (Chr(10)) & _
    "!  IDOC #  in Clipboard !"
End If


__________ ### MB1B (invetory transfer) mass transacting
If iter = 0 Then                                                                                                                 'For ! BATCH 1 only !
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").Text = material                                         'other transf. param's
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").Text = "1"
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-ERFME[0,44]").Text = "rl"
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[0,53]").Text = batch
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subBLOCK2:SAPMM07M:2401/ctxtMSEG-KDAUF").Text = orderNo
    session.findById("wnd[0]/usr/subBLOCK2:SAPMM07M:2401/txtMSEG-KDPOS").Text = orderLi                                                                                                            
    session.findById("wnd[0]").sendVKey 0                                                                                        '2x enter
    session.findById("wnd[0]").sendVKey 0
...         
  For Each cell In Workbooks("macro.xlsm").Worksheets("ZKB-T").Range("D" & (ibatchlow + 1) & ":D" & ibatchhigh1)                  'For BATCHES 2 TO BATCH24    ( ! batchcount <= 23 ! )
       ibatchlow = ibatchlow + 1                                                                                                   'next cell = next batch : 1 by 1
       Range("F1").Formula = "=D" & ibatchlow & ""                                                                                 'formula change of 1 batch for paste into SAP
       batch = Range("F1").Value                                                                                                   'as above
         session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[" & CStr(ibatchlow - 1) & ",7]").Text = material           'other transf. param's
         session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[" & CStr(ibatchlow - 1) & ",26]").Text = "1"
         session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-ERFME[" & CStr(ibatchlow - 1) & ",44]").Text = "rl"
         session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[" & CStr(ibatchlow - 1) & ",53]").Text = batch
    Next cell
    For ii = 1 To ibatchhigh1                                                                                                    'loop Enter Key for 2~23 batches
    session.findById("wnd[0]").sendVKey 0
    Next ii
   session.findById("wnd[0]/tbar[0]/btn[11]").press
   session.findById("wnd[0]/tbar[0]/btn[15]").press
ElseIf iter >= 1 Then                                                                                                             'FOR BATCHES in 1 or more full iterations (MB1B screens)
    For i = 1 To iter
        Range("F1").Formula = "=D" & ibatchlow & ""
        batch = Range("F1").Value
        session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").Text = material                                    'other transf. param's
        session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").Text = "1"
        session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-ERFME[0,44]").Text = "rl"
        session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[0,53]").Text = batch
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/subBLOCK2:SAPMM07M:2401/ctxtMSEG-KDAUF").Text = orderNo
        session.findById("wnd[0]/usr/subBLOCK2:SAPMM07M:2401/txtMSEG-KDPOS").Text = orderLi
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
            counta = counta + 1                                                                                                   'currently done batches (1 only in "header" line)
                subiter = 1                                                                                                      'FOR BATCHES(2 through 23)*iteration  ; subiter = postition inside array
        For Each cell In ....
...
Select Case lastrange                                                                                                           'ck how many batches in the last Mb1b window
    Case Is = 1        '1 batch ONLY
        ibatchlow = (iter * ScrLength) + 1                                                                                       'start pos. of 1-st and ONLY batch of the remainder
        Range("F1").Formula = "=D" & ibatchlow & ""                                                                             'formula change of 1 batch for paste into SAP
        batch = Range("F1").Value                                                                                                'as above
        session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").Text = material
        session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").Text = "1"
        session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-ERFME[0,44]").Text = "rl"
        session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[0,53]").Text = batch
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/subBLOCK2:SAPMM07M:2401/ctxtMSEG-KDAUF").Text = orderNo
        session.findById("wnd[0]/usr/subBLOCK2:SAPMM07M:2401/txtMSEG-KDPOS").Text = orderLi
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        session.findById("wnd[0]/tbar[0]/btn[15]").press
     Case Is > 1                                                                                                                  '1-st batch of more in the last Mb1b window
                 ibatchlow = (iter * ScrLength) + 1                                                                               'start pos. of 1-st batch of the remainder
                 Range("F1").Formula = "=D" & ibatchlow & ""                                                                      'formula change of 1 batch for paste into SAP
                 batch = Range("F1").Value                                                                                         'as above
...


__________ ### Array formula looking for set of digits at 1 time, and returning their positions
    Set dig = Workbooks("macro.xlsm").Worksheets("MB58").Range("E1:E27")
    Set res = Workbooks("macro.xlsm").Worksheets("MB58").Range("G1:G27")
...
    ActiveWorkbook.Names.Add Name:="results", RefersTo:=res
    ActiveWorkbook.Names.Add Name:="digits", RefersTo:=dig
...
Workbooks("macro.xlsm").Worksheets("MB58").Range("E:E").NumberFormat = "@"
    Range("E1").Value = "1.000"
    Range("E2").Value = "2.000"
    Range("E3").Value = "3.000"
...                                                                                                                      
Range("F1").Select                                                               'array formula to search for above WT endings in the contents of C col. (-3 col. index), and display their positions in col. F
Selection.FormulaArray = _
"=INDEX(results,MATCH(TRUE,ISNUMBER(SEARCH(digits,R[0]C[-3])),0))"                'col.offset 
Selection.AutoFill destination:=Range("F1:F27"), Type:=xlFillDefault
...
    Range("A7").Formula = "=COUNTIF(B1:B11,TRUE)"                                'count TRUE ship-tos
    Range("A5").Formula = "=MATCH(0,F1:F27,0)"                                   'see if array formula found position of non-zero wt somewhere, if so in which line?


END of FILE
Thank you,
