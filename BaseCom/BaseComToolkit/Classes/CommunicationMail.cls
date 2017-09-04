VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommunicationMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : CommunicationMail
'    Project    : ToolKits
'
'    Description: Stellt Mailing-Methoden zur Verfügung
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------
Private Function PrepareOutlookMail(ByVal senderAddress As String, _
                                    ByVal toAddresses As String, _
                                    ByVal ccAddresses As String, _
                                    ByVal bccAddresses As String, _
                                    ByVal Subject As String, _
                                    ByVal body As String, _
                                    ByVal attachmentFileNames As String) As Object
  
  Dim app         As Object: Set app = CreateObject("Outlook.Application")
  Dim mail        As Object: Set mail = app.CreateItem(0)
        
  With mail
    If Trim(senderAddress) <> "" Then .SentOnBehalfOfName = senderAddress
    If Trim(toAddresses) <> "" Then .To = Replace$(Replace$(toAddresses, ",", ";"), " ", "")
    If Trim(ccAddresses) <> "" Then .Cc = Replace$(Replace$(ccAddresses, ",", ";"), " ", "")
    If Trim(bccAddresses) <> "" Then .Bcc = Replace$(Replace$(bccAddresses, ",", ";"), " ", "")
    .Subject = Subject
    
    Dim signature As String: signature = .body
    .body = body & vbNewLine & signature
      
    Dim attachments() As String: attachments = Split(attachmentFileNames, ",")
    Dim i             As Integer

    For i = LBound(attachments) To UBound(attachments)
      .attachments.Add attachments(i)
    Next i

  End With
        
  Set PrepareOutlookMail = mail
End Function



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SendMailByOutlook
' Description:       Versendet eine Mail über den Outlook-Client
' Created by :       Sebastian Limke
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-16:01:01
'
' Parameters :       toAddresses (String)
'                    ccAddresses (String)
'                    bccAddresses (String)
'                    Subject (String)
'                    body (String)
'                    attachmentFileNames (String)
'--------------------------------------------------------------------------------
Public Sub SendMailByOutlook(ByVal senderAddress As String, _
                             ByVal toAddresses As String, _
                             ByVal ccAddresses As String, _
                             ByVal bccAddresses As String, _
                             ByVal Subject As String, _
                             ByVal body As String, _
                             ByVal attachmentFileNames As String)
    
  Dim mail As Object: Set mail = PrepareOutlookMail _
  (senderAddress, toAddresses, ccAddresses, bccAddresses, Subject, body, attachmentFileNames)
  mail.Send
  Set mail = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       OpenOutlookClient
' Description:       Öffnet eine neue Nachricht im Outlook-Client und füllt
'                    die übergebenen Mail-Parameter automatisch aus.
' Created by :       Project Administrator
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-16:01:50
'
' Parameters :       toAddresses (String)
'                    ccAddresses (String)
'                    bccAddresses (String)
'                    Subject (String)
'                    body (String)
'                    attachmentFileNames (String)
'--------------------------------------------------------------------------------
Public Sub OpenOutlookClient(ByVal senderAddress As String, _
                             ByVal toAddresses As String, _
                             ByVal ccAddresses As String, _
                             ByVal bccAddresses As String, _
                             ByVal Subject As String, _
                             ByVal body As String, _
                             ByVal attachmentFileNames As String)

  Dim mail As Object: Set mail = PrepareOutlookMail _
  (senderAddress, toAddresses, ccAddresses, bccAddresses, Subject, body, attachmentFileNames)
  mail.Display
  Set mail = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CheckEmailAdress
' Description:       Validates a eMail-Adress
' Created by :       Matthias Zirngibl | http://masterbootrecord.de
' Machine    :       VDI-EDV-0003
' Date-Time  :       21.01.2015-13:36:36
'
' Parameters :       mailAddress (String)
'--------------------------------------------------------------------------------
Public Function CheckMailAdress(mailAddress As String) As Boolean
  Dim whiteSpacePattern As String: whiteSpacePattern = "\s"
  Dim mailAddressPattern As String: mailAddressPattern = "([\w-+]+(?:\.[\w-+]+)*@(?:[\w-]+\.)+[a-zA-Z]{2,7})"
  
  Dim regex As regex: Set regex = New regex
  CheckMailAdress = (Not regex.IsPatternMatched(whiteSpacePattern, mailAddress)) _
  And (regex.IsPatternMatched(mailAddressPattern, mailAddress))
  Set regex = Nothing
  
'  Dim bGoodAdress As Boolean
'  Dim sTopLevelDomainsArray() As String
'  Dim sTopLevelDomains As String
'  Dim eMailSplices() As String
'  Dim i As Long
'
'  CheckMailAdress = False
'
'  If InStr(mailAddress, " ") > 0 Then Exit Function
'
'  If (InStr(mailAddress, ",") > 0) Or (InStr(mailAddress, ";") > 0) Then Exit Function
'
'  bGoodAdress = False
'
'  mailAddress = LCase(mailAddress)
'
'  sTopLevelDomains = "com,net,edu,arpa,org,gov,museum,biz,info,pro,name,aero,coop,ac,ad,ae,af,ag,ai,al,am,an,ao,aq,ar,as,at,au,aw,az,ba,bb,bd,be,bf,bg,bh,bi,bj,bm,bn,bo,br,bs,bt,bv,bw,by,bz,ca,cc,cd,cf,cg,ch,ci,ck,cl,cm,cn,co,cr,cu,cv,cx,cy,cz,de,dj,dk,dm,do,dz,ec,ee,eg,eh,er,es,et,fi,fj,fk,fm,fo,fr,ga,gd,ge,gf,gg,gh,gi,gl,gm,gn,gp,gq,gr,gs,gt,gu,gw,gy,hk,hm,hn,hr,ht,hu,id,ie,il,im,in,io,iq,ir,is,it,je,jm,jo,jp,ke,kg,kh,ki,km,kn,kp,kr,kw,ky,kz,la,lb,lc,li,lk,lr,ls,lt,lu,lv,ly,ma,mc,md,mg,mh,mk,ml,mm,mn,mo,mp,mq,mr,ms,mt,mu,mv,mw,mx,my,mz,na,nc,ne,nf,ng,ni,nl,no,np,nr,nu,nz,om,pa,pe,pf,pg,ph,pk,pl,pm,pn,pr,ps,pt,pw,py,qa,re,ro,ru,rw,sa,sb,sc,sd,se,sg,sh,si,sj,sk,sl,sm,sn,so,sr,st,sv,sy,sz,tc,td,tf,tg,th,tj,tk,tm,tn,to,tp,tr,tt,tv,tw,tz,ua,ug,uk,um,us,uy,uz,va,vc,ve,vg,vi,vn,vu,wf,ws,ye,yt,yu,za,zm,zr,zw"
'  sTopLevelDomainsArray = Split(sTopLevelDomains, ",")
'
'  eMailSplices = Split(mailAddress, "@")
'
'  If UBound(eMailSplices) <> 1 Then Exit Function
'
'  eMailSplices = Split(eMailSplices(1), ".")
'
'  If UBound(eMailSplices) < 1 Then Exit Function
'
'  For i = 0 To UBound(sTopLevelDomainsArray)
'    If eMailSplices(UBound(eMailSplices)) = sTopLevelDomainsArray(i) Then
'      bGoodAdress = True
'      Exit For
'    End If
'  Next i
'  If bGoodAdress = False Then Exit Function
'
'  CheckMailAdress = True
End Function

