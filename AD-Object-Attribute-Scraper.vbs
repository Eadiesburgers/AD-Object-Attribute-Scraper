option explicit

'Author: Rob Lawton
'Version: 2.3 back to cscript specfic version which is quicker and imho much better.
'Date: 01-Sept-2023
'Usage: Will search AD for objects of type User, Group, OU Container or Computer types, against either the samaccountname or name attribute value.  
'   - If object(s) is/are found, the schema is walked and each attribute is processed as per schema configuration.  The data is then written to the file specified by the user.
'   - Wildcard searches are supported, beware of where you situate your wildcard character "*".  Consider the walfare of the domain controllers, result set size and time taken to complete query.
'Info - IF the vbs filename extension is correctly configured on your machine, simply double-click the vbs file in Windows Explorer, which should display help text on how to execute using cscript.
'   - Failing that using a DOS command shell windows (CMD), execute script using:  "cscript.exe FILENAME.VBS"
'   - Aware that this could be scripted in a much more efficient manner, the method employed ensures it is self eductaing as the lookups are easy to see via Const and Case.
'------------------------------------------------------------

if lcase(CreateObject("Scripting.FileSystemObject").GetFileName( WScript.FullName )) <> "cscript.exe" then
    msgbox("Execute this script using cscript within a CMD shell window." & vbcrlf & vbcrlf & _
    "Click [START]" & vbcrlf & vbcrlf &_
    "Then type: CMD" & vbcrlf & vbcrlf & _
    "From within this DOS CMD window execute the script using:" & vbcrlf & vbcrlf & _
    "cscript.exe FILENAME.VBS" & vbcrlf & " (where FILENAME is the name of your script file)." & vbcrlf & vbcrlf & _
    "Example:" & vbcrlf & "cscript.exe " & chr(34) & "c:\My Files\Scripts\My AD Query_Script.vbs" & chr(34))
else
dim lineSplit,tabSplit,msgOut,i,objRootDSE,domainVar,cnt,exitDoVar,strPath
dim objFSO,adoCommand,adoConnection,strFileName,strFilter,strBase,strQuery,recCount,adoRecordset,totalRecCount,userVar,passVar
dim args,x,domain2,objDC,outInfo
tabSplit="      "
msgOut="<Enter " & chr(34) & "Quit" & chr(34) & " to quit>" & vbcrlf & _
"<Enter " & chr(34) & "Help" & chr(34) & " for help>"
for i=0 to 100
    lineSplit=lineSplit & "_"
next 
lineSplit = lineSplit & vbcrlf
dim inputArray(3)
if wscript.arguments.count < 4 then
	wscript.echo "Need you to supply the following (not case sensitive):" & vbcrlf & _ 
    tabSplit & "Param 1. - what object type to look for.  Enter one of the following Options:" & vbcrlf & tabSplit & tabSplit & tabSplit & "> user" & vbcrlf & tabSplit & tabSplit & tabSplit & "> group" & vbcrlf & tabSplit & tabSplit & tabSplit & _ 
    "> computer" & vbcrlf & tabSplit & tabSplit & tabSplit & "> container (ou container)" & vbcrlf & _
    vbcrlf & tabSplit & "Param 2. - what attribute to search against.  Options universal for all above object types.  These are: 'account' or 'name'." & vbcrlf & _
    vbcrlf & tabSplit & "Param 3. - what do you want to search for.  Wildcards thus multiple records are supported, wrap in " & chr(34) & "quotes" & chr(34) & ". Be careful where you put your wildcards, this can impact time taken and quantity of records returned." & vbcrlf & _
    vbcrlf & tabSplit & "Param 4. - where do you want to write the output file to. Wrap in " & chr(34) & "quotes" & chr(34) & " if your directory or fiilename contains spaces."
else
	Set args = WScript.Arguments
	for x=0 to 3
        inputArray(x)=args(x)
    next 
	if not args is Nothing then set args=Nothing
    Set objRootDSE = getObject("LDAP://RootDSE") 
    domainVar=replace(mid(objRootDSE.get("defaultNamingContext"),4,len(objRootDSE.get("defaultNamingContext"))),",DC=",".")
    strFileName= Mid(inputArray(3),InStrRev(inputArray(3), "\") + 1)
    if len(strFileName)=0 then 
        wscript.echo "Valid path and filename required."
    else
        On Error Goto 0
        strPath=replace(inputArray(3),strFileName,"")
        if fcnCheckFolder(strPath) Then  
            if fcnCheckFile(inputArray(3)) Then
                wscript.echo strFileName & " already exists in " & strPath & ". Delete this file or use a different filename."
            Else    
                if lcase(inputArray(1))="account" then inputArray(1)="samaccountname"
                select case lcase(inputArray(0)) 
                    case "user"
                        strFilter="(&(objectCategory=person)(objectClass=user)(" & inputArray(1) & "=" & inputArray(2) & ")(!(objectclass=contact)));distinguishedname;subtree"
                    case "group"
                        strFilter="(&(objectCategory=group)(objectClass=group)(" & inputArray(1) & "=" & inputArray(2) & "));distinguishedname;subtree"
                    case "computer"
                        if lcase(inputArray(1))="samaccountname" and right(inputArray(2),1) <> "$" then inputArray(2)=inputArray(2) & "$"
                        strFilter="(&(objectCategory=computer)(objectClass=computer)(" & inputArray(1) & "=" & inputArray(2) & "));distinguishedname;subtree"
                    case "container"
                        strFilter="(&(objectClass=organizationalUnit)(" & inputArray(1) & "=" & inputArray(2) & "));distinguishedname;subtree"
                    case Else
                        wscript.echo inputArray(0) & " is not an optional object type [User. Group or Computer]."
                end select
                Set objFSO = CreateObject("ADODB.Stream")
                objFSO.CharSet = "utf-16"
                objFSO.Open
                OutInfo="Querying aginst:" & vbcrlf & _
                tabSplit & "Object type: " & inputArray(0) & vbcrlf & _
                tabSplit & "Attribute to search against: " & inputArray(1) & vbcrlf & _
                tabSplit & "Value to search for: " & chr(34) & inputArray(2) & chr(34) & vbcrlf & _
                tabSplit & "Filename to write to: " & chr(34) & inputArray(3) & chr(34)
                wscript.echo OutInfo
                objFSO.WriteText OutInfo & vbcrlf & lineSplit
                wscript.echo "Getting domain controller list for " & domainVar & "..."
                Const adUseClient = 3
                Set adoCommand = CreateObject("ADODB.Command")
                Set adoConnection = CreateObject("ADODB.Connection")
                adoConnection.Provider = "ADsDSOObject"
                adoConnection.cursorLocation = adUseClient
                strBase="<LDAP://" & domainVar & ">"    
                adoConnection.Open "Active Directory Provider"
                Set adoCommand.ActiveConnection = adoConnection    
                domain2="DC=" & replace(domainVar,".",",DC=")
                strQuery="<LDAP://CN=Configuration," & domain2 & ">;(objectClass=nTDSDSA);AdsPath;subtree"
                adoCommand.CommandText = strQuery       
                adoCommand.Properties("Page Size") = 500
                adoCommand.Properties("Timeout") = 60
                adoCommand.Properties("Cache Results") = False
                adoCommand.Properties("Asynchronous") = False 
                Set adoRecordset = adoCommand.Execute
                if adoRecordset.EOF Then
                    objFSO.WriteText  "No DC query results returned, will continue with the main meal..."
                Else    
                    wscript.echo "Located " & adoRecordset.recordcount & " domain controllers for " & domainVar & "..."
                    objFSO.WriteText "Located " & adoRecordset.recordcount & " domain controllers for " & domainVar & ":" & vbcrlf
                     Do Until adoRecordset.EOF
                        Set objDC = GetObject(GetObject(adoRecordset.Fields("AdsPath").Value).Parent)
                        objFSO.WriteText tabSplit & objDC.DNSHostName & vbcrlf
                        adoRecordset.MoveNext
                    Loop
                end if
                objFSO.WriteText lineSplit
                call FcnDCStuff(objFSO) 
                objFSO.WriteText lineSplit
                if not adoRecordset is Nothing then set adoRecordset=Nothing
                strQuery = strBase & ";" & strFilter
                adoCommand.CommandText = strQuery   
                Set adoRecordset = adoCommand.Execute
                if adoRecordset.EOF then
                    wscript.echo "No record(s) found based on the information supplied"
                Else
                    adoRecordset.MoveFirst 
                    totalRecCount=adoRecordset.recordcount
                    outInfo="Query returned " & totalRecCount & " record(s)." 
                    wscript.echo OutInfo
                    objFSO.WriteText OutInfo & vbcrlf & lineSplit
                    adoRecordset.MoveFirst 
                    Do Until adoRecordset.EOF
                        recCount=recCount+1
                        outInfo="Wait a moment, processing Record #" & recCount & " of " & totalRecCount & "..."
                        wscript.echo outInfo
                        objFSO.WriteText outInfo & vbcrlf & lineSplit
                        call FcnProcess1(adoRecordset.Fields("distinguishedname").value,objFSO,domainVar,lineSplit,tabSplit)
                        adoRecordset.movenext
                    Loop
                    objFSO.SaveToFile inputArray(3), 2
                    objFSO.Close
                    if not objFSO is Nothing then set objFSO=Nothing
                        wscript.echo "Finished.  See file: " & inputArray(3)
                    end if
                end if
            Else
                wscript.echo strPath & " does not exist or is not accessible."
            end if 
        end if
    end if
end if
'------------------------------------------------------------
Function FcnProcess1(byref strTargetGroupDN,objFSOIN,domainVar,lineSplit,tabSplit)
    dim obj,ADSTypeNameList,value,ADSTypeNameArray,attrTypeName
    dim oMSyntaxInt,oMSyntaxChar,isReplicated,valuearray,outputVal,rowcountVar
    dim hstr,data,otherTypes
    dim objDate,FcnInteger8
    dim lngHigh,lngLow,lngAdjust,lngDate,AttribCount
    dim robj,samaccounttype
    Const ADSTYPE_INVALID=0
    Const ADSTYPE_DN_STRING=1
    Const ADSTYPE_CASE_EXACT_STRING=2
    Const ADSTYPE_CASE_IGNORE_STRING=3
    Const ADSTYPE_PRINTABLE_STRING=4
    Const ADSTYPE_NUMERIC_STRING=5
    Const ADSTYPE_BOOLEAN=6
    Const ADSTYPE_INTEGER=7
    Const ADSTYPE_OCTET_STRING=8
    Const ADSTYPE_UTC_TIME=9
    Const ADSTYPE_LARGE_INTEGER=10
    Const ADSTYPE_OBJECT_CLASS=11
    Const ADSTYPE_PROV_SPECIFIC=12
    Const ADSTYPE_CASEIGNORE_LIST=13
    Const ADSTYPE_OCTET_LIST=14
    Const ADSTYPE_PATH=15
    Const ADSTYPE_POSTALADDRESS=16
    Const ADSTYPE_TIMESTAMP=17
    Const ADSTYPE_BACKLINK=18
    Const ADSTYPE_TYPEDNAME=19
    Const ADSTYPE_HOLD=20
    Const ADSTYPE_NETADDRESS=21
    Const ADSTYPE_REPLICAPOINTER=22
    Const ADSTYPE_FAXNUMBER=23
    Const ADSTYPE_EMAIL=24
    Const ADSTYPE_NT_SECURITY_DESCRI=25
    Const ADSTYPE_DN_WITH_BINARY=26
    Const ADSTYPE_DN_WITH_STRING=27
    dim OMSyntaxArray(127) 
    OMSyntaxArray(1)="Boolean"
    OMSyntaxArray(2)="Integer"
    OMSyntaxArray(4)="OctetString"
    OMSyntaxArray(6)="OID"
    OMSyntaxArray(20)="CaseIgnoreString"
    OMSyntaxArray(23)="UTC"
    OMSyntaxArray(24)="GeneralizedTime"
    OMSyntaxArray(64)="DirectoryString"
    OMSyntaxArray(65)="Integer8"
    OMSyntaxArray(66)="NTSecurityDescriptor"
    OMSyntaxArray(127)="DN"
    Set obj = GetObject ("LDAP://" & strTargetGroupDN)
    obj.GetInfo
    objFSOIN.WriteText "Attributes found count: [" & obj.PropertyCount & "]." & vbcrlf & lineSplit
    ADSTypeNameList="ADS_DN_STRING,ADS_CASE_EXACT_STRING,ADS_CASE_IGNORE_STRING,ADS_PRINTABLE_STRING,ADS_NUMERIC_STRING,ADS_BOOLEAN,ADS_INTEGER,ADS_OCTET_STRING,ADS_UTC_TIME,ADS_LARGE_INTEGER,ADS_OBJECT_CLASS,ADS_PROV_SPECIFIC,PADS_CASEIGNORE_LIST,PADS_OCTET_LIST,PADS_PATH,PADS_POSTALADDRESS,ADS_TIMESTAMP,ADS_BACKLINK,PADS_TYPEDNAME,ADS_HOLD,PADS_NETADDRESS,PADS_REPLICAPOINTER,PADS_FAXNUMBER,ADS_EMAIL,ADS_NT_SECURITY_DESCRI,PADS_DN_WITH_BINARY,PADS_DN_WITH_STRING"
    ADSTypeNameArray=split(ADSTypeNameList,",") 
    For i = 0 To obj.PropertyCount-1
        Select Case obj.Item(i).ADStype
            Case ADSTYPE_DN_STRING
                attrTypeName = "DN String"
            Case ADSTYPE_CASE_EXACT_STRING
                attrTypeName = "CaseExact String"
            Case ADSTYPE_CASE_IGNORE_STRING
                attrTypeName = "CaseIgnore String"
            Case ADSTYPE_PRINTABLE_STRING
                attrTypeName = "Printable String"
            Case ADSTYPE_NUMERIC_STRING
                attrTypeName = "Numeric String"
            Case ADSTYPE_BOOLEAN
                attrTypeName = "Boolean"
            Case ADSTYPE_INTEGER
                attrTypeName = "Integer"
            Case ADSTYPE_OCTET_STRING
                attrTypeName = "Octet-String"
            Case ADSTYPE_UTC_TIME
                attrTypeName = "UTC Time"
            Case ADSTYPE_LARGE_INTEGER
                attrTypeName = "Large Integer" 
            Case ADSTYPE_PROV_SPECIFIC
                attrTypeName = "Provider Specific"
            Case ADSTYPE_NT_SECURITY_DESCRI
                attrTypeName = "NT Security Descriptor"
            Case Else
                attrTypeName = "Other Type:" & CStr(obj.item(i).adstype)
        End Select
    oMSyntaxInt=fcnGetSchema(obj.Item(i).Name,"oMSyntax",objFSOIN,domainVar)
    oMSyntaxChar=OMSyntaxArray(oMSyntaxInt)
    isReplicated=fcnGetSchema(obj.Item(i).Name,"isMemberOfPartialAttributeSet",objFSOIN,domainVar)
    if isnull(isReplicateD) then isReplicated="Attribute not present, not replicated."
    objFSOIN.WriteText "Is attribute replicated? (isMemberofPartialAttributeSet)=" & isReplicated & vbcrlf
    objFSOIN.WriteText "Attribute name=[" & obj.Item(i).Name & "], ADSType.int=[" & obj.Item(i).ADStype & "], ADSType.char=[" & ADSTypeNameArray((obj.Item(i).ADStype)-1) & "], oMSyntax.int=[" & oMSyntaxInt & "], oMSyntax.char=[" & oMSyntaxChar & "]" & vbcrlf & _
    "Attribute data:" & vbcrlf
        Select Case obj.Item(i).ADStype
            Case ADSTYPE_INTEGER,ADSTYPE_DN_STRING, ADSTYPE_CASE_EXACT_STRING, ADSTYPE_CASE_IGNORE_STRING, _
                ADSTYPE_PRINTABLE_STRING, ADSTYPE_NUMERIC_STRING, _
                ADSTYPE_UTC_TIME, ADSTYPE_PRINTABLE_STRING, ADSTYPE_POSTALADDRESS, _
                ADSTYPE_FAXNUMBER, ADSTYPE_EMAIL, ADSTYPE_PATH, ADSTYPE_NETADDRESS
                valuearray = obj.GetEx(obj.Item(i).name)
                outputVal=""
                rowcountVar=0
                if lcase(obj.Item(i).name)="member" then 
                    objFSOIN.WriteText  tabSplit & "Members:" & vbcrlf
                    for each value in valuearray
                        set robj=GetObject ("LDAP://" & value)
                        select case robj.samaccounttype
                            case 268435456 
                                samaccounttype="SAM_GROUP_OBJECT"
                            case 268435457 
                                samaccounttype="SAM_NON_SECURITY_GROUP_OBJECT"
                            case 536870912 
                                samaccounttype="SAM_ALIAS_OBJECT"
                            case 536870913 
                                samaccounttype="SAM_NON_SECURITY_ALIAS_OBJECT"
                            case 805306368 
                                samaccounttype="SAM_NORMAL_USER_ACCOUNT"
                            case 805306369 
                                samaccounttype="SAM_MACHINE_ACCOUNT"
                            case 805306370 
                                samaccounttype="SAM_TRUST_ACCOUNT"
                            case 1073741824 
                                samaccounttype="SAM_APP_BASIC_GROUP"
                            case 1073741825 
                                samaccounttype="SAM_APP_QUERY_GROUP"
                            case 2147483647 
                                samaccounttype="SAM_ACCOUNT_TYPE_MAX"
                        end select
                        objFSOIN.WriteText  tabSplit & value & " Type Int==>" & robj.samaccounttype & "  Type Chr==>" & samaccounttype & vbcrlf
                    next
                    if not robj is Nothing then set robj=Nothing
                else
                For Each value In valuearray 
                    rowcountVar=rowcountVar+1
                    outputVal=outputVal & value
                    if rowcountVar > 1 then outputVal=outputVal & vbcrlf
                    objFSOIN.WriteText  tabSplit & value & vbcrlf 
                    if lcase(obj.Item(i).Name) = "useraccountcontrol" then call fcnUserAccountControlLookup(value,objFSOIN,tabSplit)
                Next
            end if
            objFSOIN.WriteText lineSplit
            Case ADSTYPE_LARGE_INTEGER
        valuearray = obj.GetEx(obj.Item(i).name)
        AttribCount=0
        for each value in valuearray 
            AttribCount=AttribCount+1
            set objDate=value
            lngHigh=objDate.HighPart
            lngLow=objDate.LowPart
            objFSOIN.WriteText tabSplit & "Value #"  & AttribCount & " for attribute " & obj.Item(i).name & vbcrlf 
            objFSOIN.WriteText tabSplit & "Raw data Integer8 High part:" & lngHigh & vbcrlf
            objFSOIN.WriteText tabSplit & "Raw data Integer8 Low part:" & lngLow & vbcrlf 
            Select Case lcase(obj.Item(i).name)
                case "accountexpires","badpasswordtime","lastlogoff","lastlogon","lastlogontimestamp","lockoutduration","lockoutobservationwindow", _
                        "lockouttime","maxpwdage","minpwdage","msds-lastfailed","interactivelogontime","msds-lastsuccessful","interactivelogontime","msds-userpassword", _
                        "expirytimecomputed","pwdlastset"
                    If (lngLow < 0) Then lngHigh = lngHigh + 1
                    If (lngHigh = 0) And (lngLow = 0) Then lngAdjust = 0
                    lngDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) + lngLow) / 600000000 - lngAdjust) / 1440
                    On Error Resume Next
                    FcnInteger8= CDate(lngDate) & " UTC"
                    If (Err.Number <> 0) Then
                        On Error GoTo 0
                        FcnInteger8 = #1/1/1601#
                    End If
                    On Error GoTo 0
                    objFSOIN.WriteText  tabSplit & "Processed as Integer8 Date/Time:" & FcnInteger8 & vbcrlf
                case Else
                    objFSOIN.WriteText  tabSplit & "Processed as Integer8 Other:" & lngLow & lngHigh & vbcrlf
                End select
            objFSOIN.WriteText lineSplit
        next
            Case ADSTYPE_BOOLEAN
                valuearray = obj.GetEx(obj.Item(i).name)
                For Each value In valuearray
                    If (value) Then 
                        objFSOIN.WriteText tabSplit & "TRUE" & vbcrlf
                    Else 
                        objFSOIN.WriteText tabSplit & "FALSE" & vbcrlf 
                    end if
                Next
                objFSOIN.WriteText lineSplit
            Case ADSTYPE_OCTET_STRING
                if lcase(obj.Item(i).Name) = "thumbnailphoto" then
                    objFSOIN.WriteText  tabSplit & "Ignoring thumbnail photo stream, unremark in script if you really want this. ('_')" & vbcrlf 
                else
                    valuearray = obj.GetEx(obj.Item(i).name)
                        objFSOIN.WriteText vbcrlf
                        For Each value In valuearray
                            hstr = OctetToHexStr(value)
                            objFSOIN.WriteText fcnPrintOutHex(hstr,8,tabSplit ) & vbcrlf
                        Next
                    objFSOIN.WriteText lineSplit
                end if
            Case ADSTYPE_PROV_SPECIFIC
                Set prop = obj.GetPropertyItem(obj.Item(i).name, ADSTYPE_OCTET_STRING)
                valuearray = prop.Values
                For Each value In valuearray 
                    data = value.OctetString
                    hstr = OctetToHexStr(data)
                    objFSOIN.WriteText tabSplit & fcnprintOutHex(hstr,8,tabSplit) & vbcrlf
                Next
                objFSOIN.WriteText lineSplit
            case ADSTYPE_DN_WITH_STRING     
                objFSOIN.WriteText "ADSTYPE_DN_WITH_STRING" & vbcrlf
                valuearray = obj.GetEx(obj.Item(i).name)
                objFSOIN.WriteText tabSplit & obj.Item(i).name & "=>" & ADSTypeNameArray((obj.Item(i).ADStype)-1) & vbcrlf
                objFSOIN.WriteText lineSplit
            case ADSTYPE_NT_SECURITY_DESCRI
                objFSOIN.WriteText tabSplit & obj.Item(i).name & " => " & ADSTypeNameArray((obj.Item(i).ADStype)-1) & vbcrlf
                call fcnNTSecDescriptor(obj,objFSOIN,tabSplit)
                objFSOIN.WriteText lineSplit
            case ADSTYPE_INVALID
                objFSOIN.WriteText "Invalid type" & vbcrlf 
                objFSOIN.WriteText lineSplit
            case Else
                otherTypes="ADSTYPE_OBJECT_CLASS=11" & vbcrlf & _
                "ADSTYPE_CASEIGNORE_LIST=13" & vbcrlf & _
                "ADSTYPE_OCTET_LIST=14" & vbcrlf & _ 
                "ADSTYPE_TIMESTAMP=17" & vbcrlf & _
                "ADSTYPE_BACKLINK=18" & vbcrlf & _
                "ADSTYPE_TYPEDNAME=19" & vbcrlf & _
                "ADSTYPE_HOLD=20" & vbcrlf & _
                "ADSTYPE_REPLICAPOINTER=22" & vbcrlf & _
                "ADSTYPE_DN_WITH_BINARY=26" & vbcrlf & _
                "ADSTYPE_DN_WITH_STRING=27"
                objFSOIN.WriteText tabSplit & "Working on these other types which could most possible be processed by one of the existing defined data handlers..." & vbcrlf & otherTypes
                objFSOIN.WriteText lineSplit
        End Select
    Next
End Function
'------------------------------------------------------------
Function OctetToHexStr(var_octet)
Dim n
    OctetToHexStr = ""
    For n = 1 To lenb(var_octet)
        OctetToHexStr = OctetToHexStr & Right("0" & hex(ascb(midb(var_octet, n, 1))), 2)
    Next
End Function
'------------------------------------------------------------
Function fcnPrintoutHex(byref var_hex, width,tabSplit)
    On Error Goto 0
    Dim k1, k2, s1, s2, s3
    fcnPrintOutHex = ""
    For k1 = 1 To Len(var_hex) Step (width *2)
        s1 = Mid(var_hex, k1, (width *2))
        s2 = ""
        s3 = HexStrToAscii(s1,  False)
        For k2 = 1 To Len(s1) Step 2
            s2 = S2 & Mid(S1, k2, 2) & " "
        Next
        s2 = tabSplit & s2 & String((width *3)-Len(s2), " ")
        If (k1=1) Then 
            fcnPrintOutHex = fcnPrintOutHex & s2 & "| " & s3
        Else
            fcnPrintOutHex = fcnPrintOutHex & vbcrlf & s2 & "| " & s3
        End If
    Next
End Function
'------------------------------------------------------------
Function HexStrToAscii(byref var_hex, format)
On Error Goto 0
    Dim k,  v
    HexStrToAscii = ""
    For k = 1 To Len(var_hex) Step 2
        v = CInt("&H" & Mid(var_hex, k, 2))
        If ((v>31) And (v<128)) Then 
            HexStrToAscii = HexStrToAscii & (chr(v))
        Else
            If (format) Then
                Select Case v
                    Case 8
                        HexStrToAscii = HexStrToAscii & vbTab
                    Case 10
                        HexStrToAscii = HexStrToAscii & vbCrLf
                    Case 13
                    Case Else
                        HexStrToAscii = HexStrToAscii & "."
                End Select
            Else
                HexStrToAscii = HexStrToAscii & "."
            End If
        End If
    Next
End Function
'------------------------------------------------------------
Function fcnUserAccountControlLookup (byref intVar,objFSOIN,splitIn)
On Error Goto 0
    dim binVar,x,newbinVar,arrayposVar,resultVar,countupVar,bitVar
     objFSOIN.WriteText splitIn & "UserAccountControl flag bitwise operation evidence:" & vbcrlf
    if not isnumeric(intVar) then
        objFSOIN.WriteText splitIn & "Detected a non-integer value!" & vbcrlf
    else
        objFSOIN.WriteText splitIn & "Value: [" & intVar & "].  This equates to the following:" & vbcrlf
        binVar=DecimalToBinary(intVar)
        objFSOIN.WriteText splitIn & "Signed integer 4 byte (32-bit) two's complement binary:" & binVar & vbcrlf
        objFSOIN.WriteText splitIn & "Calculated start at offset (0-31):" & 31-len(binVar) & vbcrlf
        for x=0 to 31-len(binVar)
            newbinVar=newbinVar & 0
        next
        objFSOIN.WriteText splitIn & "Padding bits:" & newbinVar & vbcrlf
        objFSOIN.WriteText splitIn & "MSB order:" &  newbinVar & binVar & vbcrlf
        binVar = newbinVar & binVar
        binVar=strreverse(binVar)
        objFSOIN.WriteText splitIn & "LSB order:" & binVar & vbcrlf
        const useraccountcontrolVar="UNUSED_MUST_BE_ZERO-IGNORED,UNUSED_MUST_BE_ZERO-IGNORED,UNUSED_MUST_BE_ZERO-IGNORED,UNUSED_MUST_BE_ZERO-IGNORED,UNUSED_MUST_BE_ZERO-IGNORED,ADS_UF_PARTIAL_SECRETS_ACCOUNT,ADS_UF_NO_AUTH_DATA_REQUIRED,ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION,ADS_UF_PASSWORD_EXPIRED,ADS_UF_DONT_REQUIRE_PREAUTH,ADS_UF_USE_DES_KEY_ONLY,ADS_UF_NOT_DELEGATED,ADS_UF_TRUSTED_FOR_DELEGATION,ADS_UF_SMARTCARD_REQUIRED,UF_MNS_LOGON_ACCOUNT (apparently not UNUSED),ADS_UF_DONT_EXPIRE_PASSWD,UNUSED_MUST_BE_ZERO-IGNORED,UNUSED_MUST_BE_ZERO-IGNORED,ADS_UF_SERVER_TRUST_ACCOUNT,ADS_UF_WORKSTATION_TRUST_ACCOUNT,ADS_UF_INTERDOMAIN_TRUST_ACCOUNT,UNUSED_MUST_BE_ZERO-IGNORED,ADS_UF_NORMAL_ACCOUNT,UNUSED_MUST_BE_ZERO-IGNORED,ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED,ADS_UF_PASSWD_CANT_CHANGE,ADS_UF_PASSWD_NOTREQD,ADS_UF_LOCKOUT,ADS_UF_HOMEDIR_REQUIRED,UNUSED_MUST_BE_ZERO-IGNORED,ADS_UF_ACCOUNT_DISABLE,UNUSED_MUST_BE_ZERO-IGNORED"
        dim useraccountcontrolArray,useraccountControlFlagArray
        useraccountcontrolArray=split(useraccountcontrolVar,",")
        const useraccountControlFlagVar= "X_,X_,X_,X_,X_,PS,NA,TA,PE,DR,DK,ND,TD,SR,X?,DP,X_,X_,ST,WT,ID,X_,N_,X_,ET,CC,NR,L_,HR,X_,D_,X_"
        useraccountControlFlagArray=split(useraccountControlFlagVar,",")
        countupVar=0
        arrayposVar=0
        for x = len(binVar) to 1 step -1 
            bitVar = "0" & arrayposVar & " - "
            if (arrayposVar) > 9 then bitVar = arrayposVar & " - "
            if mid(binVar,x,1) = 1 then 
                'resultVar = resultVar & vbcrlf & useraccountcontrolArray(arrayposVar)
                resultVar = resultVar & "   " & useraccountcontrolArray(arrayposVar)
                objFSOIN.WriteText splitIn & " * bit " & bitVar & "[" & mid(binVar,x,1) & "] - " & useraccountControlFlagArray(arrayposVar) & " - " & useraccountcontrolArray(arrayposVar) & vbcrlf
            else
                objFSOIN.WriteText splitIn & "   bit " & bitVar & "[" & mid(binVar,x,1) & "] - " & useraccountControlFlagArray(arrayposVar) & " - " & useraccountcontrolArray(arrayposVar) & vbcrlf
            end if 
            arrayposVar = arrayposVar + 1
        next
            if arrayposVar < 31 then
                for x = arrayposVar to 31
                    bitVar = "  bit 0" & x & " - "
                    if len(x) > 1 then bitVar = "  bit " & x & " - " 
                    objFSOIN.WriteText splitIn & bitvar & "[0] - " & useraccountControlFlagArray(x) & " - " & useraccountcontrolArray(x) & vbcrlf
                next
            end if
        objFSOIN.WriteText splitIn & "Result: " & resultVar & vbcrlf
    end if
End Function
'------------------------------------------------------------
Function DecimalToBinary(byref intDecimal)
On Error Goto 0
    Dim strBinary, lngNumber1, lngNumber2, strDigit
    strBinary = ""
                intDecimal = cDbl(intDecimal)
    While (intDecimal > 1)
        lngNumber1 = intDecimal / 2
        lngNumber2 = Fix(lngNumber1)
        If (lngNumber1 > lngNumber2) Then
            strDigit = "1"
        Else
            strDigit = "0"
        End If
        strBinary = strDigit & strBinary
        intDecimal = Fix(intDecimal / 2)
    Wend
    strBinary = "1" & strBinary
    DecimalToBinary = strBinary
End Function
'------------------------------------------------------------
Function fcnGetSchema(byref AttribName,AttribToGet,objFSOIN,domainVar)
On Error Goto 0
    dim strDNSDomain
    dim objConnection,objCommand,objRecordSet
    domainVar=replace(domainVar,".",",DC=")
    strDNSDomain="CN=Schema,CN=Configuration," & "DC=" & domainVar
    Set objConnection = createObject("ADODB.Connection") 
    Set objCommand = createObject("ADODB.Command") 
    objConnection.Provider = "ADsDSOObject" 
    objConnection.Open "Active Directory Provider" 
    Set objCOmmand.ActiveConnection = objConnection 
    objCommand.CommandText ="SELECT " & AttribToGet & " FROM 'LDAP://" & strDNSDomain & "' WHERE lDAPDisplayName='" & AttribName & "'"
    Set objRecordSet = objCommand.execute
    if objRecordSet.EOF then
		fcnGetSchema="NOT FOUND"
	else
        objRecordSet.MoveFirst
        fcnGetSchema=objRecordSet.fields(AttribToGet).value
	end if
    if not objRootDSE is Nothing then set objRootDSE=Nothing
    if not objConnection is Nothing then set objConnection=Nothing
    if not objCommand is Nothing then set objCommand=Nothing
    if not objRecordSet is Nothing then set objRecordSet=Nothing
End Function
'------------------------------------------------------------
Function fcnCheckFile(byref FileIn)
	Dim FSys
    On Error Resume Next
    If (Err.Number <> 0) Then 
        On Error Goto 0
        fcnCheckFile=False
    Else
	    Set FSys = CreateObject("Scripting.FileSystemObject")
	    fcnCheckFile=false
	    if FSys.FileExists(FileIn) then	fcnCheckFile=True
	    if not FSys is Nothing then set FSys=Nothing
    end if
End Function
'------------------------------------------------------------
Function fcnCheckFolder(byref FolderIn)
	Dim FSys
	Set FSys = CreateObject("Scripting.FileSystemObject")
	fcnCheckFolder=False
	if FSys.FolderExists(FolderIn) then	fcnCheckFolder=True
	if not FSys is Nothing then set FSys=Nothing
End Function
'------------------------------------------------------------
Function FcnDCStuff(byref FSIn)
    dim oShell
    Set oShell = GetObject("LDAP://rootDse")
    FSIn.WriteText "Domain:" & oShell.get("rootDomainNamingContext") & vbcrlf
    FSIn.WriteText "DC server responding to this LDAP query via ADSI: [" & left(oShell.Get("dnsHostName"),InStr(oShell.Get("dnsHostName"),".")-1) & "]." & vbcrlf
    if not oShell is Nothing then set oShell=Nothing
    Set oShell = CreateObject( "WScript.Shell" )
    FSIn.WriteText "Configured DC server via system environment variable: [" & mid(oShell.ExpandEnvironmentStrings("%LOGONSERVER%"),3,len(oShell.ExpandEnvironmentStrings("%LOGONSERVER%"))) & "]  (these may not match)." & vbcrlf
    If not oShell is Nothing then set oShell=Nothing
End Function
'------------------------------------------------------------
Function fcnNTSecDescriptor(byref objIn,fileObj,TabIN)
    dim value,objSD,objDACL,ACType
    Const ADS_RIGHT_DELETE = &H10000
    Const ADS_RIGHT_READ_CONTROL = &H20000
    Const ADS_RIGHT_WRITE_DAC = &H40000
    Const ADS_RIGHT_OWNER = &H80000
    Const ADS_RIGHT_SYNCHRONIZE = &H100000
    Const ADS_RIGHT_ACCESS_SYSTEM_SECURITY = &H1000000
    Const ADS_RIGHT_GENERIC_READ = &H80000000
    Const ADS_RIGHT_GENERIC_WRITE = &H40000000
    Const ADS_RIGHT_GENERIC_EXECUTE = &H20000000
    Const ADS_RIGHT_GENERIC_ALL = &H10000000
    Const ADS_RIGHT_DS_CREATE_CHILD = &H1
    Const ADS_RIGHT_DS_DELETE_CHILD = &H2
    Const ADS_RIGHT_ACTRL_DS_LIST = &H4
    Const ADS_RIGHT_DS_SELF = &H8
    Const ADS_RIGHT_DS_READ_PROP = &H10
    Const ADS_RIGHT_DS_WRITE_PROP = &H20
    Const ADS_RIGHT_DS_DELETE_TREE = &H40
    Const ADS_RIGHT_DS_LIST_OBJECT = &H80
    Const ADS_RIGHT_DS_CONTROL_ACCESS = &H100
    Const ADS_ACETYPE_ACCESS_ALLOWED = &H0
    Const ADS_ACETYPE_ACCESS_DENIED = &H1
    Const ADS_ACETYPE_SYSTEM_AUDIT = &H2
    Const ADS_ACETYPE_ACCESS_ALLOWED_OBJECT = &H5
    Const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &H6
    Const ADS_ACETYPE_SYSTEM_AUDIT_OBJECT = &H7
    Set objSD = objIn.Get(objIn.Item(i).name)
    Set objDACL = objSD.DiscretionaryACL
    fileObj.WriteText TabIn & "Enumerating security descriptor..." & vbcrlf
    for each value in objDACL
        Select Case value.AceType
            case ADS_ACETYPE_ACCESS_ALLOWED 
                ACType="Access Allowed"
            case ADS_ACETYPE_ACCESS_DENIED
                ACType="Access Denied"
            case ADS_ACETYPE_SYSTEM_AUDIT 
                ACType="System Audit "
            case ADS_ACETYPE_ACCESS_ALLOWED_OBJECT 
                ACType="Access Allowed"
            case ADS_ACETYPE_ACCESS_DENIED_OBJECT
                ACType="Access Denied"
            case ADS_ACETYPE_SYSTEM_AUDIT_OBJECT 
                ACType="System Audit"
            case else
                ACType="Type Other."
        end select    
        fileObj.WriteText TabIN & TabIN & "Trustee:" & value.Trustee & "|  ACE Type.int:" & value.AceType & "|  ACE Type.char:" & ACType & "|  Mask: " & value.AccessMask & vbcrlf
    next 
End Function
