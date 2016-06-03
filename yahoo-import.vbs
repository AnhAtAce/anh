'https://github.com/AnhAtAce/anh.git

Option Explicit

Function CSVParse(ByVal strLine)
    ' Function to parse comma delimited line and return array
    ' of field values.

    Dim arrFields
    Dim blnIgnore
    Dim intFieldCount
    Dim intCursor
    Dim intStart
    Dim strChar
    Dim strValue
        
    Const QUOTE = """"
    Const QUOTE2 = """"""

    ' Check for empty string and return empty array.
    If (Len(Trim(strLine)) = 0) Then
        CSVParse = Array()
        Exit Function
    End If

    ' Initialize.
    blnIgnore = False
    intFieldCount = 0
    intStart = 1
    arrFields = Array()

    ' Add "," to delimit the last field.
    strLine = strLine & ","

    ' Walk the string.
    For intCursor = 1 To Len(strLine)
        ' Get a character.
        strChar = Mid(strLine, intCursor, 1)
        Select Case strChar
            Case QUOTE
                ' Toggle the ignore flag.
                blnIgnore = Not blnIgnore
            Case ","
                If Not blnIgnore Then
                    ' Add element to the array.
                    ReDim Preserve arrFields(intFieldCount)
                    ' Makes sure the "field" has a non-zero length.
                    If (intCursor - intStart > 0) Then
                        ' Extract the field value.
                        strValue = Mid(strLine, intStart, _
                            intCursor - intStart)
                        ' If it's a quoted string, use Mid to
                        ' remove outer quotes and replace inner
                        ' doubled quotes with single.
                        If (Left(strValue, 1) = QUOTE) Then
                            arrFields(intFieldCount) = _
                                Replace(Mid(strValue, 2, _
                                Len(strValue) - 2), QUOTE2, QUOTE)
                        Else
                            arrFields(intFieldCount) = strValue
                        End If
                    Else
                        ' An empty field is an empty array element.
                        arrFields(intFieldCount) = Empty
                    End If
                    ' increment for next field.
                    intFieldCount = intFieldCount + 1
                    intStart = intCursor + 1
                End If
        End Select
    Next
    ' Return the array.
    CSVParse = arrFields
End Function

Function EncaseQuotes(strText)

	If Len(strText) = 0 Then
		EncaseQuotes = """"""
		Exit Function
	End If

	If InStr(1, strText, """") > 0 Then

		dim a, i
		a = Split(strText)
		For i = 0 to UBound(a)
			If a(i) = """" Then
				a(i) = """"""
			End If
		Next
		strText = Join(a)
	End If
	
	EncaseQuotes = """" & strText & """"
	
End Function

'Open CSV file for parsing
dim objFSO, objFileOrders, objFileMOM, objFileItems, objFileOptions, objFileYahooOrders
dim objFileYahooItems, objWriteYahooOrders, objWriteYahooItems
dim objReadFileMOM, objReadFileOrders, objReadFileItems, objReadFileOptions, strContentsMOM, strContentsOrders, strContentsItems, strContentsOptions
dim arrCSVMOM, arrCSVOrders, arrCSVItems, arrCSVOptions, count, intOptionCount, intOptionCount2
dim linumber, x, intCheckOnce
dim arrOptionsList
redim arrOptionsList(3,0)
Dim intTax, lngItemOrderNo, intLineNo, strProductId, strProductCode, intQuantity, strUnitPrice
Dim blnItemPick, missingFile

Set objFSO = CreateObject("Scripting.FileSystemObject")



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''Notify user of missing required files'''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''Author: Anh T. Nguyen
'''''Date: 03-08-2012
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Check to see if required files exists
If objFSO.FileExists("acekaraoke_mom.csv") = False Then
	missingFile = "acekaraoke_mom.csv"
End If
If objFSO.FileExists("Orders.csv") = False Then
	missingFile = "Orders.csv"
End If
If objFSO.FileExists("Items.csv") = False Then
	missingFile = "Items.csv"
End If
If objFSO.FileExists("Options.csv") = False Then
	missingFile = "Options.csv"
End if
'Notify error if required files does not exist then exit
If missingFile <> "" Then
	Msgbox "File " & chr(34) & missingFile & chr(34) & " is not found in " & chr(34) & objFSO.GetAbsolutePathName(".") & chr(34)
	Wscript.Quit
End if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''End Notify user of missing required files'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Set objFileMOM = objFSO.GetFile("acekaraoke_mom.csv")
Set objFileOrders = objFSO.GetFile("Orders.csv")
Set objFileItems = objFSO.GetFile("Items.csv")
Set objFileOptions = objFSO.GetFile("Options.csv")
Set objFileYahooOrders = objFSO.CreateTextFile("YahooOrders.csv")
Set objFileYahooItems = objFSO.CreateTextFile("YahooItems.csv")

set objFileYahooOrders = nothing
Set objFileYahooItems = nothing

If objFileOrders.Size > 0 And objFileItems.Size > 0 And objFileMOM.Size > 0 And objFileOptions.Size > 0 Then
	Set objReadFileMOM = objFSO.OpenTextFile("acekaraoke_mom.csv", 1)
	Set objReadFileOrders = objFSO.OpenTextFile("Orders.csv", 1)
	Set objReadFileItems = objFSO.OpenTextFile("Items.csv", 1)
	Set objReadFileOptions = objFSO.OpenTextFile("Options.csv")
	Set objWriteYahooOrders = objFSO.OpenTextFile("YahooOrders.csv", 8)
	Set objWriteYahooItems = objFSO.OpenTextFile("YahooItems.csv", 8)
	
	strContentsOrders = objReadFileOrders.ReadLine
	strContentsItems = objReadFileItems.ReadLine
	strContentsOptions = objReadFileOptions.ReadLine
	count = 1
	intOptionCount = 1
	intOptionCount2 = 0
	x = 0
	intCheckOnce = 1
	
	Do While Len(strContentsOrders) <> 0
		intLineNo = 1
		If objReadFileOrders.AtEndOfStream = True Then
			Wscript.Echo "Finished"
			Wscript.Quit
		End If
		strContentsOrders = objReadFileOrders.ReadLine
		strContentsMOM = objReadFileMOM.ReadLine
		
			
		If objReadFileItems.AtEndOfStream <> True And count <> 0 Then
			strContentsItems = objReadFileItems.Readline
			count = 0
		End If

		If objReadFileOptions.AtEndOfStream <> True And intOptionCount <> 0 Then
			strContentsOptions = objReadFileOptions.Readline
			intOptionCount = 0
			intOptionCount2 = 0
		End If

		if Len(strcontentsMOM) = 0 Then
			Wscript.Echo "MOM List and Order List not matching!"
			objReadFileMOM.Close
			objReadFileOrders.Close
			Wscript.Quit
		End If
		
		
		
		arrCSVMOM = CSVParse(strContentsMOM)
		arrCSVOrders = CSVParse(strContentsOrders)
		arrCSVItems = CSVParse(strContentsItems)
		arrCSVOptions = CSVParse(strContentsOptions)
		
				
		if objReadFileOptions.AtEndOfStream <> True Then
			if intCheckOnce = 1 Then
				if Clng(arrCSVOrders(0)) <> Clng(arrCSVItems(0)) Or Clng(arrCSVOrders(0)) > Clng(arrCSVOptions(0)) Then
					MsgBox arrCSVOrders(0) & "dddd"
					MsgBox arrCSVItems(0)
					MsgBox arrCSVOptions(0)
					Wscript.Echo "Options List, Items List and Order List not matching! Redownload the files."
					Wscript.Quit
				End If
				intCheckOnce = 0
			End If
		End If
			
		if UBound(arrCSVOrders) <> 46 Then
			linumber = objReadFileOrders.Line
			Wscript.Echo "Orders.csv Error in line " & linumber - 1 & " Ubound: " & UBound(arrCSVOrders)
			Wscript.Quit
		End If
		
		Do While Len(arrCSVMOM(3)) = 0 And Len(strContentsMOM) <> 0
			strcontentsMOM = objReadFileMOM.Readline
			arrCSVMOM = CSVParse(strContentsMOM)
		Loop

		if arrCSVMOM(31) <> arrCSVOrders(0) Then
			MsgBox arrCSVMOM(31)
			MsgBox arrCSVOrders(0)
			Wscript.Echo "MOM List and Order List not matching!"
			Wscript.Quit
		End If

		'Store Order Information
		Dim lngOrderID, strDate, lngNumericTime, strShipFirstName, strShipLastName
		Dim strShipAddress1, strShipAddress2, strShipCity, strShipState, strShipCountry
		Dim arrShipCountry, strShipZipCode, strShipTelephone, strBillFirstName
		Dim strBillLastName, strBillAddress1, strBillAddress2, strBillCity
		Dim strBillState, strBillCountry, arrBillCountry, strBillZipCode, strBillTelephone
		Dim strEmail, strReferringPage, strEntryPoint, strShipping, strPaymentMethod
		Dim lngCardNumber, strCardExpiry, strComments, strTotal, strLinkFrom
		Dim strWarning, strAuthcode, strAVSCode, strGiftMessage
		Dim i
		
		'trim all fields
		for i=0 to UBound(arrCSVOrders)
			arrCSVOrders(i) = Trim(arrCSVOrders(i))
		next

		'fill in blank ship name
		if arrCSVMOM(43) = "" Then
				arrCSVMOM(43)=arrCSVMOM(3)
				arrCSVMOM(42)=arrCSVMOM(2)
		End If

		
	        'Copy Shipping Info To Billing Info if Billing info is Blank
	
		if len(arrCSVOrders(12)) = 0 Then
			if len(arrCSVOrders(4)) > 0 Then 
				arrCSVOrders(12) = arrCSVOrders(4)
				arrCSVOrders(13) = arrCSVOrders(5)
				arrCSVOrders(14) = arrCSVOrders(6)
				arrCSVOrders(15) = arrCSVOrders(7)
				arrCSVOrders(16) = arrCSVOrders(8)
				arrCSVOrders(17) = arrCSVOrders(9)
			End if
		End If

		''Copy Billing Info To Shipping Info if Shipping info is Blank
		if len(arrCSVOrders(4)) = 0 Then
			if len(arrCSVOrders(12)) > 0 Then
				arrCSVOrders(4) = arrCSVOrders(12)
				arrCSVOrders(5) = arrCSVOrders(13)
				arrCSVOrders(6) = arrCSVOrders(14)
				arrCSVOrders(7) = arrCSVOrders(15)
				arrCSVOrders(8) = arrCSVOrders(16)
				arrCSVOrders(9) = arrCSVOrders(17)
			End If
		End If
		
		
		
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''International Orders Implementation''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''Author: Anh T. Nguyen
'''''Date: 04-02-2014
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Switch shipping address for International Checkout
		if 		StrComp(LCase(Trim(arrCSVOrders(12))),LCase("7950 Woodley Ave.")) = 0 _
			And StrComp(LCase(Trim(arrCSVOrders(13))),LCase("Suite C")) = 0 _
			And StrComp(LCase(Trim(arrCSVOrders(14))),LCase("Las Vegas")) = 0 _
			And StrComp(LCase(Trim(arrCSVOrders(15))),LCase("NV")) = 0 _
			And StrComp(Trim(arrCSVOrders(17)),"89109") = 0 Then
			Msgbox("Address revised")
			arrCSVOrders(14) = "Van Nuys"	'City
			arrCSVOrders(15) = "CA"	'State
			arrCSVOrders(16) = arrCSVOrders(8)	'Country
			arrCSVOrders(17) = "91406 "	'Zip code
		End if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''End of International Orders Implementation'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		'No shipping or billing address, most likely paypal gift certificate orders to be sent to an email address
		if len(arrCSVOrders(9)) = 0 And len(arrCSVOrders(17)) = 0 Then
			arrCSVOrders(9) = "91746"
			arrCSVOrders(17) = "91746"
			arrCSVOrders(8) = "US United States"
			arrCSVOrders(16) = "US United States"
			arrCSVOrders(7) = "CA"
			arrCSVOrders(15) = "CA"
			arrCSVOrders(6) = "City Of Industry"
			arrCSVOrders(14) = "City Of Industry"
		End If
		
		'Transform Shippind Method field to correct format		
		if arrCSVOrders(22) = "UPS Next Day Air" Then
			arrCSVOrders(22) = "UPSR"
		End If
		if arrCSVOrders(22) = "2nd Day Air" Then
			arrCSVOrders(22) = "UPSB"
		End If
		if arrCSVOrders(22) = "UPS 3 Day Select" Then
			arrCSVOrders(22) = "UPSS"
		End If
		if arrCSVOrders(22) = "UPS Ground (US 48 States Only)" Then
			arrCSVOrders(22) = "UPSG"
		End If
		if arrCSVOrders(22) = "First Class Mail" Then
			arrCSVOrders(22) = "USPSF"
		End If
		if arrCSVOrders(22) = "UPS 2nd Day Air" Then
			arrCSVOrders(22) = "UPSB"
		End If
		if arrCSVOrders(22) = "U.S. Postal 4-7 Days (Canada Only)" Then
			arrCSVOrders(22) = "USPSE"
		End If
		if arrCSVOrders(22) = "U.S. Postal 4-7 Days (South America)" Then
			arrCSVOrders(22) = "USPSEMS"
		End If
		if arrCSVOrders(22) = "U.S. Postal 4-7 Days (Asia)" Then
			arrCSVOrders(22) = "USPSEMS"
		End If
		if arrCSVOrders(22) = "Next Day Air Saturday Delivery" Then
			arrCSVOrders(22) = "UPSRS"
		End If
		if arrCSVOrders(22) = "U.S. Postal 2 Weeks (Europe)" Then
			arrCSVOrders(22) = "USPSEMS"
		End If
		if arrCSVOrders(22) = "UPS Ground" Then
			arrCSVOrders(22) = "UPSG"
		End If
		if arrCSVOrders(22) = "U.S. Postal Global Priority (Europe)" Then
			arrCSVOrders(22) = "USPSGP"
		End If
		if arrCSVOrders(22) = "UPS Worldwide Express" Then
			arrCSVOrders(22) = "UPSWE"
		End If
		if arrCSVOrders(22) = "UPS Worldwide Express" Then
			arrCSVOrders(22) = "UPSWED"
		End If
		if arrCSVOrders(22) = "Free Shipping On Orders Over $99" Then
			arrCSVOrders(22) = "FREESHIP"
		End If
		if arrCSVOrders(22) = "Standard Shipping" Then
		arrCSVOrders(22) = "STDSHP"
		End If
	
		'Transform Country Field to correct format
		
		If Len(arrCSVOrders(8)) = 0 And Len(arrCSVOrders(16)) <> 0 Then
			arrCSVOrders(8) = arrCSVOrders(16)
		End If

		If Len(arrCSVOrders(16)) = 0 And Len(arrCSVOrders(8)) <> 0 Then
			arrCSVOrders(16) = arrCSVOrders(8)
		End If
		
		arrShipCountry = Split(arrCSVOrders(8), " ")
		if  UBound(arrShipCountry) > 1 Then
			for i = 2 to UBound(arrShipCountry)
				arrShipCountry(1) = arrShipCountry(1) + " " + arrShipCountry(i)
			next
		End If
		arrBillCountry = Split(arrCSVOrders(16), " ")
		if  UBound(arrBillCountry) > 1 Then
			for i = 2 to UBound(arrBillCountry)
				arrBillCountry(1) = arrBillCountry(1) + " " + arrBillCountry(i)
			next
		End If
		
		If UBound(arrShipCountry) <> -1 Then
			'Transform US Zip Codes to correct Format
			if arrShipCountry(1) = "United States" Then
				arrCSVOrders(9) = Left(arrCSVOrders(9),5)
			End If
			if arrBillCountry(1) = "United States" Then
				arrCSVOrders(17) = Left(arrCSVOrders(17),5)
			End If
		End if

		'Truncate Order Entry and Referring Point Fields if too large
		if Len(arrCSVOrders(20)) > 256 Then
			arrCSVOrders(20) = Left(arrCSVOrders(20), 255)
		End If
		if Len(arrCSVOrders(21)) > 256 Then
			arrCSVOrders(21) = Left(arrCSVOrders(21), 255)
		End If

		'Change Payment Method to match Everest
		if arrCSVOrders(23) = "Visa" Then
			arrCSVOrders(23) = "YAHOO VISA"
		End If
		if arrCSVOrders(23) = "MasterCard" Then	
			arrCSVOrders(23) = "YAHOO MASTERCARD"
		End If
		if arrCSVOrders(23) = "American Express" Then
			arrCSVOrders(23) = "YAHOO AMERICAN EXPRESS"
		End If
		if arrCSVOrders(23) = "Discover" Then
			arrCSVOrders(23) = "YAHOO DISCOVER"
		End If

		'store array contents to fields
		lngOrderID = arrCSVOrders(0)
		strDate = arrCSVOrders(1)
		lngNumericTime = arrCSVOrders(2)
		strShipFirstName = arrCSVMOM(43)
		strShipLastName = arrCSVMOM(42)
		strShipAddress1 = Replace(arrCSVOrders(4),",","")
		strShipAddress2 = arrCSVOrders(5)
		strShipCity = arrCSVOrders(6)
		strShipState = UCase(arrCSVOrders(7))
		strShipCountry = UCase(arrShipCountry(1))
		strShipZipCode = arrCSVOrders(9)
		strShipTelephone = arrCSVOrders(10)
		strBillFirstName = arrCSVMOM(3)
		strBillLastName = arrCSVMOM(2)
		strBillAddress1 = arrCSVOrders(12)
		strBillAddress2 = arrCSVOrders(13)
		strBillCity = arrCSVOrders(14)
		strBillState = UCase(arrCSVOrders(15))
		strBillCountry = UCase(arrBillCountry(1))
		strBillZipCode = arrCSVOrders(17)
		strBillTelephone = arrCSVOrders(18)
		strEmail = arrCSVOrders(19)
		strReferringPage = Replace(arrCSVOrders(20),chr(34),"")
		strEntryPoint = arrCSVOrders(21)
		strShipping = arrCSVOrders(22)
		strPaymentMethod = arrCSVOrders(23)
		
		'Yahoo Store stopped transmitting CC info. Addional changes @ line 365 - 380
		lngCardNumber = ""
		strCardExpiry = ""		
		'lngCardNumber = arrCSVOrders(24)
		'strCardExpiry = arrCSVOrders(25)
		
		strComments = Replace(arrCSVOrders(26),chr(34),"")
		strComments = Replace(arrCSVOrders(26),chr(44),".")
		
		strTotal = arrCSVOrders(27)
		strLinkFrom = arrCSVOrders(28)
		strWarning = arrCSVOrders(29)
		strAuthcode = arrCSVOrders(30)
		strAVSCode = arrCSVOrders(31)
		strGiftMessage = arrCSVOrders(32)
				
		'write to YahooOrders.csv file
		objWriteYahooOrders.WriteLine(lngOrderID & "," & strDate & "," & lngNumericTime & "," & _
			EncaseQuotes(strShipFirstName) & "," & EncaseQuotes(strShipLastName) & "," & _
			EncaseQuotes(strShipAddress1) & "," & EncaseQuotes(strShipAddress2) & "," & _
			EncaseQuotes(strShipCity) & "," & EncaseQuotes(strShipState) & "," & _ 
			EncaseQuotes(strShipCountry) & "," & EncaseQuotes(strShipZipCode) & "," & _ 
			EncaseQuotes(strShipTelephone) & "," & EncaseQuotes(strBillFirstName) & "," & _ 
			EncaseQuotes(strBillLastName) & "," & EncaseQuotes(strBillAddress1) & "," & _
			EncaseQuotes(strBillAddress2) & "," & EncaseQuotes(strBillCity) & "," & _
			EncaseQuotes(strBillState) & "," & EncaseQuotes(strBillCountry) & "," & _
			EncaseQuotes(strBillZipCode) & "," & EncaseQuotes(strBillTelephone) & "," & _
			EncaseQuotes(strEmail) & "," & EncaseQuotes(strReferringPage) & "," & _
			EncaseQuotes(strEntryPoint) & "," & EncaseQuotes(strShipping) & "," & _
			EncaseQuotes(strPaymentMethod) & "," & EncaseQuotes(lngCardNumber) & "," & _
			EncaseQuotes(strCardExpiry) & "," & EncaseQuotes(strComments) & "," & _
			EncaseQuotes(strTotal) & "," & EncaseQuotes(strLinkFrom) & "," & _
			EncaseQuotes(strWarning) & "," & EncaseQuotes(strAuthcode) & "," & _
			EncaseQuotes(strAVSCode) & "," & EncaseQuotes(strGiftMessage))

		'write to YahooItems.csv file
		Do while (arrCSVItems(0) = arrCSVOrders(0) And objReadFileItems.AtEndOfStream <> True) Or count = 2
			intTax = 0
			if arrCSVItems(2) = "Tax" Then
				if arrCSVItems(5) = "$0.00" Then
					intTax = 0
				Else
					intTax = 1
				End If
			End If
			
			if arrCSVItems(2) = "Coupon" Then
	
				arrCSVItems(2) = arrCSVItems(3)
				if intTax = 0 Then
					arrCSVItems(3) = "SALEDISC"
				End If
				if intTax = 1 Then
					arrCSVItems(3) = "SALEDISC2"
				End If
				arrCSVItems(5) = "-" & arrCSVItems(5)
			End If
				
			lngItemOrderNo = arrCSVItems(0)
			strProductCode = arrCSVItems(3)
			strProductId = arrCSVItems(2)
			intQuantity = arrCSVItems(4)
			strUnitPrice = arrCSVItems(5)
			
			If (StrComp(Cstr(arrCSVOptions(0)),Cstr(arrCSVItems(0)),1) = 0 And StrComp(arrCSVOptions(3),arrCSVItems(3)) = 0 And objReadFileOptions.AtEndOfStream <> True) Or (intOptionCount = 2 And StrComp(Cstr(arrCSVOptions(0)),Cstr(arrCSVItems(0)),1) = 0 And StrComp(arrCSVOptions(3),arrCSVItems(3)) = 0) Then
				intOptionCount2 = 1
			End if
			
			if intOptionCount2 = 0 Then
				objWriteYahooItems.WriteLine(lngItemOrderNo & "," & intLineNo & "," & _
					EncaseQuotes(strProductCode) & "," & EncaseQuotes(strProductId) & "," & _
					intQuantity & "," & strUnitPrice)
					'Add item "PICK" after "Shipping" mark boolean blnItemPick to true
				If strProductCode = "Shipping" Then
					intLineNo = intLineNo + 1
					objWriteYahooItems.WriteLine(lngItemOrderNo & "," & (intLineNo) & "," & _
						EncaseQuotes("PICK") & "," & EncaseQuotes("PICK") & "," & _
						1 & "," & "0.00")
					blnItemPick = 1
				End If
			End If
			
			Do while (StrComp(Cstr(arrCSVOptions(0)),Cstr(arrCSVItems(0)),1) = 0 And StrComp(arrCSVOptions(3),arrCSVItems(3)) = 0 And objReadFileOptions.AtEndOfStream <> True) Or (intOptionCount = 2 And StrComp(Cstr(arrCSVOptions(0)),Cstr(arrCSVItems(0)),1) = 0 And StrComp(arrCSVOptions(3),arrCSVItems(3)) = 0)


				if StrComp(Cstr(Right(arrCSVOptions(5),1)),")") = 0 Then
					dim pos1, pos2, strOptionCode, strOptionPrice, intItemLineNo
					pos1 = instrrev(arrCSVOptions(5),"(",-1,1)
					pos2 = instrrev(arrCSVOptions(5),")",-1,1)
					strOptionCode = Mid(arrCSVOptions(5), pos1+2, pos2-pos1-2)
					if instr(strOptionCode, "$") = 0 Then
						if instrrev(arrCSVOptions(5), ")", pos1, 1) Then
						else
							Msgbox "The option """ & arrCSVOptions(5) & """ for order number " & lngItemOrderNo & " is not in the correct format." & vbCr &"Please exclude this order from batch and notify webmaster to fix the add on option for item(s) in order " &  lngItemOrderNo & "."
							Wscript.Quit
						end if
						if instrrev(arrCSVOptions(5), ")", pos2, 1) Then
						else
							Msgbox "The option """ & arrCSVOptions(5) & """ for order number " & lngItemOrderNo & " is not in the correct format." & vbCr &"Please exclude this order from batch and notify webmaster to fix the add on option for item(s) in order " &  lngItemOrderNo & "."
							Wscript.Quit
						end if						
						pos2 = instrrev(arrCSVOptions(5), ")", pos1, 1)
						pos1 = instrrev(arrCSVOptions(5), "(", pos2, 1)
						strOptionPrice = Mid(arrCSVOptions(5), pos1+1, pos2-pos1-1)
						If instr(strOptionPrice, "+") > 0 Then
							strOptionPrice = Replace(strOptionPrice, "+" , "")
						End If
						intLineNo = intLineNo + 1
						arrOptionsList(0,x) = arrCSVItems(0)
						arrOptionsList(1,x) = intLineNo
						arrOptionsList(2,x) = strOptionCode
						arrOptionsList(3,x) = Ccur(strOptionPrice)
						x = x + 1
						Redim Preserve arrOptionsList(3,x)
						
					End If
				End If
				
				If intOptionCount <> 2 Then
					strContentsOptions = objReadFileOptions.ReadLine
					arrCSVOptions = CSVParse(strContentsOptions)
					intOptionCount = 0
				End If 
				
				If Cint(arrOptionsList(1,0)) > 0 Then
					intItemLineNo = Cint(arrOptionsList(1,0)) - 1
				Else
					intItemLineNo = intLineNo
				End If

				if StrComp(Cstr(arrCSVOptions(0)),Cstr(arrCSVItems(0)),1) <> 0 Or StrComp(arrCSVOptions(3),arrCSVItems(3)) <> 0  Or (intOptionCount = 2 And StrComp(Cstr(arrCSVOptions(0)),Cstr(arrCSVItems(0)),1) = 0 And StrComp(arrCSVOptions(3),arrCSVItems(3)) = 0) Then
					dim curOptionPrice
					curOptionPrice = 0
					for i = 0 to x				
						curOptionPrice = curOptionPrice + arrOptionsList(3,i)
					next
					
					strUnitPrice = Ccur(strUnitPrice)
					strUnitPrice = strUnitPrice - curOptionPrice
					objWriteYahooItems.WriteLine(lngItemOrderNo & "," & intItemLineNo & "," & _
						EncaseQuotes(strProductCode) & "," & EncaseQuotes(strProductId) & "," & _
						intQuantity & "," & strUnitPrice)
			
					for i = 0 to x
						if Len(arrOptionsList(0,i)) <> 0 Then
							objWriteYahooItems.WriteLine(arrOptionsList(0,i) & "," & arrOptionsList(1,i) & "," & EncaseQuotes(arrOptionsList(2,i)) & "," & EncaseQuotes(arrOptionsList(2,i)) & "," & "1" & "," & arrOptionsList(3,i))
						End If			
					Next
					
					intOptionCount2 = 0
					Redim arrOptionsList(3,0)
					x = 0
				End If					

				if intOptionCount = 2 Then
					intOptionCount = 3
				End If
				
				if objReadFileOptions.AtEndOfStream = True And intOptionCount <> 3 Then
					intOptionCount = 2
				End If
			Loop

			if count <> 2 Then		
				strContentsItems = objReadFileItems.ReadLine
				arrCSVItems = CSVParse(strContentsItems)
				count = 0
				intLineNo = intLineNo + 1
			End if
			
			if count = 2 Then 
				count = 3
			End If

			if objReadFileItems.AtEndOfStream = True And count <> 3 Then
				count = 2
			End If
		Loop
		
		' Add discont from yahoo order file
		If arrCSVOrders(44) > 0 Then
			intLineNo = intLineNo + 1
			objWriteYahooItems.WriteLine(lngItemOrderNo & "," & intLineNo & "," & _
				EncaseQuotes("SALEDISC2") & "," & EncaseQuotes("SALEDISC2") & "," & _
				1 & ",$" & arrCSVOrders(44)*-1)
		End If
		' Add item "PICK" if havn't already
		If blnItemPick = 0 Then
			intLineNo = intLineNo + 1
			objWriteYahooItems.WriteLine(lngItemOrderNo & "," & intLineNo & "," & _
				EncaseQuotes("PICK") & "," & EncaseQuotes("PICK") & "," & _
				1 & "," & "0.00")
		End If	
		' Add Gift message into order as an item		
		If Len(strGiftMessage) <> 0 Then
			intLineNo = intLineNo + 1
			objWriteYahooItems.WriteLine(lngItemOrderNo & "," & intLineNo & "," & _
				EncaseQuotes(strGiftMessage) & "," & EncaseQuotes("Customer's message") & "," & _
				1 & "," & "0.00")
			Wscript.echo "Gift Message for order # " & lngOrderID
		End If
		blnItemPick = 0
	Loop

	if Clng(arrCSVOrders(0)) <> Clng(arrCSVItems(0)) Or Clng(arrCSVOrders(0)) < Clng(arrCSVOptions(0)) Then
		MsgBox arrCSVOrders(0)
		MsgBox arrCSVItems(0)
		MsgBox arrCSVOptions(0)
		Wscript.Echo "Options List, Items List and Order List not matching! Redownload the files."
		Wscript.Quit
	End If

	Wscript.echo "Done"
	objReadFileMOM.close
	objReadFileOrders.close
	objReadFileItems.close
Else
    Wscript.Echo "The file is empty."
End If
