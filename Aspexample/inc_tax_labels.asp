<% 
	function GetTaxLabels()

		dim strTaxLabelSQL, rsTaxLabel, taxLabelsArray(5)
		set rsTaxLabel = Server.CreateObject("ADODB.Recordset")
		
		strTaxLabelSQL = "SELECT Top 1 LabelTax1, LabelTax2, LabelTax3, LabelTax4, LabelTax5 from Location WHERE Active = 1"

		rsTaxLabel.CursorLocation = 3
		rsTaxLabel.open strTaxLabelSQL, cnWS
		Set rsTaxLabel.ActiveConnection = Nothing

		if NOT rsTaxLabel.EOF then
			if not isNull(rsTaxLabel("LabelTax1")) then
				taxLabelsArray(0) = rsTaxLabel("LabelTax1")
			else
				taxLabelsArray(0) = "Tax Rate 1"
			end if
			if not isNull(rsTaxLabel("LabelTax2")) then
				taxLabelsArray(1) = rsTaxLabel("LabelTax2")
			else
				taxLabelsArray(1) = "Tax Rate 2"
			end if
			if not isNull(rsTaxLabel("LabelTax3")) then
				taxLabelsArray(2) = rsTaxLabel("LabelTax3")
			else
				taxLabelsArray(2) = "Tax Rate 3"
			end if
			if not isNull(rsTaxLabel("LabelTax4")) then
				taxLabelsArray(3) = rsTaxLabel("LabelTax4")
			else
				taxLabelsArray(3) = "Tax Rate 4"
			end if
			if not isNull(rsTaxLabel("LabelTax5")) then
				taxLabelsArray(4) = rsTaxLabel("LabelTax5")
			else
				taxLabelsArray(4) = "Tax Rate 5"
			end if
		end if
		
		rsTaxLabel.close

		GetTaxLabels = taxLabelsArray

	end function

	function UpdateTaxLabels(taxLabel1, taxLabel2, taxLabel3, taxLabel4, taxLabel5)
		dim strTaxLabelSQL		

		strTaxLabelSQL = "UPDATE Location Set LabelTax1 = N'"
		if taxLabel1 <>"" then
			strTaxLabelSQL = strTaxLabelSQL & sqlInjectStr(Left(CSTR(taxLabel1),14)) & "' "
		else
			strTaxLabelSQL = strTaxLabelSQL & "Tax Rate 1' "
		end if

		strTaxLabelSQL = strTaxLabelSQL & ", LabelTax2 = N'"

		if taxLabel2 <>"" then
			strTaxLabelSQL = strTaxLabelSQL & sqlInjectStr(Left(CSTR(taxLabel2),14)) & "' "
		else
			strTaxLabelSQL = strTaxLabelSQL & "Tax Rate 2' "
		end if
		
		strTaxLabelSQL = strTaxLabelSQL & ", LabelTax3 = N'"

		if taxLabel3 <>"" then
			strTaxLabelSQL = strTaxLabelSQL & sqlInjectStr(Left(CSTR(taxLabel3),14)) & "' "
		else
			strTaxLabelSQL = strTaxLabelSQL & "Tax Rate 3' "
		end if

		strTaxLabelSQL = strTaxLabelSQL & ", LabelTax4 = N'"

		if taxLabel4 <>"" then
			strTaxLabelSQL = strTaxLabelSQL & sqlInjectStr(Left(CSTR(taxLabel4),14)) & "' "
		else
			strTaxLabelSQL = strTaxLabelSQL & "Tax Rate 4' "
		end if

		strTaxLabelSQL = strTaxLabelSQL & ", LabelTax5 = N'"

		if taxLabel5 <>"" then
			strTaxLabelSQL = strTaxLabelSQL & sqlInjectStr(Left(CSTR(taxLabel5),14)) & "' "
		else
			strTaxLabelSQL = strTaxLabelSQL & "Tax Rate 5' "
		end if
		
		cnWS.Execute strTaxLabelSQL 

	end function 
%>