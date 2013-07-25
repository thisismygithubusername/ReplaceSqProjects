<%	
' This file includes functions that convert gross prices to net prices and vice versa when using Tax
' Inclusive Prices. These are just general utility functions, and their use must be intelligently
' applied.

' This class represents the adjusted net totals for a single quanity item. All of these values are
' safe to then multiply by the quantity; everything will come out great.
class ProductNetTotals
	public GrossPrice
	public GrossDiscount
	public NetPrice
	public NetDiscount
	public NetTax
	public NetTax1
	public NetTax2
	public NetTax3
	public NetTax4
	public NetTax5
	public ExtendedNetDiscount
	public ExtendedGrossPrice
	public ExtendedGrossDiscount
	public ExtendedNetTax
	public ExtendedNetPrice

	public function PopulateNetTaxes(Tax1, Tax2, Tax3, Tax4, Tax5)
		NetTax1 = GetNetTaxAmount(NetPrice, NetDiscount, Tax1)
		NetTax2 = GetNetTaxAmount(NetPrice, NetDiscount, Tax2)
		NetTax3 = GetNetTaxAmount(NetPrice, NetDiscount, Tax3)
		NetTax4 = GetNetTaxAmount(NetPrice, NetDiscount, Tax4)
		NetTax5 = GetNetTaxAmount(NetPrice, NetDiscount, Tax5)

		NetTax = NetTax1 + NetTax2 + NetTax3 + NetTax4 + NetTax5
	end function

end class

' Consturctor for ProductNetTotals 
public function GetNetTotals(GrossPrice, GrossDiscountPercentage, GrossDiscountAmount, Tax1, Tax2, Tax3, Tax4, Tax5, Quantity)
	dim netTotals : set netTotals = new ProductNetTotals

	' Cap the discount percentage at 1.
	if (GrossDiscountPercentage >= 1) then
		GrossDiscountPercentage = 1
		netTotals.NetPrice = GrossPrice
	else
		netTotals.NetPrice = GetNetPrice(GrossPrice, Tax1, Tax2, Tax3, Tax4, Tax5)
	end if
	
	' Calculate all the grosses, we use this to adjust our rounding and sanity checks
	netTotals.GrossPrice = GrossPrice
	netTotals.ExtendedGrossPrice = GrossPrice * Quantity

	netTotals.GrossDiscount = Round(GrossPrice * GrossDiscountPercentage, 2)
	netTotals.ExtendedGrossDiscount = GrossDiscountAmount

	' Calc the net discount
	netTotals.NetDiscount = GetNetDiscount(netTotals.netPrice, GrossDiscountPercentage)

	' Get the net taxes
	call netTotals.PopulateNetTaxes(Tax1, Tax2, Tax3, Tax4, Tax5)

	logIt " GrossPrice : " & GrossPrice
	logIt " NetTax1 : " & netTotals.NetTax1
	logIt " NetTax2 : " & netTotals.NetTax2
	logIt " NetTax3 : " & netTotals.NetTax3
	logIt " NetTax4 : " & netTotals.NetTax4
	logIt " NetTax5 : " & netTotals.NetTax5
	logIt " NetTax : " & netTotals.NetTax
	logIt " First NetPrice : " & netTotals.NetPrice
	logIt " NetDiscount : " &  netTotals.NetDiscount
	logIT " grossDiscountAmount : " & netTotals.GrossDiscount
	logIt " GrossDiscountPercentage : " & GrossDiscountPercentage
	logIt " (netTotals.NetPrice - netTotals.NetDiscount + netTotals.NetTax) : " & (netTotals.NetPrice - netTotals.NetDiscount + netTotals.NetTax)
	logIt " (GrossPrice - GrossDiscount) : " & (GrossPrice - netTotals.GrossDiscount)
	logIt " ExtendedGrossPrice : " & netTotals.ExtendedGrossPrice
	dim pennyRound

	'We're offloading the penny rounding issues into the price, so taxes do not have to be recalculated
	if (netTotals.NetPrice - netTotals.NetDiscount + netTotals.NetTax) <> (GrossPrice - netTotals.GrossDiscount) then
		pennyRound = GrossPrice - netTotals.GrossDiscount - netTotals.NetPrice + netTotals.NetDiscount - netTotals.NetTax
		logIt " pennyRound : " & pennyRound
		netTotals.NetPrice = netTotals.NetPrice + pennyRound
		logIt " adjusted NetPrice : " & netTotals.NetPrice
	end if

	' Now we move on to the extended values. NetPrice and NetTaxes will not change, only the discount
	netTotals.ExtendedNetTax = netTotals.NetTax * Quantity
	netTotals.ExtendedNetPrice = netTotals.NetPrice * Quantity
	netTotals.ExtendedNetDiscount = netTotals.NetDiscount * Quantity

	logIt " netTotals.ExtendedNetPrice - netTotals.ExtendedNetDiscount + netTotals.ExtendedNetTax : " &  netTotals.ExtendedNetPrice - netTotals.ExtendedNetDiscount + netTotals.ExtendedNetTax
	logIt " netTotals.ExtendedGrossPrice - netTotals.ExtendedGrossDiscount : " & netTotals.ExtendedGrossPrice - netTotals.ExtendedGrossDiscount
	logIt " ExtendedNetDiscount (Pre Penny Rounding) : " & netTotals.ExtendedNetDiscount
	
	' Now calculate the rounding on the discount
	if (netTotals.ExtendedNetPrice - netTotals.ExtendedNetDiscount + netTotals.ExtendedNetTax) <> netTotals.ExtendedGrossPrice - netTotals.ExtendedGrossDiscount then
		pennyRound = netTotals.ExtendedGrossPrice - netTotals.ExtendedGrossDiscount - (netTotals.ExtendedNetPrice - netTotals.ExtendedNetDiscount + netTotals.ExtendedNetTax)
	
		logIt " pennyRound : " & pennyRound

		netTotals.ExtendedNetDiscount = netTotals.ExtendedNetDiscount - pennyRound

		logIt " ExtendedNetDiscount (Post Penny Rounding): " & netTotals.ExtendedNetDiscount
	end if

	logIt " after (netTotals.NetPrice - netTotals.NetDiscount + netTotals.NetTax) : " & (netTotals.NetPrice - netTotals.NetDiscount + netTotals.NetTax)
	logIt " after (GrossPrice - GrossDiscount) : " & (netTotals.GrossPrice - netTotals.GrossDiscount)
	set GetNetTotals = netTotals
end function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------

'' Begin normal Tax utility functions

'' VAT Tax functions
' This function returns the net price with all taxes applied, rounded to 2 decimal places. This does not include
' any discounts that may be applied, as that is unnessessary (discounts are stored seperately)
function GetNetPrice(GrossPrice, Tax1, Tax2, Tax3, Tax4, Tax5) 
	GetNetPrice = Round(GrossPrice / (1 + Tax1 + Tax2 + Tax3 + Tax4 + Tax5), 2)
end function

' This function will give you the net discount amount. When working with flat rate discounts, convert the flat rate
' to a percentage first. The result is rounded to 2 decimal places
function GetNetDiscount(NetPrice, DiscountPercentage)
	GetNetDiscount = Round(NetPrice * DiscountPercentage, 2)
end function

' Returns the discount percentage. You could do this yourself, but having a function named exactly this makes it
' easier to read what you are doing. This percentage is NOT rounded, it is as accurate as VBScript allows
function GetDiscountPercentage(ExtendedGrossPrice, FlatRateDiscountAmount)
	GetDiscountPercentage = 0
	if (ExtendedGrossPrice > 0) then
		GetDiscountPercentage = FlatRateDiscountAmount / ExtendedGrossPrice
	end if
end function

' This function returns the net tax amount (after discount). It needs to be called 5 times to get all tax amounts.
' It is rounded to 2 decimal places
function GetNetTaxAmount(NetPrice, NetDiscount, TaxRate)
	GetNetTaxAmount = Round((NetPrice - NetDiscount) * TaxRate, 2)
end function

' Returns the gross discount on an item. Needed as we store in Sales Details the Net Discount; often we want to display
' the gross. This function does no rounding as all the parameters should be valid money ammounts.
function GetGrossDiscountAmount(NetPrice, NetDiscount, NetTaxTotal, GrossPrice)
	GetGrossDiscountAmount = -NetPrice + NetDiscount - NetTaxTotal + GrossPrice
end function

function GetTotalNetTaxAmount(NetPrice, NetDiscount, Tax1, Tax2, Tax3, Tax4, Tax5)
	dim totalNetTaxes : totalNetTaxes = 0
	totalNetTaxes = GetNetTaxAmount(NetPrice, NetDiscount, Tax1)
	totalNetTaxes = totalNetTaxes + GetNetTaxAmount(NetPrice, NetDiscount, Tax2)
	totalNetTaxes = totalNetTaxes + GetNetTaxAmount(NetPrice, NetDiscount, Tax3)
	totalNetTaxes = totalNetTaxes + GetNetTaxAmount(NetPrice, NetDiscount, Tax4)
	totalNetTaxes = totalNetTaxes + GetNetTaxAmount(NetPrice, NetDiscount, Tax5)
	GetTotalNetTaxAmount = totalNetTaxes
end function

'' End VAT Tax functions

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------

'' Begin Non VAT Tax functions

function GetTotalNonVatTaxAmount(NetPrice, Tax1, Tax2, Tax3, Tax4, Tax5)
	dim totalTaxes : totalTaxes = 0
	totalTaxes = Round((NetPrice * Tax1), 2) +  Round((NetPrice * Tax2), 2) + Round((NetPrice * Tax3), 2) + Round((NetPrice * Tax4), 2) + Round((NetPrice * Tax5), 2)
	logIt " totalTaxes : " & totalTaxes
	GetTotalNonVatTaxAmount = totalTaxes
end function 

'' End Non VAT Tax functions
%>