<%
Class BusinessRulesClass
	
	Private id
	Private rule_type
	Private label
	Private description
	Private activate
	Private voucher_id

	Private id_rule_conf
	Private id_prod_orig_conf
	Private id_prod_ref_conf
	Private rate_from_conf
	Private rate_to_conf
	Private rate_from_ref_conf
	Private rate_to_ref_conf
	Private operation_conf
	Private applyto_conf
	Private apply4qta_conf
	Private valore_conf
	
	Private order_id
	Private prod_id
	Private counter_prod
	
	
	Public Function getID()
		getID = id
	End Function
	
	Public Sub setID(strID)
		id = strID
	End Sub	
	
	Public Function getRuleType()
		getRuleType = rule_type
	End Function
	
	Public Sub setRuleType(strRuleType)
		rule_type = strRuleType
	End Sub
	
	Public Function getLabel()
		getLabel = label
	End Function
	
	Public Sub setLabel(strLabel)
		label = strLabel
	End Sub
	
	Public Function getDescrizione()
		getDescrizione = description
	End Function
	
	Public Sub setDescrizione(strDesc)
		description = strDesc
	End Sub
	
	Public Function getActivate()
		getActivate = activate
	End Function
	
	Public Sub setActivate(strActivate)
		activate = strActivate
	End Sub
	
	Public Function getVoucherID()
		getVoucherID = voucher_id
	End Function
	
	Public Sub setVoucherID(strVoucherID)
		voucher_id = strVoucherID
	End Sub


	'************ metodi GET e SET per le business rule config	
	Public Function getRuleID()
		getRuleID = id_rule_conf
	End Function
	
	Public Sub setRuleID(strRuleID)
		id_rule_conf = strRuleID
	End Sub	

	Public Function getProdOrigConfID()
		getProdOrigConfID = id_prod_orig_conf
	End Function
	
	Public Sub setProdOrigConfID(strIDPOConf)
		id_prod_orig_conf = strIDPOConf
	End Sub	

	Public Function getProdRefConfID()
		getProdRefConfID = id_prod_ref_conf
	End Function
	
	Public Sub setProdRefConfID(strIDPRConf)
		id_prod_ref_conf = strIDPRConf
	End Sub	

	Public Function getRateFromConf()
		getRateFromConf = rate_from_conf
	End Function
	
	Public Sub setRateFromConf(strRateC)
		rate_from_conf = strRateC
	End Sub

	Public Function getRateToConf()
		getRateToConf = rate_to_conf
	End Function
	
	Public Sub setRateToConf(strRatetC)
		rate_to_conf = strRatetC
	End Sub	

	Public Function getRateFromRefConf()
		getRateFromRefConf = rate_from_ref_conf
	End Function
	
	Public Sub setRateFromRefConf(strRateC)
		rate_from_ref_conf = strRateC
	End Sub

	Public Function getRateToRefConf()
		getRateToRefConf = rate_to_ref_conf
	End Function
	
	Public Sub setRateToRefConf(strRatetC)
		rate_to_ref_conf = strRatetC
	End Sub	

	Public Function getOperationConf()
		getOperationConf = operation_conf
	End Function
	
	Public Sub setOperationConf(strOperationC)
		operation_conf = strOperationC
	End Sub		

	Public Function getApplyToConf()
		getApplyToConf = applyto_conf
	End Function
	
	Public Sub setApplyToConf(strApplyToC)
		applyto_conf = strApplyToC
	End Sub		

	Public Function getApply4QtaConf()
		getApply4QtaConf = apply4qta_conf
	End Function
	
	Public Sub setApply4QtaConf(strApply4QtaC)
		apply4qta_conf = strApply4QtaC
	End Sub		

	Public Function getValoreConf()
		getValoreConf = valore_conf
	End Function
	
	Public Sub setValoreConf(strValoreC)
		valore_conf = strValoreC
	End Sub

	
	Public Function getOrderID()
		getOrderID = order_id
	End Function
	
	Public Sub setOrderID(strOrderID)
		order_id = strOrderID
	End Sub	
		
	Public Function getProdID()
		getProdID = prod_id
	End Function
	
	Public Sub setProdID(strProdID)
		prod_id = strProdID
	End Sub
		
	Public Function getCounterProd()
		getCounterProd = counter_prod
	End Function
	
	Public Sub setCounterProd(strCounterProd)
		counter_prod = strCounterProd
	End Sub	
	

	Public Function getAmountByStrategy(totaleOrder, objVoucher, idProdOrig, objListProductRulesVO)
		Set getAmountByStrategy = new StrategyResultVO
		On Error Resume Next

		'Set objLogger = New LogClass
		'response.write("totaleOrder: "&totaleOrder&"<br>")
		'response.write("getRuleType(): "&getRuleType()&"<br>")
		'response.write("getID(): "&getID()&"<br>")
		'response.end

		Select Case CInt(getRuleType())
			Case 1,4
				On Error Resume Next
				Set objListRulesConf = getListaRulesConfig(getID(), null)
				if (Instr(1, typename(objListRulesConf), "Dictionary", 1) > 0) then
					for each g in objListRulesConf
						tmpRF = objListRulesConf(g).getRateFromConf()
						tmpRT = objListRulesConf(g).getRateToConf()
						'response.write("tmpRF:"&tmpRF&"; - tmpRT:"&tmpRT&"<br>")
						if(CDbl(totaleOrder)>=CDbl(tmpRF) AND CDbl(totaleOrder)<=CDbl(tmpRT))then
							Select Case CInt(objListRulesConf(g).getOperationConf())
								Case 1
								getAmountByStrategy.foundAmount = getAmountByStrategy.foundAmount+CDbl(objListRulesConf(g).getValoreConf())
								Case 2
								getAmountByStrategy.foundAmount = getAmountByStrategy.foundAmount-CDbl(objListRulesConf(g).getValoreConf())
							End Select
							Exit for
						end if
					next
					Set objListRulesConf = nothing
				end if
				if(Err.number <> 0) then
					Set getAmountByStrategy = new StrategyResultVO
				end if			
			Case 2,5
				On Error Resume Next
				Set objListRulesConf = getListaRulesConfig(getID(), null)
				if (Instr(1, typename(objListRulesConf), "Dictionary", 1) > 0) then
					for each g in objListRulesConf
						tmpRF = objListRulesConf(g).getRateFromConf()
						tmpRT = objListRulesConf(g).getRateToConf()
						'response.write("tmpRF:"&tmpRF&"; - tmpRT:"&tmpRT&"<br>")
						if(CDbl(totaleOrder)>=CDbl(tmpRF) AND CDbl(totaleOrder)<=CDbl(tmpRT))then
							Select Case CInt(objListRulesConf(g).getOperationConf())
								Case 1
								getAmountByStrategy.foundAmount = getAmountByStrategy.foundAmount+(CDbl(totaleOrder) / 100 * CDbl(objListRulesConf(g).getValoreConf()))
								Case 2
								getAmountByStrategy.foundAmount = getAmountByStrategy.foundAmount-(CDbl(totaleOrder) / 100 * CDbl(objListRulesConf(g).getValoreConf()))
							End Select
							Exit for
						end if
					next
					Set objListRulesConf = nothing
				end if
				if(Err.number <> 0) then
					Set getAmountByStrategy = new StrategyResultVO
				end if
			Case 3
				On Error Resume Next
				if (strComp(typename(objVoucher), "VoucherClass") = 0)then
					Set objListRulesConf = getListaRulesConfig(getID(), null)
					if (Instr(1, typename(objListRulesConf), "Dictionary", 1) > 0) then
						for each g in objListRulesConf
							tmpRF = objListRulesConf(g).getRateFromConf()
							tmpRT = objListRulesConf(g).getRateToConf()
							'response.write("objVoucher.getID():"&objVoucher.getID()&"; - getVoucherID():"&getVoucherID()&"<br>")
							if(CDbl(totaleOrder)>=CDbl(tmpRF) AND CDbl(totaleOrder)<=CDbl(tmpRT) AND objVoucher.getID()=getVoucherID())then							
								'response.write("objVoucher.getOperation():"&objVoucher.getOperation()&"<br>")
								Select Case CInt(objVoucher.getOperation())
									Case 0
										getAmountByStrategy.foundAmount = getAmountByStrategy.foundAmount-(CDbl(totaleOrder) / 100 * CDbl(objVoucher.getValore()))
									Case 1
										getAmountByStrategy.foundAmount = getAmountByStrategy.foundAmount-CDbl(objVoucher.getValore())
								End Select
								Exit for
							end if
						next
						Set objListRulesConf = nothing
					end if
				end if
				if(Err.number <> 0) then
					Set getAmountByStrategy = new StrategyResultVO
				end if			
			Case 6
				Set getAmountByStrategy.objListPRVO = objListProductRulesVO
				bolHasRuleConf = false
				On Error Resume Next
				Set objListRulesConf = getListaRulesConfig(getID(), idProdOrig)
				if (Instr(1, typename(objListRulesConf), "Dictionary", 1) > 0) then
					bolHasRuleConf = true
				end if
				if(Err.number <> 0) then
					'response.write("Err.description: "&Err.description&"<br>")
					bolHasRuleConf = false
					'Set getAmountByStrategy = new StrategyResultVO
					'Set getAmountByStrategy.objListPRVO = objListProductRulesVO
				end if

				if (bolHasRuleConf) then
					'response.write(" rule 6 Before count: "& objListProductRulesVO(idProdOrig).bolAlreadyApplied.count &"<br>")
					'for each l in objListProductRulesVO(idProdOrig).bolAlreadyApplied
						'response.write("key: "&l&" - value: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied(l)&"<br>")
					'next				
					'response.write("<br>key: "& getID()&"-"&idProdOrig)
					'response.write("obj type: "&typename(objListProductRulesVO(idProdOrig).bolAlreadyApplied))
					'response.write(" -value: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied(getID()&"-"&idProdOrig))
					'response.write(" -exists: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.exists(getID()&"-"&idProdOrig))
					'response.write("obj is: "& (objListProductRulesVO(idProdOrig).bolAlreadyApplied(getID()&"-"&idProdOrig)="1"))
					
					'response.write("6 count 2: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.count)
					'keyVal = objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig)
					'response.write("<br>6 isEmpty(): "& isEmpty(objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig)))
					if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig)<>"1") then					
						tmpMaxQtadisp = objListProductRulesVO(idProdOrig).totQta
						
						'for each l in objListProductRulesVO(idProdOrig).bolAlreadyApplied
							'response.write("key: "&l&" - value: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied(l)&"<br>")
						'next
					
						'response.write("<br>6 count 3: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.count)
						
						for each g in objListRulesConf
							tmpRF = objListRulesConf(g).getRateFromConf()
							tmpRT = objListRulesConf(g).getRateToConf()
							'response.write("tmpRF:"&tmpRF&"; - tmpRT:"&tmpRT&"<br>")
							tmpApply4Qta = objListRulesConf(g).getApply4QtaConf()
							'response.write("tmpApply4Qta: "&tmpApply4Qta&"<br>")
							if(cLng(tmpApply4Qta)>cLng(tmpMaxQtadisp)) then tmpApply4Qta = tmpMaxQtadisp end if
							
							'response.write("<br>6 count 4: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.count)
							
							if(cLng(tmpMaxQtadisp)>=CDbl(tmpRF) AND cLng(tmpMaxQtadisp)<=CDbl(tmpRT))then
								tmpOpValue = CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta)
								Select Case CInt(objListRulesConf(g).getOperationConf())
									Case 1
									getAmountByStrategy.foundAmount = getAmountByStrategy.foundAmount+(tmpOpValue)
									Case 2
									getAmountByStrategy.foundAmount = getAmountByStrategy.foundAmount-(tmpOpValue)
									tmpOpValue=0-tmpOpValue
								End Select
								objListProductRulesVO(idProdOrig).listRelrulesLabel.add getID()&"-"&idProdOrig&"|"&getLabel(), tmpOpValue
								'response.write("<br>6 arrivo prima di set - count: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.count)
								if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
									'response.write("set value 1 for: "&getID()&"-"&idProdOrig&"<br>")
									' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
									'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
									'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
									objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
								else
									objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
								end if
								Exit for
							end if
						next
						Set getAmountByStrategy.objListPRVO = objListProductRulesVO
					end if
					
					'response.write("<br> rule 6 After:<br>")
					'for each l in objListProductRulesVO(idProdOrig).bolAlreadyApplied
					'	response.write("key: "&l&" - value: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied(l)&"<br>")
					'next
					
					Set objListRulesConf = nothing
				end if			
			Case 7				
				Set getAmountByStrategy.objListPRVO = objListProductRulesVO		
				bolHasRuleConf = false	
				'call objLogger.write("BusinessRulesClass: lookfor getID(): "&getID()&" -idProdOrig: "&idProdOrig, "system", "debug")
				On Error Resume Next			
				Set objListRulesConf = getListaRulesConfig(getID(), idProdOrig)
				if (Instr(1, typename(objListRulesConf), "Dictionary", 1) > 0) then
					bolHasRuleConf = true
				end if
				if(Err.number <> 0) then
					'response.write("Err.description: "&Err.description&"<br>")
					bolHasRuleConf = false
					'Set getAmountByStrategy = new StrategyResultVO
					'Set getAmountByStrategy.objListPRVO = objListProductRulesVO
				end if
				
				'call objLogger.write("BusinessRulesClass: lookfor getID(): "&getID()&" -bolHasRuleConf: "&bolHasRuleConf&" -idProdOrig: "&idProdOrig, "system", "debug")
				'for each k in objListProductRulesVO
					'call objLogger.write("k before: "&k, "system", "debug")			
				'next
				
				if (bolHasRuleConf) then		
					'response.write(" rule 7 Before count: "& objListProductRulesVO(idProdOrig).bolAlreadyApplied.count &"<br>")
					'for each l in objListProductRulesVO(idProdOrig).bolAlreadyApplied
					'	response.write("key: "&l&" - value: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied(l)&"<br>")
					'next
					'response.write("<br>key: "& getID()&"-"&idProdOrig)
					'response.write("obj type: "&typename(objListProductRulesVO(idProdOrig).bolAlreadyApplied))
					'response.write(" -value: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied(getID()&"-"&idProdOrig))
					'response.write(" -exists: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.exists(getID()&"-"&idProdOrig))
					'response.write("obj is: "& (objListProductRulesVO(idProdOrig).bolAlreadyApplied(getID()&"-"&idProdOrig)))
					
					if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig)<>"1") then					
						tmpMaxQtadisp = objListProductRulesVO(idProdOrig).totQta
						tmpImpProd = objListProductRulesVO(idProdOrig).objProd.getPrezzo()
						'response.write("tmpImpProd: "&tmpImpProd&"<br>")
						'call objLogger.write("BusinessRulesClass: tmpImpProd: "&tmpImpProd&" -idProdOrig: "&idProdOrig&" -tmpMaxQtadisp: "&tmpMaxQtadisp, "system", "debug")
						for each g in objListRulesConf
							tmpRF = objListRulesConf(g).getRateFromConf()
							tmpRT = objListRulesConf(g).getRateToConf()
							'response.write("tmpRF:"&tmpRF&"; - tmpRT:"&tmpRT&"<br>")
							tmpApply4Qta = objListRulesConf(g).getApply4QtaConf()
							if(cLng(tmpApply4Qta)>cLng(tmpMaxQtadisp)) then tmpApply4Qta = tmpMaxQtadisp end if
							'response.write("tmpApply4Qta: "&tmpApply4Qta&"<br>")
							
							if(cLng(tmpMaxQtadisp)>=CDbl(tmpRF) AND cLng(tmpMaxQtadisp)<=CDbl(tmpRT))then
								tmpOpValue = CDbl(tmpImpProd) / 100 * (CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta))	
								'call objLogger.write("BusinessRulesClass: tmpOpValue: "&tmpOpValue&" -idProdOrig: "&idProdOrig, "system", "debug")						
								Select Case CInt(objListRulesConf(g).getOperationConf())
									Case 1
									getAmountByStrategy.foundAmount = getAmountByStrategy.foundAmount+(tmpOpValue)
									Case 2
									getAmountByStrategy.foundAmount = getAmountByStrategy.foundAmount-(tmpOpValue)
									tmpOpValue=0-tmpOpValue
								End Select
								objListProductRulesVO(idProdOrig).listRelrulesLabel.add getID()&"-"&idProdOrig&"|"&getLabel(), tmpOpValue
								'call objLogger.write("BusinessRulesClass: objListProductRulesVO(idProdOrig).listRelrulesLabel: "&objListProductRulesVO(idProdOrig).listRelrulesLabel(getID()&"-"&idProdOrig&"|"&getLabel())&" -idProdOrig: "&idProdOrig, "system", "debug")	
								'response.write("7 arrivo prima di set - count: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.count)
								if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
									'response.write("set value 1 for: "&getID()&"-"&idProdOrig&"<br>")
									' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
									'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
									'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
									objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
								else
									objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
								end if
								Exit for
							end if
						next
						Set getAmountByStrategy.objListPRVO = objListProductRulesVO
					end if
					
					'response.write("<br> rule 7 After:<br>")
					'for each l in objListProductRulesVO(idProdOrig).bolAlreadyApplied
					'	response.write("key: "&l&" - value: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied(l)&"<br>")
					'next
					'for each k in objListProductRulesVO
						'call objLogger.write("k after: "&k, "system", "debug")			
					'next	
					
					Set objListRulesConf = nothing
				end if				
			Case 8
				Set getAmountByStrategy.objListPRVO = objListProductRulesVO
				bolHasRuleConf = false
				On Error Resume Next
				Set objListRulesConf = getListaRulesConfig(getID(), idProdOrig)
				if (Instr(1, typename(objListRulesConf), "Dictionary", 1) > 0) then
					bolHasRuleConf = true
				end if
				if(Err.number <> 0) then
					'response.write("Err.description: "&Err.description&"<br>")
					bolHasRuleConf = false
					'Set getAmountByStrategy = new StrategyResultVO
					'Set getAmountByStrategy.objListPRVO = objListProductRulesVO
				end if

				'call objLogger.write("BusinessRulesClass: lookfor getID(): "&getID()&" -getLabel(): "&getLabel()&" -idProdOrig: "&idProdOrig&" -bolHasRuleConf: "&bolHasRuleConf, "system", "debug")

				if (bolHasRuleConf) then
					'response.write(" rule 6 Before count: "& objListProductRulesVO(idProdOrig).bolAlreadyApplied.count &"<br>")
					'for each l in objListProductRulesVO(idProdOrig).bolAlreadyApplied
						'response.write("key: "&l&" - value: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied(l)&"<br>")
					'next				
					'response.write("<br>key: "& getID()&"-"&idProdOrig)
					'response.write("obj type: "&typename(objListProductRulesVO(idProdOrig).bolAlreadyApplied))
					'response.write(" -value: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied(getID()&"-"&idProdOrig))
					'response.write(" -exists: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.exists(getID()&"-"&idProdOrig))
					'response.write("obj is: "& (objListProductRulesVO(idProdOrig).bolAlreadyApplied(getID()&"-"&idProdOrig)="1"))
					
					'response.write("6 count 2: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.count)
					'keyVal = objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig)
					'response.write("<br>6 isEmpty(): "& isEmpty(objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig)))
					if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig)<>"1") then					
						tmpMaxQtadisp = objListProductRulesVO(idProdOrig).totQta
						tmpImpProd = objListProductRulesVO(idProdOrig).objProd.getPrezzo()
						'response.write("<br>tmpMaxQtadisp: "& tmpMaxQtadisp)
						'response.write("<br>tmpImpProd: "& tmpImpProd)
						
						for each g in objListRulesConf
							idProdRef = objListRulesConf(g).getProdRefConfID()
							
							'call objLogger.write("BusinessRulesClass: lookfor t: "&t&" -objListProductRulesVO(t).totQta:"&objListProductRulesVO(t).totQta, "system", "debug")
							'call objLogger.write("BusinessRulesClass: lookfor idProdRef: "&idProdRef&" -objListProductRulesVO.exists(idProdRef): "& objListProductRulesVO.exists(idProdRef), "system", "debug")

							'for each t in objListProductRulesVO
								'call objLogger.write("BusinessRulesClass: lookfor t: "&t&" -idProdRef: "&idProdRef&" -objListProductRulesVO(t).totQta:"&objListProductRulesVO(t).totQta&" -t=idProdRef: "& (t=idProdRef), "system", "debug")
							'next
						
							if not(objListProductRulesVO.exists(idProdRef))then
								Exit for
							end if
							
							tmpMaxRefQtadisp = objListProductRulesVO(idProdRef).totQta
							tmpImpRefProd = objListProductRulesVO(idProdRef).objProd.getPrezzo()
							tmpRF = objListRulesConf(g).getRateFromConf()
							tmpRT = objListRulesConf(g).getRateToConf()
							tmpRFref = objListRulesConf(g).getRateFromRefConf()
							tmpRTref = objListRulesConf(g).getRateToRefConf()
							tmpApplyTo = objListRulesConf(g).getApplyToConf()
							tmpApply4Qta = objListRulesConf(g).getApply4QtaConf()
							
							'response.write("<br>idProdRef: "& idProdRef)						
							'response.write("<br>tmpMaxRefQtadisp: "& tmpMaxRefQtadisp)
							'response.write("<br>tmpImpRefProd: "& tmpImpRefProd)
							'response.write("<br>tmpRF:"&tmpRF&"; - tmpRT:"&tmpRT)
							'response.write("<br>tmpRFref:"&tmpRFref&"; - tmpRTref:"&tmpRTref)
							'response.write("<br>tmpApplyTo: "& tmpApplyTo)
							'response.write("<br>tmpApply4Qta: "&tmpApply4Qta)							
							
							if(cLng(tmpMaxQtadisp)>=CDbl(tmpRF) AND cLng(tmpMaxQtadisp)<=CDbl(tmpRT))then							
								if(cLng(tmpMaxRefQtadisp)>=CDbl(tmpRFref) AND cLng(tmpMaxRefQtadisp)<=CDbl(tmpRTref))then
							
									Select Case CInt(tmpApplyTo)
										Case 1
											if(cLng(tmpApply4Qta)>cLng(tmpMaxQtadisp)) then tmpApply4Qta = tmpMaxQtadisp end if
											tmpOpValue = CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta)
											Select Case CInt(objListRulesConf(g).getOperationConf())
												Case 1
												objListProductRulesVO(idProdOrig).resultAmount = objListProductRulesVO(idProdOrig).resultAmount+(tmpOpValue)
												Case 2
												objListProductRulesVO(idProdOrig).resultAmount = objListProductRulesVO(idProdOrig).resultAmount-(tmpOpValue)
												tmpOpValue=0-tmpOpValue
											End Select
											objListProductRulesVO(idProdOrig).listRelrulesLabel.add getID()&"-"&idProdOrig&"|"&getLabel(), tmpOpValue
											'response.write("<br>6 arrivo prima di set - count: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.count)
											if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
												'response.write("set value 1 for: "&getID()&"-"&idProdOrig&"<br>")
												' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
											else
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
											end if
											Exit for
										Case 2
											'response.write("<br>entro nel case 2")
											if(cLng(tmpApply4Qta)>cLng(tmpMaxRefQtadisp)) then tmpApply4Qta = tmpMaxRefQtadisp end if
											tmpOpValue = CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta)
											Select Case CInt(objListRulesConf(g).getOperationConf())
												Case 1
												objListProductRulesVO(idProdRef).resultAmount = objListProductRulesVO(idProdRef).resultAmount+(tmpOpValue)
												Case 2
												'response.write("<br>entro nel case operation 2")
												'response.write("<br>type original: "& typename(objListProductRulesVO(idProdRef)))
												'response.write("<br>value original: "& objListProductRulesVO(idProdRef).resultAmount)
												'response.write("<br>value calculated: "& (CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta)))
												objListProductRulesVO(idProdRef).resultAmount = objListProductRulesVO(idProdRef).resultAmount-(tmpOpValue)
												tmpOpValue=0-tmpOpValue
											End Select
											objListProductRulesVO(idProdRef).listRelrulesLabel.add getID()&"-"&idProdRef&"|"&getLabel(), tmpOpValue
											'response.write("<br>objListProductRulesVO(idProdRef).resultAmount: "&objListProductRulesVO(idProdRef).resultAmount)
											'response.write("<br>6 arrivo prima di set - count: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.count)
											if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
												'response.write("set value 1 for: "&getID()&"-"&idProdOrig&"<br>")
												' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
											else
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
											end if
											Exit for
										Case 3
											if(CDbl(tmpImpProd)>CDbl(tmpImpRefProd)) then 
												tmpIdProdtoUse = idProdRef
												if(cLng(tmpApply4Qta)>cLng(tmpMaxRefQtadisp)) then tmpApply4Qta = tmpMaxRefQtadisp end if
											else
												tmpIdProdtoUse = idProdOrig
												if(cLng(tmpApply4Qta)>cLng(tmpMaxQtadisp)) then tmpApply4Qta = tmpMaxQtadisp end if												 
											end if											
											tmpOpValue = CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta)
											Select Case CInt(objListRulesConf(g).getOperationConf())
												Case 1
												objListProductRulesVO(tmpIdProdtoUse).resultAmount = objListProductRulesVO(tmpIdProdtoUse).resultAmount+(tmpOpValue)
												Case 2
												objListProductRulesVO(tmpIdProdtoUse).resultAmount = objListProductRulesVO(tmpIdProdtoUse).resultAmount-(tmpOpValue)
												tmpOpValue=0-tmpOpValue
											End Select
											objListProductRulesVO(tmpIdProdtoUse).listRelrulesLabel.add getID()&"-"&tmpIdProdtoUse&"|"&getLabel(), tmpOpValue
											'response.write("<br>6 arrivo prima di set - count: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.count)
											if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
												'response.write("set value 1 for: "&getID()&"-"&idProdOrig&"<br>")
												' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
											else
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
											end if
											Exit for										
										Case 4
											if(CDbl(tmpImpProd)>CDbl(tmpImpRefProd)) then 
												tmpIdProdtoUse = idProdOrig
												if(cLng(tmpApply4Qta)>cLng(tmpMaxQtadisp)) then tmpApply4Qta = tmpMaxQtadisp end if	
											else	
												tmpIdProdtoUse = idProdRef
												if(cLng(tmpApply4Qta)>cLng(tmpMaxRefQtadisp)) then tmpApply4Qta = tmpMaxRefQtadisp end if										 
											end if											
											tmpOpValue = CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta)
											Select Case CInt(objListRulesConf(g).getOperationConf())
												Case 1
												objListProductRulesVO(tmpIdProdtoUse).resultAmount = objListProductRulesVO(tmpIdProdtoUse).resultAmount+(tmpOpValue)
												Case 2
												objListProductRulesVO(tmpIdProdtoUse).resultAmount = objListProductRulesVO(tmpIdProdtoUse).resultAmount-(tmpOpValue)
												tmpOpValue=0-tmpOpValue
											End Select
											objListProductRulesVO(tmpIdProdtoUse).listRelrulesLabel.add getID()&"-"&tmpIdProdtoUse&"|"&getLabel(), tmpOpValue
											'response.write("<br>6 arrivo prima di set - count: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.count)
											if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
												'response.write("set value 1 for: "&getID()&"-"&idProdOrig&"<br>")
												' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
											else
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
											end if
											Exit for										
										Case 5
											tmpApply4QtaOrig=tmpApply4Qta
											tmpApply4QtaRef=tmpApply4Qta
											if(cLng(tmpApply4Qta)>cLng(tmpMaxQtadisp)) then tmpApply4QtaOrig = tmpMaxQtadisp end if
											if(cLng(tmpApply4Qta)>cLng(tmpMaxRefQtadisp)) then tmpApply4QtaRef = tmpMaxRefQtadisp end if	
											'response.write("<br>tmpApply4QtaOrig: "&tmpApply4QtaOrig)
											'response.write("<br>tmpApply4QtaRef: "&tmpApply4QtaRef)						
											tmpOpValueO = CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4QtaOrig)
											tmpOpValueR = CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4QtaRef)
											Select Case CInt(objListRulesConf(g).getOperationConf())
												Case 1
												objListProductRulesVO(idProdOrig).resultAmount = objListProductRulesVO(idProdOrig).resultAmount+(tmpOpValueO)
												objListProductRulesVO(idProdRef).resultAmount = objListProductRulesVO(idProdRef).resultAmount+(tmpOpValueR)
												Case 2
												objListProductRulesVO(idProdOrig).resultAmount = objListProductRulesVO(idProdOrig).resultAmount-(tmpOpValueO)
												objListProductRulesVO(idProdRef).resultAmount = objListProductRulesVO(idProdRef).resultAmount-(tmpOpValueR)
												tmpOpValueO=0-tmpOpValueO
												tmpOpValueR=0-tmpOpValueR
											End Select
											objListProductRulesVO(idProdOrig).listRelrulesLabel.add getID()&"-"&idProdOrig&"|"&getLabel(), tmpOpValueO
											objListProductRulesVO(idProdRef).listRelrulesLabel.add getID()&"-"&idProdRef&"|"&getLabel(), tmpOpValueR
											'response.write("<br>6 arrivo prima di set - count: "&objListProductRulesVO(idProdOrig).bolAlreadyApplied.count)
											if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
												'response.write("set value 1 for: "&getID()&"-"&idProdOrig&"<br>")
												' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
											else
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
											end if
											Exit for										
										Case else
									End Select
								end if
							end if
						next
						Set getAmountByStrategy.objListPRVO = objListProductRulesVO
					end if					
					Set objListRulesConf = nothing
				end if			
			Case 9
				Set getAmountByStrategy.objListPRVO = objListProductRulesVO
				bolHasRuleConf = false
				On Error Resume Next
				Set objListRulesConf = getListaRulesConfig(getID(), idProdOrig)
				if (Instr(1, typename(objListRulesConf), "Dictionary", 1) > 0) then
					bolHasRuleConf = true
				end if
				if(Err.number <> 0) then
					bolHasRuleConf = false
				end if

				if (bolHasRuleConf) then
					if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig)<>"1") then					
						tmpMaxQtadisp = objListProductRulesVO(idProdOrig).totQta
						tmpImpProd = objListProductRulesVO(idProdOrig).objProd.getPrezzo()
												
						for each g in objListRulesConf
							idProdRef = objListRulesConf(g).getProdRefConfID()
							if not(objListProductRulesVO.exists(idProdRef))then
								Exit for
							end if
							
							tmpMaxRefQtadisp = objListProductRulesVO(idProdRef).totQta
							tmpImpRefProd = objListProductRulesVO(idProdRef).objProd.getPrezzo()
							tmpRF = objListRulesConf(g).getRateFromConf()
							tmpRT = objListRulesConf(g).getRateToConf()
							tmpRFref = objListRulesConf(g).getRateFromRefConf()
							tmpRTref = objListRulesConf(g).getRateToRefConf()
							tmpApplyTo = objListRulesConf(g).getApplyToConf()
							tmpApply4Qta = objListRulesConf(g).getApply4QtaConf()						
							
							if(cLng(tmpMaxQtadisp)>=CDbl(tmpRF) AND cLng(tmpMaxQtadisp)<=CDbl(tmpRT))then							
								if(cLng(tmpMaxRefQtadisp)>=CDbl(tmpRFref) AND cLng(tmpMaxRefQtadisp)<=CDbl(tmpRTref))then
							
									Select Case CInt(tmpApplyTo)
										Case 1
											if(cLng(tmpApply4Qta)>cLng(tmpMaxQtadisp)) then tmpApply4Qta = tmpMaxQtadisp end if
											tmpOpValue = CDbl(tmpImpProd) / 100 * (CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta))
											Select Case CInt(objListRulesConf(g).getOperationConf())
												Case 1
												objListProductRulesVO(idProdOrig).resultAmount = objListProductRulesVO(idProdOrig).resultAmount+(tmpOpValue)
												Case 2
												objListProductRulesVO(idProdOrig).resultAmount = objListProductRulesVO(idProdOrig).resultAmount-(tmpOpValue)
												tmpOpValue=0-tmpOpValue
											End Select
											objListProductRulesVO(idProdOrig).listRelrulesLabel.add getID()&"-"&idProdOrig&"|"&getLabel(), tmpOpValue
											if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
												' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
											else
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
											end if
											Exit for
										Case 2
											if(cLng(tmpApply4Qta)>cLng(tmpMaxRefQtadisp)) then tmpApply4Qta = tmpMaxRefQtadisp end if
											tmpOpValue = CDbl(tmpImpRefProd) / 100 * (CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta))
											Select Case CInt(objListRulesConf(g).getOperationConf())
												Case 1
												objListProductRulesVO(idProdRef).resultAmount = objListProductRulesVO(idProdRef).resultAmount+(tmpOpValue)
												Case 2
												objListProductRulesVO(idProdRef).resultAmount = objListProductRulesVO(idProdRef).resultAmount-(tmpOpValue)
												tmpOpValue=0-tmpOpValue
											End Select
											objListProductRulesVO(idProdRef).listRelrulesLabel.add getID()&"-"&idProdRef&"|"&getLabel(), tmpOpValue
											if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
												' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
											else
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
											end if
											Exit for
										Case 3
											if(CDbl(tmpImpProd)>CDbl(tmpImpRefProd)) then 
												tmpIdProdtoUse = idProdRef
												if(cLng(tmpApply4Qta)>cLng(tmpMaxRefQtadisp)) then tmpApply4Qta = tmpMaxRefQtadisp end if
												tmpOpValue = CDbl(tmpImpRefProd) / 100 * (CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta))
											else
												tmpIdProdtoUse = idProdOrig
												if(cLng(tmpApply4Qta)>cLng(tmpMaxQtadisp)) then tmpApply4Qta = tmpMaxQtadisp end if
												tmpOpValue = CDbl(tmpImpProd) / 100 * (CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta))												 
											end if											
											Select Case CInt(objListRulesConf(g).getOperationConf())
												Case 1
												objListProductRulesVO(tmpIdProdtoUse).resultAmount = objListProductRulesVO(tmpIdProdtoUse).resultAmount+(tmpOpValue)
												Case 2
												objListProductRulesVO(tmpIdProdtoUse).resultAmount = objListProductRulesVO(tmpIdProdtoUse).resultAmount-(tmpOpValue)
												tmpOpValue=0-tmpOpValue
											End Select
											objListProductRulesVO(tmpIdProdtoUse).listRelrulesLabel.add getID()&"-"&tmpIdProdtoUse&"|"&getLabel(), tmpOpValue
											if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
												' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
											else
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
											end if
											Exit for										
										Case 4
											if(CDbl(tmpImpProd)>CDbl(tmpImpRefProd)) then 
												tmpIdProdtoUse = idProdOrig
												if(cLng(tmpApply4Qta)>cLng(tmpMaxQtadisp)) then tmpApply4Qta = tmpMaxQtadisp end if
												tmpOpValue = CDbl(tmpImpProd) / 100 * (CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta))	
											else	
												tmpIdProdtoUse = idProdRef
												if(cLng(tmpApply4Qta)>cLng(tmpMaxRefQtadisp)) then tmpApply4Qta = tmpMaxRefQtadisp end if
												tmpOpValue = CDbl(tmpImpRefProd) / 100 * (CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4Qta))										 
											end if											
											Select Case CInt(objListRulesConf(g).getOperationConf())
												Case 1
												objListProductRulesVO(tmpIdProdtoUse).resultAmount = objListProductRulesVO(tmpIdProdtoUse).resultAmount+(tmpOpValue)
												Case 2
												objListProductRulesVO(tmpIdProdtoUse).resultAmount = objListProductRulesVO(tmpIdProdtoUse).resultAmount-(tmpOpValue)
												tmpOpValue=0-tmpOpValue
											End Select
											objListProductRulesVO(tmpIdProdtoUse).listRelrulesLabel.add getID()&"-"&tmpIdProdtoUse&"|"&getLabel(), tmpOpValue
											if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
												' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
											else
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
											end if
											Exit for										
										Case 5
											tmpApply4QtaOrig=tmpApply4Qta
											tmpApply4QtaRef=tmpApply4Qta
											if(cLng(tmpApply4Qta)>cLng(tmpMaxQtadisp)) then tmpApply4QtaOrig = tmpMaxQtadisp end if
											if(cLng(tmpApply4Qta)>cLng(tmpMaxRefQtadisp)) then tmpApply4QtaRef = tmpMaxRefQtadisp end if
											tmpOpValueO = CDbl(tmpImpRefProd) / 100 * (CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4QtaOrig))
											tmpOpValueR = CDbl(tmpImpRefProd) / 100 * (CDbl(objListRulesConf(g).getValoreConf())*cLng(tmpApply4QtaRef))
											Select Case CInt(objListRulesConf(g).getOperationConf())
												Case 1
												objListProductRulesVO(idProdOrig).resultAmount = objListProductRulesVO(idProdOrig).resultAmount+(tmpOpValueO)
												objListProductRulesVO(idProdRef).resultAmount = objListProductRulesVO(idProdRef).resultAmount+(tmpOpValueR)
												Case 2
												objListProductRulesVO(idProdOrig).resultAmount = objListProductRulesVO(idProdOrig).resultAmount-(tmpOpValueO)
												objListProductRulesVO(idProdRef).resultAmount = objListProductRulesVO(idProdRef).resultAmount-(tmpOpValueR)
												tmpOpValueO=0-tmpOpValueO
												tmpOpValueR=0-tmpOpValueR
											End Select
											objListProductRulesVO(idProdOrig).listRelrulesLabel.add getID()&"-"&idProdOrig&"|"&getLabel(), tmpOpValueO
											objListProductRulesVO(idProdRef).listRelrulesLabel.add getID()&"-"&idProdRef&"|"&getLabel(), tmpOpValueR
											if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
												' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
												'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
											else
												objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
											end if
											Exit for										
										Case else
									End Select
								end if
							end if
						next
						Set getAmountByStrategy.objListPRVO = objListProductRulesVO
					end if					
					Set objListRulesConf = nothing
				end if			
			Case 10
				Set getAmountByStrategy.objListPRVO = objListProductRulesVO
				bolHasRuleConf = false
				On Error Resume Next
				Set objListRulesConf = getListaRulesConfig(getID(), idProdOrig)
				if (Instr(1, typename(objListRulesConf), "Dictionary", 1) > 0) then
					bolHasRuleConf = true
				end if
				if(Err.number <> 0) then
					bolHasRuleConf = false
				end if

				if (bolHasRuleConf) then
					if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig)<>"1") then					
						tmpMaxQtadisp = objListProductRulesVO(idProdOrig).totQta
						
						for each g in objListRulesConf
							tmpRF = objListRulesConf(g).getRateFromConf()
							tmpRT = objListRulesConf(g).getRateToConf()
							
							if(cLng(tmpMaxQtadisp)>=CDbl(tmpRF) AND cLng(tmpMaxQtadisp)<=CDbl(tmpRT))then
								objListProductRulesVO(idProdOrig).bolExcludeBills= true
								objListProductRulesVO(idProdOrig).listRelrulesLabel.add getID()&"-"&idProdOrig&"|"&getLabel(), 0
								if (objListProductRulesVO(idProdOrig).bolAlreadyApplied.Exists(getID()&"-"&idProdOrig)) then
									' commento vecchio modo per modificare valore esistente, uso invece il metodo obj.Item("key") = "value"
									'objListProductRulesVO(idProdOrig).bolAlreadyApplied.remove(getID()&"-"&idProdOrig)
									'objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
									objListProductRulesVO(idProdOrig).bolAlreadyApplied.Item(getID()&"-"&idProdOrig) = "1"
								else
									objListProductRulesVO(idProdOrig).bolAlreadyApplied.add getID()&"-"&idProdOrig, "1"
								end if
								Exit for
							end if
						next
						Set getAmountByStrategy.objListPRVO = objListProductRulesVO
					end if
					
					Set objListRulesConf = nothing
				end if					
			Case else
		End Select	

		if(Err.number <> 0) then
			Set getAmountByStrategy = new StrategyResultVO
		end if	
	End Function	


	Public Function getListaRules(rule_type, activate)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		getListaRules = null		
		strSQL = "SELECT * FROM business_rules"
		
		if (isNull(rule_type) AND isNull(activate)) then
			strSQL = "SELECT * FROM business_rules"
		else
			strSQL = strSQL & " WHERE"
			if not(isNull(rule_type)) then strSQL = strSQL & " AND rule_type IN("&rule_type&")"
			if not(isNull(activate)) then strSQL = strSQL & " AND activate=?"
		end if
		
		strSQL = strSQL & " ORDER BY rule_type ASC, label ASC;"
		strSQL = Replace(strSQL, "WHERE AND", "WHERE", 1, -1, 1)
		strSQL = Trim(strSQL)

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		if (isNull(activate)) then
		else
			if not(isNull(activate)) then objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,activate)			
		end if
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objRule
			do while not objRS.EOF				
				Set objRule = new BusinessRulesClass
				strID = objRS("id")
				objRule.setID(strID)
				objRule.setRuleType(objRS("rule_type"))
				objRule.setLabel(objRS("label"))	
				objRule.setDescrizione(objRS("description"))	
				objRule.setActivate(objRS("activate"))	
				objRule.setVoucherID(objRS("voucher_id"))	
				objDict.add strID, objRule
				Set objRule = nothing
				objRS.moveNext()
			loop							
			Set getListaRules = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

	
	
	Public Function findRuleByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		findRuleByID = null		
		strSQL = "SELECT * FROM business_rules WHERE id =?;"
		strSQL = Trim(strSQL)
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		Set objRS = objCommand.Execute()	
		
		if not(objRS.EOF) then			
			Dim objRule		
			Set objRule = new BusinessRulesClass
			strID = objRS("id")
			objRule.setID(strID)
			objRule.setRuleType(objRS("rule_type"))
			objRule.setLabel(objRS("label"))	
			objRule.setDescrizione(objRS("description"))	
			objRule.setActivate(objRS("activate"))	
			objRule.setVoucherID(objRS("voucher_id"))					
			Set findRuleByID = objRule			
			Set objRule = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
		
	Public Function insertRule(rule_type, label, description, activate, voucher_id, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		insertRule = -1
		
		strSQL = "INSERT INTO business_rules(rule_type, label, description, activate, voucher_id) VALUES("
		strSQL = strSQL & "?,?,?,?,"
		if(isNull(voucher_id) OR voucher_id = "") then
			strSQL = strSQL & "NULL"
		else
			strSQL = strSQL & "?"
		end if
		strSQL = strSQL & ");"
							
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,rule_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,description)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,activate)
		if not isNull(voucher_id) AND not(voucher_id = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,voucher_id)
		end if
		objCommand.Execute()
		Set objCommand = Nothing
		
		Set objRS = objConn.Execute("SELECT max(business_rules.id) as id FROM business_rules")
		if not (objRS.EOF) then
			objRS.MoveFirst()
			insertRule = objRS("id")	
		end if		
		Set objRS = Nothing	

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Function
		
	Public Sub modifyRule(id, rule_type, label, description, activate, voucher_id, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE business_rules SET "
		strSQL = strSQL & "rule_type=?,"
		strSQL = strSQL & "label=?,"
		strSQL = strSQL & "description=?,"
		strSQL = strSQL & "activate=?,"
		if(isNull(voucher_id) OR voucher_id = "") then
			strSQL = strSQL & "voucher_id=NULL"
		else
			strSQL = strSQL & "voucher_id=?"			
		end if
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,rule_type)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Parameters.Append objCommand.CreateParameter(,201,1,-1,description)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,activate)
		if not isNull(voucher_id) AND not(voucher_id = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,voucher_id)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Sub deleteRule(id)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL2 = "DELETE FROM business_rules_config WHERE id_rule=?;"
		strSQL = "DELETE FROM business_rules WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		Set objCommand2 = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand2.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand2.CommandType=1
		objCommand.CommandText = strSQL
		objCommand2.CommandText = strSQL2
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		objCommand2.Parameters.Append objCommand2.CreateParameter(,3,1,,id)
		
		objConn.BeginTrans
		
		if(Application("use_innodb_table") = 0) then
			objCommand2.Execute()
		end if
		objCommand.Execute()
		
		Set objCommand = Nothing
		Set objCommand2 = Nothing

		if objConn.Errors.Count = 0 then
			objConn.CommitTrans
		else
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		Set objDB = Nothing
				
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

'***************************************************** RULE ORDER ASSOCIATION METHODS

	Public Function findRuleOrderAssociationsByRule(id_rule)
		on error resume next
		Dim objDB, strSQL, objConn, objRS
		findRuleOrderAssociationsByRule = null	
		strSQL = "SELECT business_rules_x_ordine.*, business_rules.rule_type FROM business_rules_x_ordine LEFT JOIN business_rules ON business_rules_x_ordine.id_rule=business_rules.id WHERE id_rule=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_rule)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then							
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objRule
			do while not objRS.EOF				
				Set objRule = new BusinessRulesClass
				strID = objRS("id_rule")
				strOrderID = objRS("id_order")
				strProdID = objRS("id_prod")
				objRule.setID(strID)
				objRule.setOrderID(strOrderID)	
				objRule.setProdID(strProdID)
				objRule.setRuleType(objRS("rule_type"))	
				objRule.setCounterProd(objRS("counter_prod"))				
				objRule.setLabel(objRS("label"))	
				objRule.setValoreConf(objRS("valore"))	
				objDict.add strID&"-"&strOrderID&"-"&strProdID, objRule
				Set objRule = nothing
				objRS.moveNext()
			loop							
			Set findRuleOrderAssociationsByRule = objDict			
			Set objDict = nothing			
		end if	
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function findRuleOrderAssociationsByOrder(id_order, bolAdProd)
		on error resume next
		Dim objDB, strSQL, objConn, objRS
		findRuleOrderAssociationsByOrder = null	
		strSQL = "SELECT business_rules_x_ordine.*, business_rules.rule_type FROM business_rules_x_ordine LEFT JOIN business_rules ON business_rules_x_ordine.id_rule=business_rules.id WHERE id_order=?"
		if not(bolAdProd) then
			strSQL = strSQL&" AND id_prod=0"
		end if
		strSQL = strSQL&";"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then							
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objRule
			do while not objRS.EOF				
				Set objRule = new BusinessRulesClass
				strID = objRS("id_rule")
				strOrderID = objRS("id_order")
				strProdID = objRS("id_prod")
				objRule.setID(strID)
				objRule.setOrderID(strOrderID)	
				objRule.setProdID(strProdID)	
				objRule.setRuleType(objRS("rule_type"))
				objRule.setCounterProd(objRS("counter_prod"))
				objRule.setLabel(objRS("label"))	
				objRule.setValoreConf(objRS("valore"))	
				objDict.add strID&"-"&strOrderID&"-"&strProdID, objRule
				Set objRule = nothing
				objRS.moveNext()
			loop							
			Set findRuleOrderAssociationsByOrder = objDict			
			Set objDict = nothing				
		end if	
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function
	
	Public Function findRuleOrderAssociationsByOrderProd(id_order, id_prod)
		on error resume next
		Dim objDB, strSQL, objConn, objRS
		findRuleOrderAssociationsByOrderProd = null	
		strSQL = "SELECT business_rules_x_ordine.*, business_rules.rule_type FROM business_rules_x_ordine LEFT JOIN business_rules ON business_rules_x_ordine.id_rule=business_rules.id WHERE id_order=? AND id_prod=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prod)
		Set objRS = objCommand.Execute()	
		if not(objRS.EOF) then							
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objRule
			do while not objRS.EOF				
				Set objRule = new BusinessRulesClass
				strID = objRS("id_rule")
				strOrderID = objRS("id_order")
				strProdID = objRS("id_prod")
				objRule.setID(strID)
				objRule.setOrderID(strOrderID)	
				objRule.setProdID(strProdID)	
				objRule.setRuleType(objRS("rule_type"))
				objRule.setCounterProd(objRS("counter_prod"))
				objRule.setLabel(objRS("label"))	
				objRule.setValoreConf(objRS("valore"))	
				objDict.add strID&"-"&strOrderID&"-"&strProdID, objRule
				Set objRule = nothing
				objRS.moveNext()
			loop							
			Set findRuleOrderAssociationsByOrderProd = objDict			
			Set objDict = nothing				
		end if	
				
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

		
	Public Sub insertRuleOrder(id_rule, id_order, id_prod, counter_prod, label, valore, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO business_rules_x_ordine(id_rule, id_order, id_prod, counter_prod, label, valore) VALUES("
		strSQL = strSQL & "?,?,?,?,?,?);"
							
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_rule)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,counter_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
		objCommand.Execute()
		Set objCommand = Nothing	

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyRuleOrder(id_rule, id_order, id_prod, counter_prod, label, valore, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE business_rules_x_ordine SET "
		strSQL = strSQL & "id_rule=?,"
		strSQL = strSQL & "id_order=?,"
		strSQL = strSQL & "id_prod=?,"
		strSQL = strSQL & "counter_prod=?,"
		strSQL = strSQL & "label=?,"
		strSQL = strSQL & "valore=?"
		strSQL = strSQL & " WHERE id_rule=? AND id_order=? AND id_prod=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_rule)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,counter_prod)
		objCommand.Parameters.Append objCommand.CreateParameter(,200,1,100,label)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_rule)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prod)
		objCommand.Execute()	
		Set objCommand = Nothing

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Sub deleteRuleOrder(id_rule)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM business_rules_x_ordine WHERE id_rule=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_rule)
		objCommand.Execute()		
		Set objCommand = Nothing		
		Set objDB = Nothing
				
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Sub deleteRuleOrderByOrderIDNoTransaction(id_order)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM business_rules_x_ordine WHERE id_order=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Execute()		
		Set objCommand = Nothing		
		Set objDB = Nothing
				
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Sub deleteRuleOrderByOrderID(id_order, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM business_rules_x_ordine WHERE id_order=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Execute()		
		Set objCommand = Nothing		

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Sub deleteRuleOrderByOrderProdIDNoTransaction(id_order, id_prod)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM business_rules_x_ordine WHERE id_order=? AND id_prod=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prod)
		objCommand.Execute()		
		Set objCommand = Nothing		
		Set objDB = Nothing
				
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

	Public Sub deleteRuleOrderByOrderProdID(id_order, id_prod, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM business_rules_x_ordine WHERE id_order=? AND id_prod=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_order)
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prod)
		objCommand.Execute()		
		Set objCommand = Nothing		

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub

'***************************************************** RULE CONFIG METHODS

	Public Function getListaRulesConfig(id_rule, id_prod_orig)
		on error resume next
		Dim objDB, strSQL, objRS, objConn, objDict

		getListaRulesConfig = null		
		strSQL = "SELECT * FROM business_rules_config WHERE id_rule=?"		
		if not(isNull(id_prod_orig)) then strSQL = strSQL & " AND id_prod_orig=?"		
		strSQL = strSQL & " ORDER BY id_rule, id_prod_orig, rate_from, rate_to ASC;"
		strSQL = Trim(strSQL)

		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_rule)
		if not(isNull(id_prod_orig)) then objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id_prod_orig)
		Set objRS = objCommand.Execute()
		
		if not(objRS.EOF) then			
			Set objDict = Server.CreateObject("Scripting.Dictionary")
			Dim objRuleConf
			do while not objRS.EOF				
				Set objRuleConf = new BusinessRulesClass
				strID = objRS("id")
				objRuleConf.setID(strID)
				objRuleConf.setRuleID(objRS("id_rule"))
				objRuleConf.setProdOrigConfID(objRS("id_prod_orig"))
				objRuleConf.setProdRefConfID(objRS("id_prod_ref"))	
				objRuleConf.setRateFromConf(objRS("rate_from"))
				objRuleConf.setRateToConf(objRS("rate_to"))	
				objRuleConf.setRateFromRefConf(objRS("rate_from_ref"))
				objRuleConf.setRateToRefConf(objRS("rate_to_ref"))
				objRuleConf.setOperationConf(objRS("operation"))
				objRuleConf.setApplyToConf(objRS("applyto"))
				objRuleConf.setApply4QtaConf(objRS("apply_4_qta"))	
				objRuleConf.setValoreConf(objRS("valore"))
				objDict.add strID, objRuleConf
				objRS.moveNext()
			loop
			Set objRuleConf = nothing							
			Set getListaRulesConfig = objDict			
			Set objDict = nothing				
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function	
	
	Public Function findRuleConfigByID(id)
		on error resume next
		Dim objDB, strSQL, objRS, objConn

		findRuleConfigByID = null		
		strSQL = "SELECT * FROM business_rules_config WHERE id =?;"
		strSQL = Trim(strSQL)
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,20,1,,id)
		Set objRS = objCommand.Execute()	
		
		if not(objRS.EOF) then			
			Dim objRuleConf		
			Set objRuleConf = new BusinessRulesClass
			strID = objRS("id")
			objRuleConf.setID(strID)
			objRuleConf.setRuleID(objRS("id_rule"))
			objRuleConf.setProdOrigConfID(objRS("id_prod_orig"))
			objRuleConf.setProdRefConfID(objRS("id_prod_ref"))	
			objRuleConf.setRateFromConf(objRS("rate_from"))
			objRuleConf.setRateToConf(objRS("rate_to"))
			objRuleConf.setRateFromRefConf(objRS("rate_from_ref"))
			objRuleConf.setRateToRefConf(objRS("rate_to_ref"))
			objRuleConf.setOperationConf(objRS("operation"))
			objRuleConf.setApplyToConf(objRS("applyto"))
			objRuleConf.setApply4QtaConf(objRS("apply_4_qta"))	
			objRuleConf.setValoreConf(objRS("valore"))								
			Set findRuleConfigByID = objRuleConf			
			Set objRuleConf = nothing					
		end if
		
		Set objRS = Nothing
		Set objCommand = Nothing
		Set objDB = Nothing
 
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if		
	End Function

		
	Public Sub insertRuleConfig(id_rule, id_prod_orig, id_prod_ref, rate_from, rate_to, rate_from_ref, rate_to_ref, operation, applyto, apply_4_qta, valore, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		
		strSQL = "INSERT INTO business_rules_config(id_rule, id_prod_orig, id_prod_ref, rate_from, rate_to, rate_from_ref, rate_to_ref, operation, applyto, apply_4_qta, `valore`) VALUES("
		strSQL = strSQL & "?,"
		if(isNull(id_prod_orig) OR id_prod_orig = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if
		if(isNull(id_prod_ref) OR id_prod_ref = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if		
		strSQL = strSQL & "?,?,"
		if(isNull(rate_from_ref) OR rate_from_ref = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if
		if(isNull(rate_to_ref) OR rate_to_ref = "") then
			strSQL = strSQL & "NULL,"
		else
			strSQL = strSQL & "?,"
		end if		
		strSQL = strSQL & "?,?,?,?);"
						
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_rule)
		if not isNull(id_prod_orig) AND not(id_prod_orig = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod_orig)
		end if
		if not isNull(id_prod_ref) AND not(id_prod_ref = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod_ref)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate_from))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate_to))
		if not isNull(rate_from_ref) AND not(rate_from_ref = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate_from_ref))
		end if
		if not isNull(rate_to_ref) AND not(rate_to_ref = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate_to_ref))
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,operation)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,applyto)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,apply_4_qta)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
		objCommand.Execute()
		Set objCommand = Nothing

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub modifyRuleConfig(id, id_rule, id_prod_orig, id_prod_ref, rate_from, rate_to, rate_from_ref, rate_to_ref, operation, applyto, apply_4_qta, valore, objConn)
		on error resume next
		Dim objDB, strSQL, objRS
		strSQL = "UPDATE business_rules_config SET "
		strSQL = strSQL & "id_rule=?,"
		if(isNull(id_prod_orig) OR id_prod_orig = "") then
			strSQL = strSQL & "id_prod_orig=NULL,"
		else
			strSQL = strSQL & "id_prod_orig=?,"			
		end if
		if(isNull(id_prod_ref) OR id_prod_ref = "") then
			strSQL = strSQL & "id_prod_ref=NULL,"
		else
			strSQL = strSQL & "id_prod_ref=?,"			
		end if
		strSQL = strSQL & "rate_from=?,"
		strSQL = strSQL & "rate_to=?,"
		if(isNull(rate_from_ref) OR rate_from_ref = "") then
			strSQL = strSQL & "rate_from_ref=NULL,"
		else
			strSQL = strSQL & "rate_from_ref=?,"
		end if
		if(isNull(rate_to_ref) OR rate_to_ref = "") then
			strSQL = strSQL & "rate_to_ref=NULL,"
		else
			strSQL = strSQL & "rate_to_ref=?,"
		end if
		strSQL = strSQL & "operation=?,"	
		strSQL = strSQL & "applyto=?,"	
		strSQL = strSQL & "apply_4_qta=?,"		
		strSQL = strSQL & "`valore`=?"
		strSQL = strSQL & " WHERE id=?;"

		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL

		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_rule)
		if not isNull(id_prod_orig) AND not(id_prod_orig = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod_orig)
		end if
		if not isNull(id_prod_ref) AND not(id_prod_ref = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_prod_ref)
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate_from))
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate_to))
		if not isNull(rate_from_ref) AND not(rate_from_ref = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate_from_ref))
		end if
		if not isNull(rate_to_ref) AND not(rate_to_ref = "") then
			objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(rate_to_ref))
		end if
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,operation)
		objCommand.Parameters.Append objCommand.CreateParameter(,2,1,,applyto)
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,apply_4_qta)
		objCommand.Parameters.Append objCommand.CreateParameter(,5,1,,convertDoubleDelimiter(valore))
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()	
		Set objCommand = Nothing

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
		
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteRuleConfigNoTransaction(id)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM business_rules_config WHERE id=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
				
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteRuleConfig(id, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM business_rules_config WHERE id=?;"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id)
		objCommand.Execute()
		Set objCommand = Nothing

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
				
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteRuleConfigByRuleNoTransaction(id_rule)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM business_rules_config WHERE id_rule=?;"
		
		Set objDB = New DBManagerClass
		Set objConn = objDB.openConnection()	
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_rule)
		objCommand.Execute()
		Set objCommand = Nothing
		Set objDB = Nothing
				
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
		
	Public Sub deleteRuleConfigByRule(id_rule, objConn)
		on error resume next
		Dim objDB, strSQL, objRS 
		strSQL = "DELETE FROM business_rules_config WHERE id_rule=?;"
		
		Dim objCommand
		Set objCommand = Server.CreateObject("ADODB.Command")
		objCommand.ActiveConnection = objConn
		objCommand.CommandType=1
		objCommand.CommandText = strSQL
		objCommand.Parameters.Append objCommand.CreateParameter(,21,1,,id_rule)
		objCommand.Execute()
		Set objCommand = Nothing

		if objConn.Errors.Count > 0 then
			objConn.RollBackTrans
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
				
		if Err.number <> 0 then
			response.Redirect(Application("baseroot")&Application("error_page")&"?error="&Err.description)
		end if
	End Sub
	
	Public Function convertDoubleDelimiter(doubleValue)
		convertDoubleDelimiter = doubleValue
		
		'if (Application("dbType") = 0) then
			convertDoubleDelimiter = Replace(convertDoubleDelimiter, ".",",")
		'else		
			'convertDoubleDelimiter = Replace(convertDoubleDelimiter, ",",".")
		'end if			
	End Function
End Class


'******************************* CREO UNA INNER CLASS PER GESTIRE LA RISPOSTA DI RITORNO DEL CALCOLO BUSINESS STRATEGY
Class StrategyResultVO
	Public foundAmount
	Public objListPRVO

	Private Sub Class_Initialize()
		foundAmount = 0
		objListPRVO = null
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

'******************************* CREO UNA INNER CLASS PER GENERARE UNA COLLECTION DI PRODOTTI DA UTILIZZARE NELL'ELABORAZIONE DELLA LISTA RULES X PRODOTTO
Class ProductRulesVO
	Public idProd
	Public counterProd
	Public objProd
	Public totQta
	Public resultAmount
	Public bolAlreadyApplied
	Public listRelrulesLabel
	Public bolExcludeBills

	Private Sub Class_Initialize()
		idProd = -1
		counterProd = 0
		objProd = null
		totQta = 0
		resultAmount = 0
		bolExcludeBills = false
		Set bolAlreadyApplied = Server.CreateObject("Scripting.Dictionary")
		Set listRelrulesLabel = Server.CreateObject("Scripting.Dictionary")
	End Sub

	Private Sub Class_Terminate()
		Set bolAlreadyApplied = nothing
		Set listRelrulesLabel = nothing
	End Sub
  
	Public Function toString()
		toString="idProd: "&idProd&" - totQta: "&totQta&" - resultAmount: "&resultAmount
	End function 
End Class


'******************************* CREO UNA INNER CLASS PER GENERARE UNA COLLECTION DA UTILIZZARE NELL'ELABORAZIONE DELLA LISTA RULES
Class OrderRulesVO
	Public idRule
	Public label
	Public amount

	Private Sub Class_Initialize()
		idRule = -1
		label = ""
		amount = 0
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class
%>