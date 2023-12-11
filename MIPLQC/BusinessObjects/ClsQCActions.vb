Imports System.IO
Imports SAPbouiCOM.Framework
Public Class ClsQCActions
    Public Const Formtype = "MIQCACT"
    Dim objForm As SAPbouiCOM.Form
    Dim ObjQCForm As SAPbouiCOM.Form
    Dim strSQL As String, strQuery As String
    Dim objMatrix As SAPbouiCOM.Matrix
    Public QCACTHeader As SAPbouiCOM.DBDataSource
    Public QCACTLine As SAPbouiCOM.DBDataSource
    Dim objRecordSet As SAPbobsCOM.Recordset
    Dim Formcount As Integer = 0
    Dim objSelect As SAPbouiCOM.CheckBox

    Public Sub LoadScreen()
        Try
            objForm = objAddOn.objUIXml.LoadScreenXML("Rej_Rew_Actions.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
            objForm.PaneLevel = 1
            objMatrix = objForm.Items.Item("20").Specific
            LoadSeries(objForm.UniqueID)
            objForm.Items.Item("21").Enabled = True
            objForm.Items.Item("22").Enabled = True
        Catch ex As Exception
        End Try
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim CurRow As Integer = 0
            Dim ObjQCForm, objBPForm, objAPForm As SAPbouiCOM.Form
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim QCNum As String = ""
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
            If pVal.BeforeAction = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim flag As Boolean = False
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                                If objSelect.Checked = True Then
                                    If objMatrix.Columns.Item("11").Cells.Item(i).Specific.String <> "" Or objMatrix.Columns.Item("11A").Cells.Item(i).Specific.String <> "" Then
                                        flag = True
                                    End If
                                    'If (objMatrix.Columns.Item("9").Cells.Item(i).Specific.String = "Inventory Transfer & A/P Credit Memo" Or objMatrix.Columns.Item("9").Cells.Item(i).Specific.String = "Goods Issue & A/P Credit Memo") Then
                                    '    If objMatrix.Columns.Item("11").Cells.Item(i).Specific.String <> "" And objMatrix.Columns.Item("11A").Cells.Item(i).Specific.String <> "" Then
                                    '        flag = True
                                    '    End If
                                    'Else
                                    '    If objMatrix.Columns.Item("11").Cells.Item(i).Specific.String <> "" Then
                                    '        flag = True
                                    '    End If
                                    'End If
                                End If
                            Next
                            If flag = True Then
                                objForm.Items.Item("1").Click()
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "1" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                            If Not validate(FormUID) Then
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                        If pVal.ItemUID = "22" Then
                            Dim SelRow As Boolean = False
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                                If objSelect.Checked = True Then
                                    SelRow = True
                                    If objMatrix.Columns.Item("10").Cells.Item(i).Specific.String = "" Then
                                        objAddOn.objApplication.StatusBar.SetText("Please update the BP Code on line " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                    End If

                                    If objMatrix.Columns.Item("7C").Cells.Item(i).Specific.String = "" And objMatrix.Columns.Item("9").Cells.Item(i).Specific.String Like "Inventory*" Then
                                        objAddOn.objApplication.SetStatusBarMessage("Please select to Warehouse!!!", SAPbouiCOM.BoMessageTime.bmt_Long, True)
                                        objMatrix.Columns.Item("7C").Cells.Item(i).Click()
                                        BubbleEvent = False
                                    ElseIf objMatrix.Columns.Item("8").Cells.Item(i).Specific.String = "" Then
                                        objAddOn.objApplication.SetStatusBarMessage("Please select the user Action!!!", SAPbouiCOM.BoMessageTime.bmt_Long, True)
                                        objMatrix.Columns.Item("8").Cells.Item(i).Click()
                                        BubbleEvent = False
                                    ElseIf objMatrix.Columns.Item("9").Cells.Item(i).Specific.String = "" Then
                                        objAddOn.objApplication.SetStatusBarMessage("Please select the posting document!!!", SAPbouiCOM.BoMessageTime.bmt_Long, True)
                                        objMatrix.Columns.Item("9").Cells.Item(i).Click()
                                        BubbleEvent = False
                                    End If
                                End If
                            Next
                            If SelRow = False Then
                                objAddOn.objApplication.StatusBar.SetText("Please select atleast a line to make QC action...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                            End If
                        ElseIf pVal.ItemUID = "21A" Then
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                Dim TFlag As Boolean = False
                                For i As Integer = 1 To objMatrix.VisualRowCount
                                    If objMatrix.Columns.Item("11").Cells.Item(i).Specific.String <> "" Then
                                        TFlag = True
                                        Exit For
                                    End If
                                Next
                                If TFlag = True Then objAddOn.objApplication.StatusBar.SetText("You should not clear the selections due to transactions created...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                                ClearSelections(FormUID)
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.ItemUID = "20" And pVal.ColUID = "8" Then
                            AddCFLCondition(FormUID, pVal.Row)
                        ElseIf pVal.ItemUID = "20" And pVal.ColUID = "9" Then
                            AddCFLCondition_Tran(FormUID, pVal.Row)
                        ElseIf pVal.ItemUID = "20" And pVal.ColUID = "10A" Then
                            AddCFLCondition_BP(FormUID, pVal.Row)
                        ElseIf pVal.ItemUID = "20" And pVal.ColUID = "7C" Then
                            setCFLCond(FormUID, "CFL_TW", pVal.Row)
                        Else
                            If pVal.ItemUID = "17" Then 'QC
                                AddCFLCondition_Header(FormUID, "CFL_QC", "DocEntry", "DocEntry")
                            ElseIf pVal.ItemUID = "12" Then 'Warehouse
                                AddCFLCondition_Header(FormUID, "CFL_W", "U_RejWhse", "WhsCode")
                            ElseIf pVal.ItemUID = "8" Then 'ItemCode
                                AddCFLCondition_Header(FormUID, "CFL_I", "U_ItemCode", "ItemCode")
                            ElseIf pVal.ItemUID = "10" Then 'VendorCode
                                AddCFLCondition_Header(FormUID, "CFL_BP", "U_Vendor", "CardName")
                            ElseIf pVal.ItemUID = "tcc" Then 'CostCenter
                                CFLCondition(FormUID, "CFL_CC", "Locked", "Y")
                            ElseIf pVal.ItemUID = "tinvnum" Then 'Inv Transfer
                                CFLCondition(FormUID, "CFL_Transfer", "U_QCNum", "")
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "26" Then
                            Try
                                objAddOn.objApplication.Menus.Item("MIPLQC").Activate()
                                ObjQCForm = objAddOn.objApplication.Forms.ActiveForm ' objAddOn.objApplication.Forms.GetForm("MIPLQC", 1)
                                ObjQCForm.Freeze(True)
                                ObjQCForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                ObjQCForm.Items.Item("6E").Enabled = True
                                ObjQCForm.Items.Item("6E").Specific.String = objForm.Items.Item("17").Specific.String
                                ObjQCForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                ObjQCForm.Freeze(False)
                            Catch ex As Exception
                                ObjQCForm.Freeze(False)
                                ObjQCForm = Nothing
                            End Try
                        End If
                        If pVal.ItemUID = "20" And pVal.ColUID = "0A" Then
                            Try
                                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    'objSelect.Checked = True
                                    objMatrix.SelectRow(pVal.Row, True, True)
                                    'objMatrix.Columns.Item("10A").Cells.Item(pVal.Row).Click()
                                    objMatrix.Columns.Item("8").Cells.Item(pVal.Row).Click()
                                End If
                            Catch ex As Exception
                            End Try
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        If (pVal.ItemUID = "20" And pVal.ColUID = "1") Then
                            Try
                                objAddOn.objApplication.Menus.Item("MIPLQC").Activate()
                                ObjQCForm = objAddOn.objApplication.Forms.ActiveForm 'objAddOn.objApplication.Forms.GetForm("MIPLQC", 1)
                                ObjQCForm.Freeze(True)
                                ObjQCForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                ObjQCForm.Items.Item("6E").Enabled = True
                                ObjQCForm.Items.Item("6E").Specific.String = objMatrix.Columns.Item("1B").Cells.Item(pVal.Row).Specific.String
                                ObjQCForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                ObjQCForm.Freeze(False)
                            Catch ex As Exception
                                ObjQCForm.Freeze(False)
                                ObjQCForm = Nothing
                            End Try
                        ElseIf (pVal.ItemUID = "20" And pVal.ColUID = "10A" Or pVal.ColUID = "2B") Then
                            Try
                                objAddOn.objApplication.Menus.Item("2561").Activate()
                                objBPForm = objAddOn.objApplication.Forms.ActiveForm 'objAddOn.objApplication.Forms.GetForm("134", Formcount)
                                objBPForm.Freeze(True)
                                If pVal.ColUID = "10A" Then
                                    objBPForm.Items.Item("5").Specific.String = objMatrix.Columns.Item("10").Cells.Item(pVal.Row).Specific.String
                                ElseIf pVal.ColUID = "2B" Then
                                    objBPForm.Items.Item("5").Specific.String = objMatrix.Columns.Item("2A").Cells.Item(pVal.Row).Specific.String
                                End If
                                objBPForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objBPForm.Freeze(False)
                            Catch ex As Exception
                                objBPForm.Freeze(False)
                                objBPForm = Nothing
                            End Try
                        ElseIf (pVal.ItemUID = "20" And pVal.ColUID = "11") Then
                            Try
                                Dim ActualEntry As String = ""
                                Dim ColItem As SAPbouiCOM.Column = objMatrix.Columns.Item("11")
                                Dim link As SAPbouiCOM.LinkedButton = ColItem.ExtendedObject
                                link.LinkedObjectType = "-1"
                                If objMatrix.Columns.Item("9").Cells.Item(pVal.Row).Specific.String = "A/R Invoice" Then
                                    link.LinkedObjectType = "13"
                                ElseIf objMatrix.Columns.Item("9").Cells.Item(pVal.Row).Specific.String = "Inventory Transfer" Then
                                    If objAddOn.HANA Then
                                        ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T0.""DocEntry"" from ODRF T0 where ""ObjType""=67 and ifnull(T0.""DocStatus"",'')='O' and T0.""DocEntry""='" & objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.String & "'")
                                    Else
                                        ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T0.DocEntry from ODRF T0 where ObjType=67 and isnull(T0.DocStatus,'')='O' and T0.DocEntry='" & objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.String & "'")
                                    End If
                                    If ActualEntry = "" Then
                                        link.LinkedObjectType = "67"
                                    Else
                                        link.LinkedObjectType = "112"
                                    End If
                                ElseIf objMatrix.Columns.Item("9").Cells.Item(pVal.Row).Specific.String = "Inventory Transfer & A/P Credit Memo" Then
                                    'Dim splitval() As String = Split(objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.String, ",")
                                    If objAddOn.HANA Then
                                        ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T0.""DocEntry"" from ODRF T0 where ""ObjType""=67 and ifnull(T0.""DocStatus"",'')='O' and T0.""DocEntry""='" & objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.String & "'")
                                    Else
                                        ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T0.DocEntry from ODRF T0 where ObjType=67 and isnull(T0.DocStatus,'')='O' and T0.DocEntry='" & objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.String & "'")
                                    End If
                                    If ActualEntry = "" Then
                                        link.LinkedObjectType = "67"
                                    Else
                                        link.LinkedObjectType = "112"
                                    End If
                                    Try
                                        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        objRecordSet.DoQuery("select DocNum,Format(DocDate,'yyyyMMdd') from ORPC where DocEntry='" & objMatrix.Columns.Item("11A").Cells.Item(pVal.Row).Specific.String & "'")
                                        If objRecordSet.RecordCount > 0 Then
                                            objAddOn.objApplication.Menus.Item("2309").Activate()
                                            objAPForm = objAddOn.objApplication.Forms.ActiveForm 'objAddOn.objApplication.Forms.GetForm("MIPLQC", 1)
                                            objAPForm.Freeze(True)
                                            objAPForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                            'objAPForm.Items.Item("8").Enabled = True
                                            objAPForm.Items.Item("8").Specific.String = objRecordSet.Fields.Item(0).Value.ToString
                                            objAPForm.Items.Item("10").Specific.String = objRecordSet.Fields.Item(1).Value.ToString
                                            objAPForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            objAPForm.Freeze(False)
                                        Else
                                            Exit Sub
                                        End If
                                    Catch ex As Exception
                                        objAPForm.Freeze(False)
                                        objAPForm = Nothing
                                    End Try

                                ElseIf objMatrix.Columns.Item("9").Cells.Item(pVal.Row).Specific.String = "A/P Credit Memo" Then
                                    Try
                                        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        objRecordSet.DoQuery("select DocNum,Format(DocDate,'yyyyMMdd') from ORPC where DocEntry='" & objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.String & "'")
                                        If objRecordSet.RecordCount > 0 Then
                                            objAddOn.objApplication.Menus.Item("2309").Activate()
                                            objAPForm = objAddOn.objApplication.Forms.ActiveForm 'objAddOn.objApplication.Forms.GetForm("MIPLQC", 1)
                                            objAPForm.Freeze(True)
                                            objAPForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                            'objAPForm.Items.Item("8").Enabled = True
                                            objAPForm.Items.Item("8").Specific.String = objRecordSet.Fields.Item(0).Value.ToString
                                            objAPForm.Items.Item("10").Specific.String = objRecordSet.Fields.Item(1).Value.ToString
                                            objAPForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            objAPForm.Freeze(False)
                                        Else
                                            Exit Sub
                                        End If
                                    Catch ex As Exception
                                        objAPForm.Freeze(False)
                                        objAPForm = Nothing
                                    End Try
                                ElseIf objMatrix.Columns.Item("9").Cells.Item(pVal.Row).Specific.String = "Goods Issue & A/P Credit Memo" Then
                                    'Dim splitval() As String = Split(objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.String, ",")
                                    If objAddOn.HANA Then
                                        ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T0.""DocEntry"" from ODRF T0 where ""ObjType""=60 and ifnull(T0.""DocStatus"",'')='O' and T0.""DocEntry""='" & objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.String & "'")
                                    Else
                                        ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T0.DocEntry from ODRF T0 where ObjType=60 and isnull(T0.DocStatus,'')='O' and T0.DocEntry='" & objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.String & "'")
                                    End If
                                    If ActualEntry = "" Then
                                        link.LinkedObjectType = "60"
                                    Else
                                        link.LinkedObjectType = "112"
                                    End If
                                    Try
                                        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        objRecordSet.DoQuery("select DocNum,Format(DocDate,'yyyyMMdd') from ORPC where DocEntry='" & objMatrix.Columns.Item("11A").Cells.Item(pVal.Row).Specific.String & "'")
                                        If objRecordSet.RecordCount > 0 Then
                                            objAddOn.objApplication.Menus.Item("2309").Activate()
                                            objAPForm = objAddOn.objApplication.Forms.ActiveForm 'objAddOn.objApplication.Forms.GetForm("MIPLQC", 1)
                                            objAPForm.Freeze(True)
                                            objAPForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                            'objAPForm.Items.Item("8").Enabled = True
                                            objAPForm.Items.Item("8").Specific.String = objRecordSet.Fields.Item(0).Value.ToString
                                            objAPForm.Items.Item("10").Specific.String = objRecordSet.Fields.Item(1).Value.ToString
                                            objAPForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            objAPForm.Freeze(False)
                                        End If
                                    Catch ex As Exception
                                        objAPForm.Freeze(False)
                                        objAPForm = Nothing
                                    End Try
                                ElseIf objMatrix.Columns.Item("9").Cells.Item(pVal.Row).Specific.String = "Goods Issue" Then
                                    Try
                                        link.LinkedObjectType = "60"
                                    Catch ex As Exception

                                    End Try

                                ElseIf objMatrix.Columns.Item("9").Cells.Item(pVal.Row).Specific.String = "Sub-Contracting" Then
                                    Try
                                        objAddOn.objApplication.Menus.Item("SUBCTPO").Activate()
                                        ObjQCForm = objAddOn.objApplication.Forms.ActiveForm 'objAddOn.objApplication.Forms.GetForm("MIPLQC", 1)
                                        ObjQCForm.Freeze(True)
                                        ObjQCForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        ObjQCForm.Items.Item("txtentry").Enabled = True
                                        ObjQCForm.Items.Item("txtentry").Specific.String = objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.String
                                        ObjQCForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        ObjQCForm.Freeze(False)
                                    Catch ex As Exception
                                        ObjQCForm.Freeze(False)
                                        ObjQCForm = Nothing
                                    End Try
                                End If

                            Catch ex As Exception
                            End Try
                        End If

                End Select
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        objForm.Items.Item("21").Left = objForm.Items.Item("17").Left
                        objForm.Items.Item("21A").Left = objForm.Items.Item("21").Left + objForm.Items.Item("21").Width + 3
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If (pVal.ItemUID = "20" And pVal.ColUID = "10A") Then
                            Try
                                objForm.Freeze(True)
                                objMatrix.AutoResizeColumns()
                                objForm.Freeze(False)
                            Catch ex As Exception
                                objForm.Freeze(False)
                            End Try
                        ElseIf (pVal.ItemUID = "20" And (pVal.ColUID = "8" Or pVal.ColUID = "9")) Then
                            Try
                                If objMatrix.Columns.Item("9").Cells.Item(pVal.Row).Specific.String = "" Then Exit Sub
                                If objMatrix.Columns.Item("7C").Cells.Item(pVal.Row).Specific.String = "" Then
                                    If objMatrix.Columns.Item("10").Cells.Item(pVal.Row).Specific.string <> "" And objMatrix.Columns.Item("9").Cells.Item(pVal.Row).Specific.String <> "Sub-Contracting" Then 'And Trim(objDataTable.GetValue("U_TType", 0)) <> "Sub-Contracting"
                                        objMatrix.Columns.Item("7C").Cells.Item(pVal.Row).Specific.String = objAddOn.objGenFunc.getSingleValue("select U_WAREHOUSE from OCRD where CardCode='" & objMatrix.Columns.Item("10").Cells.Item(pVal.Row).Specific.string & "' ")
                                    End If
                                End If
                            Catch ex As Exception
                            End Try
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "21" Then
                            If objForm.Items.Item("21").Enabled = False Then Exit Sub
                            Dim TFlag As Boolean = False
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If objMatrix.Columns.Item("11").Cells.Item(i).Specific.String <> "" Then
                                    TFlag = True
                                    Exit For
                                End If
                            Next
                            If TFlag = True Then objAddOn.objApplication.StatusBar.SetText("You should not reload the data due to transactions created...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                            If objForm.Items.Item("21").Enabled = True Then
                                LoadQCDetails(FormUID)
                            End If

                        ElseIf pVal.ItemUID = "22" Then
                            If objForm.Items.Item("22").Enabled = True Then
                                For i As Integer = 1 To objMatrix.VisualRowCount
                                    objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                                    If objSelect.Checked = True Then
                                        If objMatrix.Columns.Item("9").Cells.Item(i).Specific.String <> "" Then
                                            CurRow = i
                                            Exit For
                                        End If
                                    End If
                                Next
                                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    Dim FuncFlag As Boolean = False
                                    If objMatrix.Columns.Item("9").Cells.Item(CurRow).Specific.String = "A/R Invoice" Then
                                        'Create_AR_Invoice(FormUID)
                                    ElseIf objMatrix.Columns.Item("9").Cells.Item(CurRow).Specific.String = "Inventory Transfer" Then
                                        If Create_InventoryTransfer(FormUID) Then
                                            FuncFlag = True
                                        End If
                                    ElseIf objMatrix.Columns.Item("9").Cells.Item(CurRow).Specific.String = "Inventory Transfer & A/P Credit Memo" Then
                                        strQuery = objAddOn.objGenFunc.getSingleValue("Select 1 as Status FROM [@MIQCACT1] T2 INNER JOIN [@MIQCACT] T3 ON T3.DocEntry=T2.DocEntry where  T2.U_Select='Y' and T2.U_TranEntry<>'' and T2.U_QCEntry='" & objMatrix.Columns.Item("1B").Cells.Item(CurRow).Specific.String & "'")
                                        If strQuery = "" Then
                                            If Create_InventoryTransfer(FormUID) Then
                                                FuncFlag = True
                                            Else
                                                Exit Sub
                                            End If
                                        Else
                                            If objMatrix.Columns.Item("11").Cells.Item(CurRow).Specific.String = "" Then
                                                strSQL = objAddOn.objGenFunc.getSingleValue("Select T2.U_TranEntry FROM [@MIQCACT1] T2 INNER JOIN [@MIQCACT] T3 ON T3.DocEntry=T2.DocEntry where T2.U_Select='Y' and T2.U_QCEntry='" & objMatrix.Columns.Item("1B").Cells.Item(CurRow).Specific.String & "' order by T3.DocEntry Desc")
                                                If strSQL <> "" Then
                                                    objAddOn.objApplication.MessageBox("Inventory Transfer already posted for the Line No : " & CStr(CurRow), , "OK")
                                                    objMatrix.Columns.Item("11").Cells.Item(CurRow).Specific.String = strSQL
                                                End If
                                            End If
                                        End If
                                        'If Create_APCreditMemo_Manual(FormUID) Then
                                        '    FuncFlag = True
                                        'End If
                                        If Create_AP_Invoice_CreditMemo_Manual(FormUID) Then
                                            FuncFlag = True
                                        End If
                                    ElseIf objMatrix.Columns.Item("9").Cells.Item(CurRow).Specific.String = "Sub-Contracting" Then
                                        If Create_SubContracting(FormUID) Then
                                            FuncFlag = True
                                        End If
                                    ElseIf objMatrix.Columns.Item("9").Cells.Item(CurRow).Specific.String = "A/P Credit Memo" Then
                                        'Create_AP_Invoice_CreditMemo(FormUID)
                                        If Create_AP_Invoice_CreditMemo_Manual(FormUID) Then
                                            FuncFlag = True
                                        End If
                                    ElseIf objMatrix.Columns.Item("9").Cells.Item(CurRow).Specific.String = "Goods Issue & A/P Credit Memo" Then
                                        strQuery = objAddOn.objGenFunc.getSingleValue("Select 1 as Status FROM [@MIQCACT1] T2 INNER JOIN [@MIQCACT] T3 ON T3.DocEntry=T2.DocEntry where  T2.U_Select='Y' and T2.U_TranEntry<>'' and T2.U_QCEntry='" & objMatrix.Columns.Item("1B").Cells.Item(CurRow).Specific.String & "'")
                                        If strQuery = "" Then
                                            If Create_GoodsIssue(FormUID) Then
                                                FuncFlag = True
                                            Else
                                                Exit Sub
                                            End If
                                        Else
                                            If objMatrix.Columns.Item("11").Cells.Item(CurRow).Specific.String = "" Then
                                                strSQL = objAddOn.objGenFunc.getSingleValue("Select T2.U_TranEntry FROM [@MIQCACT1] T2 INNER JOIN [@MIQCACT] T3 ON T3.DocEntry=T2.DocEntry where T2.U_Select='Y' and T2.U_QCEntry='" & objMatrix.Columns.Item("1B").Cells.Item(CurRow).Specific.String & "' order by T3.DocEntry Desc")
                                                If strSQL <> "" Then
                                                    objAddOn.objApplication.MessageBox("Goods Issue already posted for the Line No : " & CStr(CurRow), , "OK")
                                                    objMatrix.Columns.Item("11").Cells.Item(CurRow).Specific.String = strSQL
                                                End If
                                            End If
                                        End If
                                        If Create_APCreditMemo_Manual(FormUID) Then
                                            FuncFlag = True
                                        End If
                                    ElseIf objMatrix.Columns.Item("9").Cells.Item(CurRow).Specific.String = "Goods Issue" Then
                                        If Create_GoodsIssue(FormUID) Then
                                            FuncFlag = True
                                        Else
                                            Exit Sub
                                        End If
                                    End If
                                    If FuncFlag = True Then
                                        objMatrix.Columns.Item("0A").Editable = False
                                    Else
                                        objForm.Items.Item("21").Enabled = True
                                        objForm.Items.Item("22").Enabled = True
                                    End If
                                End If
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If pVal.ItemUID = "14" Then
                            QCACTHeader.SetValue("DocNum", 0, objAddOn.objGenFunc.GetDocNum(Formtype, CInt(objForm.Items.Item("14").Specific.Selected.value)))
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.ItemUID = "8" Or pVal.ItemUID = "10" Or pVal.ItemUID = "12" Or pVal.ItemUID = "17" Or pVal.ItemUID = "tcc" Or pVal.ItemUID = "tinvnum" Or (pVal.ItemUID = "20" And pVal.ColUID = "10A" Or pVal.ColUID = "8" Or pVal.ColUID = "9" Or pVal.ColUID = "7C") Then
                            CFL(FormUID, pVal)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "20" And pVal.ColUID = "0A" Then
                            Dim Row As Integer = 0
                            Try
                                objSelect = objMatrix.Columns.Item("0A").Cells.Item(pVal.Row).Specific
                                'QCACTLine = objForm.DataSources.DBDataSources.Item("@MIQCACT1")
                                QCACTLine.SetValue("U_Select", pVal.Row, "Y")
                                Row = pVal.Row
                                QCNum = objMatrix.Columns.Item("1B").Cells.Item(Row).Specific.String
                                objForm.Freeze(True)

                                For i As Integer = 1 To objMatrix.VisualRowCount
                                    objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                                    If i <> Row And objMatrix.Columns.Item("1B").Cells.Item(i).Specific.String <> QCNum Then
                                        If objSelect.Checked = True Then
                                            objSelect.Checked = False : objMatrix.SelectRow(i, False, True)
                                        End If
                                    End If
                                Next
                            Catch ex As Exception
                            Finally
                                objForm.Freeze(False)
                            End Try
                        End If
                        If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Action_Success = True Then
                            LoadSeries(objForm.UniqueID)
                            objForm.Items.Item("21").Enabled = True
                            objForm.Items.Item("22").Enabled = True
                        End If
                End Select
            End If
        Catch ex As Exception
            'objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Sub DeleteRow(ByVal FormUID As String)
        Try
            Dim Flag As Boolean = False
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            'objForm.Freeze(True)
            objMatrix = objForm.Items.Item("20").Specific
            'objMatrix.FlushToDataSource()
            objMatrix.Columns.Item("0A").TitleObject.Click(SAPbouiCOM.BoCellClickType.ct_Double)
            For i As Integer = objMatrix.VisualRowCount To 1 Step -1
                objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                If objSelect.Checked = False Then
                    objMatrix.DeleteRow(i)
                    QCACTLine.RemoveRecord(i)
                    Flag = True

                End If
            Next
            If Flag = True Then
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                    If objSelect.Checked = True Then
                        objMatrix.Columns.Item("0").Cells.Item(i).Specific.String = i
                    End If
                Next
                objForm.Freeze(False)
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                'objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
            End If
        Catch ex As Exception
            objForm.Freeze(False)
            'objAddOn.objApplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
            If BusinessObjectInfo.BeforeAction = True Then
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        objForm.EnableMenu("1282", True)
                End Select
            Else
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        Try
                            'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            Dim InvFlag As Boolean = False
                            Dim ActualEntry As String = "", ActualEntry1 As String = ""

                            DeleteRow(BusinessObjectInfo.FormUID)
                            If objMatrix.VisualRowCount > 0 Then
                                For j = 1 To objMatrix.VisualRowCount
                                    objSelect = objMatrix.Columns.Item("0A").Cells.Item(j).Specific
                                    If objSelect.Checked = True Then
                                        If objMatrix.Columns.Item("9").Cells.Item(j).Specific.String = "Inventory Transfer" Then
                                            If objAddOn.HANA Then
                                                ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T1.""DocEntry"" from ODRF T0 inner join OWTR T1 on T1.""ObjType""=T0.""ObjType"" and T0.""DocEntry""=T1.""draftKey"" where ifnull(T0.""DocStatus"",'')='C' and T1.""draftKey""='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            Else
                                                ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T1.DocEntry from ODRF T0 inner join OWTR T1 on T1.ObjType=T0.ObjType and T0.DocEntry=T1.draftKey where isnull(T0.DocStatus,'')='C' and T1.draftKey='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            End If
                                            If ActualEntry <> "" Then
                                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.String = ActualEntry
                                                InvFlag = True
                                            End If
                                        ElseIf objMatrix.Columns.Item("9").Cells.Item(j).Specific.String = "Goods Issue & A/P Credit Memo" Then

                                            If objAddOn.HANA Then
                                                ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T1.""DocEntry"" from ODRF T0 inner join OIGE T1 on T1.""ObjType""=T0.""ObjType"" and T0.""DocEntry""=T1.""draftKey"" where ifnull(T0.""DocStatus"",'')='C' and T1.""draftKey""='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            Else
                                                ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T1.DocEntry from ODRF T0 inner join OIGE T1 on T1.ObjType=T0.ObjType and T0.DocEntry=T1.draftKey where isnull(T0.DocStatus,'')='C' and T1.draftKey='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            End If
                                            If objAddOn.HANA Then
                                                ActualEntry1 = objAddOn.objGenFunc.getSingleValue("Select T1.""DocEntry"" from ODRF T0 inner join ORPC T1 on T1.""ObjType""=T0.""ObjType"" and T0.""DocEntry""=T1.""draftKey"" where ifnull(T0.""DocStatus"",'')='C' and T1.""draftKey""='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            Else
                                                ActualEntry1 = objAddOn.objGenFunc.getSingleValue("Select T1.DocEntry from ODRF T0 inner join ORPC T1 on T1.ObjType=T0.ObjType and T0.DocEntry=T1.draftKey where isnull(T0.DocStatus,'')='C' and T1.draftKey='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            End If
                                            If ActualEntry <> "" Or ActualEntry1 <> "" Then
                                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.String = ActualEntry
                                                objMatrix.Columns.Item("11A").Cells.Item(j).Specific.String = ActualEntry1
                                                InvFlag = True
                                            End If
                                        ElseIf objMatrix.Columns.Item("9").Cells.Item(j).Specific.String = "A/P Credit Memo" Then
                                            If objAddOn.HANA Then
                                                ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T1.""DocEntry"" from ODRF T0 inner join ORPC T1 on T1.""ObjType""=T0.""ObjType"" and T0.""DocEntry""=T1.""draftKey"" where ifnull(T0.""DocStatus"",'')='C' and T1.""draftKey""='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            Else
                                                ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T1.DocEntry from ODRF T0 inner join ORPC T1 on T1.ObjType=T0.ObjType and T0.DocEntry=T1.draftKey where isnull(T0.DocStatus,'')='C' and T1.draftKey='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            End If
                                            If ActualEntry <> "" Then
                                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.String = ActualEntry
                                                InvFlag = True
                                            End If
                                        ElseIf objMatrix.Columns.Item("9").Cells.Item(j).Specific.String = "Inventory Transfer & A/P Credit Memo" Then

                                            If objAddOn.HANA Then
                                                ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T1.""DocEntry"" from ODRF T0 inner join OWTR T1 on T1.""ObjType""=T0.""ObjType"" and T0.""DocEntry""=T1.""draftKey"" where ifnull(T0.""DocStatus"",'')='C' and T1.""draftKey""='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            Else
                                                ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T1.DocEntry from ODRF T0 inner join OWTR T1 on T1.ObjType=T0.ObjType and T0.DocEntry=T1.draftKey where isnull(T0.DocStatus,'')='C' and T1.draftKey='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            End If

                                            If objAddOn.HANA Then
                                                ActualEntry1 = objAddOn.objGenFunc.getSingleValue("Select T1.""DocEntry"" from ODRF T0 inner join ORPC T1 on T1.""ObjType""=T0.""ObjType"" and T0.""DocEntry""=T1.""draftKey"" where ifnull(T0.""DocStatus"",'')='C' and T1.""draftKey""='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            Else
                                                ActualEntry1 = objAddOn.objGenFunc.getSingleValue("Select T1.DocEntry from ODRF T0 inner join ORPC T1 on T1.ObjType=T0.ObjType and T0.DocEntry=T1.draftKey where isnull(T0.DocStatus,'')='C' and T1.draftKey='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            End If
                                            If ActualEntry <> "" Or ActualEntry1 <> "" Then
                                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.String = ActualEntry
                                                objMatrix.Columns.Item("11A").Cells.Item(j).Specific.String = ActualEntry1
                                                InvFlag = True
                                            End If
                                        ElseIf objMatrix.Columns.Item("9").Cells.Item(j).Specific.String = "Sub-Contracting" Then
                                            'If objAddOn.HANA Then
                                            '    ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T1.""DocEntry"" from ODRF T0 inner join ""@MIPL_OPOR"" T1 on T1.""ObjType""=T0.""ObjType"" and T0.""DocEntry""=T1.""draftKey"" where ifnull(T0.""DocStatus"",'')='C' and T1.""draftKey""='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            'Else
                                            '    ActualEntry = objAddOn.objGenFunc.getSingleValue("Select T1.DocEntry from ODRF T0 inner join [@MIPL_OPOR] T1 on T1.ObjType=T0.ObjType and T0.DocEntry=T1.draftKey where isnull(T0.DocStatus,'')='C' and T1.draftKey='" & objMatrix.Columns.Item("11").Cells.Item(j).Specific.String & "'")
                                            'End If
                                            'If ActualEntry <> "" Then
                                            '    objMatrix.Columns.Item("11").Cells.Item(j).Specific.String = ActualEntry
                                            '    InvFlag = True
                                            'End If
                                        End If
                                    End If
                                Next
                                If InvFlag = True Then
                                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    objForm.Items.Item("1").Click()
                                End If
                            End If
                            objMatrix = objForm.Items.Item("20").Specific
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                            objMatrix.AutoResizeColumns()
                        Catch ex As Exception
                            'objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                End Select
            End If

        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Private Sub AddCFLCondition_Header(ByVal FormUID As String, ByVal CFLID As String, ByVal QueryCol As String, ByVal ColAlias As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item(CFLID)
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim objQCType, objQCStatus As SAPbouiCOM.ComboBox
            Dim StrQuery As String = ""
            Dim BranchID As String = "", LocCode As String = ""
            Dim rsetCFL As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            objQCType = objForm.Items.Item("4").Specific  'QC Type
            objQCStatus = objForm.Items.Item("6").Specific  'QC Status
            If objQCStatus.Selected.Value = "Rej" Then
                If objAddOn.HANA Then
                Else
                    'StrQuery = "Select distinct A." & QueryCol & " from (Select T0.DocEntry,T0.DocNum,Case when T1.U_RejQty>0 then T1.U_RejQty end as QCQty,T0.U_Type as QCType,T1.U_RejWhse,T1.U_ItemCode ,(Select CardName from OCRD where CardName= T0.U_Vendor and CardType='S') U_Vendor "
                    'StrQuery += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry where T0.U_AccStk is not null)A where A.QCQty>0  and A.U_RejWhse<>'' and  A.U_Vendor<>'' and A.QCType='" & objQCType.Selected.Value & "' "

                    StrQuery = "Select distinct A." & QueryCol & " from (Select T0.DocEntry,T0.DocNum,Case when T1.U_RejQty>0 then T1.U_RejQty end as QCQty,T0.U_Type as QCType,T1.U_RejWhse,T1.U_ItemCode ,(Select CardName from OCRD where CardName= T0.U_Vendor and CardType='S') U_Vendor "
                    StrQuery += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry and T0.DocEntry not in (Select U_QCEntry FROM [@MIQCACT1] A where A.U_QCEntry=T0.DocEntry and A.U_QCEntry =case when A.U_TranType like '%&%' then case when A.U_TranEntry1 is not null then A.U_QCEntry else null end else case when A.U_TranEntry is not null then A.U_QCEntry else null end end) and T0.U_AccStk is not null)A "
                    StrQuery += vbCrLf + "where A.QCQty>0  and A.U_RejWhse<>'' and  A.U_Vendor<>'' and A.QCType='" & objQCType.Selected.Value & "' "
                End If
            ElseIf objQCStatus.Selected.Value = "Rew" Then
                If objAddOn.HANA Then
                Else
                    'StrQuery = "Select distinct A." & QueryCol & " from (Select T0.DocEntry,T0.DocNum,Case when T1.U_RewQty>0 then T1.U_RewQty end as QCQty,T0.U_Type as QCType,T1.U_RewWhse ,T1.U_ItemCode,(Select CardName from OCRD where CardName= T0.U_Vendor and CardType='S') U_Vendor "
                    'StrQuery += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry where T0.U_AccStk is not null)A where A.QCQty>0 and A.U_RewWhse<>'' and  A.U_Vendor<>'' and A.QCType='" & objQCType.Selected.Value & "'  "

                    StrQuery = "Select distinct A." & QueryCol & " from (Select T0.DocEntry,T0.DocNum,Case when T1.U_RewQty>0 then T1.U_RewQty end as QCQty,T0.U_Type as QCType,T1.U_RewWhse ,T1.U_ItemCode,(Select CardName from OCRD where CardName= T0.U_Vendor and CardType='S') U_Vendor "
                    StrQuery += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry and T0.DocEntry not in (Select U_QCEntry FROM [@MIQCACT1] A where A.U_QCEntry=T0.DocEntry and A.U_QCEntry =case when A.U_TranType like '%&%' then case when A.U_TranEntry1 is not null then A.U_QCEntry else null end else case when A.U_TranEntry is not null then A.U_QCEntry else null end end) and T0.U_AccStk is not null)A  "
                    StrQuery += vbCrLf + "where A.QCQty>0 And A.U_RewWhse<>'' and  A.U_Vendor<>'' and A.QCType='" & objQCType.Selected.Value & "'  "
                End If
            End If
            If StrQuery <> "" Then rsetCFL.DoQuery(StrQuery)
            If rsetCFL.RecordCount > 0 Then
                For i As Integer = 0 To rsetCFL.RecordCount - 1
                    If i = (rsetCFL.RecordCount - 1) Then
                        oCond = oConds.Add()
                        oCond.Alias = ColAlias ' "DocEntry"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                    Else
                        oCond = oConds.Add()
                        oCond.Alias = ColAlias '"DocEntry"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If
                    rsetCFL.MoveNext()
                Next
            Else
                oCond = oConds.Add()
                oCond.Alias = ColAlias ' "DocEntry"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = ""
            End If
            oCFL.SetConditions(oConds)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub AddCFLCondition(ByVal FormUID As String, ByVal Row As Integer)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_UA")
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim rsetCFL As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            oCond = oConds.Add
            oCond.Alias = "U_QCType"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = objMatrix.Columns.Item("1A").Cells.Item(Row).Specific.String
            If objMatrix.Columns.Item("7A").Cells.Item(Row).Specific.String <> "" Then 'And (objMatrix.Columns.Item("1A").Cells.Item(Row).Specific.String <> "G" And objMatrix.Columns.Item("7").Cells.Item(Row).Specific.String <> "Rej")
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCond = oConds.Add
                oCond.Alias = "U_ProcessBy"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = objMatrix.Columns.Item("7A").Cells.Item(Row).Specific.String
            End If
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCond = oConds.Add
            oCond.Alias = "U_QCStat"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = objMatrix.Columns.Item("7").Cells.Item(Row).Specific.String ' objForm.Items.Item("6").Specific.Selected.value
            oCFL.SetConditions(oConds)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub AddCFLCondition_BP(ByVal FormUID As String, ByVal Row As Integer)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_BP1")
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim rsetCFL As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            If objMatrix.Columns.Item("9").Cells.Item(Row).Specific.String = "A/R Invoice" Then
                oCond = oConds.Add
                oCond.Alias = "CardType"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "C"
            End If
            oCFL.SetConditions(oConds)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CFLCondition(ByVal FormUID As String, ByVal CFLID As String, ByVal ColAlias As String, ByVal ColValue As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item(CFLID)
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim rsetCFL As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            oCond = oConds.Add
            oCond.Alias = ColAlias ' "Locked"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCond.CondVal = ColValue '"N"
            oCFL.SetConditions(oConds)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub AddCFLCondition_Tran(ByVal FormUID As String, ByVal Row As Integer)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_PD")
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim rsetCFL As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            oCond = oConds.Add
            oCond.Alias = "U_QCType"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = objMatrix.Columns.Item("1A").Cells.Item(Row).Specific.String
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCond = oConds.Add
            oCond.Alias = "U_QCStat"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = Trim(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.String)
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCond = oConds.Add
            oCond.Alias = "U_QCAction"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = Trim(objMatrix.Columns.Item("8").Cells.Item(Row).Specific.String)
            If objMatrix.Columns.Item("7A").Cells.Item(Row).Specific.String <> "" Then ' And (objMatrix.Columns.Item("1A").Cells.Item(Row).Specific.String <> "G" And objMatrix.Columns.Item("7").Cells.Item(Row).Specific.String <> "Rej")
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCond = oConds.Add
                oCond.Alias = "U_ProcessBy"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = objMatrix.Columns.Item("7A").Cells.Item(Row).Specific.String
            End If
            oCFL.SetConditions(oConds)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub setCFLCond(ByVal FormUID As String, ByVal CFLId As String, ByVal Row As Integer)
        Try
            Dim objCFL As SAPbouiCOM.ChooseFromList
            Dim objCondition As SAPbouiCOM.Condition
            Dim objConditions As SAPbouiCOM.Conditions
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objCFL = objForm.ChooseFromLists.Item(CFLId)
            For i As Integer = 0 To objCFL.GetConditions.Count - 1
                objCFL.SetConditions(Nothing)
            Next
            objConditions = objCFL.GetConditions()
            'Dim Location As String = ""
            'If objAddOn.HANA Then
            '    Location = objAddOn.objGenFunc.getSingleValue("select ""Location"" from OWHS where ""WhsCode""='" & objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string & "'")
            'Else
            '    Location = objAddOn.objGenFunc.getSingleValue("select Location from OWHS where WhsCode='" & objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string & "'")
            'End If
            objCondition = objConditions.Add()
            objCondition.Alias = "Inactive"
            objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            objCondition.CondVal = "Y"
            'If Location <> "" Then
            '    objCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            '    objCondition = objConditions.Add()
            '    objCondition.Alias = "Location"
            '    objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '    objCondition.CondVal = Location
            '    objCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            '    objCondition = objConditions.Add()
            '    objCondition.Alias = "WhsCode"
            '    objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            '    objCondition.CondVal = objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string
            'End If
            objCFL.SetConditions(objConditions)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ClearSelections(ByVal FormUID As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            objMatrix = objForm.Items.Item("20").Specific
            objMatrix.Clear()
            objForm.Items.Item("12").Specific.string = "" 'Warehouse
            objForm.Items.Item("17").Specific.string = "" 'QC Number
            objForm.Items.Item("8").Specific.string = "" 'Itemcode
            objForm.Items.Item("10").Specific.string = "" 'Vendor Code
            objForm.Items.Item("tcc").Specific.string = "" 'Cost Center
            objForm.Items.Item("tqcdat").Specific.string = "" 'QC Date
            objForm.Items.Item("tinvnum").Specific.string = "" 'Inv Transfer
            objForm.Items.Item("4").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index) 'QC Type
            objForm.Items.Item("6").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index) 'QC Status
            objForm.Items.Item("21").Enabled = True
            objForm.Items.Item("22").Enabled = True
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
        End Try
    End Sub

    Public Sub LoadSeries(ByVal FormUID As String)
        Try
            Dim j As Integer = 0
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            QCACTHeader = objForm.DataSources.DBDataSources.Item("@MIQCACT")
            QCACTLine = objForm.DataSources.DBDataSources.Item("@MIQCACT1")
            objForm.Items.Item("14").Specific.validvalues.loadseries(Formtype, SAPbouiCOM.BoSeriesMode.sf_Add)
            objForm.Items.Item("14").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            QCACTHeader.SetValue("DocNum", 0, objAddOn.objGenFunc.GetDocNum(Formtype, CInt(objForm.Items.Item("14").Specific.Selected.value)))
            objForm.Items.Item("19").Specific.String = Now.Date.ToString("dd/MM/yy") ' A" 
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CFL(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Try
            'Dim GetVal As String = ""
            Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
            Dim objDataTable As SAPbouiCOM.DataTable
            objCFLEvent = pval
            objDataTable = objCFLEvent.SelectedObjects
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("20").Specific
            Select Case objCFLEvent.ChooseFromListUID
                Case "CFL_BP"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objForm.Items.Item("10").Specific.string = objDataTable.GetValue("CardCode", 0)
                        End If
                    Catch ex As Exception
                        objForm.Items.Item("10").Specific.string = objDataTable.GetValue("CardCode", 0)
                    End Try
                Case "CFL_I"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objForm.Items.Item("8").Specific.string = objDataTable.GetValue("ItemCode", 0)
                        End If
                    Catch ex As Exception
                        objForm.Items.Item("8").Specific.string = objDataTable.GetValue("ItemCode", 0)
                    End Try
                Case "CFL_CC"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objForm.Items.Item("tcc").Specific.string = objDataTable.GetValue("PrcCode", 0)
                        End If
                    Catch ex As Exception
                        objForm.Items.Item("tcc").Specific.string = objDataTable.GetValue("PrcCode", 0)
                    End Try
                Case "CFL_Transfer"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objForm.Items.Item("tinvnum").Specific.string = objDataTable.GetValue("DocEntry", 0)
                        End If
                    Catch ex As Exception
                        objForm.Items.Item("tinvnum").Specific.string = objDataTable.GetValue("DocEntry", 0)
                    End Try
                Case "CFL_QC"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objForm.Items.Item("17").Specific.string = objDataTable.GetValue("DocEntry", 0)
                        End If
                    Catch ex As Exception
                        objForm.Items.Item("17").Specific.string = objDataTable.GetValue("DocEntry", 0)
                    End Try
                Case "CFL_W"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objForm.Items.Item("12").Specific.string = objDataTable.GetValue("WhsCode", 0)
                        End If
                    Catch ex As Exception
                        objForm.Items.Item("12").Specific.string = objDataTable.GetValue("WhsCode", 0)
                    End Try
                Case "CFL_TW"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("7C").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)
                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("7C").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)
                    End Try
                Case "CFL_BP1"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("10").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("CardCode", 0)
                            objMatrix.Columns.Item("10A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("CardName", 0)

                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("10").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("CardCode", 0)
                        objMatrix.Columns.Item("10A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("CardName", 0)
                    End Try
                    If objMatrix.Columns.Item("10").Cells.Item(pval.Row).Specific.string <> "" And objMatrix.Columns.Item("9").Cells.Item(pval.Row).Specific.String <> "Sub-Contracting" Then
                        objMatrix.Columns.Item("7C").Cells.Item(pval.Row).Specific.String = objAddOn.objGenFunc.getSingleValue("select U_WAREHOUSE from OCRD where CardCode='" & objMatrix.Columns.Item("10").Cells.Item(pval.Row).Specific.string & "' ")
                    End If
                Case "CFL_UA"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            'If objAddOn.HANA Then

                            'Else
                            '    GetVal = objAddOn.objGenFunc.getSingleValue("select T1.Descr from CUFD T0 inner join UFD1 T1 on T0.TableID=T1.TableID and T0.FieldID=T1.FieldID where T0.TableID='@QCACTION' and T0.AliasID='QCAction' and T1.FldValue='" & objDataTable.GetValue("U_QCAction", 0) & "' ")
                            '    GetVal = ""
                            'End If
                            objMatrix.Columns.Item("8").Cells.Item(pval.Row).Specific.String = Trim(objDataTable.GetValue("U_QCAction", 0))
                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("8").Cells.Item(pval.Row).Specific.String = Trim(objDataTable.GetValue("U_QCAction", 0))
                    End Try
                    objMatrix.Columns.Item("9").Cells.Item(pval.Row).Specific.String = objAddOn.objGenFunc.getSingleValue(" select U_TType from [@QCACTION] where  U_QCType='" & objMatrix.Columns.Item("1A").Cells.Item(pval.Row).Specific.string & "' and U_QCStat='" & objMatrix.Columns.Item("7").Cells.Item(pval.Row).Specific.string & "' and U_QCAction='" & objMatrix.Columns.Item("8").Cells.Item(pval.Row).Specific.string & "' and U_ProcessBy='" & objMatrix.Columns.Item("7A").Cells.Item(pval.Row).Specific.string & "' ")

                Case "CFL_PD"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            'Dim vv As String = Trim(objDataTable.GetValue("U_TType", 0))
                            objMatrix.Columns.Item("9").Cells.Item(pval.Row).Specific.String = Trim(objDataTable.GetValue("U_TType", 0))
                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("9").Cells.Item(pval.Row).Specific.String = Trim(objDataTable.GetValue("U_TType", 0))
                        'objMatrix.Columns.Item("9").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("U_TType", 0)
                    End Try

            End Select
            objForm.Freeze(True)
            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            'MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace, MsgBoxStyle.Critical)
        End Try

    End Sub

    Public Sub LoadQCDetails(ByVal FormUID As String)
        Try
            Dim j As Integer = 0
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim objQCType, objQCStatus As SAPbouiCOM.ComboBox
            objQCType = objForm.Items.Item("4").Specific
            objQCStatus = objForm.Items.Item("6").Specific
            If objAddOn.HANA Then
            Else
                strSQL = "Select  (select Top 1 U_ToWhse from [@WHSE] where U_ObjType='MIQCACT' and U_QcResult like A.QCType + '%' and U_Division=A.OcrCode and U_QcType=A.U_Type and U_FromWhse=A.WhsCode and U_ProceesedBy=A.BPProcess) as ToWhse,"
                strSQL += vbCrLf + "Case when A.QCType='Rej' and A.U_Type='R' and A.TranType='A/R Invoice' then '' else A.CardCode end as BPCode,case when A.QCType='Rej' and A.U_Type='R' and A.TranType='A/R Invoice' then '' else A.U_Vendor end as BPName,* from (Select T0.U_QCAReq QCAReq, T0.U_Vendor,T0.U_Type,T0.DocEntry,T0.DocNum,Format(T0.U_DocDate,'yyyyMMdd') as U_DocDate,T1.U_ItemCode,T1.U_ItemName,"
                strSQL += vbCrLf + "(Select Top 1 case when T2.U_TranType like '%&%' then case when T2.U_TranEntry1 is not null then T2.U_QCEntry else null end else case when T2.U_TranEntry is not null then T2.U_QCEntry else null end end FROM [@MIQCACT1] T2 INNER JOIN [@MIQCACT] T3 ON T3.DocEntry=T2.DocEntry where T2.U_QCEntry=T1.DocEntry and T2.U_BaseLine=T1.U_BaseLinNum and T2.U_Select='Y' and T2.U_QCStat='Rej' order by T3.DocEntry Desc) as QCACT,T1.U_BaseLinNum,"
                strSQL += vbCrLf + "(Select Top 1 CardCode from OCRD where CardName=T0.U_Vendor and CardType='S') as CardCode,(Select Top 1 U_Process from OCRD where CardName=T0.U_Vendor) as BPProcess,"
                strSQL += vbCrLf + "(select T2.DocEntry from ODRF T4 left join OWTR T2 on T4.DocEntry=T2.draftKey join DRF1 T3 on T4.DocEntry=T3.DocEntry and T2.WddStatus='P' and T2.draftkey=T0.U_AccStk and T3.ItemCode=T1.U_ItemCode and T3.U_BaseLine=T1.U_BaseLinNum and T2.U_QCNum =T0.DocNum and T3.WhsCode=T1.U_RejWhse and T3.Quantity=T1.U_RejQty "
                strSQL += vbCrLf + "union all select T2.DocEntry from OWTR T2 join WTR1 T3 on T2.DocEntry=T3.DocEntry and T2.DocNum=T0.U_AccStk and T3.ItemCode=T1.U_ItemCode and T3.U_BaseLine=T1.U_BaseLinNum and T2.U_QCNum =T0.DocNum and T3.WhsCode=T1.U_RejWhse and T3.Quantity=T1.U_RejQty"
                strSQL += vbCrLf + ") as InvNum," 'T2.DocNum=T0.U_AccStk and T2.DocEntry=T0.U_AccStkD 
                strSQL += vbCrLf + "Case when T1.U_RejQty>0 then T1.U_RejQty end as QCQty,'Rej' as QCType,T1.U_RejWhse as WhsCode,T1.U_OcrCode,"
                strSQL += vbCrLf + "Case when T0.U_Type='G' then (Select T3.OcrCode from OPDN T2 join PDN1 T3 ON T3.DocEntry=T2.DocEntry where T3.ItemCode=T1.U_ItemCode and T3.LineNum=T1.U_BaseLinNum and T2.DocEntry=T0.U_GRNEntry and  T3.OcrCode<>'') "
                strSQL += vbCrLf + "else (Select T3.OcrCode from OIGN T2 join IGN1 T3 ON T3.DocEntry=T2.DocEntry where T3.ItemCode=T1.U_ItemCode and T3.LineNum=T1.U_BaseLinNum and T2.DocEntry=T0.U_GRNum and  T3.OcrCode<>'') end as OcrCode,"

                strSQL += vbCrLf + "(Select case when B.Stat=1 then (Select U_QCAction as Stat from [@QCACTION] where U_QCStat='Rej' and U_QCType=T0.U_Type and U_ProcessBy=(Select Top 1 U_Process from OCRD where CardName=T0.U_Vendor)) else '' end as QCAction from (Select count(*) as Stat from [@QCACTION] where U_QCStat='Rej' and U_QCType=T0.U_Type and U_ProcessBy=(Select Top 1 U_Process from OCRD where CardName=T0.U_Vendor))B) as QCAction,"
                strSQL += vbCrLf + "(Select case when B.Stat=1 then (Select U_TType as Stat from [@QCACTION] where U_QCStat='Rej' and U_QCType=T0.U_Type and U_ProcessBy=(Select Top 1 U_Process from OCRD where CardName=T0.U_Vendor)) else '' end as QCAction from (Select count(*) as Stat from [@QCACTION] where U_QCStat='Rej' and U_QCType=T0.U_Type and U_ProcessBy=(Select Top 1 U_Process from OCRD where CardName=T0.U_Vendor))B) as TranType"
                strSQL += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry"
                strSQL += vbCrLf + "Union all"
                strSQL += vbCrLf + "Select T0.U_QCAReq,T0.U_Vendor,T0.U_Type,T0.DocEntry,T0.DocNum,Format(T0.U_DocDate,'yyyyMMdd') as U_DocDate,T1.U_ItemCode,T1.U_ItemName,"
                strSQL += vbCrLf + "(Select Top 1 case when T2.U_TranType like '%&%' then case when T2.U_TranEntry1 is not null then T2.U_QCEntry else null end else case when T2.U_TranEntry is not null then T2.U_QCEntry else null end end FROM [@MIQCACT1] T2 INNER JOIN [@MIQCACT] T3 ON T3.DocEntry=T2.DocEntry where T2.U_QCEntry=T1.DocEntry and T2.U_BaseLine=T1.U_BaseLinNum and T2.U_Select='Y' and T2.U_QCStat='Rew' order by T3.DocEntry Desc) as QCACT,T1.U_BaseLinNum,"
                strSQL += vbCrLf + "(Select Top 1 CardCode from OCRD where CardName=T0.U_Vendor and CardType='S') as CardCode,(Select Top 1 U_Process from OCRD where CardName=T0.U_Vendor) as BPProcess,"
                strSQL += vbCrLf + "(select T2.DocEntry from ODRF T4 left join OWTR T2 on T4.DocEntry=T2.draftKey join DRF1 T3 on T4.DocEntry=T3.DocEntry and T2.WddStatus='P' and T2.draftkey=T0.U_AccStk  and T3.ItemCode=T1.U_ItemCode and T3.U_BaseLine=T1.U_BaseLinNum and T2.U_QCNum =T0.DocNum and T3.WhsCode=T1.U_RewWhse and T3.Quantity=T1.U_RewQty"
                strSQL += vbCrLf + "union all select T2.DocEntry from  OWTR T2  join WTR1 T3 on T2.DocEntry=T3.DocEntry and T2.DocNum=T0.U_AccStk and T3.ItemCode=T1.U_ItemCode and T3.U_BaseLine=T1.U_BaseLinNum and T2.U_QCNum =T0.DocNum and T3.WhsCode=T1.U_RewWhse and T3.Quantity=T1.U_RewQty"
                strSQL += vbCrLf + ") as InvNum," 'T2.DocNum=T0.U_AccStk and T2.DocEntry=T0.U_AccStkD
                strSQL += vbCrLf + "case when T1.U_RewQty>0 then T1.U_RewQty end as QCQty,'Rew' as QCType,T1.U_RewWhse as WhsCode,T1.U_OcrCode,"
                strSQL += vbCrLf + "Case when T0.U_Type='G' then (Select T3.OcrCode from OPDN T2 join PDN1 T3 ON T3.DocEntry=T2.DocEntry where T3.ItemCode=T1.U_ItemCode and T3.LineNum=T1.U_BaseLinNum and T2.DocEntry=T0.U_GRNEntry and  T3.OcrCode<>'') "
                strSQL += vbCrLf + "else (Select T3.OcrCode from OIGN T2 join IGN1 T3 ON T3.DocEntry=T2.DocEntry where T3.ItemCode=T1.U_ItemCode and T3.LineNum=T1.U_BaseLinNum and T2.DocEntry=T0.U_GRNum and  T3.OcrCode<>'') end as OcrCode,"
                strSQL += vbCrLf + "(Select case when B.Stat=1 then (Select U_QCAction as Stat from [@QCACTION] where U_QCStat='Rew' and U_QCType=T0.U_Type and U_ProcessBy=(Select Top 1 U_Process from OCRD where CardName=T0.U_Vendor)) else '' end as QCAction from (Select count(*) as Stat from [@QCACTION] where U_QCStat='Rew' and U_QCType=T0.U_Type and U_ProcessBy=(Select Top 1 U_Process from OCRD where CardName=T0.U_Vendor))B) as QCAction,"
                strSQL += vbCrLf + "(Select case when B.Stat=1 then (Select U_TType as Stat from [@QCACTION] where U_QCStat='Rew' and U_QCType=T0.U_Type and U_ProcessBy=(Select Top 1 U_Process from OCRD where CardName=T0.U_Vendor)) else '' end as QCAction from (Select count(*) as Stat from [@QCACTION] where U_QCStat='Rew' and U_QCType=T0.U_Type and U_ProcessBy=(Select Top 1 U_Process from OCRD where CardName=T0.U_Vendor))B) as TranType"
                strSQL += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry) A where A.QCACT is null and A.QCQty>0 and A.WhsCode<>'' and A.U_Type='" & objQCType.Selected.Value & "' and A.QCType='" & objQCStatus.Selected.Value & "' and A.InvNum<>'' and isnull(A.QCAReq,'')='Y'"
                If objForm.Items.Item("17").Specific.string <> "" Then 'QC Number
                    strSQL += vbCrLf + "and A.DocEntry='" & objForm.Items.Item("17").Specific.string & "'"
                End If
                If objForm.Items.Item("12").Specific.string <> "" Then 'Warehouse
                    strSQL += vbCrLf + "and A.WhsCode='" & objForm.Items.Item("12").Specific.string & "'"
                End If
                If objForm.Items.Item("8").Specific.string <> "" Then 'Itemcode
                    strSQL += vbCrLf + "and A.U_ItemCode='" & objForm.Items.Item("8").Specific.string & "'"
                End If
                If objForm.Items.Item("10").Specific.string <> "" Then 'Vendor Code
                    strSQL += vbCrLf + "and A.U_Vendor=(Select CardName from OCRD Where CardCode='" & objForm.Items.Item("10").Specific.string & "')"
                End If
                If objForm.Items.Item("tcc").Specific.string <> "" Then 'Cost Center
                    strSQL += vbCrLf + "and A.OcrCode='" & objForm.Items.Item("tcc").Specific.string & "'"
                End If
                If objForm.Items.Item("tinvnum").Specific.string <> "" Then 'Inv Transfer Entry
                    strSQL += vbCrLf + "and A.InvNum='" & objForm.Items.Item("tinvnum").Specific.string & "'"
                End If
                If objForm.Items.Item("tqcdat").Specific.string <> "" Then 'QC Date
                    Dim objedit As SAPbouiCOM.EditText
                    objedit = objForm.Items.Item("tqcdat").Specific
                    Dim DocDate As Date = Date.ParseExact(objedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    strSQL += vbCrLf + "and A.U_DocDate='" & DocDate.ToString("yyyyMMdd") & "'"
                End If
            End If
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecordSet.DoQuery(strSQL)
            QCACTLine = objForm.DataSources.DBDataSources.Item("@MIQCACT1")
            Dim Row As Integer = 0
            If objRecordSet.RecordCount > 0 Then
                objAddOn.objApplication.StatusBar.SetText("Loading Details.Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objForm.Freeze(True)
                objMatrix.Clear()
                For Rec As Integer = 0 To objRecordSet.RecordCount - 1
                    If objMatrix.RowCount = 0 Then
                        objMatrix.AddRow()
                    ElseIf objMatrix.Columns.Item("3").Cells.Item(objMatrix.RowCount).Specific.String <> "" Then
                        objMatrix.AddRow()
                    End If
                    QCACTLine.Clear()
                    'If Validate_Batch_Serial(FormUID, Trim(objRecordSet.Fields.Item("U_ItemCode").Value), Trim(objRecordSet.Fields.Item("WhsCode").Value), Trim(objRecordSet.Fields.Item("InvNum").Value), CDbl(objRecordSet.Fields.Item("QCQty").Value)) Then Continue For
                    Row += 1
                    objMatrix.GetLineData(objMatrix.RowCount)
                    QCACTLine.SetValue("LineId", 0, Row)  'objRecordSet.Fields.Item("LineId").Value
                    QCACTLine.SetValue("U_BaseLine", 0, objRecordSet.Fields.Item("U_BaseLinNum").Value)
                    QCACTLine.SetValue("U_QCNum", 0, objRecordSet.Fields.Item("DocNum").Value)
                    QCACTLine.SetValue("U_QCEntry", 0, objRecordSet.Fields.Item("DocEntry").Value)
                    QCACTLine.SetValue("U_QCDate", 0, objRecordSet.Fields.Item("U_DocDate").Value)
                    QCACTLine.SetValue("U_ItemCode", 0, objRecordSet.Fields.Item("U_ItemCode").Value)
                    QCACTLine.SetValue("U_ItemName", 0, objRecordSet.Fields.Item("U_ItemName").Value)
                    QCACTLine.SetValue("U_WhsCode", 0, objRecordSet.Fields.Item("WhsCode").Value)
                    QCACTLine.SetValue("U_QCQty", 0, objRecordSet.Fields.Item("QCQty").Value)
                    QCACTLine.SetValue("U_QCStat", 0, objRecordSet.Fields.Item("QCType").Value)
                    QCACTLine.SetValue("U_QCType", 0, objRecordSet.Fields.Item("U_Type").Value)
                    QCACTLine.SetValue("U_Process", 0, objRecordSet.Fields.Item("BPProcess").Value)
                    QCACTLine.SetValue("U_BPCode", 0, objRecordSet.Fields.Item("BPCode").Value)
                    QCACTLine.SetValue("U_InvEntry", 0, objRecordSet.Fields.Item("InvNum").Value)
                    QCACTLine.SetValue("U_BPName", 0, objRecordSet.Fields.Item("BPName").Value)
                    QCACTLine.SetValue("U_UserAct", 0, objRecordSet.Fields.Item("QCAction").Value)
                    QCACTLine.SetValue("U_TranType", 0, objRecordSet.Fields.Item("TranType").Value)
                    QCACTLine.SetValue("U_VCode", 0, objRecordSet.Fields.Item("BPCode").Value)
                    QCACTLine.SetValue("U_VName", 0, objRecordSet.Fields.Item("BPName").Value)
                    QCACTLine.SetValue("U_OcrCode", 0, objRecordSet.Fields.Item("OcrCode").Value)
                    QCACTLine.SetValue("U_ToWhse", 0, objRecordSet.Fields.Item("ToWhse").Value)
                    QCACTLine.SetValue("U_Select", 0, "N")
                    objMatrix.SetLineData(objMatrix.RowCount)
                    objRecordSet.MoveNext()
                Next
                objMatrix.AutoResizeColumns()
                If objMatrix.Columns.Item("0A").Editable = False Then
                    objMatrix.Columns.Item("0A").Editable = True
                End If
                If objQCType.Selected.Value = "R" And objQCStatus.Selected.Value = "Rew" Then
                    objMatrix.Columns.Item("10A").Editable = True
                Else
                    objMatrix.Columns.Item("10A").Editable = False
                End If
                'For ii As Integer = 1 To objMatrix.VisualRowCount
                '    objMatrix.CommonSetting.MergeCell(ii, 23, True)
                '    objMatrix.CommonSetting.MergeCell(ii, 22, True)
                'Next
                objForm.Freeze(False)
                objAddOn.objApplication.StatusBar.SetText("Successfully loaded entries...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objForm.Freeze(True)
                objMatrix.Clear()
                objForm.Freeze(False)
                objAddOn.objApplication.StatusBar.SetText("No Records found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
        Catch ex As Exception
            objForm.Freeze(False)
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Function Validate_Batch_Serial(ByVal FormUID As String, ByVal ItemCode As String, ByVal WhsCode As String, ByVal DocEntry As String, ByVal Qty As Double) As Boolean
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim ErrCount As Integer = 0
            Dim Batch As String = ""

            If objAddOn.HANA Then
                'Serial = objAddOn.objGenFunc.getSingleValue("select ""ManSerNum"" from OITM WHERE ""ItemCode""='" & ItemCode & "'")
                Batch = objAddOn.objGenFunc.getSingleValue("select ""ManBtchNum"" from OITM WHERE ""ItemCode""='" & ItemCode & "'")
            Else
                'Serial = objAddOn.objGenFunc.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & ItemCode & "'")
                Batch = objAddOn.objGenFunc.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & ItemCode & "'")
            End If
            If Batch = "Y" Then
                If objAddOn.HANA Then
                    Batch = objAddOn.objGenFunc.getSingleValue("Select 1 as ""Status"" from ibt1 T inner join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"" " &
                                           " where T.""BaseType""='67' and T.""Direction""=0 and T.""BaseEntry""='" & DocEntry & "' and T.""ItemCode""='" & ItemCode & "'  and T.""WhsCode""='" & WhsCode & "'  having Sum(T.""Quantity"")>=" & Qty & " ")

                Else
                    Batch = objAddOn.objGenFunc.getSingleValue("Select 1 as Status from ibt1 T inner join oibt T1 on T.ItemCode=T1.ItemCode and T.BatchNum=T1.BatchNum and T.WhsCode=T1.WhsCode " &
                                           " where T.BaseType='67' and T.Direction=0 and T.BaseEntry='" & DocEntry & "' and T.ItemCode='" & ItemCode & "'  and T.WhsCode='" & WhsCode & "'  having Sum(T.Quantity)>=" & Qty & " ")
                End If
            End If
            If Batch <> "" Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End Try
    End Function

    Private Sub Create_AR_Invoice(ByVal FormUID As String)
        Try
            Dim objARMatrix As SAPbouiCOM.Matrix
            Dim objARform As SAPbouiCOM.Form
            Dim objrecset As SAPbobsCOM.Recordset
            Dim Lineflag As Boolean = False
            Dim WhsCode As String = ""
            Dim Row As Integer = 1
            objMatrix = objForm.Items.Item("20").Specific
            objrecset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            For i As Integer = 1 To objMatrix.VisualRowCount
                objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                If objSelect.Checked = True Then
                    Lineflag = True
                    Row = i
                    Exit For
                End If
            Next
            If Lineflag = True Then
                objAddOn.objApplication.Menus.Item("2053").Activate()
                objARform = objAddOn.objApplication.Forms.ActiveForm
                objARform = objAddOn.objApplication.Forms.Item(objARform.UniqueID)
                objARform.Visible = True
                objARMatrix = objARform.Items.Item("38").Specific
                objAddOn.objApplication.StatusBar.SetText("Data Loading to A/R Invoice Screen Please wait ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Dim AcctCode As String = ""
                Try
                    objForm.Freeze(True)
                    objARform.Freeze(True)
                    objARform.Items.Item("4").Specific.String = objMatrix.Columns.Item("10").Cells.Item(Row).Specific.string
                    objARform.Items.Item("t_qcanum").Specific.String = objForm.Items.Item("15").Specific.String 'objMatrix.Columns.Item("1B").Cells.Item(Row).Specific.string
                    If objARMatrix.Columns.Item("U_Item").Editable = False Or objARMatrix.Columns.Item("U_Whse").Editable = False Or objARMatrix.Columns.Item("U_Qty").Editable = False Then
                        objARMatrix.Columns.Item("U_Item").Editable = True
                        objARMatrix.Columns.Item("U_Whse").Editable = True
                        objARMatrix.Columns.Item("U_Qty").Editable = True
                    End If
                    Row = 1
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                        If objSelect.Checked = True Then
                            objARMatrix.Columns.Item("1").Cells.Item(Row).Specific.String = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string
                            objARMatrix.Columns.Item("11").Cells.Item(Row).Specific.String = objMatrix.Columns.Item("6").Cells.Item(i).Specific.string
                            objARMatrix.Columns.Item("24").Cells.Item(Row).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.string
                            objARMatrix.Columns.Item("U_Item").Cells.Item(Row).Specific.String = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string
                            objARMatrix.Columns.Item("U_Whse").Cells.Item(Row).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.string
                            objARMatrix.Columns.Item("U_Qty").Cells.Item(Row).Specific.String = objMatrix.Columns.Item("6").Cells.Item(i).Specific.string
                            Row += 1
                        End If
                    Next
                    objARMatrix.Columns.Item("160").Cells.Item(1).Click()
                    objARMatrix.Columns.Item("U_Item").Editable = False
                    objARMatrix.Columns.Item("U_Whse").Editable = False
                    objARMatrix.Columns.Item("U_Qty").Editable = False
                    objrecset = Nothing
                    objAddOn.objApplication.StatusBar.SetText("Data Loaded to A/R Invoice Screen ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Catch ex As Exception
                    objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Finally
                    objForm.Freeze(False)
                    objARform.Freeze(False)
                End Try
            Else
                objAddOn.objApplication.SetStatusBarMessage("No more Data for posting the Goods Receipt ...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Exit Sub
            End If
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Function Create_SubContracting(ByVal FormUID As String) As Boolean
        Try
            Dim objSubPOMatrix As SAPbouiCOM.Matrix
            Dim objSubPOform As SAPbouiCOM.Form
            Dim objrecset As SAPbobsCOM.Recordset
            Dim Lineflag As Boolean = False
            Dim WhsCode As String = ""
            Dim Row As Integer = 1
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("20").Specific
            objrecset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            For i As Integer = 1 To objMatrix.VisualRowCount
                objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                If objSelect.Checked = True And objMatrix.Columns.Item("11").Cells.Item(i).Specific.String = "" Then
                    Lineflag = True
                    Row = i
                    Exit For
                End If
            Next
            If Lineflag = True Then
                objAddOn.objApplication.Menus.Item("SUBCTPO").Activate()
                objSubPOform = objAddOn.objApplication.Forms.ActiveForm
                objSubPOform = objAddOn.objApplication.Forms.Item(objSubPOform.UniqueID)
                objSubPOform.Visible = True
                Try
                    Dim objButton As SAPbouiCOM.StaticText
                    Dim objItem As SAPbouiCOM.Item
                    'objSubPOform = objAddOn.objApplication.Forms.Item(FormUID)
                    objItem = objSubPOform.Items.Add("l_qcanum", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                    objItem.Left = 280 ' objSubPOform.Items.Item("tInvUom").Left '+ objSubPOform.Items.Item("tInvUom").Width + 10 '280
                    objItem.Width = 50
                    objItem.Top = 115 'objSubPOform.Items.Item("tInvUom").Top ' 128 'objSubPOform.Items.Item("tpodoc").Top '+ objSubPOform.Items.Item("tpodoc").Height + 2
                    objItem.Height = 14 'objSubPOform.Items.Item("2").Height
                    objButton = objItem.Specific
                    objButton.Caption = "QCA"

                    Dim objedit As SAPbouiCOM.EditText
                    objItem = objSubPOform.Items.Add("t_qcanum", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                    objItem.Left = objSubPOform.Items.Item("l_qcanum").Left + objSubPOform.Items.Item("l_qcanum").Width + 15
                    objItem.Width = 60 '
                    objItem.Top = objSubPOform.Items.Item("l_qcanum").Top
                    objItem.Height = 14 'objSubPOform.Items.Item("l_qcanum").Height
                    objItem.LinkTo = "l_qcanum"
                    objedit = objItem.Specific
                    objedit.Item.Enabled = False
                    objItem.Enabled = False
                    objedit.DataBind.SetBound(True, "@MIPL_OPOR", "U_QCANum")


                    Dim objlink As SAPbouiCOM.LinkedButton
                    objItem = objSubPOform.Items.Add("lnk_qca", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
                    objItem.Left = objSubPOform.Items.Item("l_qcanum").Left + objSubPOform.Items.Item("l_qcanum").Width + 2
                    objItem.Width = 12
                    objItem.Top = objSubPOform.Items.Item("l_qcanum").Top + 3
                    objItem.Height = 10 'objSubPOform.Items.Item("l_qcanum").Height
                    'objItem.LinkTo = "t_qcanum"
                    objlink = objItem.Specific
                    objlink.LinkedObjectType = "-1"
                    objlink.Item.LinkTo = "t_qcanum"

                    'objAddOn.objApplication.SetStatusBarMessage("Button Created", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Catch ex As Exception
                End Try
                objSubPOMatrix = objSubPOform.Items.Item("MtxinputN").Specific
                objAddOn.objApplication.StatusBar.SetText("Data Loading to Sub-Contracting Screen Please wait ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Dim AcctCode As String = ""
                Try
                    objForm.Freeze(True)
                    objSubPOform.Freeze(True)
                    objSubPOform.Items.Item("txtcode").Specific.String = objMatrix.Columns.Item("10").Cells.Item(Row).Specific.string
                    objSubPOform.Items.Item("t_qcanum").Specific.String = objForm.Items.Item("15").Specific.String 'objMatrix.Columns.Item("1B").Cells.Item(Row).Specific.string
                    objrecset.DoQuery("select U_GREntry DocNum, U_GRNum DocEntry from [@MIPLQC] where DocEntry='" & objMatrix.Columns.Item("1B").Cells.Item(Row).Specific.string & "'")
                    If objrecset.RecordCount > 0 Then
                        objSubPOform.Items.Item("txtremark").Specific.String = "Created from QC Action By " & objAddOn.objCompany.UserName & " on " & Now.ToString & " for the Receipt Num " & objrecset.Fields.Item(0).Value.ToString & " and Receipt Entry " & objrecset.Fields.Item(1).Value.ToString
                    End If
                    Dim BOMItem As String = ""
                    BOMItem = objAddOn.objGenFunc.getSingleValue("Select T0.U_BOMCode from [@MIPL_OPOR] T0 where T0.DocEntry=(Select T2.U_SubConNo from OIGN T2 where T2.DocEntry=(Select T1.U_GRNum from [@MIPLQC] T1 where T1.DocEntry='" & Trim(objMatrix.Columns.Item("1B").Cells.Item(Row).Specific.string) & "'))")
                    objSubPOform.Items.Item("txtbitem").Specific.String = BOMItem ' objMatrix.Columns.Item("3").Cells.Item(Row).Specific.string
                    objSubPOform.Items.Item("SQty").Specific.String = objMatrix.Columns.Item("6").Cells.Item(Row).Specific.string
                    Row = 1
                    For i As Integer = 1 To objSubPOMatrix.VisualRowCount
                        If objSubPOMatrix.Columns.Item("Code").Cells.Item(i).Specific.String <> "" Then
                            objSubPOMatrix.Columns.Item("Whse").Cells.Item(i).Specific.String = objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string
                        End If
                    Next
                    objSubPOform.ActiveItem = "deldate"
                    objSubPOform.Items.Item("txtcode").Enabled = False
                    objSubPOform.Items.Item("txtbitem").Enabled = False
                    objSubPOform.Items.Item("t_qcanum").Enabled = False
                    objSubPOform.Items.Item("SQty").Enabled = False
                    objrecset = Nothing
                    objAddOn.objApplication.StatusBar.SetText("Data Loaded to Sub-Contracting Screen ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    objForm.Items.Item("21").Enabled = False
                    objForm.Items.Item("22").Enabled = False
                    Return True
                Catch ex As Exception
                    objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Finally
                    objForm.Freeze(False)
                    objSubPOform.Freeze(False)
                End Try
            Else
                objAddOn.objApplication.SetStatusBarMessage("No more Data for posting the Sub-Contracting ...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return False
            End If
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Public Function Create_SubContractingDraft(ByVal FormUID As String) As Boolean
        Try
            Dim DocEntry As String, TranDocEntry As String = ""
            Dim objSubPODraft As SAPbobsCOM.Documents
            Dim objrs As SAPbobsCOM.Recordset
            'Dim Quantity As Double
            Dim QCDocNum As Long
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                Dim Lineflag As Boolean = False
                Dim ToWhse As String = ""
                Dim Row As Integer = 1
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                    If objSelect.Checked = True And objMatrix.Columns.Item("11").Cells.Item(i).Specific.String = "" Then
                        Lineflag = True
                        Row = i
                        Exit For
                    End If
                Next
                If Lineflag = True Then
                    objAddOn.objApplication.StatusBar.SetText("Sub-Contracting Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objSubPODraft = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                    If Not objAddOn.objCompany.InTransaction Then objAddOn.objCompany.StartTransaction()
                    Dim DocDate As Date = Date.ParseExact(objForm.Items.Item("19").Specific.string, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    QCDocNum = objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("14").Specific.Selected.value)
                    objSubPODraft.DocDate = DocDate
                    objSubPODraft.JournalMemo = "Auto-Gen-> QC Action " & Now.ToString
                    objSubPODraft.Comments = "QCA DocNum-> " & CStr(QCDocNum) 'QCHeader.GetValue("DocNum", 0) ' & objForm.Items.Item("4").Specific.string
                    objSubPODraft.UserFields.Fields.Item("U_QCANum").Value = CStr(QCDocNum)
                    objSubPODraft.DocObjectCode = CInt("SUBPO")
                    'objMatrix = objForm.Items.Item("20").Specific
                    'objGoodsIssue.CardCode = objMatrix.Columns.Item("10").Cells.Item(Row).Specific.string
                    'If objAddOn.HANA Then
                    '    BranchEnabled = objAddOn.objGenFunc.getSingleValue("select ""MltpBrnchs"" from OADM")
                    'Else
                    '    BranchEnabled = objAddOn.objGenFunc.getSingleValue("select MltpBrnchs from OADM")
                    'End If
                    'If BranchEnabled = "Y" Then
                    '    Branch = objAddOn.objGenFunc.getSingleValue("Select Top 1 BPLid from OWHS where WhsCode='" & objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string & "'")
                    '    objSubPODraft.BPL_IDAssignedToInvoice = Branch
                    'End If

                    If objSubPODraft.Add() <> 0 Then
                        If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        objAddOn.objApplication.SetStatusBarMessage("Sub-Contracting: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        objAddOn.objApplication.MessageBox("Sub-Contracting: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode,, "OK")
                        Return False
                    Else
                        If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        DocEntry = objAddOn.objCompany.GetNewObjectKey()
                        For j = 1 To objMatrix.VisualRowCount
                            objSelect = objMatrix.Columns.Item("0A").Cells.Item(j).Specific
                            If objSelect.Checked = True Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.String = DocEntry
                            End If
                        Next
                        objAddOn.objApplication.StatusBar.SetText("Sub-Contracting Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Return True
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objSubPODraft)
                    GC.Collect()
                Else
                    objAddOn.objApplication.StatusBar.SetText("No more data for Sub-Contracting...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If
            Catch ex As Exception
                'objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                objAddOn.objApplication.MessageBox(ex.Message, , "OK")
                Return False
                'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End Try
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.Message, , "OK")
            Return False
            'objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    Public Sub Create_AP_Invoice_CreditMemo(ByVal FormUID As String)
        Try
            Dim Batch As String, Serial As String, DocEntry As String, BranchEnabled As String, Branch As String
            Dim objPurchaseReturn As SAPbobsCOM.Documents
            Dim objrs As SAPbobsCOM.Recordset
            Dim Quantity As Double
            Dim QCDocNum As Long
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                Dim Lineflag As Boolean = False
                Dim ToWhse As String = ""
                Dim Row As Integer = 1
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                    If objSelect.Checked = True Then
                        Lineflag = True
                        Row = i
                        Exit For
                    End If
                Next
                If Lineflag = True Then
                    If objMatrix.Columns.Item("1A").Cells.Item(Row).Specific.string = "G" Then
                        strSQL = "select distinct T0.DocEntry,T1.TaxCode,T1.Price,T1.LineNum,Format(T0.DocDate,'yyyyMMdd') DocDate,T0.DocNum from OPCH T0 join PCH1 T1 on T0.DocEntry=T1.DocEntry where T0.DocStatus='O' and T1.LineStatus='O' and T1.BaseType='20' " &
                                                                       " and T1.BaseEntry=(select distinct T0.U_GRPONum from OWTR T0 join WTR1 T1 on T0.DocEntry=T1.DocEntry where T0.DocEntry=" & objMatrix.Columns.Item("7B").Cells.Item(Row).Specific.string & ") order by T0.DocEntry desc"
                    Else
                        strSQL = "select distinct T0.DocEntry,T1.TaxCode,T1.Price,T1.LineNum,Format(T0.DocDate,'yyyyMMdd') DocDate,T0.DocNum from OPCH T0 join PCH1 T1 on T0.DocEntry=T1.DocEntry where T0.DocStatus='O' and T1.LineStatus='O' and T1.BaseType='20' " &
                                                                      " and T1.BaseEntry=(select distinct T0.U_REntry from OWTR T0 join WTR1 T1 on T0.DocEntry=T1.DocEntry where T0.DocEntry=" & objMatrix.Columns.Item("7B").Cells.Item(Row).Specific.string & ") order by T0.DocEntry desc"
                    End If
                    objrs.DoQuery(strSQL)
                    If objrs.RecordCount = 0 Then objAddOn.objApplication.StatusBar.SetText("A/P Invoice not found for the specified QC...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub
                    objAddOn.objApplication.StatusBar.SetText("A/P CreditMemo Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objPurchaseReturn = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                    If Not objAddOn.objCompany.InTransaction Then objAddOn.objCompany.StartTransaction()
                    Dim DocDate As Date = Date.ParseExact(objForm.Items.Item("19").Specific.string, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    QCDocNum = objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("14").Specific.Selected.value)
                    objPurchaseReturn.CardCode = objMatrix.Columns.Item("10").Cells.Item(Row).Specific.string
                    objPurchaseReturn.DocDate = DocDate
                    objPurchaseReturn.JournalMemo = "Auto-Gen-> QC Action " & Now.ToString
                    objPurchaseReturn.Comments = "QCA DocNum-> " & CStr(QCDocNum) 'QCHeader.GetValue("DocNum", 0) ' & objForm.Items.Item("4").Specific.string
                    objPurchaseReturn.Series = 130
                    objPurchaseReturn.UserFields.Fields.Item("U_QCANum").Value = CStr(QCDocNum)
                    objMatrix = objForm.Items.Item("20").Specific
                    Dim OrigDate As Date = Date.ParseExact(Trim(objrs.Fields.Item("DocDate").Value.ToString), "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    'objPurchaseReturn.OriginalRefDate = OrigDate ' Trim(objrs.Fields.Item("DocDate").Value)
                    'objPurchaseReturn.OriginalRefNo = Trim(objrs.Fields.Item("DocNum").Value.ToString)

                    If objAddOn.HANA Then
                        BranchEnabled = objAddOn.objGenFunc.getSingleValue("select ""MltpBrnchs"" from OADM")
                    Else
                        BranchEnabled = objAddOn.objGenFunc.getSingleValue("select MltpBrnchs from OADM")
                    End If
                    If BranchEnabled = "Y" Then
                        Branch = objAddOn.objGenFunc.getSingleValue("Select Top 1 BPLid from OWHS where WhsCode='" & objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string & "'")
                        objPurchaseReturn.BPL_IDAssignedToInvoice = Branch
                    End If
                    'objPurchaseReturn.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                    'objPurchaseReturn.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_GSTTaxInvoice
                    'objPurchaseReturn.GSTTransactionType = SAPbobsCOM.GSTTransactionTypeEnum.gsttrantyp_BillOfSupply
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                        If objSelect.Checked = True Then
                            Serial = objAddOn.objGenFunc.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & objMatrix.Columns.Item("3").Cells.Item(i).Specific.string & "'")
                            Batch = objAddOn.objGenFunc.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & objMatrix.Columns.Item("3").Cells.Item(i).Specific.string & "'")
                            Quantity = CDbl(objMatrix.Columns.Item("6").Cells.Item(i).Specific.string)
                            objPurchaseReturn.Lines.ItemCode = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string
                            objPurchaseReturn.Lines.Quantity = 1 'Quantity
                            objPurchaseReturn.Lines.WarehouseCode = "Main" '"ReJGRN" ' Trim(objMatrix.Columns.Item("5").Cells.Item(i).Specific.string)
                            objPurchaseReturn.Lines.BaseType = 18
                            objPurchaseReturn.Lines.BaseEntry = 10 'Trim(objrs.Fields.Item("DocEntry").Value.ToString)
                            objPurchaseReturn.Lines.BaseLine = 0 'Trim(objrs.Fields.Item("LineNum").Value.ToString)
                            objPurchaseReturn.Lines.UnitPrice = Trim(objrs.Fields.Item("Price").Value.ToString)
                            objPurchaseReturn.Lines.TaxCode = Trim(objrs.Fields.Item("TaxCode").Value.ToString)
                            objPurchaseReturn.Lines.LineTotal = Trim(objrs.Fields.Item("Price").Value.ToString)

                            'If Batch = "Y" And Serial = "N" Then
                            '    Dim BQty As Double = 0, TotBatchQty As Double = 0, LastBQty As Double = 0, PendQty As Double
                            '    BQty = Quantity
                            '    strSQL = "SELECT A.BatchNum as BatchSerial,  SUM(A.Quantity) as Qty FROM ("
                            '    strSQL += vbCrLf + "select T.BatchNum,T.Quantity from ibt1 T inner join oibt T1 on T.ItemCode=T1.ItemCode and T.BatchNum=T1.BatchNum and T.WhsCode=T1.WhsCode"
                            '    strSQL += vbCrLf + "inner join wtr1 T2 on T2.DocEntry=T.BaseEntry and T2.ItemCode=T.ItemCode and T2.LineNum=T.BaseLinNum"
                            '    strSQL += vbCrLf + "inner join owtr T3 on T2.DocEntry=T3.DocEntry"
                            '    strSQL += vbCrLf + "where T.BaseType='67' and T.Direction=0 and T2.ItemCode='" & objMatrix.Columns.Item("3").Cells.Item(i).Specific.string & "' and T3.DocEntry='" & objMatrix.Columns.Item("7B").Cells.Item(i).Specific.string & "')A "
                            '    strSQL += vbCrLf + "GROUP BY A.BatchNum having SUM(A.Quantity) >0"
                            '    objrs.DoQuery(strSQL)
                            '    If objrs.RecordCount > 0 Then
                            '        For j As Integer = 0 To objrs.RecordCount - 1
                            '            If (BQty - TotBatchQty) - CDbl(objrs.Fields.Item("Qty").Value) > 0 Then
                            '                PendQty = CDbl(objrs.Fields.Item("Qty").Value)
                            '            Else
                            '                PendQty = BQty - TotBatchQty
                            '            End If
                            '            objPurchaseReturn.Lines.BatchNumbers.BatchNumber = CStr(objrs.Fields.Item("BatchSerial").Value)
                            '            objPurchaseReturn.Lines.BatchNumbers.Quantity = PendQty ' BQty ' Quantity
                            '            objPurchaseReturn.Lines.BatchNumbers.Add()
                            '            TotBatchQty += PendQty  '2
                            '            If BQty - TotBatchQty > 0 Then
                            '                objrs.MoveNext()
                            '            Else
                            '                Exit For
                            '            End If
                            '        Next
                            '    End If
                            'ElseIf Batch = "N" And Serial = "Y" Then
                            '    Dim SQty As Double = 0, TotSerialQty As Double = 0
                            '    SQty = Quantity
                            '    strSQL = "Select * from (SELECT distinct T4.IntrSerial BatchSerial,T1.DocEntry,T1.ItemCode, T4.Quantity,T4.WhsCode,T4.Status,T1.LineNum from OWTR T0 inner join WTR1 T1 on T0.DocEntry=T1.DocEntry"
                            '    strSQL += vbCrLf + "left outer join SRI1 I1 on T1.ItemCode=I1.ItemCode and (T1.DocEntry=I1.BaseEntry and T1.ObjType=I1.BaseType) and T1.LineNum=I1.BaseLinNum"
                            '    strSQL += vbCrLf + "left outer join OSRI T4 on T4.ItemCode=I1.ItemCode and I1.SysSerial=T4.SysSerial and I1.WhsCode = T4.WhsCode ) A "
                            '    strSQL += vbCrLf + " Where A. DocEntry  = '" & objMatrix.Columns.Item("7B").Cells.Item(i).Specific.string & "' and A. ItemCode ='" & objMatrix.Columns.Item("3").Cells.Item(i).Specific.string & "' and A. BatchSerial  <>'' and A. Status =0 and A.WhsCode='" & objMatrix.Columns.Item("5").Cells.Item(i).Specific.string & "' "
                            '    objrs.DoQuery(strSQL)
                            '    If objrs.RecordCount > 0 Then
                            '        For j As Integer = 0 To objrs.RecordCount - 1
                            '            objPurchaseReturn.Lines.SerialNumbers.InternalSerialNumber = CStr(objrs.Fields.Item("BatchSerial").Value)
                            '            objPurchaseReturn.Lines.SerialNumbers.Quantity = CDbl(1)
                            '            objPurchaseReturn.Lines.SerialNumbers.Add()
                            '            TotSerialQty += CDbl(1)  '2
                            '            If SQty - TotSerialQty > 0 Then
                            '                objrs.MoveNext()
                            '            Else
                            '                Exit For
                            '            End If
                            '        Next
                            '    End If
                            'End If
                            objPurchaseReturn.Lines.Add()
                        End If
                    Next

                    If objPurchaseReturn.Add() <> 0 Then
                        If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        objAddOn.objApplication.SetStatusBarMessage("A/P CreditMemo: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        'objAddOn.objApplication.MessageBox("A/P CreditMemo: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode,, "OK")
                    Else
                        If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        DocEntry = objAddOn.objCompany.GetNewObjectKey()
                        For j = 1 To objMatrix.VisualRowCount
                            objSelect = objMatrix.Columns.Item("0A").Cells.Item(j).Specific
                            If objSelect.Checked = True Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.String = DocEntry
                            End If
                        Next
                        objAddOn.objApplication.StatusBar.SetText("A/P CreditMemo Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objForm.Items.Item("21").Enabled = False
                        objForm.Items.Item("22").Enabled = False
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objPurchaseReturn)
                    GC.Collect()
                End If
            Catch ex As Exception
                'objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                objAddOn.objApplication.MessageBox(ex.Message, , "OK")
                'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End Try
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.Message, , "OK")
            'objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Private Function Create_APCreditMemo_Draft() As Boolean
        Try
            'Under Development
            Dim objAPCreditMemo As SAPbobsCOM.Documents
            Dim DocEntry, QCDocNum As String
            objAPCreditMemo = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
            Dim Lineflag As Boolean = False
            Dim Row As Integer = 1
            For i As Integer = 1 To objMatrix.VisualRowCount
                objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                If objSelect.Checked = True And objMatrix.Columns.Item("11").Cells.Item(i).Specific.String = "" Then
                    Lineflag = True
                    Row = i
                    Exit For
                End If
            Next

            QCDocNum = objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("14").Specific.Selected.value)
            Dim DocDate As Date = Date.ParseExact(objForm.Items.Item("19").Specific.string, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            objMatrix = objForm.Items.Item("20").Specific


            objAPCreditMemo.DocDate = DocDate
            objAPCreditMemo.JournalMemo = "Auto-Gen-> QC Action " & Now.ToString
            objAPCreditMemo.Comments = "QCA DocNum-> " & CStr(QCDocNum)
            objAPCreditMemo.UserFields.Fields.Item("U_QCANum").Value = CStr(QCDocNum)
            objAPCreditMemo.DocObjectCode = 19

            For i As Integer = 1 To objMatrix.VisualRowCount
                objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                If objSelect.Checked = True And objMatrix.Columns.Item("3").Cells.Item(i).Specific.String <> "" Then
                    objAPCreditMemo.Lines.ItemCode = objMatrix.Columns.Item("3").Cells.Item(i).Specific.String
                    objAPCreditMemo.Lines.Quantity = CDbl(objMatrix.Columns.Item("6").Cells.Item(i).Specific.String)
                    'objAPCreditMemo.Lines.AccountCode = Trim(objRs.Fields.Item("AcctCode").Value.ToString)
                    'objAPCreditMemo.Lines.TaxCode = Trim(objRs.Fields.Item("TaxCode").Value.ToString)
                    'objAPCreditMemo.Lines.BaseType = 18
                    'objAPCreditMemo.Lines.BaseEntry = CInt(objRs.Fields.Item("DocEntry").Value.ToString)
                    'objAPCreditMemo.Lines.BaseLine = CInt(objRs.Fields.Item("LineNum").Value.ToString)
                    'objAPCreditMemo.Lines.UnitPrice = Trim(objRs.Fields.Item("StockPrice").Value.ToString)
                    'objAPCreditMemo.Lines.WarehouseCode = Trim(objRs.Fields.Item("WhsCode").Value.ToString)
                    'If Loc() <> "" Then objAPCreditMemo.Lines.LocationCode = Loc()
                End If
            Next
            If objAPCreditMemo.Add() <> 0 Then
                If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objAddOn.objApplication.SetStatusBarMessage("A/P Credit Memo: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                objAddOn.objApplication.MessageBox("A/P Credit Memo: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode,, "OK")
                Return False
            Else
                If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                DocEntry = objAddOn.objCompany.GetNewObjectKey()
                For j = 1 To objMatrix.VisualRowCount
                    objSelect = objMatrix.Columns.Item("0A").Cells.Item(j).Specific
                    If objSelect.Checked = True Then
                        objMatrix.Columns.Item("11").Cells.Item(j).Specific.String = DocEntry
                    End If
                Next
                objAddOn.objApplication.StatusBar.SetText("A/P Credit Memo Draft Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return True
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objAPCreditMemo)
            GC.Collect()
        Catch ex As Exception

        End Try

    End Function

    Public Function Create_AP_Invoice_CreditMemo_Manual(ByVal FormUID As String) As Boolean
        Try
            Dim objAPCMatrix As SAPbouiCOM.Matrix
            Dim objAPCform As SAPbouiCOM.Form
            Dim objrecset As SAPbobsCOM.Recordset
            Dim copyto As SAPbouiCOM.ComboBox
            Dim Lineflag As Boolean = False
            Dim WhsCode As String = ""
            Dim Row As Integer = 1, IRow As Integer
            objMatrix = objForm.Items.Item("20").Specific
            objrecset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            For i As Integer = 1 To objMatrix.VisualRowCount
                objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                If objSelect.Checked = True Then
                    Lineflag = True
                    Row = i
                    Exit For
                End If
            Next
            If Lineflag = True Then
                If objMatrix.Columns.Item("1A").Cells.Item(Row).Specific.string = "G" Then
                    strSQL = "select distinct T0.DocEntry,T1.TaxCode,T1.Price,T1.LineNum,Format(T0.DocDate,'yyyyMMdd') DocDate,T0.DocNum,T0.DocStatus, T1.LineStatus,T1.ItemCode from OPCH T0 join PCH1 T1 on T0.DocEntry=T1.DocEntry where T1.BaseType='20' " &
                                                                       " and T1.BaseEntry=(select distinct T0.U_GRPONum from OWTR T0 join WTR1 T1 on T0.DocEntry=T1.DocEntry where T0.DocEntry=" & objMatrix.Columns.Item("7B").Cells.Item(Row).Specific.string & ") order by T0.DocEntry desc"
                Else
                    strSQL = "select distinct T0.DocEntry,T1.TaxCode,T1.Price,T1.LineNum,Format(T0.DocDate,'yyyyMMdd') DocDate,T0.DocNum,T0.DocStatus, T1.LineStatus,T1.ItemCode from OPCH T0 join PCH1 T1 on T0.DocEntry=T1.DocEntry where T1.U_GRNEntry=(select distinct T0.U_REntry from OWTR T0 " &
                                                                      " join WTR1 T1 on T0.DocEntry=T1.DocEntry where T0.DocEntry=" & objMatrix.Columns.Item("7B").Cells.Item(Row).Specific.string & ") order by T0.DocEntry desc"
                End If
                objrecset.DoQuery(strSQL)
                If objrecset.RecordCount = 0 Then objAddOn.objApplication.StatusBar.SetText("A/P Invoice not found for the specified QC...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
                If objrecset.Fields.Item("DocStatus").Value = "C" Then objAddOn.objApplication.StatusBar.SetText("A/P Invoice Document Status Closed for the specified QC...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
                objAddOn.objApplication.Menus.Item("2308").Activate()
                objAPCform = objAddOn.objApplication.Forms.ActiveForm
                objAPCform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objAPCform.Items.Item("8").Specific.String = objrecset.Fields.Item("DocNum").Value
                objAPCform.Items.Item("10").Specific.String = objrecset.Fields.Item("DocDate").Value
                objAPCform.Items.Item("1").Click()
                copyto = objAPCform.Items.Item("10000329").Specific
                copyto.Item.Click()
                'copyto.SelectExclusive(0, SAPbouiCOM.BoSearchKey.psk_Index)
                copyto.SelectExclusive("19", SAPbouiCOM.BoSearchKey.psk_ByDescription)
                objAPCform.Items.Item("10").Click()
                'copyto.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                objAPCform.Close()
                objAPCform = objAddOn.objApplication.Forms.ActiveForm
                objAPCform = objAddOn.objApplication.Forms.Item(objAPCform.UniqueID)
                objAPCform.Visible = True
                objAPCMatrix = objAPCform.Items.Item("38").Specific
                objAddOn.objApplication.StatusBar.SetText("Data Loading to A/P Credit Memo Screen Please wait ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Dim AcctCode As String = ""
                Try
                    objForm.Freeze(True)
                    objAPCform.Freeze(True)
                    objAPCform.Items.Item("4").Specific.String = objMatrix.Columns.Item("10").Cells.Item(Row).Specific.string
                    objAPCform.Items.Item("16").Specific.String = objAPCform.Items.Item("16").Specific.String + " Auto Gen thro' QC Action"
                    objAPCform.Items.Item("t_qcanum").Specific.String = objForm.Items.Item("15").Specific.String 'objMatrix.Columns.Item("1B").Cells.Item(Row).Specific.string
                    IRow = 1
                    If objAPCMatrix.Columns.Item("U_Item").Editable = False Or objAPCMatrix.Columns.Item("U_Whse").Editable = False Or objAPCMatrix.Columns.Item("U_Qty").Editable = False Or objAPCMatrix.Columns.Item("U_QCEntry").Editable = False Then
                        objAPCMatrix.Columns.Item("U_Item").Editable = True
                        objAPCMatrix.Columns.Item("U_Whse").Editable = True
                        objAPCMatrix.Columns.Item("U_Qty").Editable = True
                        objAPCMatrix.Columns.Item("U_QCEntry").Editable = True
                    End If
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                        If objSelect.Checked = True Then
                            'objAPCMatrix.Columns.Item("1").Cells.Item(IRow).Specific.String = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string
                            objAPCMatrix.Columns.Item("11").Cells.Item(IRow).Specific.String = objMatrix.Columns.Item("6").Cells.Item(i).Specific.string
                            objAPCMatrix.Columns.Item("24").Cells.Item(IRow).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.string
                            'objAPCMatrix.Columns.Item("U_Item").Cells.Item(IRow).Specific.String = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string
                            objAPCMatrix.Columns.Item("U_Whse").Cells.Item(IRow).Specific.String = objMatrix.Columns.Item("5").Cells.Item(i).Specific.string
                            objAPCMatrix.Columns.Item("U_Qty").Cells.Item(IRow).Specific.String = objMatrix.Columns.Item("6").Cells.Item(i).Specific.string
                            objAPCMatrix.Columns.Item("U_QCEntry").Cells.Item(IRow).Specific.String = objMatrix.Columns.Item("1B").Cells.Item(i).Specific.string
                            IRow += 1
                        End If
                    Next
                    objAPCMatrix.Columns.Item("1").Cells.Item(1).Click()
                    objAPCMatrix.Columns.Item("U_Item").Editable = False
                    objAPCMatrix.Columns.Item("U_Whse").Editable = False
                    objAPCMatrix.Columns.Item("U_Qty").Editable = False
                    objAPCMatrix.Columns.Item("U_QCEntry").Editable = False
                    objAPCform.PaneLevel = "8"
                    objAPCform.Items.Item("254000035").Specific.String = objrecset.Fields.Item("DocNum").Value 'Original RefNo
                    objAPCform.Items.Item("254000036").Specific.String = objrecset.Fields.Item("DocDate").Value 'Original RefDate
                    objAPCform.PaneLevel = "1"
                    objrecset = Nothing
                    If objMatrix.Columns.Item("1A").Cells.Item(Row).Specific.string = "G" Then

                    Else
                        objAddOn.objApplication.StatusBar.SetText("Removing the Invalid Items. Please wait... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        For i As Integer = objAPCMatrix.VisualRowCount To 1 Step -1
                            If objAPCMatrix.Columns.Item("1").Cells.Item(i).Specific.String = "" Then Continue For
                            strSQL = objAddOn.objGenFunc.getSingleValue(" Select T0.U_GRNum from [@MIPLQC] T0 join [@MIPLQC1] T1 on T0.DocEntry=T1.DocEntry where T0.DocEntry=" & objAPCMatrix.Columns.Item("U_QCEntry").Cells.Item(i).Specific.string & " and T0.U_Type='R'")
                            If objAPCMatrix.Columns.Item("U_GRNEntry").Cells.Item(i).Specific.String <> strSQL Then
                                objAPCMatrix.Columns.Item("14").Cells.Item(i).Specific.String = 0
                                objAPCMatrix.DeleteRow(i)
                            End If
                        Next
                        objAddOn.objApplication.StatusBar.SetText("Removed the Invalid Items... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    End If
                    objAPCMatrix.Columns.Item("1").Cells.Item(1).Click()
                    objAddOn.objApplication.StatusBar.SetText("Data Loaded to  A/P Credit Memo Screen ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    objForm.Items.Item("21").Enabled = False
                    objForm.Items.Item("22").Enabled = False
                    Return True
                Catch ex As Exception
                    objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Finally
                    objForm.Freeze(False)
                    objAPCform.Freeze(False)
                End Try
            Else
                objAddOn.objApplication.SetStatusBarMessage("No more Data for the transaction...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return False
            End If
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("A/P Credit Memo: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function Create_APCreditMemo_Manual(ByVal FormUID As String) As Boolean
        Try
            Dim objAPCform As SAPbouiCOM.Form
            Dim objAPCMatrix As SAPbouiCOM.Matrix
            Dim Lineflag As Boolean = False
            Dim WhsCode As String = ""
            Dim Row As Integer = 1
            objMatrix = objForm.Items.Item("20").Specific
            For i As Integer = 1 To objMatrix.VisualRowCount
                objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                If objSelect.Checked = True And objMatrix.Columns.Item("11A").Cells.Item(Row).Specific.String = "" Then
                    Lineflag = True
                    Row = i
                    Exit For
                End If
            Next
            If Lineflag = True Then
                objAddOn.objApplication.Menus.Item("2309").Activate()
                objAPCform = objAddOn.objApplication.Forms.ActiveForm
                objAPCform = objAddOn.objApplication.Forms.Item(objAPCform.UniqueID)
                objAPCform.Visible = True
                objAddOn.objApplication.StatusBar.SetText("Data Loading to A/P Credit Memo Screen Please wait ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Try

                    objForm.Freeze(True)
                    objAPCform.Freeze(True)
                    objAPCMatrix = objAPCform.Items.Item("39").Specific
                    objAPCform.Items.Item("4").Specific.String = objMatrix.Columns.Item("10").Cells.Item(Row).Specific.string
                    objAPCform.Items.Item("16").Specific.String = objAPCform.Items.Item("16").Specific.String + " Auto Gen thro' QC Action"
                    objAPCform.Items.Item("t_qcanum").Specific.String = objForm.Items.Item("15").Specific.String
                    objAPCform.Items.Item("3").Specific.select("S", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    objRecordSet.DoQuery("select Code,Name,U_TaxCode from [@MISCODE]")
                    'For i As Integer = 1 To objMatrix.VisualRowCount

                    'Next
                    If objRecordSet.RecordCount > 0 Then
                        Row = 1
                        For Rec As Integer = 0 To objRecordSet.RecordCount - 1
                            objAPCMatrix.Columns.Item("1").Cells.Item(Row).Specific.String = objRecordSet.Fields.Item(1).Value
                            objAPCMatrix.Columns.Item("95").Cells.Item(Row).Specific.String = objRecordSet.Fields.Item(2).Value
                            objRecordSet.MoveNext()
                            Row += 1
                        Next
                        objAPCMatrix.AutoResizeColumns()
                    End If
                    objAddOn.objApplication.StatusBar.SetText("Data Loaded to A/P Credit Memo Screen ...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    objForm.Items.Item("21").Enabled = False
                    objForm.Items.Item("22").Enabled = False
                    Return True
                Catch ex As Exception
                    objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Finally
                    objForm.Freeze(False)
                    objAPCform.Freeze(False)
                End Try
            Else
                objAddOn.objApplication.SetStatusBarMessage("No more Data for the transaction...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return False
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function Create_APCreditMemo(ByVal FormUID As String) As Boolean
        Try
            Dim DocEntry As String, BranchEnabled As String, Branch As String
            Dim objPurchaseReturn As SAPbobsCOM.Documents
            Dim objrs As SAPbobsCOM.Recordset
            Dim Quantity As Double
            Dim QCDocNum As Long
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                Dim Lineflag As Boolean = False
                Dim ToWhse As String = ""
                Dim Row As Integer = 1
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                    If objSelect.Checked = True Then
                        Lineflag = True
                        Row = i
                        Exit For
                    End If
                Next
                If Lineflag = True Then
                    objAddOn.objApplication.StatusBar.SetText("A/P CreditMemo Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objPurchaseReturn = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                    If Not objAddOn.objCompany.InTransaction Then objAddOn.objCompany.StartTransaction()
                    Dim DocDate As Date = Date.ParseExact(objForm.Items.Item("19").Specific.string, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    QCDocNum = objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("14").Specific.Selected.value)
                    objPurchaseReturn.CardCode = objMatrix.Columns.Item("10").Cells.Item(Row).Specific.string
                    objPurchaseReturn.DocDate = DocDate
                    objPurchaseReturn.JournalMemo = "Auto-Gen-> QC Action " & Now.ToString
                    objPurchaseReturn.Comments = "QCA DocNum-> " & CStr(QCDocNum) 'QCHeader.GetValue("DocNum", 0) ' & objForm.Items.Item("4").Specific.string
                    objPurchaseReturn.Series = 130
                    objPurchaseReturn.UserFields.Fields.Item("U_QCANum").Value = CStr(QCDocNum)
                    objMatrix = objForm.Items.Item("20").Specific
                    If objAddOn.HANA Then
                        BranchEnabled = objAddOn.objGenFunc.getSingleValue("select ""MltpBrnchs"" from OADM")
                    Else
                        BranchEnabled = objAddOn.objGenFunc.getSingleValue("select MltpBrnchs from OADM")
                    End If
                    If BranchEnabled = "Y" Then
                        Branch = objAddOn.objGenFunc.getSingleValue("Select Top 1 BPLid from OWHS where WhsCode='" & objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string & "'")
                        objPurchaseReturn.BPL_IDAssignedToInvoice = Branch
                    End If
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                        If objSelect.Checked = True Then
                            Quantity = CDbl(objMatrix.Columns.Item("6").Cells.Item(i).Specific.string)
                            objPurchaseReturn.Lines.ItemCode = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string
                            objPurchaseReturn.Lines.Quantity = Quantity
                            objPurchaseReturn.Lines.WarehouseCode = Trim(objMatrix.Columns.Item("5").Cells.Item(i).Specific.string)
                            objPurchaseReturn.Lines.UnitPrice = 10 ' Trim(objrs.Fields.Item("Price").Value.ToString)
                            objPurchaseReturn.Lines.TaxCode = Trim(objrs.Fields.Item("TaxCode").Value.ToString)
                            'objPurchaseReturn.Lines.LineTotal = Trim(objrs.Fields.Item("Price").Value.ToString)
                            objPurchaseReturn.Lines.Add()
                        End If
                    Next

                    If objPurchaseReturn.Add() <> 0 Then
                        If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        objAddOn.objApplication.SetStatusBarMessage("A/P CreditMemo: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        'objAddOn.objApplication.MessageBox("A/P CreditMemo: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode,, "OK")
                    Else
                        If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        DocEntry = objAddOn.objCompany.GetNewObjectKey()
                        For j = 1 To objMatrix.VisualRowCount
                            objSelect = objMatrix.Columns.Item("0A").Cells.Item(j).Specific
                            If objSelect.Checked = True Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.String = DocEntry
                            End If
                        Next
                        objAddOn.objApplication.StatusBar.SetText("A/P CreditMemo Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objForm.Items.Item("21").Enabled = False
                        objForm.Items.Item("22").Enabled = False
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objPurchaseReturn)
                    GC.Collect()
                End If
            Catch ex As Exception
                'objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                objAddOn.objApplication.MessageBox(ex.Message, , "OK")
                'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End Try
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.Message, , "OK")
            'objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    Public Function Create_InventoryTransfer(ByVal FormUID As String) As Boolean
        Try
            Dim Batch As String, Serial As String, DocEntry As String, TranDocEntry As String = ""
            Dim objstocktransfer As SAPbobsCOM.StockTransfer
            Dim objrs, objcc As SAPbobsCOM.Recordset
            Dim Quantity As Double
            Dim QCDocNum As Long
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                Dim Lineflag As Boolean = False
                Dim ToWhse As String = ""
                'Dim objSelect As SAPbouiCOM.CheckBox
                Dim Row As Integer = 1
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                    If objSelect.Checked = True And objMatrix.Columns.Item("11").Cells.Item(i).Specific.String = "" Then
                        Lineflag = True
                        Row = i
                        Exit For
                    End If
                Next
                If Lineflag = True Then
                    'If objMatrix.Columns.Item("10").Cells.Item(Row).Specific.string <> "" Then
                    '    ToWhse = objAddOn.objGenFunc.getSingleValue("select U_WAREHOUSE from OCRD where CardCode='" & objMatrix.Columns.Item("10").Cells.Item(Row).Specific.string & "' ")
                    'End If
                    'If ToWhse = "" Then objAddOn.objApplication.StatusBar.SetText("Please update the warehouse for specified BP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Function
                    objAddOn.objApplication.StatusBar.SetText("Inventory Transfer Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objstocktransfer = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                    If Not objAddOn.objCompany.InTransaction Then objAddOn.objCompany.StartTransaction()
                    Dim DocDate As Date = Date.ParseExact(objForm.Items.Item("19").Specific.string, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    QCDocNum = objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("14").Specific.Selected.value)
                    objstocktransfer.DocDate = DocDate
                    objstocktransfer.JournalMemo = "Auto-Gen-> QC Action " & Now.ToString
                    objstocktransfer.Comments = "QCA DocNum-> " & CStr(QCDocNum) 'QCHeader.GetValue("DocNum", 0) ' & objForm.Items.Item("4").Specific.string
                    objstocktransfer.UserFields.Fields.Item("U_QCANum").Value = CStr(QCDocNum)
                    objMatrix = objForm.Items.Item("20").Specific
                    objstocktransfer.CardCode = objMatrix.Columns.Item("10").Cells.Item(Row).Specific.string
                    objstocktransfer.FromWarehouse = objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string
                    objstocktransfer.ToWarehouse = objMatrix.Columns.Item("7C").Cells.Item(Row).Specific.string 'ToWhse
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                        If objSelect.Checked = True Then
                            Serial = objAddOn.objGenFunc.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & objMatrix.Columns.Item("3").Cells.Item(i).Specific.string & "'")
                            Batch = objAddOn.objGenFunc.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & objMatrix.Columns.Item("3").Cells.Item(i).Specific.string & "'")
                            Quantity = CDbl(objMatrix.Columns.Item("6").Cells.Item(i).Specific.string)
                            objstocktransfer.Lines.ItemCode = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string
                            objstocktransfer.Lines.Quantity = Quantity
                            objstocktransfer.Lines.FromWarehouseCode = objMatrix.Columns.Item("5").Cells.Item(i).Specific.string 'objForm.Items.Item("25").Specific.string
                            objstocktransfer.Lines.WarehouseCode = objMatrix.Columns.Item("7C").Cells.Item(i).Specific.string ' ToWhse
                            objcc = GetCostCenter(objMatrix.Columns.Item("1A").Cells.Item(i).Specific.string, i, objMatrix.Columns.Item("1B").Cells.Item(i).Specific.string)
                            If objcc.Fields.Item(0).Value <> "" Then objstocktransfer.Lines.DistributionRule = objcc.Fields.Item(0).Value
                            If objcc.Fields.Item(1).Value <> "" Then objstocktransfer.Lines.DistributionRule2 = objcc.Fields.Item(1).Value
                            If objcc.Fields.Item(2).Value <> "" Then objstocktransfer.Lines.DistributionRule3 = objcc.Fields.Item(2).Value
                            If objcc.Fields.Item(3).Value <> "" Then objstocktransfer.Lines.DistributionRule4 = objcc.Fields.Item(3).Value
                            If objcc.Fields.Item(4).Value <> "" Then objstocktransfer.Lines.DistributionRule5 = objcc.Fields.Item(4).Value

                            If Batch = "Y" And Serial = "N" Then
                                Dim BQty As Double = 0, TotBatchQty As Double = 0, LastBQty As Double = 0, PendQty As Double
                                BQty = Quantity
                                strSQL = "SELECT A.BatchNum as BatchSerial,  SUM(A.Quantity) as Qty FROM ("
                                strSQL += vbCrLf + "select T.BatchNum,T.Quantity from ibt1 T inner join oibt T1 on T.ItemCode=T1.ItemCode and T.BatchNum=T1.BatchNum and T.WhsCode=T1.WhsCode"
                                strSQL += vbCrLf + "inner join wtr1 T2 on T2.DocEntry=T.BaseEntry and T2.ItemCode=T.ItemCode and T2.LineNum=T.BaseLinNum"
                                strSQL += vbCrLf + "inner join owtr T3 on T2.DocEntry=T3.DocEntry"
                                strSQL += vbCrLf + "where T.BaseType='67' and T.Direction=0 and T2.ItemCode='" & objMatrix.Columns.Item("3").Cells.Item(i).Specific.string & "' and T2.U_BaseLine='" & objMatrix.Columns.Item("0B").Cells.Item(i).Specific.string & "' and T3.DocEntry='" & objMatrix.Columns.Item("7B").Cells.Item(i).Specific.string & "')A "
                                strSQL += vbCrLf + "GROUP BY A.BatchNum having SUM(A.Quantity) >0"
                                objrs.DoQuery(strSQL)
                                If objrs.RecordCount > 0 Then
                                    For j As Integer = 0 To objrs.RecordCount - 1
                                        If (BQty - TotBatchQty) - CDbl(objrs.Fields.Item("Qty").Value) > 0 Then
                                            PendQty = CDbl(objrs.Fields.Item("Qty").Value)
                                        Else
                                            PendQty = BQty - TotBatchQty
                                        End If
                                        objstocktransfer.Lines.BatchNumbers.BatchNumber = CStr(objrs.Fields.Item("BatchSerial").Value)
                                        objstocktransfer.Lines.BatchNumbers.Quantity = PendQty ' BQty ' Quantity
                                        objstocktransfer.Lines.BatchNumbers.Add()
                                        TotBatchQty += PendQty  '2
                                        If BQty - TotBatchQty > 0 Then
                                            objrs.MoveNext()
                                        Else
                                            Exit For
                                        End If
                                    Next
                                End If
                            ElseIf Batch = "N" And Serial = "Y" Then
                                Dim SQty As Double = 0, TotSerialQty As Double = 0
                                SQty = Quantity
                                strSQL = "Select * from (SELECT distinct T4.IntrSerial BatchSerial,T1.DocEntry,T1.ItemCode, T4.Quantity,T4.WhsCode,T4.Status,T1.LineNum,T1.U_BaseLine from OWTR T0 inner join WTR1 T1 on T0.DocEntry=T1.DocEntry"
                                strSQL += vbCrLf + "left outer join SRI1 I1 on T1.ItemCode=I1.ItemCode and (T1.DocEntry=I1.BaseEntry and T1.ObjType=I1.BaseType) and T1.LineNum=I1.BaseLinNum"
                                strSQL += vbCrLf + "left outer join OSRI T4 on T4.ItemCode=I1.ItemCode and I1.SysSerial=T4.SysSerial and I1.WhsCode = T4.WhsCode ) A "
                                strSQL += vbCrLf + " Where A. DocEntry  = '" & objMatrix.Columns.Item("7B").Cells.Item(i).Specific.string & "' and A.U_BaseLine='" & objMatrix.Columns.Item("0B").Cells.Item(i).Specific.string & "' and A. ItemCode ='" & objMatrix.Columns.Item("3").Cells.Item(i).Specific.string & "' and A. BatchSerial  <>'' and A. Status =0 and A.WhsCode='" & objMatrix.Columns.Item("5").Cells.Item(i).Specific.string & "' "
                                objrs.DoQuery(strSQL)
                                If objrs.RecordCount > 0 Then
                                    For j As Integer = 0 To objrs.RecordCount - 1
                                        objstocktransfer.Lines.SerialNumbers.InternalSerialNumber = CStr(objrs.Fields.Item("BatchSerial").Value)
                                        objstocktransfer.Lines.SerialNumbers.Quantity = CDbl(1)
                                        objstocktransfer.Lines.SerialNumbers.Add()
                                        TotSerialQty += CDbl(1)  '2
                                        If SQty - TotSerialQty > 0 Then
                                            objrs.MoveNext()
                                        Else
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                            objstocktransfer.Lines.Add()
                        End If
                    Next
                    If objstocktransfer.Add() <> 0 Then
                        If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        objAddOn.objApplication.SetStatusBarMessage("Inventory Transfer: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        objAddOn.objApplication.MessageBox("Inventory Transfer: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode,, "OK")
                        Return False
                    Else
                        If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        DocEntry = objAddOn.objCompany.GetNewObjectKey()
                        For j = 1 To objMatrix.VisualRowCount
                            objSelect = objMatrix.Columns.Item("0A").Cells.Item(j).Specific
                            If objSelect.Checked = True Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.String = DocEntry
                            End If
                        Next
                        objAddOn.objApplication.StatusBar.SetText("Inventory Transfer Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objForm.Items.Item("21").Enabled = False
                        objForm.Items.Item("22").Enabled = False
                        Return True
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objstocktransfer)
                    GC.Collect()
                Else
                    objAddOn.objApplication.StatusBar.SetText("No more data for posting Inventory Transfer...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If
            Catch ex As Exception
                'objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                objAddOn.objApplication.MessageBox(ex.Message, , "OK")
                Return False
                'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End Try
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.Message, , "OK")
            Return False
            'objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    Public Function Create_GoodsIssue(ByVal FormUID As String) As Boolean
        Try
            Dim Batch As String, Serial As String, DocEntry As String, BranchEnabled As String, Branch As String, TranDocEntry As String = ""
            Dim objGoodsIssue As SAPbobsCOM.Documents
            Dim objrs, objcc As SAPbobsCOM.Recordset
            Dim Quantity As Double
            Dim QCDocNum As Long
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                Dim Lineflag As Boolean = False
                Dim ToWhse As String = ""
                Dim Row As Integer = 1
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                    If objSelect.Checked = True And objMatrix.Columns.Item("11").Cells.Item(i).Specific.String = "" Then
                        Lineflag = True
                        Row = i
                        Exit For
                    End If
                Next
                If Lineflag = True Then
                    objAddOn.objApplication.StatusBar.SetText("Goods Issue Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objGoodsIssue = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                    If Not objAddOn.objCompany.InTransaction Then objAddOn.objCompany.StartTransaction()
                    Dim DocDate As Date = Date.ParseExact(objForm.Items.Item("19").Specific.string, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    QCDocNum = objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("14").Specific.Selected.value)
                    objGoodsIssue.DocDate = DocDate
                    objGoodsIssue.JournalMemo = "Auto-Gen-> QC Action " & Now.ToString
                    objGoodsIssue.Comments = "QCA DocNum-> " & CStr(QCDocNum) 'QCHeader.GetValue("DocNum", 0) ' & objForm.Items.Item("4").Specific.string
                    objGoodsIssue.UserFields.Fields.Item("U_QCANum").Value = CStr(QCDocNum)
                    objMatrix = objForm.Items.Item("20").Specific
                    'objGoodsIssue.CardCode = objMatrix.Columns.Item("10").Cells.Item(Row).Specific.string
                    If objAddOn.HANA Then
                        BranchEnabled = objAddOn.objGenFunc.getSingleValue("select ""MltpBrnchs"" from OADM")
                    Else
                        BranchEnabled = objAddOn.objGenFunc.getSingleValue("select MltpBrnchs from OADM")
                    End If
                    If BranchEnabled = "Y" Then
                        Branch = objAddOn.objGenFunc.getSingleValue("Select Top 1 BPLid from OWHS where WhsCode='" & objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string & "'")
                        objGoodsIssue.BPL_IDAssignedToInvoice = Branch
                    End If
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
                        If objSelect.Checked = True Then
                            Serial = objAddOn.objGenFunc.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & Trim(objMatrix.Columns.Item("3").Cells.Item(i).Specific.string) & "'")
                            Batch = objAddOn.objGenFunc.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & Trim(objMatrix.Columns.Item("3").Cells.Item(i).Specific.string) & "'")
                            Quantity = CDbl(Trim(objMatrix.Columns.Item("6").Cells.Item(i).Specific.string))
                            objGoodsIssue.Lines.ItemCode = Trim(objMatrix.Columns.Item("3").Cells.Item(i).Specific.string)
                            objGoodsIssue.Lines.Quantity = Quantity
                            ' objGoodsIssue.Lines.FromWarehouseCode = objMatrix.Columns.Item("5").Cells.Item(i).Specific.string 'objForm.Items.Item("25").Specific.string
                            objGoodsIssue.Lines.WarehouseCode = Trim(objMatrix.Columns.Item("5").Cells.Item(i).Specific.string) ' ToWhse)
                            objcc = GetCostCenter(objMatrix.Columns.Item("1A").Cells.Item(i).Specific.string, i, objMatrix.Columns.Item("1B").Cells.Item(i).Specific.string)
                            If objcc.Fields.Item(0).Value <> "" Then objGoodsIssue.Lines.CostingCode = objcc.Fields.Item(0).Value
                            If objcc.Fields.Item(1).Value <> "" Then objGoodsIssue.Lines.CostingCode2 = objcc.Fields.Item(1).Value
                            If objcc.Fields.Item(2).Value <> "" Then objGoodsIssue.Lines.CostingCode3 = objcc.Fields.Item(2).Value
                            If objcc.Fields.Item(3).Value <> "" Then objGoodsIssue.Lines.CostingCode4 = objcc.Fields.Item(3).Value
                            If objcc.Fields.Item(4).Value <> "" Then objGoodsIssue.Lines.CostingCode5 = objcc.Fields.Item(4).Value

                            If Batch = "Y" And Serial = "N" Then
                                Dim BQty As Double = 0, TotBatchQty As Double = 0, PendQty As Double = 0
                                BQty = Quantity
                                strSQL = "SELECT A.BatchNum as BatchSerial,  SUM(A.Quantity) as Qty FROM ("
                                strSQL += vbCrLf + "select T.BatchNum,T.Quantity from ibt1 T inner join oibt T1 on T.ItemCode=T1.ItemCode and T.BatchNum=T1.BatchNum and T.WhsCode=T1.WhsCode"
                                strSQL += vbCrLf + "inner join wtr1 T2 on T2.DocEntry=T.BaseEntry and T2.ItemCode=T.ItemCode and T2.LineNum=T.BaseLinNum"
                                strSQL += vbCrLf + "inner join owtr T3 on T2.DocEntry=T3.DocEntry"
                                strSQL += vbCrLf + "where T.BaseType='67' and T.Direction=0 and T2.ItemCode='" & Trim(objMatrix.Columns.Item("3").Cells.Item(i).Specific.string) & "' and T2.U_BaseLine='" & Trim(objMatrix.Columns.Item("0B").Cells.Item(i).Specific.string) & "' and T3.DocEntry='" & Trim(objMatrix.Columns.Item("7B").Cells.Item(i).Specific.string) & "')A "
                                strSQL += vbCrLf + "GROUP BY A.BatchNum having SUM(A.Quantity) >0"
                                objrs.DoQuery(strSQL)
                                'objAddOn.objGenFunc.WriteErrorLog("Item :" + Trim(objMatrix.Columns.Item("3").Cells.Item(i).Specific.string) + "Quant : " + Trim(objMatrix.Columns.Item("6").Cells.Item(i).Specific.string) + "Whse : " + Trim(objMatrix.Columns.Item("5").Cells.Item(i).Specific.string) + "Query " + strSQL + " Rec Count " + CStr(objrs.RecordCount))
                                If objrs.RecordCount > 0 Then
                                    For j As Integer = 0 To objrs.RecordCount - 1
                                        'objAddOn.objGenFunc.WriteErrorLog("batch1 " + Trim(CStr(objrs.Fields.Item("BatchSerial").Value)) + "Qty " + CStr(PendQty) + "RecQty " + Trim(objrs.Fields.Item("Qty").Value))
                                        If (BQty - TotBatchQty) - CDbl(Trim(objrs.Fields.Item("Qty").Value)) > 0 Then
                                            PendQty = CDbl(Trim(objrs.Fields.Item("Qty").Value))
                                        Else
                                            'objAddOn.objGenFunc.WriteErrorLog("batch1A " + Trim(CStr(objrs.Fields.Item("BatchSerial").Value)) + "Qty " + CStr(PendQty) + "RecQty " + Trim(objrs.Fields.Item("Qty").Value))
                                            PendQty = BQty - TotBatchQty
                                        End If
                                        objGoodsIssue.Lines.BatchNumbers.BatchNumber = Trim(CStr(objrs.Fields.Item("BatchSerial").Value))
                                        objGoodsIssue.Lines.BatchNumbers.Quantity = PendQty ' BQty ' Quantity
                                        objGoodsIssue.Lines.BatchNumbers.Add()
                                        TotBatchQty += PendQty  '2
                                        'objAddOn.objGenFunc.WriteErrorLog("batch2 " + Trim(CStr(objrs.Fields.Item("BatchSerial").Value)) + "Qty " + CStr(PendQty) + "RecQty " + Trim(objrs.Fields.Item("Qty").Value))
                                        If BQty - TotBatchQty > 0 Then
                                            objrs.MoveNext()
                                        Else
                                            'objAddOn.objGenFunc.WriteErrorLog("batch3 " + Trim(CStr(objrs.Fields.Item("BatchSerial").Value)) + "Qty " + CStr(PendQty) + "RecQty " + Trim(objrs.Fields.Item("Qty").Value))
                                            Exit For
                                        End If
                                    Next
                                End If

                            ElseIf Batch = "N" And Serial = "Y" Then
                                Dim SQty As Double = 0, TotSerialQty As Double = 0
                                SQty = Quantity
                                strSQL = "Select * from (SELECT distinct T4.IntrSerial BatchSerial,T1.DocEntry,T1.ItemCode, T4.Quantity,T4.WhsCode,T4.Status,T1.LineNum,T1.U_BaseLine from OWTR T0 inner join WTR1 T1 on T0.DocEntry=T1.DocEntry"
                                strSQL += vbCrLf + "left outer join SRI1 I1 on T1.ItemCode=I1.ItemCode and (T1.DocEntry=I1.BaseEntry and T1.ObjType=I1.BaseType) and T1.LineNum=I1.BaseLinNum"
                                strSQL += vbCrLf + "left outer join OSRI T4 on T4.ItemCode=I1.ItemCode and I1.SysSerial=T4.SysSerial and I1.WhsCode = T4.WhsCode ) A "
                                strSQL += vbCrLf + " Where A. DocEntry  = '" & objMatrix.Columns.Item("7B").Cells.Item(i).Specific.string & "' and A.U_BaseLine='" & objMatrix.Columns.Item("0B").Cells.Item(i).Specific.string & "' and A. ItemCode ='" & objMatrix.Columns.Item("3").Cells.Item(i).Specific.string & "' and A. BatchSerial  <>'' and A. Status =0 and A.WhsCode='" & objMatrix.Columns.Item("5").Cells.Item(i).Specific.string & "' "
                                objrs.DoQuery(strSQL)
                                If objrs.RecordCount > 0 Then
                                    For j As Integer = 0 To objrs.RecordCount - 1
                                        objGoodsIssue.Lines.SerialNumbers.InternalSerialNumber = CStr(objrs.Fields.Item("BatchSerial").Value)
                                        objGoodsIssue.Lines.SerialNumbers.Quantity = CDbl(1)
                                        objGoodsIssue.Lines.SerialNumbers.Add()
                                        TotSerialQty += CDbl(1)  '2
                                        If SQty - TotSerialQty > 0 Then
                                            objrs.MoveNext()
                                        Else
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                            objGoodsIssue.Lines.Add()
                        End If
                    Next
                    If objGoodsIssue.Add() <> 0 Then
                        If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        objAddOn.objApplication.SetStatusBarMessage("Goods Issue: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        objAddOn.objApplication.MessageBox("Goods Issue: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode,, "OK")
                        Return False
                    Else
                        If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        DocEntry = objAddOn.objCompany.GetNewObjectKey()
                        For j = 1 To objMatrix.VisualRowCount
                            objSelect = objMatrix.Columns.Item("0A").Cells.Item(j).Specific
                            If objSelect.Checked = True Then
                                objMatrix.Columns.Item("11").Cells.Item(j).Specific.String = DocEntry
                            End If
                        Next
                        objAddOn.objApplication.StatusBar.SetText("Goods Issue Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        objForm.Items.Item("21").Enabled = False
                        objForm.Items.Item("22").Enabled = False
                        Return True
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objGoodsIssue)
                    GC.Collect()
                Else
                    objAddOn.objApplication.StatusBar.SetText("No more data for posting Goods Issue...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If
            Catch ex As Exception
                'objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                objAddOn.objApplication.MessageBox(ex.Message, , "OK")
                Return False
                'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End Try
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.Message, , "OK")
            Return False
            'objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    Function validate(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("20").Specific
        If objMatrix.VisualRowCount = 0 Then
            objAddOn.objApplication.SetStatusBarMessage("Row is Empty!!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End If
        Dim Flag As Boolean = False
        For i As Integer = 1 To objMatrix.VisualRowCount
            'objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
            If objMatrix.Columns.Item("11").Cells.Item(i).Specific.String <> "" Or objMatrix.Columns.Item("11A").Cells.Item(i).Specific.String <> "" Then
                Flag = True
            End If
            'If (objMatrix.Columns.Item("9").Cells.Item(i).Specific.String = "Inventory Transfer & A/P Credit Memo" Or objMatrix.Columns.Item("9").Cells.Item(i).Specific.String = "Goods Issue & A/P Credit Memo") Then
            '    If objMatrix.Columns.Item("11").Cells.Item(i).Specific.String <> "" And objMatrix.Columns.Item("11A").Cells.Item(i).Specific.String <> "" Then
            '        Flag = True
            '    End If
            'Else
            '    If objMatrix.Columns.Item("11").Cells.Item(i).Specific.String <> "" Or objMatrix.Columns.Item("11A").Cells.Item(i).Specific.String <> "" Then
            '        Flag = True
            '    End If
            'End If
        Next
        If Flag = False Then
            objForm.Items.Item("21").Enabled = True
            objForm.Items.Item("22").Enabled = True
            objAddOn.objApplication.SetStatusBarMessage("Please Select & generate the document!!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End If
        'For i As Integer = 1 To objMatrix.VisualRowCount
        '    objSelect = objMatrix.Columns.Item("0A").Cells.Item(i).Specific
        '    If objSelect.Checked = True Then
        '        If objMatrix.Columns.Item("11").Cells.Item(i).Specific.String = "" Then
        '            objAddOn.objApplication.SetStatusBarMessage("Please generate the document!!!", SAPbouiCOM.BoMessageTime.bmt_Short, True)
        '            Return False
        '        End If
        '    End If
        'Next
        Return True
    End Function

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            objForm = objAddOn.objApplication.Forms.Item(objForm.UniqueID)
            objMatrix = objForm.Items.Item("20").Specific
            Dim ItemUID As String = ""
            Select Case pVal.MenuUID
                Case "1284"

                Case "1281" 'Find Mode
                    If pVal.BeforeAction = False Then
                        objForm.Items.Item("15").Enabled = True
                        objForm.Items.Item("19").Enabled = True
                        objMatrix.Item.Enabled = False
                    End If
                Case "1282"

                Case "1293"  'delete Row
                    For i As Integer = objMatrix.VisualRowCount To 1 Step -1
                        objMatrix.Columns.Item("0").Cells.Item(i).Specific.String = i
                    Next
            End Select

        Catch ex As Exception
            ' objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub

    Private Function test(ByVal Input As String, ByVal pass As String)
        Dim AES As New System.Security.Cryptography.RijndaelManaged
        Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim decrypted As String = ""
        Try
            Dim hash(31) As Byte
            Dim temp As Byte() = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(pass))
            Array.Copy(temp, 0, hash, 0, 16)
            Array.Copy(temp, 0, hash, 15, 16)
            AES.Key = hash
            AES.Mode = Security.Cryptography.CipherMode.ECB
            Dim DESDecrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateDecryptor
            Dim Buffer As Byte() = Convert.FromBase64String(Input)
            decrypted = System.Text.ASCIIEncoding.ASCII.GetString(DESDecrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Return decrypted
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Private Function GetCostCenter(ByVal Type As String, ByVal Row As Integer, ByVal DocEntry As String) As SAPbobsCOM.Recordset
        Try
            Dim objrs As SAPbobsCOM.Recordset
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim THeader As String = "", TLine As String = "", TranDocEntry As String
            If Type = "G" Then
                THeader = "OPDN"
                TLine = "PDN1"
                TranDocEntry = objAddOn.objGenFunc.getSingleValue("Select T0.U_GRNEntry FROM [@MIPLQC] T0 where T0.DocEntry='" & DocEntry & "'")
            ElseIf Type = "P" Or Type = "R" Then
                THeader = "OIGN"
                TLine = "IGN1"
                TranDocEntry = objAddOn.objGenFunc.getSingleValue("Select T0.U_GRNum FROM [@MIPLQC] T0 where T0.DocEntry='" & DocEntry & "'")
            ElseIf Type = "T" Then
                THeader = "OWTR"
                TLine = "WTR1"
                TranDocEntry = objAddOn.objGenFunc.getSingleValue("Select T0.U_TransEntry FROM [@MIPLQC] T0 where T0.DocEntry='" & DocEntry & "'")
            Else
                Exit Function
            End If

            If objAddOn.HANA Then
                strSQL = "Select T1.""OcrCode"",T1.""OcrCode2"",T1.""OcrCode3"",T1.""OcrCode4"",T1.""OcrCode5"" from " & THeader & " T0 join " & TLine & " T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""DocEntry""='" & TranDocEntry & "'"
            Else
                strSQL = "Select T1.OcrCode,T1.OcrCode2,T1.OcrCode3,T1.OcrCode4,T1.OcrCode5 from " & THeader & " T0 join " & TLine & " T1 on T0.DocEntry=T1.DocEntry where T0.DocEntry='" & TranDocEntry & "'"
            End If
            objrs.DoQuery(strSQL)
            If objrs.RecordCount = 0 Then Return Nothing
            Return objrs
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

End Class
