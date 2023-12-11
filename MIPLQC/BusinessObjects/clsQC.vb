Imports System.IO
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework
Public Class clsQC
    Public Const Formtype = "MIPLQC"
    Dim objForm As SAPbouiCOM.Form
    Dim strSQL As String, strQuery As String
    Dim objMatrix, oattachMatrix As SAPbouiCOM.Matrix
    Dim QCHeader As SAPbouiCOM.DBDataSource
    Public QCLine, AttachLine As SAPbouiCOM.DBDataSource
    Dim InWhse As String
    Dim AccWhse As String
    Dim RejWhse As String
    Dim RewWhse As String
    Dim objRecordSet As SAPbobsCOM.Recordset
    Dim AccQty, RejQty, RewQty As Double
    Dim TotQty, InspQty, QtyInsp, Disposable As Double
    Dim PendQty As Double
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim TranFlag As Boolean = False
    Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg

    Public Sub LoadScreen()
        Try
            objForm = objAddOn.objUIXml.LoadScreenXML("QC.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, Formtype)
            objForm.PaneLevel = 1
            objMatrix = objForm.Items.Item("20").Specific
            oattachMatrix = objForm.Items.Item("mtxattach").Specific
            AddCFLConditionEMP(objForm.UniqueID)
            'objForm.Items.Item("2A").Enabled = False
            'objMatrix.CommonSetting.EnableArrowKey = True
            BranchFlag = BranchEnabled(objForm.UniqueID)
            'BranchEnabled(objForm.UniqueID)
            objForm.EnableMenu("772", True)
            objForm.EnableMenu("771", True)
            objForm.EnableMenu("784", True)
            LoadSeries(objForm.UniqueID)
            If objAddOn.HANA Then
                DisposableGoodsIssue = objAddOn.objGenFunc.getSingleValue("select ifnull(""U_DispGI"",'N') from OADM")
            Else
                DisposableGoodsIssue = objAddOn.objGenFunc.getSingleValue("select isnull(U_DispGI,'N') from OADM")
            End If
            If DisposableGoodsIssue = "N" Then
                objForm.PaneLevel = 2
                objForm.Items.Item("33A").Visible = False
                objForm.Items.Item("34B").Visible = False
                objForm.Items.Item("34C").Visible = False
                objForm.Items.Item("BtnGI").Visible = False
                objMatrix.Columns.Item("9B").Visible = False
                objForm.PaneLevel = 1
            End If
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6B").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("13B").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("15").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("15B").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("23B").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("51").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("34").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("2A").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("2A").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objMatrix.Columns.Item("6").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("9").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("7").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("7D").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("7E").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix.Columns.Item("8").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objForm.Items.Item("1000001").Click()
            'objForm.Items.Item("BtnInv").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("6").Specific.Active = True
        Catch ex As Exception

        End Try


    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("20").Specific
            Dim objcombo As SAPbouiCOM.ComboBox
            objcombo = objForm.Items.Item("8").Specific
            If pVal.BeforeAction = True Then
                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If bModal And (objAddOn.objApplication.Forms.ActiveForm.TypeEx = "MIPLQC") Then
                            Try
                                objAddOn.objApplication.Forms.Item("OCFL").Select()
                                BubbleEvent = False
                            Catch ex As Exception
                            End Try
                        End If
                        If pVal.ItemUID = "1" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                            If Not validate(FormUID) Then
                                BubbleEvent = False
                                Exit Sub
                            End If

                        End If
                        'If pVal.ItemUID = "1" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        '    If Not validate(FormUID) Then
                        '        BubbleEvent = False
                        '        Exit Sub
                        '    Else
                        '        objAddOn.objCompany.StartTransaction()
                        '        If StockPosting_GRN(FormUID) Then
                        '            objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        '        Else
                        '            BubbleEvent = False
                        '            objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        '            Exit Sub
                        '        End If
                        '    End If
                        'End If
                        'If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        '    If objForm.Items.Item("34").Specific.String = "" Or objForm.Items.Item("34").Specific.String = "0" Then
                        '        objForm.Items.Item("34").Specific.String = getQCEntry(FormUID)
                        '    Else
                        '        BubbleEvent = False
                        '        Exit Sub
                        '    End If
                        'End If
                    'Case SAPbouiCOM.BoEventTypes.et_CLICK
                    '    If pVal.ItemUID = "1" Then
                    '        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    '            'If validate(FormUID) = False Then
                    '            '    System.Media.SystemSounds.Asterisk.Play()
                    '            '    BubbleEvent = False
                    '            '    objAddOn.objApplication.StatusBar.SetText("1", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            '    Exit Sub
                    '            'End If
                    '            objMatrix.Columns.Item("6").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                    '            objMatrix.Columns.Item("9").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                    '            objMatrix.Columns.Item("7").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                    '            objMatrix.Columns.Item("7D").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                    '            objMatrix.Columns.Item("7E").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                    '            objMatrix.Columns.Item("8").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                    '        End If
                    '    End If
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                        'BubbleEvent = False
                        If pVal.ItemUID = "6B" Then
                            AddCFLConditionEMP(FormUID)
                        End If
                        If pVal.ItemUID = "20" Then
                            If pVal.ColUID = "6A" Then
                                setCFLCond(FormUID, "CFL_1", pVal.Row)
                            ElseIf pVal.ColUID = "7A" Then
                                setCFLCond(FormUID, "CFL_2", pVal.Row)
                            ElseIf pVal.ColUID = "8A" Then
                                setCFLCond(FormUID, "CFL_3", pVal.Row)
                            End If
                        End If
                        objcombo = objForm.Items.Item("8").Specific

                        If objcombo.Selected.Value = "R" Then
                            ChooseFromListFilteration(FormUID, objForm, "CFL_GR", "", "")
                        ElseIf objcombo.Selected.Value = "P" Then
                            AddCFLConditionPO(FormUID)
                            'If objForm.Items.Item("15B").Specific.String <> "" Then
                            If objAddOn.HANA Then
                                ChooseFromListFilteration(FormUID, objForm, "CFL_GR", "DocEntry", "select ""DocEntry"" from IGN1 where ""BaseType""='202' and ""BaseEntry""='" & objForm.Items.Item("15B").Specific.String & "'")
                            Else
                                ChooseFromListFilteration(FormUID, objForm, "CFL_GR", "DocEntry", "select DocEntry from IGN1 where BaseType='202' and BaseEntry='" & objForm.Items.Item("15B").Specific.String & "'")
                            End If
                            'End If
                        ElseIf objcombo.Selected.Value = "T" Then
                            'AddCFLConditionInvTransfer(objForm.UniqueID)
                            ChooseFromListFilterationINV(objForm, "CFL_INV")
                        ElseIf objcombo.Selected.Value = "G" Then
                            AddCFLConditionGRPO(FormUID)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            If objForm.Items.Item("34").Specific.string <> "" Then
                                'objForm.Items.Item("1").Click()
                                'objAddOn.objApplication.MessageBox("Please update")
                                'BubbleEvent = False
                            End If
                        End If
                End Select
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        objMatrix.AutoResizeColumns()
                        oattachMatrix.AutoResizeColumns()
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            'If validate(FormUID) = True Then
                            'If objForm.Items.Item("34").Specific.string = "" Then
                            '    Auto_StockTransfer()
                            'Else
                            '    objAddOn.objApplication.StatusBar.SetText("Already Posted Inventory Transfer", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            'End If
                            'Else
                            '    Exit Sub
                            'End If
                            'ElseIf pVal.ItemUID = "2B" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            '    DeleteRow()
                        End If
                        If pVal.ItemUID = "2A" Then
                            'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objForm.Items.Item("2A").Enabled = False Then Exit Sub
                            'If objForm.Items.Item("34").Specific.string = "" Then
                            '    Auto_StockTransfer(objForm.UniqueID)
                            'Else

                            '    objForm.Items.Item("2A").Enabled = False
                            '    objAddOn.objApplication.StatusBar.SetText("Already Posted Inventory Transfer", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            'End If
                            ''StockTransfer_BinLocation(FormUID)
                            ''' BatchUpdate()
                        ElseIf pVal.ItemUID = "Btnbrowse" Then
                            If objForm.Items.Item(pVal.ItemUID).Enabled = False Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                            If objAddOn.objGenFunc.SetAttachMentFile(objForm, QCHeader, oattachMatrix, AttachLine) = False Then
                                BubbleEvent = False
                            End If
                        ElseIf pVal.ItemUID = "Btndisp" Then
                            If objForm.Items.Item(pVal.ItemUID).Enabled = False Then Exit Sub
                            If pVal.ActionSuccess Then objAddOn.objGenFunc.OpenAttachment(oattachMatrix, AttachLine, pVal.Row)
                        ElseIf pVal.ItemUID = "Btndel" Then
                            If objForm.Items.Item(pVal.ItemUID).Enabled = False Then Exit Sub
                            If pVal.ActionSuccess Then
                                objAddOn.objGenFunc.DeleteRowAttachment(objForm, oattachMatrix, AttachLine, oattachMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder))
                            End If
                        ElseIf pVal.ItemUID = "1000001" Then
                            objForm.PaneLevel = 1
                        ElseIf pVal.ItemUID = "32" Then
                            objForm.PaneLevel = 2
                        ElseIf pVal.ItemUID = "fattach" Then
                            objForm.PaneLevel = 3
                            oattachMatrix.AutoResizeColumns()
                        ElseIf pVal.ItemUID = "BtnInv" And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                            'CreateMySimpleForm("InvTransfer", "Inventory Transfer Details", "OWTR", "67")
                            CreateMySimpleForm("VIEWDATA", "Inventory Transfer Details", "", "67")
                        ElseIf pVal.ItemUID = "BtnGI" And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            CreateMySimpleForm("VIEWDATA", "Disposable Goods Issue Details", "OIGE", "60")
                        End If
                        If pVal.ItemUID = "20" Then
                            objMatrix.SelectRow(pVal.Row, True, False)
                            'objForm.ActiveItem = "20"
                        ElseIf pVal.ItemUID = "mtxattach" Then
                            oattachMatrix.SelectRow(pVal.Row, True, False)
                            If pVal.Row > 0 Then
                                If oattachMatrix.IsRowSelected(pVal.Row) Then
                                    objForm.Items.Item("Btndisp").Enabled = True
                                    objForm.Items.Item("Btndel").Enabled = True
                                End If
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        If pVal.ItemUID = "mtxattach" Then
                            If pVal.ActionSuccess Then objAddOn.objGenFunc.OpenAttachment(oattachMatrix, AttachLine, pVal.Row)
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                        If pVal.ActionSuccess = False Then Exit Sub
                        'If CDbl(objMatrix.Columns.Item("7").Cells.Item(pVal.Row).Specific.string) = 0 Then If objMatrix.Columns.Item("7_1").Cells.Item(pVal.Row).Specific.string <> "" Then objMatrix.Columns.Item("7_1").Cells.Item(pVal.Row).Specific.string = ""
                        If (pVal.ItemUID = "13") Or (pVal.ItemUID = "23") Or (pVal.ItemUID = "15") Or (pVal.ItemUID = "51" And objcombo.Selected.Value = "R") Then
                            'getDocumentEntry(FormUID, objcombo.Selected.Value)
                            If objcombo.Selected.Value = "G" Then
                                getDocumentEntry(FormUID, "G", objForm.Items.Item("13B").Specific.string)
                                'AssignDocEntry(FormUID, "G", objForm.Items.Item("13").Specific.string)
                            ElseIf objcombo.Selected.Value = "T" Then
                                getDocumentEntry(FormUID, "T", objForm.Items.Item("23B").Specific.string)
                                'AssignDocEntry(FormUID, "T", objForm.Items.Item("23").Specific.string)
                            ElseIf objcombo.Selected.Value = "R" Then
                                getDocumentEntry(FormUID, "R", objForm.Items.Item("51B").Specific.string)
                            End If
                            If Validate_Batch_Serial() = False Then
                                objAddOn.objApplication.StatusBar.SetText("Batch_Serial Not found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                            If pVal.ItemUID <> "15" Then
                                LoadInspectionMatrix(FormUID, objForm.Items.Item("8").Specific.selected.value)
                            End If
                        ElseIf (pVal.ItemUID = "51" And (objcombo.Selected.Value = "P")) Then
                            If Validate_Batch_Serial() = False Then
                                objAddOn.objApplication.StatusBar.SetText("Batch_Serial Not found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                            LoadInspectionMatrix(FormUID, objForm.Items.Item("8").Specific.selected.value)
                            getDocumentEntry(FormUID, "P", objForm.Items.Item("51B").Specific.string)
                        ElseIf pVal.ItemUID = "20" And pVal.ColUID = "7" Then
                            objMatrix.Columns.Item("7E").Cells.Item(pVal.Row).Specific.string = CDbl(objMatrix.Columns.Item("7").Cells.Item(pVal.Row).Specific.string) * CDbl(objMatrix.Columns.Item("7D").Cells.Item(pVal.Row).Specific.string)
                            objForm.Update()
                        ElseIf pVal.ItemUID = "20" And pVal.ColUID = "9" Then
                            objMatrix.Columns.Item("3").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                            objMatrix.Columns.Item("4").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                            objMatrix.Columns.Item("5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If pVal.ItemUID = "8" Then
                            TypeSelection(FormUID)
                        ElseIf pVal.ItemUID = "21" Then
                            objForm.Items.Item("4").Specific.String = objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("21").Specific.selected.value, Formtype)
                        ElseIf pVal.ItemUID = "50" Then
                            LoadLocation(FormUID)
                            'CFLConditions(FormUID)
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                        Try
                            If pVal.CharPressed = 13 Then
                                Dim cflForm As SAPbouiCOM.Form
                                If objAddOn.objApplication.Forms.Count > 0 Then
                                    For frm As Integer = 0 To objAddOn.objApplication.Forms.Count - 1
                                        If objAddOn.objApplication.Forms.Item(frm).UniqueID = "OCFL" Then
                                            cflForm = objAddOn.objApplication.Forms.Item("OCFL")
                                            cflForm.Close()
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                        Catch ex As Exception
                        End Try

                        If (pVal.ItemUID = "13" And pVal.CharPressed = 13) Then
                            If objForm.Items.Item("13").Specific.string = "" Then
                                objAddOn.objCFLList.LoadScreen(FormUID, "G")
                            End If
                        ElseIf (pVal.ItemUID = "15" And pVal.CharPressed = 13) Then
                            If objForm.Items.Item("15").Specific.string = "" Then
                                objAddOn.objCFLList.LoadScreen(FormUID, "P")
                            End If
                        ElseIf (pVal.ItemUID = "23" And pVal.CharPressed = 13) Then
                            If objForm.Items.Item("23").Specific.string = "" Then
                                objAddOn.objCFLList.LoadScreen(FormUID, "T")
                            End If
                        ElseIf (pVal.ItemUID = "51" And pVal.CharPressed = 13) Then
                            If objForm.Items.Item("51").Specific.string = "" Then
                                objAddOn.objCFLList.LoadScreen(FormUID, "R")
                            End If
                        End If
                        objMatrix = objForm.Items.Item("20").Specific
                        If pVal.ItemUID = "20" And pVal.ColUID = "7_1" And (pVal.CharPressed = 37 Or pVal.CharPressed = 39 Or pVal.CharPressed = 9) Then
                            If objMatrix.Columns.Item("7_1").Cells.Item(pVal.Row).Specific.string = "" And CDbl(objMatrix.Columns.Item("7").Cells.Item(pVal.Row).Specific.string) > 0 Then
                                objAddOn.objRejDetails.LoadScreen(FormUID, pVal.Row)
                            Else
                                Exit Sub
                                'objAddOn.objApplication.StatusBar.SetText("Please Enter the Rejection Qty to Open Rejection details Screen...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                        If pVal.ItemUID = "20" And pVal.ColUID = "7_1" Then
                            objMatrix.Columns.Item("7_1").Cells.Item(pVal.Row).Specific.string = ""
                            BubbleEvent = False : Exit Sub
                        End If
                        Try
                            Dim ColID As Integer = objMatrix.GetCellFocus().ColumnIndex
                            If pVal.ItemUID = "20" And pVal.CharPressed = 38 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then  'up
                                objMatrix.SetCellFocus(pVal.Row - 1, ColID)
                                objMatrix.SelectRow(pVal.Row - 1, True, False)
                            ElseIf pVal.ItemUID = "20" And pVal.CharPressed = 40 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then 'down
                                objMatrix.SetCellFocus(pVal.Row + 1, ColID)
                                objMatrix.SelectRow(pVal.Row + 1, True, False)
                            ElseIf pVal.ItemUID = "20" And pVal.CharPressed = 37 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then 'Left
                                objMatrix.SetCellFocus(pVal.Row, ColID - 1)
                            ElseIf pVal.ItemUID = "20" And pVal.CharPressed = 39 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then 'Right
                                objMatrix.SetCellFocus(pVal.Row, ColID + 1)
                            End If
                        Catch ex As Exception
                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        If pVal.ItemUID = "20" And pVal.ColUID = "7_1" Then
                            objMatrix = objForm.Items.Item("20").Specific
                            objAddOn.objRejDetails.LoadScreen(FormUID, pVal.Row, objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.ItemUID = "20" And (pVal.ColUID = "6A" Or pVal.ColUID = "7A" Or pVal.ColUID = "8A" Or pVal.ColUID = "AccDet" Or pVal.ColUID = "12") Then
                            CFL(FormUID, pVal)
                        End If
                        If pVal.ItemUID = "15" Or pVal.ItemUID = "13" Or pVal.ItemUID = "6B" Or pVal.ItemUID = "51" Or pVal.ItemUID = "23" Then
                            CFL(FormUID, pVal)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Action_Success = True Then
                            'TypeSelection(FormUID)
                            'BranchEnabled(objForm.UniqueID)
                            LoadSeries(objForm.UniqueID)
                        End If
                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
            objMatrix = objForm.Items.Item("20").Specific
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    Try
                        Dim errorCode As Integer

                        If BusinessObjectInfo.ActionSuccess = False And BusinessObjectInfo.BeforeAction = True Then
                            If objMatrix.VisualRowCount = 0 Then BubbleEvent = False : objAddOn.objApplication.StatusBar.SetText("Row Data is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub
                            If Not objAddOn.objCompany.InTransaction Then objAddOn.objCompany.StartTransaction()
                            If Auto_StockTransfer(BusinessObjectInfo.FormUID) Then TranFlag = True Else TranFlag = False : BubbleEvent = False : Exit Sub
                            If DisposableGoodsIssue = "Y" Then If TranFlag = True Then If Create_GoodsIssue(objForm.UniqueID) Then TranFlag = True Else TranFlag = False Else TranFlag = False : BubbleEvent = False : Exit Sub
                            If TranFlag = True Then
                                If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            Else
                                If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                objAddOn.objCompany.GetLastError(errorCode, strSQL)
                                objForm.Items.Item("34A").Specific.string = "" : objForm.Items.Item("34").Specific.string = ""
                                If DisposableGoodsIssue = "Y" Then objForm.Items.Item("34B").Specific.string = "" : objForm.Items.Item("34C").Specific.string = ""
                                BubbleEvent = False
                                objAddOn.objApplication.MessageBox("RolledBack transactions... " & strSQL, , "OK")
                                objAddOn.objApplication.StatusBar.SetText("RolledBack transactions..." & strSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        Else
                            If BusinessObjectInfo.ActionSuccess = True Then
                                UpdateDocEntry_InvTransfer("N")
                                objAddOn.objApplication.StatusBar.SetText("Committed transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If

                    Catch ex As Exception
                        objAddOn.objApplication.StatusBar.SetText("FormDataEvent Exception: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                        objAddOn.objApplication.MessageBox("Exception: " & ex.Message.ToString, , "OK")
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = True Then
                        objForm.EnableMenu("1282", True)
                    Else
                        Try
                            Dim objedit As SAPbouiCOM.EditText
                            objedit = objForm.Items.Item("27").Specific
                            Dim Fieldsize As Size = TextRenderer.MeasureText(objedit.Value, New Font("Arial", 12.0F))
                            If Fieldsize.Width <= 135 Then
                                objedit.Item.Width = 135
                            Else
                                objedit.Item.Width = Fieldsize.Width
                            End If
                            TypeSelection(objForm.UniqueID)

                            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("BtnInv").Enabled = True
                            Dim oUDFForm As SAPbouiCOM.Form
                            oUDFForm = objAddOn.objApplication.Forms.Item(objForm.UDFFormUID)
                            oUDFForm.Items.Item("U_QCAReq").Enabled = True
                            If objForm.Items.Item("34").Specific.string <> "" Then  ' objForm.Items.Item("6G").Specific.Selected.Value = "O" 
                                objForm.Items.Item("2A").Enabled = False
                            End If
                        Catch ex As Exception
                            'objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If


            End Select
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Form_DataEvent: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
            objAddOn.objApplication.MessageBox("Exception: " & ex.Message.ToString, , "OK")
        End Try

    End Sub

    Private Sub UpdateDocEntry_InvTransfer(ByVal RolledBack As String)
        Try
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If RolledBack = "Y" Then
                'If objAddOn.HANA Then
                'strSQL = "Update OWTR set ""U_QCEntry""='', ""U_QCNum""='' where ""DocEntry""='" & objForm.Items.Item("34A").Specific.String & "'"
                'objRecordSet.DoQuery(strSQL)
                '    strSQL = "Update ""@MIPLQC"" set ""U_AccStk""='',""U_AccStkD""='',""Status""='O where ""DocEntry""='" & QCHeader.GetValue("DocEntry", 0) & "' "
                '    objRecordSet.DoQuery(strSQL)
                '    If objForm.Items.Item("34C").Specific.String <> "" Then
                '        strSQL = "Update OIGE set ""U_QCEntry""='', ""U_QCNum""='' where ""DocEntry""='" & objForm.Items.Item("34C").Specific.String & "'"
                '        objRecordSet.DoQuery(strSQL)
                '    End If
                'Else
                '    strSQL = "Update OWTR set U_QCEntry='', U_QCNum='' where DocEntry='" & objForm.Items.Item("34A").Specific.String & "'"
                '    objRecordSet.DoQuery(strSQL)
                '    strSQL = "Update [@MIPLQC] set U_AccStk='',U_AccStkD='',Status='O where DocEntry='" & QCHeader.GetValue("DocEntry", 0) & "' "
                '    objRecordSet.DoQuery(strSQL)
                '    If objForm.Items.Item("34C").Specific.String <> "" Then
                '        strSQL = "Update OIGE set U_QCEntry='', U_QCNum='' where DocEntry='" & objForm.Items.Item("34C").Specific.String & "'"
                '        objRecordSet.DoQuery(strSQL)
                '    End If
                'End If

            Else
                If objAddOn.HANA Then
                    strSQL = "Update OWTR set ""U_QCEntry""='" & QCHeader.GetValue("DocEntry", 0) & "', ""U_QCNum""='" & QCHeader.GetValue("DocNum", 0) & "' where ""DocEntry""='" & objForm.Items.Item("34A").Specific.String & "'"
                    objRecordSet.DoQuery(strSQL)
                    strSQL = "Update ""@MIPLQC"" set ""U_AccStk""='" & objForm.Items.Item("34").Specific.String & "',""U_AccStkD""='" & objForm.Items.Item("34A").Specific.String & "',""Status""='C' where ""DocEntry""='" & QCHeader.GetValue("DocEntry", 0) & "' "
                    objRecordSet.DoQuery(strSQL)
                    If objForm.Items.Item("34C").Specific.String <> "" Then
                        strSQL = "Update OIGE set ""U_QCEntry""='" & QCHeader.GetValue("DocEntry", 0) & "', ""U_QCNum""='" & QCHeader.GetValue("DocNum", 0) & "' where ""DocEntry""='" & objForm.Items.Item("34C").Specific.String & "'"
                        objRecordSet.DoQuery(strSQL)
                    End If
                Else
                    strSQL = "Update OWTR set U_QCEntry='" & QCHeader.GetValue("DocEntry", 0) & "', U_QCNum='" & QCHeader.GetValue("DocNum", 0) & "' where DocEntry='" & objForm.Items.Item("34A").Specific.String & "'"
                    objRecordSet.DoQuery(strSQL)
                    strSQL = "Update [@MIPLQC] set U_AccStk='" & objForm.Items.Item("34").Specific.String & "',U_AccStkD='" & objForm.Items.Item("34A").Specific.String & "',Status='C' where DocEntry='" & QCHeader.GetValue("DocEntry", 0) & "' "
                    objRecordSet.DoQuery(strSQL)
                    If objForm.Items.Item("34C").Specific.String <> "" Then
                        strSQL = "Update OIGE set U_QCEntry='" & QCHeader.GetValue("DocEntry", 0) & "', U_QCNum='" & QCHeader.GetValue("DocNum", 0) & "' where DocEntry='" & objForm.Items.Item("34C").Specific.String & "'"
                        objRecordSet.DoQuery(strSQL)
                    End If
                End If

            End If



        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            objForm = objAddOn.objApplication.Forms.Item(objForm.UniqueID)
            objMatrix = objForm.Items.Item("20").Specific
            Dim ItemUID As String = ""
            Select Case pVal.MenuUID
                Case "1284"

                Case "1281" 'Find Mode
                    If pVal.BeforeAction = False Then
                        objForm.Items.Item("4").Enabled = True
                        objForm.Items.Item("6").Enabled = True
                        objForm.Items.Item("6E").Enabled = True
                        objForm.Items.Item("6C").Enabled = True
                        objForm.Items.Item("25").Enabled = True
                        objForm.Items.Item("27").Enabled = True
                        objForm.Items.Item("13B").Enabled = True
                        objForm.Items.Item("23B").Enabled = True
                        objForm.Items.Item("15B").Enabled = True
                        objForm.Items.Item("51B").Enabled = True
                        objForm.Items.Item("34").Enabled = True
                        objForm.Items.Item("6G").Enabled = True
                        objMatrix.Item.Enabled = False
                        objForm.Items.Item("34B").Enabled = True
                        objForm.Items.Item("BtnInv").Enabled = False
                        objForm.Items.Item("BtnGI").Enabled = False
                        objForm.Items.Item("Btnbrowse").Enabled = False
                        objForm.Items.Item("mtxattach").Enabled = False
                    End If
                    Dim oUDFForm As SAPbouiCOM.Form
                    If objAddOn.objApplication.Forms.ActiveForm.TypeEx.Contains("940") Then
                        oUDFForm = objAddOn.objApplication.Forms.Item(objAddOn.objApplication.Forms.ActiveForm.UDFFormUID)
                        oUDFForm.Items.Item("U_QCNum").Enabled = True
                        oUDFForm.Items.Item("U_QCEntry").Enabled = True
                        oUDFForm.Items.Item("U_GRPONum").Enabled = True
                        oUDFForm.Items.Item("U_GRNEntry").Enabled = True
                        oUDFForm.Items.Item("U_PORNum").Enabled = True
                        oUDFForm.Items.Item("U_ProdEntry").Enabled = True
                        oUDFForm.Items.Item("U_REntry").Enabled = True
                        oUDFForm.Items.Item("U_GREntry").Enabled = True
                        oUDFForm.Items.Item("U_StkEntry").Enabled = True
                        oUDFForm.Items.Item("U_StkNum").Enabled = True

                    End If
                Case "1282"
                    If pVal.BeforeAction = False Then
                        objCombo = objForm.Items.Item("50").Specific
                        objCombo.Select("-1", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        objCombo = objForm.Items.Item("10").Specific
                        objCombo.Select("-", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If

                Case "1293"  'delete Row
                    'For i As Integer = objMatrix.VisualRowCount To 1 Step -1
                    '    objMatrix.Columns.Item("0").Cells.Item(i).Specific.String = i
                    'Next
                    DeleteRow()

            End Select

        Catch ex As Exception
            ' objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub

    Sub RightClickEvent(ByRef EventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objAddOn.objApplication.Forms.Item(objForm.UniqueID)
            objMatrix = objForm.Items.Item("20").Specific
            If EventInfo.BeforeAction Then
                Select Case EventInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK
                        objForm.EnableMenu("1283", False)
                        Select Case EventInfo.ItemUID
                            Case "20"
                                If (EventInfo.ColUID = "0" Or EventInfo.ColUID = "1") And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    objForm.EnableMenu("1293", True)
                                Else
                                    objForm.EnableMenu("1293", False)
                                End If
                            Case Else
                                objForm.EnableMenu("1293", False)
                        End Select
                End Select
            Else

            End If

            'If EventInfo.BeforeAction Then
            'Else
            '    Select Case EventInfo.ItemUID
            '        Case "20"
            '            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '                objForm.EnableMenu("1293", True)
            '            End If

            '    End Select
            'End If

        Catch ex As Exception
            'objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub

    Public Sub RemoveLastrow(ByVal Columname_check As String)
        Try
            objMatrix = objForm.Items.Item("20").Specific
            If objMatrix.VisualRowCount = 0 Then Exit Sub
            If Columname_check.ToString = "" Then Exit Sub
            For i As Integer = 1 To objMatrix.VisualRowCount
                If objMatrix.IsRowSelected(i) Then
                    If objMatrix.Columns.Item(Columname_check).Cells.Item(i).Specific.string = "" Then
                        objMatrix.DeleteRow(i)
                    End If
                End If
            Next

        Catch ex As Exception

        End Try
    End Sub

    Sub DeleteRow()
        Try
            objMatrix = objForm.Items.Item("20").Specific
            objMatrix.FlushToDataSource()
            QCLine = objForm.DataSources.DBDataSources.Item("@MIPLQC1")
            For i As Integer = 1 To objMatrix.VisualRowCount
                'If objMatrix.IsRowSelected(i) Then
                objMatrix.GetLineData(i)
                QCLine.Offset = i - 1
                QCLine.SetValue("LineId", QCLine.Offset, i)
                objMatrix.SetLineData(i)
                objMatrix.FlushToDataSource()
                'End If
            Next
            QCLine.RemoveRecord(QCLine.Size - 1)
            objMatrix.LoadFromDataSource()

        Catch ex As Exception
            ' objAddOn.objApplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Private Sub AddCFLConditionPO(ByVal FormUID As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_PO")
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim BranchID As String = "", LocCode As String = ""
            Dim rsetCFL As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            oCond = oConds.Add
            oCond.Alias = "Status"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = "R"
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCond = oConds.Add
            oCond.Alias = "Status"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCond.CondVal = "C"
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCond = oConds.Add
            oCond.Alias = "Status"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCond.CondVal = "L"

            ''If BranchFlag = "Y" Then
            ''    objCombo = objForm.Items.Item("50").Specific
            ''    If Not objCombo.Selected.Value = "-1" Then
            ''        BranchID = objCombo.Selected.Value
            ''    End If
            ''End If
            ''objCombo = objForm.Items.Item("10").Specific
            ''If Not objCombo.Selected.Value = "-" Then
            ''    LocCode = objCombo.Selected.Value
            ''End If
            ''If objAddOn.HANA Then
            ''    strQuery = "Select distinct A.""BaseEntry"",A.""ItemCode"" from (SELECT distinct T2.""BaseEntry"", T2.""ItemCode"",T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1  INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" "
            ''    strQuery += vbCrLf + "AND  T0.""U_GRNum"" = cast(T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""PendQty"" FROM OIGN T5 inner join IGN1 T2 on T5.""DocEntry""=T2.""DocEntry"" INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" "
            ''    strQuery += vbCrLf + "left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" where T5.""DocStatus""='O' and T2.""BaseType"" ='202' and T5.""CANCELED""='N' and T5.""U_QCReq""='Y'  and T3.""U_InspReq""='Y' "
            ''    If BranchID <> "" Then
            ''        strQuery += vbCrLf + "and T5.""BPLId""='" & BranchID & "' "
            ''    End If
            ''    If LocCode <> "" Then
            ''        strQuery += vbCrLf + "and T2.""LocCode""='" & LocCode & "' "
            ''    End If
            ''    strQuery += vbCrLf + "GROUP BY T4.""BaseQty"",T4.""AltQty"",T2.""BaseEntry"",T2.""ItemCode"") as A where A.""PendQty"">0 and A.""BaseEntry"" is not null"
            ''Else
            ''    strQuery = "Select distinct A.BaseEntry,A.ItemCode from (SELECT distinct T2.BaseEntry,T2.ItemCode, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (SELECT ISNULL(SUM(T1.U_QtyInsp), 0) FROM @MIPLQC1 T1  INNER JOIN @MIPLQC T0 ON T0.DocEntry=T1.DocEntry "
            ''    strQuery += vbCrLf + "AND  T0.U_GRNum = cast(T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum) AS PendQty FROM OIGN T5 inner join IGN1 T2 on T5.DocEntry=T2.DocEntry INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode "
            ''    strQuery += vbCrLf + "left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T5.DocStatus='O' and T2.BaseType ='202' and T5.CANCELED='N' and T5.U_QCReq='Y'  and T3.U_InspReq='Y' "
            ''    If BranchID <> "" Then
            ''        strQuery += vbCrLf + "and T5.BPLId='" & BranchID & "' "
            ''    End If
            ''    If LocCode <> "" Then
            ''        strQuery += vbCrLf + "and T2.LocCode='" & LocCode & "' "
            ''    End If
            ''    strQuery += vbCrLf + "GROUP BY T4.BaseQty,T4.AltQty,T2.BaseEntry,T2.ItemCode) as A where A.PendQty>0 and A.BaseEntry is not null"
            ''End If
            ''rsetCFL.DoQuery(strQuery)
            ''If rsetCFL.RecordCount > 0 Then
            ''    For i As Integer = 1 To rsetCFL.RecordCount
            ''        If i = (rsetCFL.RecordCount) Then
            ''            oCond = oConds.Add()
            ''            oCond.Alias = "DocEntry"
            ''            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            ''            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
            ''        Else
            ''            oCond = oConds.Add()
            ''            oCond.Alias = "DocEntry"
            ''            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            ''            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
            ''            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            ''        End If
            ''        rsetCFL.MoveNext()
            ''    Next
            ''Else
            ''    oCond = oConds.Add()
            ''    oCond.Alias = "DocEntry"
            ''    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            ''    oCond.CondVal = ""
            ''End If

            'objCombo = objForm.Items.Item("50").Specific
            'If Not objCombo.Selected.Value = "-1" Then
            '    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            '    oCond = oConds.Add
            '    oCond.Alias = "BPLId"
            '    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '    oCond.CondVal = objCombo.Selected.Value
            'End If
            oCFL.SetConditions(oConds)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub AddCFLConditionEMP(ByVal FormUID As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("EMP_1")
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim rsetCFL As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            oCond = oConds.Add
            oCond.Alias = "Active"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = "Y"
            oCFL.SetConditions(oConds)
        Catch ex As Exception

        End Try
    End Sub

    Sub ChooseFromListFilteration(ByVal FormUID As String, ByVal oForm As SAPbouiCOM.Form, ByVal strCFL_ID As String, ByVal strCFL_Alies As String, ByVal strQuery As String)
        Try
            Dim oCFL As SAPbouiCOM.ChooseFromList = oForm.ChooseFromLists.Item(strCFL_ID)
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim rsetCFL As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            Dim BranchID As String = "", LocCode As String = ""
            If strQuery <> "" And strCFL_Alies <> "" Then
                If objForm.Items.Item("15B").Specific.String = "" Then
                    oCond = oConds.Add
                    oCond.Alias = "DocEntry"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = ""
                    oCFL.SetConditions(oConds)
                    Exit Sub
                End If
                rsetCFL.DoQuery(strQuery)
                rsetCFL.MoveFirst()
                oCond = oConds.Add
                oCond.Alias = "JrnlMemo"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Receipt from Production"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                If rsetCFL.RecordCount > 0 Then
                    For i As Integer = 1 To rsetCFL.RecordCount
                        If i = (rsetCFL.RecordCount) Then
                            oCond = oConds.Add()
                            oCond.Alias = strCFL_Alies  'DocEntry
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                        Else
                            oCond = oConds.Add()
                            oCond.Alias = strCFL_Alies
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        End If
                        rsetCFL.MoveNext()
                    Next
                Else
                    oCond = oConds.Add()
                    oCond.Alias = "DocEntry"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = ""
                End If
            End If
            If strQuery = "" And strCFL_Alies = "" Then
                oCond = oConds.Add
                oCond.Alias = "DocStatus"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "O"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCond = oConds.Add
                oCond.Alias = "CANCELED"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "N"
                '    If BranchFlag = "Y" Then
                '        objCombo = objForm.Items.Item("50").Specific
                '        If Not objCombo.Selected.Value = "-1" Then
                '            BranchID = objCombo.Selected.Value
                '        End If
                '    End If
                '    objCombo = objForm.Items.Item("10").Specific
                '    If Not objCombo.Selected.Value = "-" Then
                '        LocCode = objCombo.Selected.Value
                '    End If
                '    If objAddOn.HANA Then
                '        strQuery = "Select distinct A.""DocEntry"",A.""DocNum"",A.""ItemCode"" from (SELECT distinct T5.""DocEntry"",T5.""DocNum"",T2.""ItemCode"",T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1  INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" "
                '        strQuery += vbCrLf + "AND  T0.""U_GRNum"" = cast(T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""PendQty"" FROM OIGN T5 inner join IGN1 T2 on T5.""DocEntry""=T2.""DocEntry"" INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" "
                '        strQuery += vbCrLf + "left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" where T5.""DocStatus""='O' and T2.""BaseType"" <>'202' and T5.""CANCELED""='N' and T5.""U_QCReq""='Y'  and T3.""U_InspReq""='Y' "
                '        If BranchID <> "" Then
                '            strQuery += vbCrLf + "and T5.""BPLId""='" & BranchID & "' "
                '        End If
                '        If LocCode <> "" Then
                '            strQuery += vbCrLf + "and T2.""LocCode""='" & LocCode & "' "
                '        End If
                '        strQuery += vbCrLf + "GROUP BY T5.""DocEntry"",T5.""DocNum"",T2.""ItemCode"",T4.""BaseQty"",T4.""AltQty"") as A where A.""PendQty"">0"
                '    Else
                '        strQuery = "Select distinct A.DocEntry,A.DocNum,A.ItemCode from (SELECT distinct T5.DocEntry,T5.DocNum,T2.ItemCode,T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (SELECT ISNULL(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1  INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry "
                '        strQuery += vbCrLf + "AND  T0.U_GRNum = cast(T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum) AS PendQty FROM OIGN T5 inner join IGN1 T2 on T5.DocEntry=T2.DocEntry INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode "
                '        strQuery += vbCrLf + "left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T5.DocStatus='O' and T2.BaseType <>'202' and T5.CANCELED='N' and T5.U_QCReq='Y'  and T3.U_InspReq='Y' "
                '        If BranchID <> "" Then
                '            strQuery += vbCrLf + "and T5.BPLId='" & BranchID & "' "
                '        End If
                '        If LocCode <> "" Then
                '            strQuery += vbCrLf + "and T2.LocCode='" & LocCode & "' "
                '        End If
                '        strQuery += vbCrLf + "GROUP BY T5.DocEntry,T5.DocNum, T2.ItemCode,T4.BaseQty,T4.AltQty) as A where A.PendQty>0"
                '    End If
                '    rsetCFL.DoQuery(strQuery)
                '    If rsetCFL.RecordCount > 0 Then
                '        For i As Integer = 1 To rsetCFL.RecordCount
                '            If i = (rsetCFL.RecordCount) Then
                '                oCond = oConds.Add()
                '                oCond.Alias = "DocEntry"
                '                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                '                oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                '            Else
                '                oCond = oConds.Add()
                '                oCond.Alias = "DocEntry"
                '                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                '                oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                '                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                '            End If
                '            rsetCFL.MoveNext()
                '        Next
                '    Else
                '        oCond = oConds.Add()
                '        oCond.Alias = "DocEntry"
                '        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                '        oCond.CondVal = ""
                '    End If
                'End If
                'If BranchFlag = "Y" Then
                '    objCombo = objForm.Items.Item("50").Specific
                '    If Not objCombo.Selected.Value = "-1" Then
                '        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                '        oCond = oConds.Add
                '        oCond.Alias = "BPLId"
                '        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                '        oCond.CondVal = objCombo.Selected.Value
                '    End If
            End If
            'If oConds.Count > 0 Then
            '    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            'End If
            'If objAddOn.HANA Then
            '    strSQL = objAddOn.objGenFunc.getSingleValue("select ""Indicator"" from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between ""F_RefDate"" and ""T_RefDate""")
            'Else
            '    strSQL = objAddOn.objGenFunc.getSingleValue("select Indicator from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between F_RefDate and T_RefDate")
            'End If
            'oCond = oConds.Add
            'oCond.Alias = "PIndicator"
            'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCond.CondVal = strSQL


            oCFL.SetConditions(oConds)
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Choose FromList Filter Global Fun. Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub ChooseFromListFilterationINV(ByVal oForm As SAPbouiCOM.Form, ByVal strCFL_ID As String)
        Try
            Dim oCFL As SAPbouiCOM.ChooseFromList = oForm.ChooseFromLists.Item(strCFL_ID)
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim BranchID As String = "", LocCode As String = ""
            Dim rsetCFL As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            oCond = oConds.Add
            oCond.Alias = "DocStatus"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = "O"
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCond = oConds.Add
            oCond.Alias = "CANCELED"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = "N"
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCond = oConds.Add
            oCond.BracketOpenNum = 1
            oCond.Alias = "JrnlMemo"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_CONTAIN
            oCond.CondVal = "Auto Generated"
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            oCond = oConds.Add
            oCond.Alias = "JrnlMemo"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN
            oCond.CondVal = "Auto-Gen-> QC Action"
            oCond.BracketCloseNum = 1


            ''If objAddOn.HANA Then
            ''    strSQL = objAddOn.objGenFunc.getSingleValue("select ""Indicator"" from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between ""F_RefDate"" and ""T_RefDate""")
            ''Else
            ''    strSQL = objAddOn.objGenFunc.getSingleValue("select Indicator from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between F_RefDate and T_RefDate")
            ''End If
            ''oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            ''oCond = oConds.Add
            ''oCond.Alias = "PIndicator"
            ''oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            ''oCond.CondVal = strSQL
            'If BranchFlag = "Y" Then
            '    objCombo = objForm.Items.Item("50").Specific
            '    If Not objCombo.Selected.Value = "-1" Then
            '        BranchID = objCombo.Selected.Value
            '    End If
            'End If
            'objCombo = objForm.Items.Item("10").Specific
            'If Not objCombo.Selected.Value = "-" Then
            '    LocCode = objCombo.Selected.Value
            'End If
            'If objAddOn.HANA Then
            '    strQuery = "Select distinct A.""DocEntry"",A.""DocNum"",A.""ItemCode"" from (SELECT distinct T5.""DocEntry"",T5.""DocNum"",T2.""ItemCode"", T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1  INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" "
            '    strQuery += vbCrLf + "AND  T0.""U_TransEntry"" = cast(T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""PendQty"" FROM OWTR T5 inner join WTR1 T2 on T5.""DocEntry""=T2.""DocEntry"" INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" "
            '    strQuery += vbCrLf + "left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" where T5.""DocStatus""='O' and T5.""CANCELED""='N' and T5.""U_QCReq""='Y' and ifnull(T5.""U_QCNum"",'')='' and T3.""U_InspReq""='Y' "
            '    If BranchID <> "" Then
            '        strQuery += vbCrLf + "and T5.""BPLId""='" & BranchID & "' "
            '    End If
            '    If LocCode <> "" Then
            '        strQuery += vbCrLf + "and T2.""LocCode""='" & LocCode & "' "
            '    End If
            '    strQuery += vbCrLf + "GROUP BY T5.""DocEntry"",T5.""DocNum"", T2.""ItemCode"",T4.""BaseQty"",T4.""AltQty"") as A where A.""PendQty"">0"
            'Else
            '    strQuery = "Select distinct A.DocEntry,A.DocNum,T2.ItemCode from (SELECT distinct T5.DocEntry,T5.DocNum,T2.ItemCode,T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (SELECT ISNULL(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1  INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry "
            '    strQuery += vbCrLf + "AND  T0.U_TransEntry = cast(T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum) AS PendQty FROM OWTR T5 inner join WTR1 T2 on T5.DocEntry=T2.DocEntry INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode "
            '    strQuery += vbCrLf + "left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T5.DocStatus='O' and T5.CANCELED='N' and T5.U_QCReq='Y' and isnull(T5.""U_QCNum"",'')='' and T3.U_InspReq='Y' "
            '    If BranchID <> "" Then
            '        strQuery += vbCrLf + "and T5.BPLId='" & BranchID & "' "
            '    End If
            '    If LocCode <> "" Then
            '        strQuery += vbCrLf + "and T2.LocCode='" & LocCode & "' "
            '    End If
            '    strQuery += vbCrLf + "GROUP BY T5.DocEntry,T5.DocNum,T2.ItemCode,T4.BaseQty,T4.AltQty) as A where A.PendQty>0"
            'End If
            'rsetCFL.DoQuery(strQuery)
            'If rsetCFL.RecordCount > 0 Then
            '    For i As Integer = 1 To rsetCFL.RecordCount
            '        If i = (rsetCFL.RecordCount) Then
            '            oCond = oConds.Add()
            '            oCond.Alias = "DocEntry"
            '            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
            '        Else
            '            oCond = oConds.Add()
            '            oCond.Alias = "DocEntry"
            '            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
            '            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            '        End If
            '        rsetCFL.MoveNext()
            '    Next
            'Else
            '    oCond = oConds.Add()
            '    oCond.Alias = "DocEntry"
            '    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '    oCond.CondVal = ""
            'End If
            'If BranchFlag = "Y" Then
            '    objCombo = objForm.Items.Item("50").Specific
            '    If Not objCombo.Selected.Value = "-1" Then
            '        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            '        oCond = oConds.Add
            '        oCond.Alias = "BPLId"
            '        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '        oCond.CondVal = objCombo.Selected.Value
            '    End If
            'End If
            oCFL.SetConditions(oConds)
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText("Choose FromList Filter Global Fun. Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Private Sub AddCFLConditionGRPO(ByVal FormUID As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim oCFL As SAPbouiCOM.ChooseFromList = objForm.ChooseFromLists.Item("CFL_GRPO")
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            'Dim objCombo As SAPbouiCOM.ComboBox
            Dim StrQuery As String = ""
            Dim BranchID As String = "", LocCode As String = ""
            Dim rsetCFL As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            oCond = oConds.Add
            oCond.Alias = "DocStatus"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = "O"
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCond = oConds.Add
            oCond.Alias = "CANCELED"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = "N"
            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCond = oConds.Add
            oCond.Alias = "U_QCReq"
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.CondVal = "Y"
            'If objAddOn.HANA Then
            '    strSQL = objAddOn.objGenFunc.getSingleValue("select ""Indicator"" from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between ""F_RefDate"" and ""T_RefDate""")
            'Else
            '    strSQL = objAddOn.objGenFunc.getSingleValue("select Indicator from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between F_RefDate and T_RefDate")
            'End If
            'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            'oCond = oConds.Add
            'oCond.Alias = "PIndicator"
            'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCond.CondVal = strSQL
            objCombo = objForm.Items.Item("50").Specific
            If Not objCombo.Selected.Value = "-1" Then
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCond = oConds.Add
                oCond.Alias = "BPLId"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = objCombo.Selected.Value
            End If
            'If BranchFlag = "Y" Then
            '    objCombo = objForm.Items.Item("50").Specific
            '    If Not objCombo.Selected.Value = "-1" Then
            '        BranchID = objCombo.Selected.Value
            '    End If
            'End If
            'objCombo = objForm.Items.Item("10").Specific
            'If Not objCombo.Selected.Value = "-" Then
            '    LocCode = objCombo.Selected.Value
            'End If
            'If objAddOn.HANA Then
            '    StrQuery = "Select distinct A.""DocEntry"",A.""DocNum"",A.""ItemCode"" from (SELECT distinct T5.""DocEntry"",T5.""DocNum"",T2.""ItemCode"",T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1  INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" "
            '    StrQuery += vbCrLf + "AND  T0.""U_GRNEntry"" = cast(T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""PendQty"" FROM OPDN T5 inner join PDN1 T2 on T5.""DocEntry""=T2.""DocEntry"" INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" "
            '    StrQuery += vbCrLf + "left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" where T5.""DocStatus""='O' and T5.""CANCELED""='N' and T5.""U_QCReq""='Y' and T3.""U_InspReq""='Y' "
            '    If BranchID <> "" Then
            '        StrQuery += vbCrLf + "and T5.""BPLId""='" & BranchID & "' "
            '    End If
            '    If LocCode <> "" Then
            '        StrQuery += vbCrLf + "and T2.""LocCode""='" & LocCode & "' "
            '    End If
            '    StrQuery += vbCrLf + "GROUP BY T5.""DocEntry"",T5.""DocNum"",T2.""ItemCode"", T4.""BaseQty"",T4.""AltQty"") as A where A.""PendQty"">0"
            'Else
            '    StrQuery = "Select distinct A.DocEntry,A.DocNum,A.ItemCode from (SELECT distinct T5.DocEntry,T5.DocNum,T2.ItemCode, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (SELECT ISNULL(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1  INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry "
            '    StrQuery += vbCrLf + "AND  T0.U_GRNEntry = cast(T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum) AS PendQty FROM OPDN T5 inner join PDN1 T2 on T5.DocEntry=T2.DocEntry INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode "
            '    StrQuery += vbCrLf + "left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T5.DocStatus='O' and T5.CANCELED='N' and T5.U_QCReq='Y' and T3.U_InspReq='Y' "
            '    If BranchID <> "" Then
            '        StrQuery += vbCrLf + "and T5.BPLId='" & BranchID & "' "
            '    End If
            '    If LocCode <> "" Then
            '        StrQuery += vbCrLf + "and T2.LocCode='" & LocCode & "' "
            '    End If
            '    StrQuery += vbCrLf + "GROUP BY T5.DocEntry,T5.DocNum,T2.ItemCode, T4.BaseQty,T4.AltQty) as A where A.PendQty>0"
            'End If
            'rsetCFL.DoQuery(StrQuery)
            'QCHeader.Query(oConds)
            'If rsetCFL.RecordCount > 0 Then
            '    For i As Integer = 1 To rsetCFL.RecordCount
            '        If i = (rsetCFL.RecordCount) Then
            '            oCond = oConds.Add()
            '            oCond.Alias = "DocEntry"
            '            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
            '        Else
            '            oCond = oConds.Add()
            '            oCond.Alias = "DocEntry"
            '            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
            '            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            '        End If
            '        rsetCFL.MoveNext()
            '    Next
            'Else
            '    oCond = oConds.Add()
            '    oCond.Alias = "DocEntry"
            '    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '    oCond.CondVal = ""
            'End If
            'If BranchFlag = "Y" Then
            '    objCombo = objForm.Items.Item("50").Specific
            '    If Not objCombo.Selected.Value = "-1" Then
            '        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            '        oCond = oConds.Add
            '        oCond.Alias = "BPLId"
            '        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '        oCond.CondVal = objCombo.Selected.Value
            '    End If
            'End If
            oCFL.SetConditions(oConds)
        Catch ex As Exception

        End Try

    End Sub

    Public Function GetListQuery(ByVal FormUID As String, ByVal DocName As String, ByVal LikeText As String, ByVal LikeFieldName As String)
        Try
            Dim BranchID As String = "", LocCode As String = ""
            objForm = objAddOn.objApplication.Forms.Item(FormUID) ' objAddOn.objApplication.Forms.GetForm("MIPLQC", FormUID)
            If BranchFlag = "Y" Then
                objCombo = objForm.Items.Item("50").Specific
                If Not objCombo.Selected.Value = "-1" Then
                    BranchID = objCombo.Selected.Value
                End If
            End If
            objCombo = objForm.Items.Item("10").Specific
            If Not objCombo.Selected.Value = "-" Then
                LocCode = objCombo.Selected.Value
            End If
            Dim ObjCFLForm As SAPbouiCOM.Form
            ObjCFLForm = objAddOn.objApplication.Forms.GetForm("OCFL", 1)
            Dim InvExceptionInWhse As String = ""
            If objAddOn.HANA Then
                InvExceptionInWhse = objAddOn.objGenFunc.getSingleValue("Select ""U_Whse"" from OADM")
            Else
                InvExceptionInWhse = objAddOn.objGenFunc.getSingleValue("Select U_Whse from OADM")
            End If

            If DocName = "G" Then
                ObjCFLForm.Title = "GRPO Entries"
                If objAddOn.HANA Then
                    strQuery = "Select ROW_NUMBER() OVER (order by A.""DocEntry"" desc) as ""LineId"",* from (SELECT T5.""DocNum"",T5.""DocEntry"",T5.""CardCode"",T2.""ItemCode"",T2.""Dscription"",TO_VARCHAR(T5.""DocDate"",'dd/MM/yy') as ""DocDate"",TO_VARCHAR(T5.""DocDueDate"",'dd/MM/yy') as ""DocDueDate"","
                    strQuery += vbCrLf + "(Select ""BPLName"" from OBPL where ""BPLId""=T5.""BPLId"") as ""BPLId"",T2.""WhsCode"",(Select ""Location"" from OLCT where ""Code""=T2.""LocCode"")as ""LocCode"",T5.""Comments"" as ""Remarks"","
                    strQuery += vbCrLf + "T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1  INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" "
                    strQuery += vbCrLf + "AND  T0.""U_GRNEntry"" = cast(T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""PendQty"" "
                    strQuery += vbCrLf + "FROM OPDN T5 inner join PDN1 T2 on T5.""DocEntry""=T2.""DocEntry"" INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" "
                    strQuery += vbCrLf + "left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" where T5.""DocStatus""='O' and T5.""CANCELED""='N' and T5.""U_QCReq""='Y'  "
                    'and T3.""U_InspReq""='Y'
                    If InvExceptionInWhse <> "" Then
                        strQuery += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
                    Else
                        strQuery += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
                    End If
                    If BranchID <> "" Then
                        strQuery += vbCrLf + "and T5.""BPLId""='" & BranchID & "' "
                    End If
                    If LocCode <> "" Then
                        strQuery += vbCrLf + "and T2.""LocCode""='" & LocCode & "' "
                    End If
                    strQuery += vbCrLf + "GROUP BY T5.""DocEntry"",T5.""Comments"",T5.""DocNum"",T5.""CardCode"",T2.""ItemCode"",T2.""Dscription"",T5.""DocDate"",T5.""DocDueDate"",T5.""BPLId"",T2.""WhsCode"",T2.""LocCode"","
                    strQuery += vbCrLf + "T2.""ItemCode"", T2.""Dscription"",T4.""BaseQty"",T4.""AltQty"" Order by T5.""DocDate"" desc)as A where A.""PendQty"">0"
                    If LikeText <> "" Then
                        'strQuery += vbCrLf + "and A.""" & LikeFieldName & """ like '" & LikeText & "%' "
                        strQuery += vbCrLf + "and (UPPER(A.""" & LikeFieldName & """) like '" & LikeText & "%' or UPPER(A.""" & LikeFieldName & """) like '%" & LikeText & "%' or UPPER(A.""" & LikeFieldName & """) like '" & LikeText & "') "
                    End If
                Else
                    strQuery = "Select ROW_NUMBER() OVER (order by A.DocEntry desc) as lineId,* from (SELECT T5.DocNum,T5.DocEntry,(Select CardName from OCRD where CardCode=T5.CardCode) as CardName,T2.ItemCode,T2.Dscription,T5.DocDate as DocDate,Format(T5.DocDueDate,'dd/MM/yy') as DocDueDate,"
                    strQuery += vbCrLf + "(Select BPLName from OBPL where BPLId=T5.BPLId) as BPLId,T2.WhsCode,(Select Location from OLCT where Code=T2.LocCode)as LocCode,T5.Comments as Remarks,"
                    strQuery += vbCrLf + "T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (SELECT isnull(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1  INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry "
                    strQuery += vbCrLf + "AND  T0.U_GRNEntry = cast(T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum) AS PendQty "
                    strQuery += vbCrLf + "FROM OPDN T5 inner join PDN1 T2 on T5.DocEntry=T2.DocEntry INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode "
                    strQuery += vbCrLf + "left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T5.DocStatus='O' and T5.CANCELED='N' and T5.U_QCReq='Y' "
                    'And T3.U_InspReq='Y' "
                    If InvExceptionInWhse <> "" Then
                        strQuery += vbCrLf + " AND (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
                    Else
                        strQuery += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
                    End If
                    If BranchID <> "" Then
                        strQuery += vbCrLf + "and T5.BPLId='" & BranchID & "' "
                    End If
                    If LocCode <> "" Then
                        strQuery += vbCrLf + "and T2.LocCode='" & LocCode & "' "
                    End If
                    strQuery += vbCrLf + "GROUP BY T5.DocEntry,T2.DocEntry,T2.LineNum,T5.Comments,T5.DocNum,T5.CardCode,T5.DocDate,T5.DocDueDate,T5.BPLId,T2.WhsCode,T2.LocCode,"
                    strQuery += vbCrLf + "T2.ItemCode,T2.Dscription, T4.BaseQty,T4.AltQty " 'Order by T5.DocDate desc"
                    strQuery += vbCrLf + ") as A where A.PendQty>0"
                    If LikeText <> "" Then
                        'strQuery += vbCrLf + "and A." & LikeFieldName & " like '" & LikeText & "%' "
                        strQuery += vbCrLf + "and (UPPER(A." & LikeFieldName & ") like '" & LikeText & "%' or UPPER(A." & LikeFieldName & ") like '%" & LikeText & "%' or UPPER(A." & LikeFieldName & ") like '" & LikeText & "') "
                    End If
                    strQuery += vbCrLf + "Order by A.DocDate desc"
                End If
            ElseIf DocName = "P" Then
                ObjCFLForm.Title = "Production Entries"
                If objAddOn.HANA Then
                    strQuery = "Select ROW_NUMBER() OVER (order by A.""DocEntry"" desc) as ""LineId"",* from (SELECT (Select ""DocNum"" from OWOR where ""DocEntry""=T2.""BaseEntry"") as ""DocNum"",T2.""BaseEntry"" as ""DocEntry"",T5.""CardCode"",T2.""ItemCode"",T2.""Dscription"",TO_VARCHAR(T5.""DocDate"",'dd/MM/yy') as ""DocDate"",TO_VARCHAR(T5.""DocDueDate"",'dd/MM/yy') as ""DocDueDate"","
                    strQuery += vbCrLf + "(Select ""BPLName"" from OBPL where ""BPLId""=T5.""BPLId"") as ""BPLId"",T2.""WhsCode"",(Select ""Location"" from OLCT where ""Code""=T2.""LocCode"")as ""LocCode"",T5.""Comments"" as ""Remarks"","
                    strQuery += vbCrLf + "T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1  INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" "
                    strQuery += vbCrLf + "AND  T0.""U_GRNum"" =cast(T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""PendQty"" "
                    strQuery += vbCrLf + "FROM OIGN T5 inner join IGN1 T2 on T5.""DocEntry""=T2.""DocEntry"" INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" "
                    strQuery += vbCrLf + "left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry""  left outer join OWOR T6 on T6.""DocEntry""=T2.""BaseEntry"" where T5.""DocStatus""='O' and T2.""BaseType"" ='202' and T5.""CANCELED""='N' and T5.""U_QCReq""='Y' and T6.""Status"" not in ('C','L')   "
                    'and T3.""U_InspReq""='Y' 
                    If InvExceptionInWhse <> "" Then
                        strQuery += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
                    Else
                        strQuery += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
                    End If
                    If BranchID <> "" Then
                        strQuery += vbCrLf + "and T5.""BPLId""='" & BranchID & "' "
                    End If
                    If LocCode <> "" Then
                        strQuery += vbCrLf + "and T2.""LocCode""='" & LocCode & "' "
                    End If
                    strQuery += vbCrLf + "GROUP BY T2.""BaseEntry"",T5.""Comments"",T5.""DocNum"",T5.""CardCode"",T5.""DocDate"",T5.""DocDueDate"",T5.""BPLId"",T2.""WhsCode"",T2.""LocCode"",T6.""Status"","
                    strQuery += vbCrLf + "T2.""ItemCode"", T2.""Dscription"",T4.""BaseQty"",T4.""AltQty"" Order by T5.""DocDate"" desc )as A where A.""PendQty"">0"
                    If LikeText <> "" Then
                        'strQuery += vbCrLf + "and A.""" & LikeFieldName & """ like '" & LikeText & "%' "
                        strQuery += vbCrLf + "and (UPPER(A.""" & LikeFieldName & """) like '" & LikeText & "%' or UPPER(A.""" & LikeFieldName & """) like '%" & LikeText & "%' or UPPER(A.""" & LikeFieldName & """) like '" & LikeText & "') "
                    End If
                Else
                    strQuery = "Select ROW_NUMBER() OVER (order by A.DocEntry desc) as lineId,* from (SELECT (Select DocNum from OWOR where DocEntry=T2.BaseEntry) as DocNum,T2.BaseEntry as DocEntry,case when T5.U_BpMaster is null then (Select CardName from OCRD where CardCode=T2.U_CardCode) else T5.U_BpMaster end as CardName,T2.ItemCode,T2.Dscription,T5.DocDate as DocDate,Format(T5.DocDueDate,'dd/MM/yy') as DocDueDate,"
                    strQuery += vbCrLf + "(Select BPLName from OBPL where BPLId=T5.BPLId) as BPLId,T2.WhsCode,(Select Location from OLCT where Code=T2.LocCode)as LocCode,T5.Comments as Remarks,"
                    strQuery += vbCrLf + "T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (SELECT isnull(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1  INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry "
                    strQuery += vbCrLf + "AND  T0.U_GRNum =cast(T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum) AS PendQty "
                    strQuery += vbCrLf + "FROM OIGN T5 inner join IGN1 T2 on T5.DocEntry=T2.DocEntry INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode "
                    strQuery += vbCrLf + "left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry left outer join OWOR T6 on T6.DocEntry=T2.BaseEntry where T5.DocStatus='O' and T2.BaseType ='202' and T5.CANCELED='N' and T5.U_QCReq='Y' and T6.Status not in ('C','L') "
                    'and T3.U_InspReq='Y'
                    If InvExceptionInWhse <> "" Then
                        strQuery += vbCrLf + " AND (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
                    Else
                        strQuery += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
                    End If
                    If BranchID <> "" Then
                        strQuery += vbCrLf + "and T5.BPLId='" & BranchID & "' "
                    End If
                    If LocCode <> "" Then
                        strQuery += vbCrLf + "and T2.LocCode='" & LocCode & "' "
                    End If
                    strQuery += vbCrLf + "GROUP BY T2.BaseEntry,T2.DocEntry,T2.LineNum,T5.Comments,T5.DocNum,T5.CardCode,T5.DocDate,T5.DocDueDate,T5.BPLId,T2.WhsCode,T2.LocCode,T6.Status,T5.U_BpMaster,T2.U_CardCode,"
                    strQuery += vbCrLf + "T2.ItemCode,T2.Dscription, T4.BaseQty,T4.AltQty" ' Order by T5.DocDate desc" 
                    strQuery += vbCrLf + " ) as A where A.PendQty>0"
                    If LikeText <> "" Then
                        'strQuery += vbCrLf + "and A." & LikeFieldName & " like '" & LikeText & "%' "
                        strQuery += vbCrLf + "and (UPPER(A." & LikeFieldName & ") like '" & LikeText & "%' or UPPER(A." & LikeFieldName & ") like '%" & LikeText & "%' or UPPER(A." & LikeFieldName & ") like '" & LikeText & "') "
                    End If
                    strQuery += vbCrLf + "Order by A.DocDate desc"
                End If

            ElseIf DocName = "T" Then
                ObjCFLForm.Title = "Inventory Transfer Entries"
                If objAddOn.HANA Then
                    strQuery = "Select ROW_NUMBER() OVER (order by A.""DocEntry"" desc) as ""LineId"",* from (SELECT T5.""DocNum"",T5.""DocEntry"",T5.""CardCode"",T2.""ItemCode"",T2.""Dscription"",TO_VARCHAR(T5.""DocDate"",'dd/MM/yy') as ""DocDate"",TO_VARCHAR(T5.""DocDueDate"",'dd/MM/yy') as ""DocDueDate"","
                    strQuery += vbCrLf + "(Select ""BPLName"" from OBPL where ""BPLId""=T5.""BPLId"") as ""BPLId"",T2.""WhsCode"",(Select ""Location"" from OLCT where ""Code""=T2.""LocCode"")as ""LocCode"",T5.""Comments"" as ""Remarks"","
                    strQuery += vbCrLf + "T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1  INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" "
                    strQuery += vbCrLf + "AND  T0.""U_TransEntry"" =cast(T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""PendQty"" "
                    strQuery += vbCrLf + "FROM OWTR T5 inner join WTR1 T2 on T5.""DocEntry""=T2.""DocEntry"" INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" "
                    strQuery += vbCrLf + "left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" where T5.""DocStatus""='O' and T5.""CANCELED""='N' and T5.""U_QCReq""='Y' and ifnull(T5.""U_QCNum"",'')=''  "
                    If InvExceptionInWhse <> "" Then
                        strQuery += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
                    Else
                        strQuery += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
                    End If
                    If BranchID <> "" Then
                        strQuery += vbCrLf + "and T5.""BPLId""='" & BranchID & "' "
                    End If
                    If LocCode <> "" Then
                        strQuery += vbCrLf + "and T2.""LocCode""='" & LocCode & "' "
                    End If
                    strQuery += vbCrLf + "GROUP BY T5.""DocEntry"",T5.""Comments"",T5.""DocNum"",T5.""CardCode"",T5.""DocDate"",T5.""DocDueDate"",T5.""BPLId"",T2.""WhsCode"",T2.""LocCode"","
                    strQuery += vbCrLf + "T2.""ItemCode"",T2.""Dscription"", T4.""BaseQty"",T4.""AltQty"" Order by T5.""DocDate"" desc )as A where A.""PendQty"">0"
                    If LikeText <> "" Then
                        'strQuery += vbCrLf + "and A.""" & LikeFieldName & """ like '" & LikeText & "%' "
                        strQuery += vbCrLf + "and (UPPER(A.""" & LikeFieldName & """) like '" & LikeText & "%' or UPPER(A.""" & LikeFieldName & """) like '%" & LikeText & "%' or UPPER(A.""" & LikeFieldName & """) like '" & LikeText & "') "
                    End If
                Else
                    strQuery = "Select ROW_NUMBER() OVER (order by A.DocEntry desc) as lineId,* from (SELECT T5.DocNum,T5.DocEntry,(Select CardName from OCRD where CardCode=T5.CardCode) as CardName,T2.ItemCode,T2.Dscription,T5.DocDate as DocDate,Format(T5.DocDueDate,'dd/MM/yy') as DocDueDate,"
                    strQuery += vbCrLf + "(Select BPLName from OBPL where BPLId=T5.BPLId) as BPLId,T2.WhsCode,(Select Location from OLCT where Code=T2.LocCode)as LocCode,T5.Comments as Remarks,"
                    strQuery += vbCrLf + "T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (SELECT isnull(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1  INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry "
                    strQuery += vbCrLf + "AND  T0.U_TransEntry =cast(T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum) AS PendQty "
                    strQuery += vbCrLf + "FROM OWTR T5 inner join WTR1 T2 on T5.DocEntry=T2.DocEntry INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode "
                    strQuery += vbCrLf + "left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T5.DocStatus='O' and T5.CANCELED='N' and T5.U_QCReq='Y' and isnull(T5.U_QCNum,'')=''  "
                    'and T3.U_InspReq='Y'
                    If InvExceptionInWhse <> "" Then
                        strQuery += vbCrLf + " AND (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
                    Else
                        strQuery += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
                    End If
                    If BranchID <> "" Then
                        strQuery += vbCrLf + "and T5.BPLId='" & BranchID & "' "
                    End If
                    If LocCode <> "" Then
                        strQuery += vbCrLf + "and T2.LocCode='" & LocCode & "' "
                    End If
                    strQuery += vbCrLf + "GROUP BY T5.DocEntry,T2.DocEntry,T2.LineNum,T5.Comments,T5.DocNum,T5.CardCode,T5.DocDate,T5.DocDueDate,T5.BPLId,T2.WhsCode,T2.LocCode,"
                    strQuery += vbCrLf + "T2.ItemCode,T2.Dscription, T4.BaseQty,T4.AltQty" ' Order by T5.DocDate desc"
                    strQuery += vbCrLf + ") as A where A.PendQty>0"
                    If LikeText <> "" Then
                        'strQuery += vbCrLf + "and A." & LikeFieldName & " like '" & LikeText & "%' "
                        strQuery += vbCrLf + "and (UPPER(A." & LikeFieldName & ") like '" & LikeText & "%' or UPPER(A." & LikeFieldName & ") like '%" & LikeText & "%' or UPPER(A." & LikeFieldName & ") like '" & LikeText & "') "
                    End If
                    strQuery += vbCrLf + "Order by A.DocDate desc"
                End If
            ElseIf DocName = "R" Then
                ObjCFLForm.Title = "Receipt Entries"
                objCombo = objForm.Items.Item("8").Specific
                If objCombo.Selected.Value = "P" Then
                    ObjCFLForm.Title = "Receipt From Production Entries"
                    If objForm.Items.Item("15B").Specific.String = "" Then
                        If objAddOn.HANA Then
                            strQuery = "Select '' from dummy"
                        Else
                            strQuery = "Select '' "
                        End If
                    Else
                        If objAddOn.HANA Then
                            strQuery = "SELECT ROW_NUMBER() OVER (order by T5.""DocEntry"" desc) as ""LineId"",T5.""DocNum"",T5.""DocEntry"",T5.""CardCode"",T2.""ItemCode"",T2.""Dscription"",TO_VARCHAR(T5.""DocDate"",'dd/MM/yy') as ""DocDate"",TO_VARCHAR(T5.""DocDueDate"",'dd/MM/yy') as ""DocDueDate"","
                            strQuery += vbCrLf + "(Select ""BPLName"" from OBPL where ""BPLId""=T5.""BPLId"") as ""BPLId"",T2.""WhsCode"",(Select ""Location"" from OLCT where ""Code""=T2.""LocCode"")as ""LocCode"",T5.""Comments"" as ""Remarks"""
                            strQuery += vbCrLf + "FROM OIGN T5 inner join IGN1 T2 on T5.""DocEntry""=T2.""DocEntry"" INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" "
                            strQuery += vbCrLf + "left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" "
                            strQuery += vbCrLf + "where T5.""DocStatus""='O' and T5.""JrnlMemo""='Receipt from Production' and T2.""BaseType"" ='202' "
                            strQuery += vbCrLf + "and T2.""BaseEntry""='" & objForm.Items.Item("15B").Specific.String & "' and T5.""CANCELED""='N' and T5.""U_QCReq""='Y'   "
                            'and T3.""U_InspReq""='Y'
                            If InvExceptionInWhse <> "" Then
                                strQuery += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
                            Else
                                strQuery += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
                            End If
                        Else
                            strQuery = "SELECT ROW_NUMBER() OVER (Order by T5.DocEntry desc) AS LineId,T5.DocNum,T5.DocEntry,case when T5.U_BpMaster is null then (Select CardName from OCRD where CardCode=T2.U_CardCode) else T5.U_BpMaster end as CardName,T2.ItemCode,T2.Dscription,Format(T5.DocDate,'dd/MM/yy') as DocDate,Format(T5.DocDueDate,'dd/MM/yy') as DocDueDate,"
                            strQuery += vbCrLf + "(Select BPLName from OBPL where BPLId=T5.BPLId) as BPLId,T2.WhsCode,(Select Location from OLCT where Code=T2.LocCode)as LocCode,T5.Comments as Remarks"
                            strQuery += vbCrLf + "FROM OIGN T5 inner join IGN1 T2 on T5.DocEntry=T2.DocEntry INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode "
                            strQuery += vbCrLf + "left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry "
                            strQuery += vbCrLf + "where T5.DocStatus='O' and T5.JrnlMemo='Receipt from Production' and T2.BaseType ='202' "
                            strQuery += vbCrLf + "and T2.BaseEntry='" & objForm.Items.Item("15B").Specific.String & "' and T5.CANCELED='N' and T5.U_QCReq='Y'   "
                            'and T3.U_InspReq='Y'
                            If InvExceptionInWhse <> "" Then
                                strQuery += vbCrLf + " AND (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
                            Else
                                strQuery += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
                            End If
                        End If
                    End If
                Else
                    If objAddOn.HANA Then
                        strQuery = "Select ROW_NUMBER() OVER (order by A.""DocEntry"" desc) as ""LineId"",* from (SELECT T5.""DocNum"",T5.""DocEntry"",T5.""CardCode"",T2.""ItemCode"",T2.""Dscription"",TO_VARCHAR(T5.""DocDate"",'dd/MM/yy') as ""DocDate"",TO_VARCHAR(T5.""DocDueDate"",'dd/MM/yy') as ""DocDueDate"","
                        strQuery += vbCrLf + "(Select ""BPLName"" from OBPL where ""BPLId""=T5.""BPLId"") as ""BPLId"",T2.""WhsCode"",(Select ""Location"" from OLCT where ""Code""=T2.""LocCode"")as ""LocCode"",T5.""Comments"" as ""Remarks"","
                        strQuery += vbCrLf + "T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1  INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" "
                        strQuery += vbCrLf + "AND  T0.""U_GRNum"" =cast(T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""PendQty"" "
                        strQuery += vbCrLf + "FROM OIGN T5 inner join IGN1 T2 on T5.""DocEntry""=T2.""DocEntry"" INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" "
                        strQuery += vbCrLf + "left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" where T5.""DocStatus""='O' and T2.""BaseType"" <>'202' and T5.""CANCELED""='N' and T5.""U_QCReq""='Y'  "
                        'and T3.""U_InspReq""='Y' 
                        If InvExceptionInWhse <> "" Then
                            strQuery += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
                        Else
                            strQuery += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
                        End If
                        If BranchID <> "" Then
                            strQuery += vbCrLf + "and T5.""BPLId""='" & BranchID & "' "
                        End If
                        If LocCode <> "" Then
                            strQuery += vbCrLf + "and T2.""LocCode""='" & LocCode & "' "
                        End If
                        strQuery += vbCrLf + "GROUP BY T5.""DocEntry"",T5.""Comments"",T5.""DocNum"",T5.""CardCode"",T5.""DocDate"",T5.""DocDueDate"",T5.""BPLId"",T2.""WhsCode"",T2.""LocCode"","
                        strQuery += vbCrLf + "T2.""ItemCode"",T2.""Dscription"", T4.""BaseQty"",T4.""AltQty"" Order by T5.""DocDate"" desc )as A where A.""PendQty"">0"
                        If LikeText <> "" Then
                            'strQuery += vbCrLf + "and A.""" & LikeFieldName & """ like '" & LikeText & "%' "
                            strQuery += vbCrLf + "and (UPPER(A.""" & LikeFieldName & """) like '" & LikeText & "%' or UPPER(A.""" & LikeFieldName & """) like '%" & LikeText & "%' or UPPER(A.""" & LikeFieldName & """) like '" & LikeText & "') "
                        End If
                    Else
                        strQuery = "Select ROW_NUMBER() OVER (order by A.DocEntry desc) as lineId,* from (SELECT T5.DocNum,T5.DocEntry,case when T5.U_BpMaster is null then (Select CardName from OCRD where CardCode=T2.U_CardCode) else T5.U_BpMaster end as CardName,T2.ItemCode,T2.Dscription,T5.DocDate as DocDate,Format(T5.DocDueDate,'dd/MM/yy') as DocDueDate,"
                        strQuery += vbCrLf + "(Select BPLName from OBPL where BPLId=T5.BPLId) as BPLId,T2.WhsCode,(Select Location from OLCT where Code=T2.LocCode)as LocCode,T5.Comments as Remarks,"
                        strQuery += vbCrLf + "T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (SELECT isnull(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1  INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry "
                        strQuery += vbCrLf + "AND  T0.U_GRNum =cast(T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum and T0.U_AccStkD is not null) AS PendQty "
                        strQuery += vbCrLf + "FROM OIGN T5 inner join IGN1 T2 on T5.DocEntry=T2.DocEntry INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode "
                        strQuery += vbCrLf + "left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T5.DocStatus='O' and T2.BaseType <>'202' and T5.CANCELED='N' and T5.U_QCReq='Y'   "
                        'and T3.U_InspReq='Y'
                        If InvExceptionInWhse <> "" Then
                            strQuery += vbCrLf + " AND (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
                        Else
                            strQuery += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
                        End If
                        If BranchID <> "" Then
                            strQuery += vbCrLf + "and T5.BPLId='" & BranchID & "' "
                        End If
                        If LocCode <> "" Then
                            strQuery += vbCrLf + "and T2.LocCode='" & LocCode & "' "
                        End If
                        strQuery += vbCrLf + "GROUP BY T5.DocEntry,T2.DocEntry,T2.LineNum,T5.Comments,T5.DocNum,T5.U_BpMaster,T2.U_CardCode,T5.DocDate,T5.DocDueDate,T5.BPLId,T2.WhsCode,T2.LocCode,"
                        strQuery += vbCrLf + "T2.ItemCode,T2.Dscription, T4.BaseQty,T4.AltQty" ' Order by T5.DocDate desc "
                        strQuery += vbCrLf + ") as A where A.PendQty>0"
                        If LikeText <> "" Then
                            'strQuery += vbCrLf + "and A." & LikeFieldName & " like '" & LikeText & "%' "
                            strQuery += vbCrLf + "and (UPPER(A." & LikeFieldName & ") like '" & LikeText & "%' or UPPER(A." & LikeFieldName & ") like '%" & LikeText & "%' or UPPER(A." & LikeFieldName & ") like '" & LikeText & "') "
                        End If
                        strQuery += vbCrLf + "Order by A.DocDate desc"
                    End If
                End If

            End If

            Return strQuery
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Sub CreateMySimpleForm(ByVal FormID As String, ByVal FormTitle As String, ByVal TableName As String, LinkedID As String)
        Dim oCreationParams As SAPbouiCOM.FormCreationParams
        Dim objTempForm As SAPbouiCOM.Form
        Dim objcombo As SAPbouiCOM.ComboBox
        Dim objrs As SAPbobsCOM.Recordset
        Try
            objAddOn.objApplication.Forms.Item(FormID).Visible = True
        Catch ex As Exception
            Dim str_sql As String = ""
            If TableName = "" Then
                objcombo = objForm.Items.Item("8").Specific
                If objcombo.Selected.Value = "G" Then
                    If objAddOn.HANA Then
                        str_sql = "select distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from OWTR T0  join ""@MIPLQC"" T1 on cast (T0.""U_GRPONum"" as varchar)= cast (T1.""U_GRNEntry"" as varchar) where  T0.""U_GRPONum"" ='" & objForm.Items.Item("13B").Specific.String & "'"
                    Else
                        str_sql = "select distinct T0.DocEntry,T0.DocNum,T0.DocDate from OWTR T0  join [@MIPLQC] T1 on cast (T0.U_GRPONum as varchar) = cast (T1.U_GRNEntry as varchar) where  T0.U_GRPONum ='" & objForm.Items.Item("13B").Specific.String & "'"
                    End If
                ElseIf objcombo.Selected.Value = "P" Then
                    If objAddOn.HANA Then
                        str_sql = "select distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from OWTR T0  join ""@MIPLQC"" T1 on cast (T0.""U_PORNum"" as varchar) = cast (T1.""U_ProdEntry"" as varchar)  and T1.""U_GRNum""=T0.""U_REntry"" where  T0.""U_PORNum""='" & objForm.Items.Item("15B").Specific.String & "'"
                    Else
                        str_sql = "select distinct T0. DocEntry ,T0. DocNum ,T0. DocDate  from OWTR T0  join  [@MIPLQC]  T1 on cast (T0. U_PORNum as varchar)  = cast (T1. U_ProdEntry as varchar)  and T1. U_GRNum =T0. U_REntry  where  T0. U_PORNum ='" & objForm.Items.Item("15B").Specific.String & "'"
                    End If

                ElseIf objcombo.Selected.Value = "T" Then
                    If objAddOn.HANA Then
                        str_sql = "select distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from OWTR T0  join ""@MIPLQC"" T1 on cast (T0.""U_StkEntry"" as varchar) = cast (T1.""U_TransEntry"" as varchar) where  T0.""U_StkEntry"" ='" & objForm.Items.Item("23B").Specific.String & "'"
                    Else
                        str_sql = "select distinct T0.DocEntry,T0.DocNum,T0.DocDate from OWTR T0  join [@MIPLQC] T1 on cast (T0.U_StkEntry as varchar)=cast ( T1.U_TransEntry as varchar) where  T0.U_StkEntry ='" & objForm.Items.Item("23B").Specific.String & "'"
                    End If
                ElseIf objcombo.Selected.Value = "R" Then
                    If objAddOn.HANA Then
                        str_sql = "select distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from OWTR T0  join ""@MIPLQC"" T1 on cast (T0.""U_REntry"" as varchar) = cast (T1.""U_GRNum"" as varchar) where  T0.""U_REntry"" ='" & objForm.Items.Item("51B").Specific.String & "'"
                    Else
                        str_sql = "select distinct T0.DocEntry,T0.DocNum,T0.DocDate from OWTR T0  join [@MIPLQC] T1 on cast (T0.U_REntry as varchar)= cast (T1.U_GRNum as varchar) where  T0.U_REntry ='" & objForm.Items.Item("51B").Specific.String & "'"
                    End If
                Else
                    Exit Sub
                End If
            Else
                If objAddOn.HANA Then
                    str_sql = "select distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from OIGE T0  join ""@MIPLQC"" T1 on cast (T0.""U_QCEntry"" as varchar)= cast (T1.""DocEntry"" as varchar) where  T0.""U_GRPONum"" ='" & objForm.Items.Item("13B").Specific.String & "' or T0.""U_REntry"" ='" & objForm.Items.Item("51B").Specific.String & "' "
                Else
                    str_sql = "select distinct T0.DocEntry,T0.DocNum,T0.DocDate from OIGE T0  join [@MIPLQC] T1 on cast (T0.U_QCEntry as varchar) = cast (T1.DocEntry as varchar) where  T0.U_GRPONum ='" & objForm.Items.Item("13B").Specific.String & "' or T0.U_REntry ='" & objForm.Items.Item("51B").Specific.String & "' "
                End If
            End If

            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(str_sql)
            'objrs = GetTransactions(objcombo.Selected.Value)
            If objrs.RecordCount = 0 Then objAddOn.objApplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : objrs = Nothing : Exit Sub
            oCreationParams = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            oCreationParams.UniqueID = FormID
            objTempForm = objAddOn.objApplication.Forms.AddEx(oCreationParams)
            objTempForm.Title = FormTitle
            objTempForm.Left = 400
            objTempForm.Top = 100
            objTempForm.ClientHeight = 200 '335
            objTempForm.ClientWidth = 400
            objTempForm.Left = objForm.Left + 100
            objTempForm.Top = objForm.Top + 100
            objTempForm = objAddOn.objApplication.Forms.Item(FormID)
            Dim oitm As SAPbouiCOM.Item

            Dim oGrid As SAPbouiCOM.Grid
            oitm = objTempForm.Items.Add("Grid", SAPbouiCOM.BoFormItemTypes.it_GRID)
            oitm.Top = 30
            oitm.Left = 2
            oitm.Width = 400
            oitm.Height = 130
            oGrid = objTempForm.Items.Item("Grid").Specific
            objTempForm.DataSources.DataTables.Add("DataTable")
            oitm = objTempForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oitm.Top = objTempForm.Items.Item("Grid").Top + objTempForm.Items.Item("Grid").Height + 5
            oitm.Left = 10

            Dim objDT As SAPbouiCOM.DataTable
            objDT = objTempForm.DataSources.DataTables.Item("DataTable")
            objDT.ExecuteQuery(str_sql)

            objTempForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(str_sql)

            oGrid.DataTable = objTempForm.DataSources.DataTables.Item("DataTable")

            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).TitleObject.Sortable = True
                oGrid.Columns.Item(i).Editable = False
            Next

            oGrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            Dim col As SAPbouiCOM.EditTextColumn
            col = oGrid.Columns.Item(0)
            col.LinkedObjectType = LinkedID
            objTempForm.Visible = True
            oGrid.AutoResizeColumns()
            objTempForm.Update()
            bModal = True
        End Try
    End Sub

    'Private Sub CFLConditions(ByVal FormUID As String)
    '    setCFLCond(FormUID, "CFL_1")
    '    setCFLCond(FormUID, "CFL_2")
    '    setCFLCond(FormUID, "CFL_3")
    'End Sub

    Private Sub setCFLCond(ByVal FormUID As String, ByVal CFLId As String, ByVal Row As Integer)
        Dim objCFL As SAPbouiCOM.ChooseFromList
        Dim objCondition As SAPbouiCOM.Condition
        Dim objConditions As SAPbouiCOM.Conditions
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objCombo = objForm.Items.Item("50").Specific
        objCFL = objForm.ChooseFromLists.Item(CFLId)
        For i As Integer = 0 To objCFL.GetConditions.Count - 1
            objCFL.SetConditions(Nothing)
        Next
        objConditions = objCFL.GetConditions()
        'If Not objCombo.Selected.Value = "-1" Then
        '    objCondition = objConditions.Add()
        '    objCondition.Alias = "BPLid"
        '    objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        '    objCondition.CondVal = objCombo.Selected.Value
        'End If
        objCombo = objForm.Items.Item("10").Specific
        Dim Location As String = ""
        If objAddOn.HANA Then
            Location = objAddOn.objGenFunc.getSingleValue("select ""Location"" from OWHS where ""WhsCode""='" & objMatrix.Columns.Item("3B").Cells.Item(Row).Specific.string & "'")
        Else
            Location = objAddOn.objGenFunc.getSingleValue("select Location from OWHS where WhsCode='" & objMatrix.Columns.Item("3B").Cells.Item(Row).Specific.string & "'")
        End If
        objCondition = objConditions.Add()
        objCondition.Alias = "Inactive"
        objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
        objCondition.CondVal = "Y"
        If Location <> "" Then 'Not objCombo.Selected.Value = "-"
            objCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            objCondition = objConditions.Add()
            objCondition.Alias = "Location"
            objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            objCondition.CondVal = Location 'objCombo.Selected.Value
            objCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            objCondition = objConditions.Add()
            objCondition.Alias = "WhsCode"
            objCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            objCondition.CondVal = objMatrix.Columns.Item("3B").Cells.Item(Row).Specific.string
        End If
        objCFL.SetConditions(objConditions)
    End Sub

    Private Sub LoadLocation(ByVal FormUID As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim StrSql As String = ""
            Dim j As Integer = 0
            objCombo = objForm.Items.Item("50").Specific
            If objCombo.Selected.Value = "-1" Then
                If objAddOn.HANA Then
                    StrSql = "Select Distinct T2.""Code"",T2.""Location"" from OWHS T0 join OBPL T1 on T1.""BPLId""=T0.""BPLid"" join OLCT T2 on T2.""Code""=T0.""Location"" "
                Else
                    StrSql = "Select Distinct T2. Code ,T2. Location  from OWHS T0 join OBPL T1 on T1. BPLId =T0. BPLid  join OLCT T2 on T2. Code =T0. Location  "
                End If

            Else
                If objAddOn.HANA Then
                    StrSql = "Select Distinct T2.""Code"",T2.""Location"" from OWHS T0 join OBPL T1 on T1.""BPLId""=T0.""BPLid"" join OLCT T2 on T2.""Code""=T0.""Location"" Where T1.""BPLId""='" & objCombo.Selected.Value & "'"
                Else
                    StrSql = "Select Distinct T2. Code ,T2. Location  from OWHS T0 join OBPL T1 on T1. BPLId =T0. BPLid  join OLCT T2 on T2. Code =T0. Location  Where T1. BPLId ='" & objCombo.Selected.Value & "'"
                End If

            End If

            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecordSet.DoQuery(StrSql)
            objCombo = objForm.Items.Item("10").Specific
            If objCombo.ValidValues.Count > 0 Then
                While j <= objCombo.ValidValues.Count - 1
                    objCombo.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                End While
            End If
            objCombo.ValidValues.Add("-", "-")
            While Not objRecordSet.EoF
                objCombo.ValidValues.Add(objRecordSet.Fields.Item(0).Value, objRecordSet.Fields.Item(1).Value)
                objRecordSet.MoveNext()
            End While
            objCombo.Select("-", SAPbouiCOM.BoSearchKey.psk_ByValue)
            objRecordSet = Nothing
        Catch ex As Exception

        End Try
    End Sub

    Public Sub LoadSeries(ByVal FormUID As String)
        Dim j As Integer = 0
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        QCHeader = objForm.DataSources.DBDataSources.Item("@MIPLQC")
        QCLine = objForm.DataSources.DBDataSources.Item("@MIPLQC1")
        AttachLine = objForm.DataSources.DBDataSources.Item("@MIPLQC2")
        objForm.Items.Item("21").Specific.validvalues.loadseries(Formtype, SAPbouiCOM.BoSeriesMode.sf_Add)
        objForm.Items.Item("21").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        QCHeader.SetValue("DocNum", 0, objAddOn.objGenFunc.GetDocNum(Formtype, CInt(objForm.Items.Item("21").Specific.Selected.value)))
        'QCHeader.SetValue("DocEntry", 0, QCHeader.GetValue("DocEntry", 0))
        'QCHeader.SetValue("DocNum", 0, objAddOn.objGenFunc.GetNextDocNum_Value("@MIPLQC", CInt(objForm.Items.Item("21").Specific.Selected.value)))
        'QCHeader.SetValue("DocEntry", 0, objAddOn.objGenFunc.GetNextDocEntry_Value("@MIPLQC", CInt(objForm.Items.Item("21").Specific.Selected.value)))

        objForm.Items.Item("6").Specific.String = Now.Date.ToString("dd/MM/yy") '"A"
        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
            objCombo = objForm.Items.Item("8").Specific
            objCombo.Select("G", SAPbouiCOM.BoSearchKey.psk_ByValue)
        End If
        BranchFlag = BranchEnabled(FormUID)
        If BranchFlag = "Y" Then
            objCombo = objForm.Items.Item("50").Specific
            If objCombo.ValidValues.Count > 0 Then
                While j <= objCombo.ValidValues.Count - 1
                    objCombo.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                End While
            End If
            'If objCombo.ValidValues.Count = 0 Then
            If objAddOn.HANA Then
                'strSQL = "select ""BPLId"", ""BPLName"" from OBPL"
                strSQL = "Select T0.""BPLId"", T0.""BPLName"" from OBPL T0 Join USR6 T1 on T0.""BPLId""=T1.""BPLId"" where T1.""UserCode""='" & objAddOn.objCompany.UserName & "'"
            Else
                'strSQL = "select BPLId, BPLName from OBPL"
                strSQL = "Select T0.BPLId, T0.BPLName from OBPL T0 Join USR6 T1 on T0.BPLId=T1.BPLId where T1.UserCode='" & objAddOn.objCompany.UserName & "'"
            End If
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecordSet.DoQuery(strSQL)
            objCombo.ValidValues.Add("-1", "All")
            While Not objRecordSet.EoF
                objCombo.ValidValues.Add(objRecordSet.Fields.Item(0).Value, objRecordSet.Fields.Item(1).Value)
                objRecordSet.MoveNext()
            End While
            objCombo.Select("-1", SAPbouiCOM.BoSearchKey.psk_ByValue)
            objRecordSet = Nothing
            'End If
        End If
        objCombo = objForm.Items.Item("10").Specific
        If objCombo.ValidValues.Count > 0 Then
            While j <= objCombo.ValidValues.Count - 1
                objCombo.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
            End While
        End If
        'If objCombo.ValidValues.Count = 0 Then
        If objAddOn.HANA Then
            strSQL = "select ""Code"", ""Location"" from OLCT"
        Else
            strSQL = "select Code, Location from OLCT"
        End If
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRecordSet.DoQuery(strSQL)
        objCombo.ValidValues.Add("-", "-")
        While Not objRecordSet.EoF
            objCombo.ValidValues.Add(objRecordSet.Fields.Item(0).Value, objRecordSet.Fields.Item(1).Value)
            objRecordSet.MoveNext()
        End While
        objCombo.Select("-", SAPbouiCOM.BoSearchKey.psk_ByValue)
        'End If
        objRecordSet = Nothing
    End Sub

    Private Function BranchEnabled(ByVal FormUID As String) As String
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        If objAddOn.HANA Then
            strSQL = "SELECT ""MltpBrnchs"" FROM OADM"
        Else
            strSQL = "SELECT MltpBrnchs FROM OADM"
        End If
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRecordSet.DoQuery(strSQL)
        If objRecordSet.EoF Then Return ""
        If UCase(Trim(CStr(objRecordSet.Fields.Item("MltpBrnchs").Value))) = "Y" Then
            objForm.Items.Item("50").Visible = True
            objForm.Items.Item("47").Visible = True
            Return objRecordSet.Fields.Item(0).Value.ToString
        Else
            objForm.Items.Item("50").Visible = False
            objForm.Items.Item("47").Visible = False
            Return ""
        End If
    End Function

    Private Sub CFL(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Try
            Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
            Dim objDataTable As SAPbouiCOM.DataTable
            objCFLEvent = pval
            objDataTable = objCFLEvent.SelectedObjects
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("20").Specific
            Select Case objCFLEvent.ChooseFromListUID
                Case "CFL_1"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("6A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)

                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("6A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)
                    End Try
                Case "CFL_2"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("7A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)

                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("7A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)
                    End Try
                Case "CFL_3"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("8A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)
                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("8A").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("WhsCode", 0)
                    End Try
                Case "EMP_Mat"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("12").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("empID", 0)
                            objMatrix.Columns.Item("13").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("lastName", 0) & "," & Trim(objDataTable.GetValue("firstName", 0))
                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("12").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("empID", 0)
                        objMatrix.Columns.Item("13").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("lastName", 0) & "," & Trim(objDataTable.GetValue("firstName", 0))
                    End Try
                Case "CFL_ACC"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objMatrix.Columns.Item("AccDet").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("Name", 0)
                        End If
                    Catch ex As Exception
                        objMatrix.Columns.Item("AccDet").Cells.Item(pval.Row).Specific.String = objDataTable.GetValue("Name", 0)
                    End Try
                Case "CFL_PO"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objForm.Items.Item("15").Specific.String = objDataTable.GetValue("DocNum", 0)
                            objForm.Items.Item("15B").Specific.String = objDataTable.GetValue("DocEntry", 0)
                            'objForm.Items.Item("25").Specific.string = objDataTable.GetValue("Warehouse", 0)
                        End If
                    Catch ex As Exception
                        objForm.Items.Item("15").Specific.String = objDataTable.GetValue("DocNum", 0)
                        objForm.Items.Item("15B").Specific.String = objDataTable.GetValue("DocEntry", 0)
                    End Try
                Case "CFL_GRPO"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objForm.Items.Item("13").Specific.String = objDataTable.GetValue("DocNum", 0)
                            objForm.Items.Item("13B").Specific.String = objDataTable.GetValue("DocEntry", 0)
                            'objForm.Items.Item("25").Specific.string = objDataTable.GetValue("WhsCode", 0)
                            'objForm.Items.Item("27").Specific.string = objDataTable.GetValue("CardName", 0)
                        End If
                    Catch ex As Exception
                        objForm.Items.Item("13").Specific.String = objDataTable.GetValue("DocNum", 0)
                        objForm.Items.Item("13B").Specific.String = objDataTable.GetValue("DocEntry", 0)
                        'objForm.Items.Item("25").Specific.string = objDataTable.GetValue("WhsCode", 0)
                        'objForm.Items.Item("27").Specific.string = objDataTable.GetValue("CardName", 0)
                    End Try
                Case "EMP_1"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objForm.Items.Item("6B").Specific.String = objDataTable.GetValue("empID", 0)
                            objForm.Items.Item("6C").Specific.String = objDataTable.GetValue("lastName", 0) & "," & Trim(objDataTable.GetValue("firstName", 0))
                        End If
                    Catch ex As Exception
                        objForm.Items.Item("6B").Specific.String = objDataTable.GetValue("empID", 0)
                        objForm.Items.Item("6C").Specific.String = objDataTable.GetValue("lastName", 0) & "," & Trim(objDataTable.GetValue("firstName", 0))
                    End Try
                Case "CFL_GR"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objForm.Items.Item("51").Specific.String = objDataTable.GetValue("DocNum", 0)
                            objForm.Items.Item("51B").Specific.String = objDataTable.GetValue("DocEntry", 0)
                            'objForm.Items.Item("25").Specific.string = objDataTable.GetValue("WhsCode", 0)
                        End If
                    Catch ex As Exception
                        objForm.Items.Item("51").Specific.String = objDataTable.GetValue("DocNum", 0)
                        objForm.Items.Item("51B").Specific.String = objDataTable.GetValue("DocEntry", 0)
                        'objForm.Items.Item("25").Specific.string = objDataTable.GetValue("WhsCode", 0)
                    End Try
                Case "CFL_INV"
                    Try
                        If objDataTable Is Nothing Then
                        Else
                            objForm.Items.Item("23").Specific.String = objDataTable.GetValue("DocNum", 0)
                            objForm.Items.Item("23B").Specific.String = objDataTable.GetValue("DocEntry", 0)
                            'objForm.Items.Item("25").Specific.string = objDataTable.GetValue("WhsCode", 0)
                            'objForm.Items.Item("27").Specific.string = objDataTable.GetValue("CardName", 0)
                        End If
                    Catch ex As Exception
                        objForm.Items.Item("23").Specific.String = objDataTable.GetValue("DocNum", 0)
                        objForm.Items.Item("23B").Specific.String = objDataTable.GetValue("DocEntry", 0)
                        'objForm.Items.Item("25").Specific.string = objDataTable.GetValue("WhsCode", 0)
                        'objForm.Items.Item("27").Specific.string = objDataTable.GetValue("CardName", 0)
                    End Try
            End Select
        Catch ex As Exception
            'MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Function InspectedQtyCorrect(ByVal FormUID As String, ByVal Row As Integer) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("20").Specific
        If objMatrix.RowCount = 0 Then
            objAddOn.objApplication.SetStatusBarMessage("Inspection Matrix is Empty", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return False
        End If
        Try
            TotQty = IIf(objMatrix.Columns.Item("3").Cells.Item(Row).Specific.string.trim = "", 0, CInt(objMatrix.Columns.Item("3").Cells.Item(Row).Specific.string))

            InspQty = IIf(objMatrix.Columns.Item("4").Cells.Item(Row).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("4").Cells.Item(Row).Specific.string))

            PendQty = IIf(objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string))

            AccQty = IIf(objMatrix.Columns.Item("6").Cells.Item(Row).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("6").Cells.Item(Row).Specific.string))

            RejQty = IIf(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string))

            RewQty = IIf(objMatrix.Columns.Item("8").Cells.Item(Row).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("8").Cells.Item(Row).Specific.string))
            Dim TTQty As Integer = AccQty + RejQty + RewQty
            If AccQty > 0 Then
                If objMatrix.Columns.Item("6A").Cells.Item(Row).Specific.string.Trim = "" Then
                    objAddOn.objApplication.SetStatusBarMessage("Please Update the Accepted Warehouse in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
                If objMatrix.Columns.Item("AccDet").Cells.Item(Row).Specific.string.Trim = "" Then
                    objAddOn.objApplication.SetStatusBarMessage("Please Update the Accepted Reason in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
            End If
            If RejQty > 0 Then
                If objMatrix.Columns.Item("7A").Cells.Item(Row).Specific.string.Trim = "" Then
                    objAddOn.objApplication.SetStatusBarMessage("Please Update the Rejected Warehouse in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
                If objMatrix.Columns.Item("7_1").Cells.Item(Row).Specific.string.Trim = "" Then
                    objAddOn.objApplication.SetStatusBarMessage("Please Update the Accepted Reason in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
            End If
            If RewQty > 0 Then
                If objMatrix.Columns.Item("8A").Cells.Item(Row).Specific.string.Trim = "" Then
                    objAddOn.objApplication.SetStatusBarMessage("Please Update the Reworked Warehouse in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End If
            End If
            If (CInt(objMatrix.Columns.Item("9").Cells.Item(Row).Specific.String) > TotQty) Then
                objAddOn.objApplication.SetStatusBarMessage("QtyInspected should not exceed of Total Qty.Please check in Line No :" & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return False
            End If
            If PendQty < TTQty Then
                objAddOn.objApplication.SetStatusBarMessage("Please check quantity")
                Return False
            End If
            If (CInt(objMatrix.Columns.Item("9").Cells.Item(Row).Specific.String) <> TTQty) Then
                objAddOn.objApplication.SetStatusBarMessage("Please check the QtyInspected in Line No :" & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Return False
            End If
            If CInt(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string) > 0 Then
                objMatrix.Columns.Item("7B").Cells.Item(Row).Specific.string = CStr((CInt(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string) / TTQty) * 100)
                objMatrix.Columns.Item("7C").Cells.Item(Row).Specific.string = CDbl(objMatrix.Columns.Item("7B").Cells.Item(Row).Specific.string) * 10000
                objMatrix.Columns.Item("7E").Cells.Item(Row).Specific.string = CDbl(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string) * CDbl(objMatrix.Columns.Item("7D").Cells.Item(Row).Specific.string)
            End If
            Return True
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message)
            Return False
        End Try
    End Function

    Public Sub TypeSelection(ByVal FormUID As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim objCombo As SAPbouiCOM.ComboBox
            objCombo = objForm.Items.Item("8").Specific
            objMatrix = objForm.Items.Item("20").Specific
            Select Case objCombo.Selected.Value
                Case "-"
                    objForm.Items.Item("12").Visible = False
                    objForm.Items.Item("13").Visible = False
                    objForm.Items.Item("13A").Visible = False
                    objForm.Items.Item("13B").Visible = False
                    objForm.Items.Item("14").Visible = False
                    objForm.Items.Item("15").Visible = False
                    objForm.Items.Item("15A").Visible = False
                    objForm.Items.Item("15B").Visible = False
                    objForm.Items.Item("22").Visible = False
                    objForm.Items.Item("23").Visible = False
                    objForm.Items.Item("23A").Visible = False
                    objForm.Items.Item("23B").Visible = False
                    objForm.Items.Item("49").Visible = False
                    objForm.Items.Item("51").Visible = False
                    objForm.Items.Item("51A").Visible = False
                    objForm.Items.Item("51B").Visible = False
                Case "G"
                    objForm.Items.Item("12").Visible = True
                    objForm.Items.Item("13").Visible = True
                    objForm.Items.Item("13A").Visible = True
                    objForm.Items.Item("13B").Visible = True
                    objForm.Items.Item("14").Visible = False
                    objForm.Items.Item("15").Visible = False
                    objForm.Items.Item("15A").Visible = False
                    objForm.Items.Item("15B").Visible = False
                    objForm.Items.Item("22").Visible = False
                    objForm.Items.Item("23").Visible = False
                    objForm.Items.Item("23A").Visible = False
                    objForm.Items.Item("23B").Visible = False
                    objForm.Items.Item("49").Visible = False
                    objForm.Items.Item("51").Visible = False
                    objForm.Items.Item("51A").Visible = False
                    objForm.Items.Item("51B").Visible = False
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objForm.ActiveItem = "13"
                        objForm.Items.Item("15").Specific.String = ""
                        objForm.Items.Item("15B").Specific.String = ""
                        objForm.Items.Item("23").Specific.String = ""
                        objForm.Items.Item("23B").Specific.String = ""
                        objForm.Items.Item("51").Specific.String = ""
                        objForm.Items.Item("51B").Specific.String = ""
                    End If

                Case "P"
                    objForm.Items.Item("12").Visible = False
                    objForm.Items.Item("13").Visible = False
                    objForm.Items.Item("13A").Visible = False
                    objForm.Items.Item("13B").Visible = False
                    objForm.Items.Item("14").Visible = True
                    objForm.Items.Item("15").Visible = True
                    objForm.Items.Item("15A").Visible = True
                    objForm.Items.Item("15B").Visible = True
                    objForm.Items.Item("22").Visible = False
                    objForm.Items.Item("23").Visible = False
                    objForm.Items.Item("23A").Visible = False
                    objForm.Items.Item("23B").Visible = False
                    objForm.Items.Item("49").Visible = True
                    objForm.Items.Item("51").Visible = True
                    objForm.Items.Item("51A").Visible = True
                    objForm.Items.Item("51B").Visible = True

                    objForm.Freeze(True)
                    objForm.Items.Item("49").Left = objForm.Items.Item("15B").Left + objForm.Items.Item("15B").Width
                    objForm.Items.Item("49").Top = objForm.Items.Item("15B").Top
                    objForm.Items.Item("51A").Left = objForm.Items.Item("49").Left + objForm.Items.Item("49").Width
                    objForm.Items.Item("51A").Top = objForm.Items.Item("49").Top
                    objForm.Items.Item("51").Left = objForm.Items.Item("51A").Left + objForm.Items.Item("51A").Width + 2
                    objForm.Items.Item("51").Top = objForm.Items.Item("51A").Top
                    objForm.Items.Item("51B").Left = objForm.Items.Item("51").Left + objForm.Items.Item("51").Width + 2
                    objForm.Items.Item("51B").Top = objForm.Items.Item("51").Top
                    objForm.Freeze(False)
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objForm.ActiveItem = "15"
                        'objForm.Items.Item("15").Click()
                        objForm.Items.Item("13").Specific.String = ""
                        objForm.Items.Item("13B").Specific.String = ""
                        objForm.Items.Item("23").Specific.String = ""
                        objForm.Items.Item("23B").Specific.String = ""
                        objForm.Items.Item("51").Specific.String = ""
                        objForm.Items.Item("51B").Specific.String = ""
                    End If
                Case "T"
                    objForm.Items.Item("12").Visible = False
                    objForm.Items.Item("13").Visible = False
                    objForm.Items.Item("13A").Visible = False
                    objForm.Items.Item("13B").Visible = False
                    objForm.Items.Item("14").Visible = False
                    objForm.Items.Item("15").Visible = False
                    objForm.Items.Item("15A").Visible = False
                    objForm.Items.Item("15B").Visible = False
                    objForm.Items.Item("22").Visible = True
                    objForm.Items.Item("23").Visible = True
                    objForm.Items.Item("23A").Visible = True
                    objForm.Items.Item("23B").Visible = True
                    objForm.Items.Item("49").Visible = False
                    objForm.Items.Item("51").Visible = False
                    objForm.Items.Item("51A").Visible = False
                    objForm.Items.Item("51B").Visible = False
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objForm.ActiveItem = "23"
                        objForm.Items.Item("13").Specific.String = ""
                        objForm.Items.Item("13B").Specific.String = ""
                        objForm.Items.Item("15").Specific.String = ""
                        objForm.Items.Item("15B").Specific.String = ""
                        objForm.Items.Item("51").Specific.String = ""
                        objForm.Items.Item("51B").Specific.String = ""
                    End If
                Case "R"
                    objForm.Items.Item("12").Visible = False
                    objForm.Items.Item("13").Visible = False
                    objForm.Items.Item("13A").Visible = False
                    objForm.Items.Item("13B").Visible = False
                    objForm.Items.Item("14").Visible = False
                    objForm.Items.Item("15").Visible = False
                    objForm.Items.Item("15A").Visible = False
                    objForm.Items.Item("15B").Visible = False
                    objForm.Items.Item("22").Visible = False
                    objForm.Items.Item("23").Visible = False
                    objForm.Items.Item("23A").Visible = False
                    objForm.Items.Item("23B").Visible = False
                    objForm.Items.Item("49").Visible = True
                    objForm.Items.Item("51").Visible = True
                    objForm.Items.Item("51A").Visible = True
                    objForm.Items.Item("51B").Visible = True

                    objForm.Freeze(True)
                    objForm.Items.Item("49").Left = objForm.Items.Item("22").Left
                    objForm.Items.Item("49").Top = objForm.Items.Item("22").Top + objForm.Items.Item("22").Height
                    objForm.Items.Item("51").Left = objForm.Items.Item("23").Left
                    objForm.Items.Item("51").Top = objForm.Items.Item("23").Top + objForm.Items.Item("23").Height
                    objForm.Items.Item("51A").Left = objForm.Items.Item("23A").Left
                    objForm.Items.Item("51A").Top = objForm.Items.Item("23A").Top + objForm.Items.Item("23A").Height
                    objForm.Items.Item("51B").Left = objForm.Items.Item("23B").Left
                    objForm.Items.Item("51B").Top = objForm.Items.Item("23B").Top + objForm.Items.Item("23B").Height
                    objForm.Freeze(False)
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objForm.ActiveItem = "51"
                        objForm.Items.Item("13").Specific.String = ""
                        objForm.Items.Item("13B").Specific.String = ""
                        objForm.Items.Item("23").Specific.String = ""
                        objForm.Items.Item("23B").Specific.String = ""
                        objForm.Items.Item("15").Specific.String = ""
                        objForm.Items.Item("15B").Specific.String = ""
                        objForm.Items.Item("51").Specific.String = ""
                        objForm.Items.Item("51B").Specific.String = ""
                    End If

            End Select
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                objForm.Items.Item("27").Specific.String = ""
            End If
            setWhse(objCombo.Selected.Value)
            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If objMatrix.RowCount > 0 Then
                    objMatrix.Clear()
                End If
            End If

        Catch ex As Exception
            objForm.Freeze(False)
        End Try
      
    End Sub

    Private Sub setWhse(ByVal doctype As String)
        'If objAddOn.HANA Then
        '    strSQL = "SELECT * FROM ""@QCWHSE"" WHERE ""U_Type"" = '" & doctype & "';"

        'Else
        '    strSQL = "select Code,Name,U_Type,U_InWhse,U_AccWhse,U_RejWhse,U_RewWhse from [@QCWHSE] where U_Type='" & doctype & "'"
        'End If
        'objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'objRecordSet = objAddOn.objGenFunc.DoQuery(strSQL)
        'If Not objRecordSet.EoF Then
        '    InWhse = objRecordSet.Fields.Item("U_InWhse").Value
        '    AccWhse = objRecordSet.Fields.Item("U_AccWhse").Value
        '    RejWhse = objRecordSet.Fields.Item("U_RejWhse").Value
        '    RewWhse = objRecordSet.Fields.Item("U_RewWhse").Value
        'End If

    End Sub

    Public Sub LoadInspectionMatrix(ByVal FormUID As String, ByVal Type As String)
        Try
            Dim DocEntry As String = ""
            Dim objcombo As SAPbouiCOM.ComboBox
            Dim Row As Integer = 0

            objcombo = objForm.Items.Item("8").Specific
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("20").Specific
            QCLine = objForm.DataSources.DBDataSources.Item("@MIPLQC1")
            Select Case objcombo.Selected.Value
                Case "P"
                    DocEntry = objForm.Items.Item("51B").Specific.String
                Case "G"
                    DocEntry = objForm.Items.Item("13B").Specific.String
                Case "T"
                    DocEntry = objForm.Items.Item("23B").Specific.String
                Case "R"
                    DocEntry = objForm.Items.Item("51B").Specific.String
            End Select
            If DocEntry = "" Then
                'objAddOn.objApplication.MessageBox("Select valid DocEntry")
                'objAddOn.objApplication.StatusBar.SetText("Please Select valid Document Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If
            If QCLine.GetValue("U_DEntry", 0) = DocEntry Then
                If objMatrix.VisualRowCount > 0 Then If objAddOn.objApplication.MessageBox("Do you want to Reload the data?", 2, "Yes", "No") <> 1 Then Exit Sub
            End If
            objMatrix.Clear()
            strSQL = getDetailsQuery(DocEntry, Type)
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecordSet.DoQuery(strSQL)

            ' If objRecordSet.RecordCount = 0 Then objAddOn.objApplication.StatusBar.SetText("No Records Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub
            If objRecordSet.RecordCount > 0 Then
                getItemRemovalNotify(Type)
            End If
            If objRecordSet.RecordCount > 0 Then
                If objRecordSet.RecordCount > 50 Then
                    objAddOn.objApplication.StatusBar.SetText("Loading Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If
                While Not objRecordSet.EoF
                    Dim asss As Double = CDbl(objRecordSet.Fields.Item("PendQty").Value)
                    If CDbl(objRecordSet.Fields.Item("PendQty").Value) > 0 Then
                        If objMatrix.RowCount = 0 Then
                            objMatrix.AddRow()
                        ElseIf objMatrix.Columns.Item("1").Cells.Item(objMatrix.RowCount).Specific.String <> "" Then
                            objMatrix.AddRow()
                        End If
                        QCLine.Clear()
                        Row += 1
                        objMatrix.GetLineData(objMatrix.RowCount)
                        QCLine.SetValue("U_DEntry", 0, DocEntry)
                        QCLine.SetValue("LineId", 0, Row)  'objRecordSet.Fields.Item("LineId").Value
                        QCLine.SetValue("U_BaseLinNum", 0, objRecordSet.Fields.Item("LineNum").Value)
                        QCLine.SetValue("U_ItemCode", 0, objRecordSet.Fields.Item("ItemCode").Value)
                        Dim strItemQry As String
                        Dim objItemRS As SAPbobsCOM.Recordset
                        If objAddOn.HANA Then
                            strItemQry = "SELECT T1.""ItmsGrpNam"" FROM OITM T0  INNER JOIN OITB T1 ON T0.""ItmsGrpCod"" = T1.""ItmsGrpCod"" WHERE T0.""ItemCode"" = '" & objRecordSet.Fields.Item("ItemCode").Value & "'"
                        Else
                            strItemQry = "select T1.ItmsGrpNam from OITM T0 join  OITB T1 on T0.ItmsGrpCod=T1.ItmsGrpCod where T0.ItemCode='" & objRecordSet.Fields.Item("ItemCode").Value & "'"
                        End If

                        objItemRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objItemRS.DoQuery(strItemQry)
                        Dim groupname As String = objItemRS.Fields.Item("ItmsGrpNam").Value
                        objItemRS = Nothing
                        QCLine.SetValue("U_ItemName", 0, objRecordSet.Fields.Item("Dscription").Value)
                        QCLine.SetValue("U_FrmWhse", 0, objRecordSet.Fields.Item("FrmWhse").Value)
                        QCLine.SetValue("U_ItGrpNm", 0, groupname)
                        '   QCLine.SetValue("U_TotQty", 0, objRecordSet.Fields.Item("TotQty").Value)
                        Dim ConvTotQty As Double = objRecordSet.Fields.Item("TotQty").Value
                        QCLine.SetValue("U_TotQty", 0, ConvTotQty)
                        'QCLine.SetValue("U_InvUom", 0, objRecordSet.Fields.Item("InvntryUom").Value)
                        QCLine.SetValue("U_InvUom", 0, objRecordSet.Fields.Item("unitMsr").Value)
                        'QCLine.SetValue("U_InspQty", 0, objRecordSet.Fields.Item("InspQty").Value)
                        'QCLine.SetValue("U_PendQty", 0, objRecordSet.Fields.Item("PendQty").Value)
                        QCLine.SetValue("U_InspQty", 0, objRecordSet.Fields.Item("InspQty").Value)
                        QCLine.SetValue("U_OcrCode", 0, objRecordSet.Fields.Item("OcrCode").Value)
                        QCLine.SetValue("U_PendQty", 0, objRecordSet.Fields.Item("PendQty").Value)
                        QCLine.SetValue("U_AccQty", 0, 0)
                        QCLine.SetValue("U_RejQty", 0, 0)
                        QCLine.SetValue("U_RewQty", 0, 0)
                        QCLine.SetValue("U_QtyInsp", 0, 0)
                        QCLine.SetValue("U_SmplQty", 0, 0)
                        QCLine.SetValue("U_AccWhse", 0, AccWhse)
                        QCLine.SetValue("U_RejWhse", 0, RejWhse)
                        QCLine.SetValue("U_ItemCost", 0, objRecordSet.Fields.Item("Price").Value)
                        QCLine.SetValue("U_Remarks", 0, objRecordSet.Fields.Item("U_Rem").Value)
                        QCLine.SetValue("U_RewWhse", 0, RewWhse)
                        objMatrix.SetLineData(objMatrix.RowCount)
                    End If
                    objRecordSet.MoveNext()
                End While
                If objMatrix.Item.Enabled = False Then objMatrix.Item.Enabled = True
            End If

            If objMatrix.VisualRowCount > 0 Then
                objMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
                objMatrix.SelectRow(1, True, False)
                objMatrix.Columns.Item("3").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                objMatrix.Columns.Item("4").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                objMatrix.Columns.Item("5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If objMatrix.CommonSetting.GetCellEditable(i, 19) = False Then
                        objMatrix.CommonSetting.SetCellEditable(i, 19, True)
                    End If
                Next
            End If
            If objMatrix.VisualRowCount = 0 Then
                objAddOn.objApplication.StatusBar.SetText("No Records found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                objAddOn.objApplication.StatusBar.SetText("QC Line details successfully loaded...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            objAddOn.objApplication.Menus.Item("1300").Activate()
            objForm.Refresh()
            objRecordSet = Nothing
        Catch ex As Exception
            ' MsgBox(ex.ToString)
        End Try
    End Sub

    Public Function GetUOMQty(ByVal FormUID As String, ByVal DocEntry As String, ByVal Type As String, ByVal ItemCode As String)
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim objMatrix As SAPbouiCOM.Matrix
            Dim objRecordSet As SAPbobsCOM.Recordset
            objMatrix = objForm.Items.Item("20").Specific
            Dim StrQuery As String = ""
            Dim convertedQty As Double = 0
            If objAddOn.HANA Then
                Select Case Type
                    Case "G"
                        StrQuery = "Select ""NumPerMsr"" from PDN1 where ""DocEntry""=" & DocEntry & " and ""ItemCode"" ='" & ItemCode & "'"
                    Case "P"
                        StrQuery = "Select ""NumPerMsr"" from IGN1 where ""DocEntry""=" & DocEntry & " and ""ItemCode"" ='" & ItemCode & "'"
                    Case "T"
                        StrQuery = "Select ""NumPerMsr"" from WTR1 where ""DocEntry""=" & DocEntry & " and ""ItemCode"" ='" & ItemCode & "'"
                    Case "R"
                        StrQuery = "Select ""NumPerMsr"" from IGN1 where ""DocEntry""=" & DocEntry & " and ""ItemCode"" ='" & ItemCode & "'"
                End Select
            Else
                Select Case Type
                    Case "G"
                        StrQuery = "Select NumPerMsr from PDN1 where DocEntry=" & DocEntry & " and ItemCode ='" & ItemCode & "'"
                        'StrQuery = "Select T4.BaseQty/T4.AltQty NumPerMsr FROM PDN1 T2 INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry=" & DocEntry & " and T2.ItemCode ='" & ItemCode & "'"
                    Case "P"
                        StrQuery = "Select NumPerMsr from IGN1 where DocEntry=" & DocEntry & " and ItemCode ='" & ItemCode & "'"
                        'StrQuery = "Select T4.BaseQty/T4.AltQty NumPerMsr FROM IGN1 T2 INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry=" & DocEntry & " and T2.ItemCode ='" & ItemCode & "'"
                    Case "T"
                        StrQuery = "Select NumPerMsr from WTR1 where DocEntry=" & DocEntry & " and ItemCode ='" & ItemCode & "'"
                        'StrQuery = "Select T4.BaseQty/T4.AltQty NumPerMsr FROM WTR1 T2 INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry=" & DocEntry & " and T2.ItemCode ='" & ItemCode & "'"
                    Case "R"
                        StrQuery = "Select NumPerMsr from IGN1 where DocEntry=" & DocEntry & " and ItemCode ='" & ItemCode & "'"
                        'StrQuery = "Select T4.BaseQty/T4.AltQty NumPerMsr FROM IGN1 T2 INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry=" & DocEntry & " and T2.ItemCode ='" & ItemCode & "'"
                End Select
            End If
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecordSet.DoQuery(StrQuery)

            If objRecordSet.EoF Then Return 1
            convertedQty = objRecordSet.Fields.Item("NumPerMsr").Value

            Return convertedQty
        Catch ex As Exception

        End Try
    End Function

    Public Function GetConvertedQty(ByVal FormUID As String, ByVal DocEntry As String, ByVal Type As String, ByVal ItemCode As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim objMatrix As SAPbouiCOM.Matrix
        Dim objRecordSet As SAPbobsCOM.Recordset
        objMatrix = objForm.Items.Item("20").Specific
        Dim StrQuery As String = ""
        Dim convertedQty As Double = 0
        If objAddOn.HANA Then
            Select Case Type
                Case "G"
                    StrQuery = "select  T0. ""ItemCode"", T0.""UomCode"", T0.""UomEntry"", T1.""UgpEntry"" , T1. ""InvntryUom"" ,T2.""AltQty"" ,T2.""BaseQty"" from PDN1 T0  left outer join OITM T1 on T1.""ItemCode"" = T0.""ItemCode""" &
                                                   "left outer join UGP1 T2 on T2.""UomEntry"" = T0.""UomEntry""  and T2.""UgpEntry"" = T1.""UgpEntry""  where  T0.""DocEntry"" ='" & DocEntry & "' and T0.""ItemCode"" ='" & ItemCode & "'"

                Case "P"
                    StrQuery = "select  T0. ""ItemCode"", T0.""UomCode"", T0.""UomEntry"", T1.""UgpEntry"" , T1. ""InvntryUom"" ,T2.""AltQty"" ,T2.""BaseQty"" from IGN1 T0  left outer join OITM T1 on T1.""ItemCode"" = T0.""ItemCode""" &
                                                   "left outer join UGP1 T2 on T2.""UomEntry"" = T0.""UomEntry""  and T2.""UgpEntry"" = T1.""UgpEntry""  where  T0.""DocEntry"" ='" & DocEntry & "' and T0.""ItemCode"" ='" & ItemCode & "'"

                Case "T"
                    StrQuery = "select  T0. ""ItemCode"", T0.""UomCode"", T0.""UomEntry"", T1.""UgpEntry"" , T1. ""InvntryUom"" ,T2.""AltQty"" ,T2.""BaseQty"" from WTR1 T0  left outer join OITM T1 on T1.""ItemCode"" = T0.""ItemCode""" &
                                                   "left outer join UGP1 T2 on T2.""UomEntry"" = T0.""UomEntry""  and T2.""UgpEntry"" = T1.""UgpEntry""  where  T0.""DocEntry"" ='" & DocEntry & "' and T0.""ItemCode"" ='" & ItemCode & "'"
                Case "R"
                    StrQuery = "select  T0. ""ItemCode"", T0.""UomCode"", T0.""UomEntry"", T1.""UgpEntry"" , T1. ""InvntryUom"" ,T2.""AltQty"" ,T2.""BaseQty"" from IGN1 T0  left outer join OITM T1 on T1.""ItemCode"" = T0.""ItemCode""" &
                                                   "left outer join UGP1 T2 on T2.""UomEntry"" = T0.""UomEntry""  and T2.""UgpEntry"" = T1.""UgpEntry""  where  T0.""DocEntry"" ='" & DocEntry & "' and T0.""ItemCode"" ='" & ItemCode & "'"

            End Select
        Else
            Select Case Type
                Case "G"
                    StrQuery = "select  T0. ItemCode, T0.UomCode, T0.UoMEntry, T1.UgpEntry , T1. InvntryUom ,T2.AltQty ,T2.BaseQty from PDN1 T0  left outer join OITM T1 on T1.ItemCode = T0.ItemCode " &
                               "left outer join UGP1 T2 on T2.UomEntry = T0.UomEntry  and T2.UgpEntry = T1.UgpEntry  where  T0.DocEntry ='" & DocEntry & "' and t0.ItemCode ='" & ItemCode & "'"

                Case "P"
                    StrQuery = "select  T0. ItemCode, T0.UomCode, T0.UoMEntry, T1.UgpEntry , T1. InvntryUom ,T2.AltQty ,T2.BaseQty from IGN1 T0  left outer join OITM T1 on T1.ItemCode = T0.ItemCode " &
                               "left outer join UGP1 T2 on T2.UomEntry = T0.UomEntry  and T2.UgpEntry = T1.UgpEntry  where  T0.DocEntry ='" & DocEntry & "' and t0.ItemCode ='" & ItemCode & "'"

                Case "T"
                    StrQuery = "select  T0. ItemCode, T0.UomCode, T0.UoMEntry, T1.UgpEntry , T1. InvntryUom ,T2.AltQty ,T2.BaseQty from WTR1 T0  left outer join OITM T1 on T1.ItemCode = T0.ItemCode " &
                               "left outer join UGP1 T2 on T2.UomEntry = T0.UomEntry  and T2.UgpEntry = T1.UgpEntry  where  T0.DocEntry ='" & DocEntry & "' and t0.ItemCode ='" & ItemCode & "'"
                Case "R"
                    StrQuery = "select  T0. ItemCode, T0.UomCode, T0.UoMEntry, T1.UgpEntry , T1. InvntryUom ,T2.AltQty ,T2.BaseQty from IGN1 T0  left outer join OITM T1 on T1.ItemCode = T0.ItemCode " &
                               "left outer join UGP1 T2 on T2.UomEntry = T0.UomEntry  and T2.UgpEntry = T1.UgpEntry  where  T0.DocEntry ='" & DocEntry & "' and t0.ItemCode ='" & ItemCode & "'"
            End Select
        End If
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRecordSet.DoQuery(StrQuery)

        If objRecordSet.EoF Then Return 1
        Dim AltQty As Double = objRecordSet.Fields.Item("AltQty").Value
        Dim BaseQty As Double = objRecordSet.Fields.Item("BaseQty").Value
        convertedQty = BaseQty / AltQty
        'convertedQty = BaseQty / AltQty * Quantity


        Return convertedQty
    End Function

    Private Sub getUOMValue(ByVal ItemCode As String, ByVal Type As String) ' not required

        Dim UOMCode As String
        Dim UgpEntry As Integer
        Select Case Type
            Case "G"
                strQuery = "select T1.UomCode,T0.UgpEntry from OITM T0 join PDN1 T1 on T0.ItemCode= T1.ItemCode where T1.itemCode='" & ItemCode & "'"

            Case "P"
                strQuery = "select T1.UomCode,T0.UgpEntry from OITM T0 join IGN1 T1 on T0.ItemCode= T1.ItemCode where T1.itemCode='" & ItemCode & "'"

            Case "T"
                strQuery = "select T1.UomCode,T0.UgpEntry from OITM T0 join WTR1 T1 on T0.ItemCode= T1.ItemCode where T1.itemCode='" & ItemCode & "'"
            Case "R"
                strQuery = "select T1.UomCode,T0.UgpEntry from OITM T0 join IGN1 T1 on T0.ItemCode= T1.ItemCode where T1.itemCode='" & ItemCode & "'"
        End Select
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRecordSet.DoQuery(strQuery)

        UOMCode = objRecordSet.Fields.Item("UomCode").Value
        UgpEntry = objRecordSet.Fields.Item("UgpEntry").Value

    End Sub

    Public Sub getItemRemovalNotify(ByVal Type As String)
        Dim strQuery As String = ""
        Dim objItemres As SAPbobsCOM.Recordset
        If objAddOn.HANA Then
            Select Case Type
                Case "G"
                    strQuery = "select COUNT(*) From  PDN1 T0  join OITM T1 on T0.""ItemCode""=T1.""ItemCode"" where ifnull(T1.""U_InspReq"",'N')='N'AND T0.""DocEntry""='" & objForm.Items.Item("13B").Specific.String & "'"
                Case "P"
                    strQuery = "select COUNT(*) From  IGN1 T0  join OITM T1 on T0.""ItemCode""=T1.""ItemCode"" where ifnull(T1.""U_InspReq"",'N')='N'AND T0.""DocEntry""='" & objForm.Items.Item("51B").Specific.String & "'"
                Case "T"
                    strQuery = "select COUNT(*) From  WTR1 T0  join OITM T1 on T0.""ItemCode""=T1.""ItemCode"" where ifnull(T1.""U_InspReq"",'N')='N'AND T0.""DocEntry""='" & objForm.Items.Item("23B").Specific.String & "'"
                Case "R"
                    strQuery = "select COUNT(*) From  IGN1 T0  join OITM T1 on T0.""ItemCode""=T1.""ItemCode"" where ifnull(T1.""U_InspReq"",'N')='N'AND T0.""DocEntry""='" & objForm.Items.Item("51B").Specific.String & "'"
            End Select
        Else
            Select Case Type
                Case "G"
                    strQuery = "select COUNT(*) From  PDN1 T0  join OITM T1 on T0.itemcode=T1.itemcode where isnull(T1.U_inspreq,'N')='N'AND T0.DocEntry='" & objForm.Items.Item("13B").Specific.String & "'"
                Case "P"
                    strQuery = "select COUNT(*) From  IGN1 T0  join OITM T1 on T0.itemcode=T1.itemcode where isnull(T1.U_inspreq,'N')='N'AND T0.DocEntry='" & objForm.Items.Item("51B").Specific.String & "'"
                Case "T"
                    strQuery = "select COUNT(*) From  WTR1 T0  join OITM T1 on T0.itemcode=T1.itemcode where isnull(T1.U_inspreq,'N')='N'AND T0.DocEntry='" & objForm.Items.Item("23B").Specific.String & "'"
                Case "R"
                    strQuery = "select COUNT(*) From  IGN1 T0  join OITM T1 on T0.itemcode=T1.itemcode where isnull(T1.U_inspreq,'N')='N'AND T0.DocEntry='" & objForm.Items.Item("51B").Specific.String & "'"
            End Select
        End If
        objItemres = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objItemres.DoQuery(strQuery)
        If CInt(objItemres.Fields.Item(0).Value) > 0 Then
            objAddOn.objApplication.SetStatusBarMessage(CStr(objItemres.Fields.Item(0).Value) & " items Removed... Inspection Not Required", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End If

    End Sub

    Function validate(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("20").Specific
        If objMatrix.VisualRowCount = 0 Then
            ' objAddOn.objApplication.SetStatusBarMessage("Inspection Matrix is Empty!!!", SAPbouiCOM.BoMessageTime.bmt_Long, True)
            'MsgBox("Inspection Matrix is Empty!!!", MsgBoxStyle.OkOnly, "QC Error")
            objAddOn.objApplication.MessageBox("Inspection Matrix is Empty!!!", , "OK")
            'MsgBox("Inspection Matrix is Empty!!!", MsgBoxStyle.OkCancel)
            Return False
        End If
        Try
            Dim Emptyrow As Integer = 0
            For Row = 1 To objMatrix.VisualRowCount
                'If Not InspectedQtyCorrect(FormUID, Row) Then
                '    Return False
                'End If
                'If Not RejDetailAvailable(FormUID) Then
                '    Return False
                'End If 
                TotQty = IIf(objMatrix.Columns.Item("3").Cells.Item(Row).Specific.string.trim = "", 0, CDbl(objMatrix.Columns.Item("3").Cells.Item(Row).Specific.string))

                InspQty = IIf(objMatrix.Columns.Item("4").Cells.Item(Row).Specific.string.Trim = "", 0, CDbl(objMatrix.Columns.Item("4").Cells.Item(Row).Specific.string))

                PendQty = IIf(objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string.Trim = "", 0, CDbl(objMatrix.Columns.Item("5").Cells.Item(Row).Specific.string))

                AccQty = IIf(objMatrix.Columns.Item("6").Cells.Item(Row).Specific.string.Trim = "", 0, CDbl(objMatrix.Columns.Item("6").Cells.Item(Row).Specific.string))

                RejQty = IIf(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string.Trim = "", 0, CDbl(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string))

                RewQty = IIf(objMatrix.Columns.Item("8").Cells.Item(Row).Specific.string.Trim = "", 0, CDbl(objMatrix.Columns.Item("8").Cells.Item(Row).Specific.string))

                QtyInsp = IIf(objMatrix.Columns.Item("9").Cells.Item(Row).Specific.string.Trim = "", 0, CDbl(objMatrix.Columns.Item("9").Cells.Item(Row).Specific.string))

                Disposable = IIf(objMatrix.Columns.Item("9B").Cells.Item(Row).Specific.string.Trim = "", 0, CDbl(objMatrix.Columns.Item("9B").Cells.Item(Row).Specific.string))
                If (AccQty = 0 And RejQty = 0 And RewQty = 0) Then
                    Emptyrow += 1
                End If
                If IIf(DisposableGoodsIssue = "Y", Trim(CDbl(AccQty + RejQty + RewQty + Disposable)), Trim(CDbl(AccQty + RejQty + RewQty))) <> QtyInsp Then
                    'objAddOn.objApplication.SetStatusBarMessage("Please check the editable Quantity in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    objAddOn.objApplication.MessageBox("Please check the QC Quantity in Line No : " & CStr(Row), , "OK")
                    Return False
                End If
                If (AccQty > 0 Or RejQty > 0 Or RewQty > 0) Then
                    Dim TTQty As Double
                    TTQty = IIf(DisposableGoodsIssue = "Y", Trim(CDbl(AccQty + RejQty + RewQty + Disposable)), Trim(CDbl(AccQty + RejQty + RewQty)))
                    If (AccQty = 0 Or RejQty = 0 Or RewQty = 0) And QtyInsp = 0 Then
                        'objAddOn.objApplication.SetStatusBarMessage("Please check the editable Quantity in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        objAddOn.objApplication.MessageBox("Please check the editable Quantity in Line No : " & CStr(Row), , "OK")
                        Return False
                    End If
                    If objMatrix.Columns.Item("6A").Cells.Item(Row).Specific.string.Trim <> "" Then
                        If AccQty = 0 Then
                            'objAddOn.objApplication.SetStatusBarMessage("Please Update the Accepted Quantity in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            objAddOn.objApplication.MessageBox("Please Update the Accepted Quantity in Line No : " & CStr(Row), , "OK")
                            Return False
                        End If
                    End If
                    If objMatrix.Columns.Item("7A").Cells.Item(Row).Specific.string.Trim <> "" Then
                        If RejQty = 0 Then
                            'objAddOn.objApplication.SetStatusBarMessage("Please Update the Rejected Quantity in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            objAddOn.objApplication.MessageBox("Please Update the Rejected Quantity in Line No : " & CStr(Row),, "OK")
                            Return False
                        End If
                    End If
                    If objMatrix.Columns.Item("8A").Cells.Item(Row).Specific.string.Trim <> "" Then
                        If RewQty = 0 Then
                            'objAddOn.objApplication.SetStatusBarMessage("Please Update the Rework Quantity in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            objAddOn.objApplication.MessageBox("Please Update the Rework Quantity in Line No : " & CStr(Row),, "OK")
                            Return False
                        End If
                    End If
                    If AccQty > 0 Then
                        If objMatrix.Columns.Item("6A").Cells.Item(Row).Specific.string.Trim = "" Then
                            'objAddOn.objApplication.SetStatusBarMessage("Please Update the Accepted Warehouse in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            objAddOn.objApplication.MessageBox("Please Update the Accepted Warehouse in Line No : " & CStr(Row), , "OK")
                            Return False
                        End If
                        If objMatrix.Columns.Item("AccDet").Cells.Item(Row).Specific.string.Trim = "" Then
                            'objAddOn.objApplication.SetStatusBarMessage("Please Update the Accepted Reason in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            objAddOn.objApplication.MessageBox("Please Update the Accepted Reason in Line No : " & CStr(Row),, "OK")
                            Return False
                        End If
                    End If
                    If RejQty > 0 Then
                        If objMatrix.Columns.Item("7A").Cells.Item(Row).Specific.string.Trim = "" Then
                            'objAddOn.objApplication.SetStatusBarMessage("Please Update the Rejected Warehouse in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            objAddOn.objApplication.MessageBox("Please Update the Rejected Warehouse in Line No : " & CStr(Row), , "OK")
                            Return False
                        End If
                        If objMatrix.Columns.Item("7_1").Cells.Item(Row).Specific.string.Trim = "" Then
                            'objAddOn.objApplication.SetStatusBarMessage("Please Update the Rejected Reason in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            objAddOn.objApplication.MessageBox("Please Update the Rejected Reason in Line No : " & CStr(Row),, "OK")
                            Return False
                        End If
                    End If
                    If RewQty > 0 Then
                        If objMatrix.Columns.Item("8A").Cells.Item(Row).Specific.string.Trim = "" Then
                            'objAddOn.objApplication.SetStatusBarMessage("Please Update the Reworked Warehouse in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            objAddOn.objApplication.MessageBox("Please Update the Reworked Warehouse in Line No : " & CStr(Row), , "OK")
                            Return False
                        End If
                    End If
                    If (CDbl(objMatrix.Columns.Item("9").Cells.Item(Row).Specific.String) > TotQty) Then
                        'objAddOn.objApplication.SetStatusBarMessage("QtyInspected should not exceed of Total Qty.Please check in Line No :" & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        objAddOn.objApplication.MessageBox("QtyInspected should not exceed of Total Qty.Please check in Line No :" & CStr(Row), , "OK")
                        Return False
                    End If
                    If PendQty < TTQty Then
                        ' objAddOn.objApplication.SetStatusBarMessage("Please check quantity")
                        objAddOn.objApplication.MessageBox("Please check quantity in Line No :" & CStr(Row), , "OK")
                        Return False
                    End If
                    If (CDbl(objMatrix.Columns.Item("9").Cells.Item(Row).Specific.String) <> TTQty) Then
                        'objAddOn.objApplication.SetStatusBarMessage("Please check the QtyInspected in Line No : " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        objAddOn.objApplication.MessageBox("Please check the QtyInspected in Line No :" & CStr(Row),, "OK")
                        Return False
                    End If
                    If CDbl(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string) > 0 Then
                        objMatrix.Columns.Item("7E").Cells.Item(Row).Specific.string = CDbl(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string) * CDbl(objMatrix.Columns.Item("7D").Cells.Item(Row).Specific.string)
                        objMatrix.Columns.Item("7B").Cells.Item(Row).Specific.string = CStr((CInt(objMatrix.Columns.Item("7").Cells.Item(Row).Specific.string) / TTQty) * 100)
                        objMatrix.Columns.Item("7C").Cells.Item(Row).Specific.string = CDbl(objMatrix.Columns.Item("7B").Cells.Item(Row).Specific.string) * 10000
                    End If
                End If
            Next
            If Emptyrow = objMatrix.VisualRowCount Then
                objAddOn.objApplication.MessageBox("Minimum one line required. Please update...", , "OK")
                Return False
            End If
            Return True
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Function RejDetailAvailable(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("20").Specific
        For intloop = 1 To objMatrix.RowCount
            If CInt(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) > 0 And objMatrix.Columns.Item("7_1").Cells.Item(intloop).Specific.string = "" Then
                objAddOn.objApplication.SetStatusBarMessage("Rejection Detail has to be entered")
                Return False
            End If
        Next
        Return True
    End Function

    'Function getDetailsQuery(ByVal DocumentEntry As String, ByVal Type As String) As String
    '    ' should return one of below query
    '    Dim strSQL1 As String = ""
    '    Dim InvExceptionInWhse As String = ""
    '    If objAddOn.HANA Then
    '        InvExceptionInWhse = objAddOn.objGenFunc.getSingleValue("Select ""U_Whse"" from OADM")
    '    Else
    '        InvExceptionInWhse = objAddOn.objGenFunc.getSingleValue("Select U_Whse from OADM")
    '    End If
    '    Select Case Type
    '        Case "G"
    '            If objAddOn.HANA Then
    '                strSQL1 = "SELECT ROW_NUMBER() OVER () AS ""LineId"",T2.""LineNum"", T2.""U_Rem"", T2.""ItemCode"", T2.""WhsCode"" as ""FrmWhse"",T2.""Dscription"", T2.""Price"",T2.""OcrCode"", T2.""unitMsr"",IFNULL(T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity""),0) AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 "
    '                strSQL1 += vbCrLf + " INNER JOIN ""@MIPLQC"" T0 ON  T0.""DocEntry""=T1.""DocEntry"" AND T0.""U_GRNEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""InspQty"", IFNULL( T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 "
    '                strSQL1 += vbCrLf + " INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_GRNEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum""),0) AS ""PendQty"" FROM PDN1 T2 "
    '                strSQL1 += vbCrLf + " INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "' "
    '                'strSQL1 += vbCrLf + " And IFNULL(T3.""U_InspReq"",'') = 'Y' "
    '                If InvExceptionInWhse <> "" Then
    '                    strSQL1 += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
    '                Else
    '                    strSQL1 += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
    '                End If
    '                strSQL1 += vbCrLf + "GROUP BY T2.""DocEntry"", T2.""ItemCode"", T2.""LineNum"", T2.""ItemCode"", T2.""U_Rem"",T2.""WhsCode"", T2.""Dscription"",T2.""Price"",T2.""OcrCode"", T2.""unitMsr"",T4.""BaseQty"",T4.""AltQty"";"

    '            Else
    '                ' strSQL = "Select LineNum, ItemCode, Dscription,Quantity from PDN1 where DocEntry= '" & objForm.Items.Item("13B").Specific.string & "'"
    '                strSQL1 = "SELECT T2.LineNum,T2.U_Rem, T2.ItemCode, T2.WhsCode as FrmWhse,T2.Dscription, T2.Price,T2.OcrCode, T2.unitMsr,ISNULL(T4.BaseQty/T4.AltQty * Sum(T2.Quantity),0) AS TotQty, (SELECT ISNULL(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1  "
    '                strSQL1 += vbCrLf + "INNER JOIN [@MIPLQC] T0 ON  T0.DocEntry=T1.DocEntry AND T0.U_GRNEntry = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum) AS InspQty,  ISNULL(T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (SELECT ISNULL(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1  "
    '                strSQL1 += vbCrLf + "INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry AND  T0.U_GRNEntry = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum),0) AS PendQty FROM PDN1 T2 "
    '                strSQL1 += vbCrLf + "INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry "
    '                strSQL1 += vbCrLf + "WHERE T2.DocEntry = '" & DocumentEntry & "'"
    '                'strSQL1 += vbCrLf + " And ISNULL(T3.U_InspReq,'') = 'Y'"
    '                If InvExceptionInWhse <> "" Then
    '                    strSQL1 += vbCrLf + " AND (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
    '                Else
    '                    strSQL1 += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
    '                End If
    '                strSQL1 += vbCrLf + "GROUP BY T2.DocEntry, T2.ItemCode,T2.U_Rem, T2.LineNum, T2.ItemCode,T2.WhsCode, T2.Dscription,T2.Price,T2.OcrCode, T2.unitMsr,T4.BaseQty,T4.AltQty"

    '                'strSQL1 = "select T2.LineNum, T2.ItemCode,T2. Dscription,T2.Price, T3.BuyUnitMsr,T4.BaseQty/T4.AltQty * Sum(T2.Quantity) TotQty , (select isnull(sum(T1.U_QtyInsp),0) from [@MIPLQC1] T1 join [@MIPLQC] T0 on T0.DocEntry=T1.DocEntry and T0. U_GRNEntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) InspQty ,T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (select isnull(sum(T1.U_QtyInsp),0)   from [@MIPLQC1] T1 join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and T0. U_GRNEntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) PendQty" & _
    '                '" from PDN1 T2 inner join OITM T3 on T3.ItemCode= T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry ='" & DocumentEntry & "' AND isnull(T3.U_InspReq,'')='Y' group by T2.DocEntry ,T2.ItemCode ,T2.LineNum, T2.ItemCode,T2. Dscription, T2.price,T3.BuyUnitMsr,T4.BaseQty,T4.AltQty"
    '            End If
    '        Case "P"
    '            If objAddOn.HANA Then
    '                strSQL1 = " SELECT ROW_NUMBER() OVER () AS ""LineId"", T2.""LineNum"", T2.""U_Rem"",T2.""ItemCode"", T2.""WhsCode"" as ""FrmWhse"",T2.""Dscription"", T2.""Price"", T2.""OcrCode"",T2.""unitMsr"",  IFNULL(T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity""),0) AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) "
    '                strSQL1 += vbCrLf + " FROM ""@MIPLQC1"" T1 INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry"" = T1.""DocEntry"" AND T0.""U_GRNum"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" "
    '                strSQL1 += vbCrLf + " AND T1.""U_BaseLinNum"" = T2.""LineNum"") AS ""InspQty"", IFNULL( T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) "
    '                strSQL1 += vbCrLf + " FROM ""@MIPLQC1"" T1 INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry"" = T1.""DocEntry"" AND T0.""U_GRNum"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" "
    '                strSQL1 += vbCrLf + " AND T1.""U_BaseLinNum"" = T2.""LineNum""),0) AS ""PendQty"" FROM IGN1 T2  INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "' AND T2.""BaseType"" = '202' "
    '                ' strSQL1 += vbCrLf + "AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
    '                If InvExceptionInWhse <> "" Then
    '                    strSQL1 += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
    '                Else
    '                    strSQL1 += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
    '                End If
    '                strSQL1 += vbCrLf + "GROUP BY T2.""DocEntry"", T2.""ItemCode"",  T2.""U_Rem"",T2.""LineNum"", T2.""ItemCode"",T2.""WhsCode"", T2.""Dscription"", T2.""Price"", T2.""OcrCode"",T2.""unitMsr"",T4.""BaseQty"",T4.""AltQty"";"

    '            Else
    '                strSQL1 = "SELECT T2.LineNum,T2.U_Rem, T2.ItemCode, T2.WhsCode as FrmWhse,T2.Dscription, T2.Price, T2.OcrCode,T2.unitMsr, ISNULL(T4.BaseQty/T4.AltQty * Sum(T2.Quantity),0) AS TotQty, (SELECT ISNULL(SUM(T1.U_QtyInsp), 0)  "
    '                strSQL1 += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry = T1.DocEntry AND T0.U_GRNum = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode  "
    '                strSQL1 += vbCrLf + "AND T1.U_BaseLinNum = T2.LineNum) AS InspQty,  ISNULL(T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (SELECT ISNULL(SUM(T1.U_QtyInsp), 0)  "
    '                strSQL1 += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry = T1.DocEntry AND T0.U_GRNum = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode  "
    '                strSQL1 += vbCrLf + "AND T1.U_BaseLinNum = T2.LineNum),0) AS PendQty FROM IGN1 T2  INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry"
    '                strSQL1 += vbCrLf + "WHERE T2.DocEntry = '" & DocumentEntry & "' AND T2.BaseType = '202' "
    '                ' strSQL1 += vbCrLf + "And ISNULL(T3.U_InspReq,'') = 'Y'"
    '                If InvExceptionInWhse <> "" Then
    '                    strSQL1 += vbCrLf + " AND (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
    '                Else
    '                    strSQL1 += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
    '                End If
    '                strSQL1 += vbCrLf + " GROUP BY T2.DocEntry, T2.ItemCode,T2.U_Rem, T2.LineNum, T2.ItemCode,T2.WhsCode, T2.Dscription, T2.Price, T2.OcrCode,T2.unitMsr,T4.BaseQty,T4.AltQty"

    '                '                  strSQL1 = "select T2.LineNum, T2.ItemCode,T2. Dscription,T2.Price, T3.BuyUnitMsr, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) TotQty , (select isnull(sum(T1.U_QtyInsp),0) from [@MIPLQC1] T1 " & _
    '                '          " join [@MIPLQC] T0 on T0.DocEntry=T1.DocEntry And T0. U_GREntry = T2.DocEntry And T1.U_ItemCode = T2.ItemCode And T1.U_BaseLinNum =T2.LineNum) InspQty, " & _
    '                '" T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (select isnull(sum(T1.U_QtyInsp),0)   from [@MIPLQC1] T1 join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry And " & _
    '                ' " T0. U_GREntry = T2.DocEntry And T1.U_ItemCode = T2.ItemCode And T1.U_BaseLinNum =T2.LineNum) PendQty " & _
    '                '       " from IGN1 T2  inner join OITM T3 on T3.ItemCode= T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  And T4.UgpEntry = T3.UgpEntry where T2.DocEntry ='" & DocumentEntry & "'  and T2.BaseType='202'  AND isnull(T3.U_InspReq,'')='Y' group by T2.DocEntry ,T2.ItemCode ,T2.LineNum, T2.ItemCode,T2. Dscription, T2.price, T3.BuyUnitMsr ,T4.BaseQty,T4.AltQty"


    '            End If
    '        Case "T"
    '            If objAddOn.HANA Then
    '                strSQL1 = "SELECT ROW_NUMBER() OVER () AS ""LineId"", T2.""LineNum"", T2.""U_Rem"",T2.""ItemCode"", T2.""WhsCode"" as ""FrmWhse"",T2.""Dscription"", T2.""StockPrice"" AS ""Price"",T2.""OcrCode"", T2.""unitMsr"",  IFNULL(T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity""),0) AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 "
    '                strSQL1 += vbCrLf + " INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_TransEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""InspQty"", IFNULL( T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 "
    '                strSQL1 += vbCrLf + "  INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_TransEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum""),0) AS ""PendQty"" FROM WTR1 T2 "
    '                strSQL1 += vbCrLf + " INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "'   "
    '                If InvExceptionInWhse <> "" Then
    '                    'strSQL1 += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
    '                    'strSQL1 += vbCrLf + " AND (IFNULL(T3.""U_InspReq"",'') = 'Y' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
    '                    strSQL1 += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
    '                Else
    '                    strSQL1 += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
    '                End If
    '                strSQL1 += vbCrLf + "  GROUP BY T2.""DocEntry"", T2.""ItemCode"", T2.""LineNum"",  T2.""U_Rem"",T2.""ItemCode"", T2.""WhsCode"",T2.""Dscription"", T2.""StockPrice"",T2.""OcrCode"", T2.""unitMsr"",T4.""BaseQty"",T4.""AltQty"";"

    '                'strSQL1 = "SELECT ROW_NUMBER() OVER () AS ""LineId"", T2.""LineNum"", T2.""U_Rem"",T2.""ItemCode"", T2.""WhsCode"" as ""FrmWhse"",T2.""Dscription"", T2.""StockPrice"" AS ""Price"", T2.""unitMsr"",  IFNULL(T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity""),0) AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 " &
    '                '    " INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_TransEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""InspQty"", IFNULL( T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 " &
    '                '    " INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_TransEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum""),0) AS ""PendQty"" FROM WTR1 T2 " &
    '                '    "  INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "'  AND IFNULL(T3.""U_InspReq"",'') = 'Y' " &
    '                '    " GROUP BY T2.""DocEntry"", T2.""ItemCode"", T2.""LineNum"",  T2.""U_Rem"",T2.""ItemCode"", T2.""WhsCode"",T2.""Dscription"", T2.""StockPrice"", T2.""unitMsr"",T4.""BaseQty"",T4.""AltQty"";"
    '            Else

    '                strSQL1 = "SELECT T2.LineNum,T2.U_Rem, T2.ItemCode, T2.WhsCode as FrmWhse,T2.Dscription, T2.StockPrice AS Price,T2.OcrCode, T2.unitMsr,  ISNULL(T4.BaseQty/T4.AltQty * Sum(T2.Quantity),0) AS TotQty, (SELECT ISNULL(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1 "
    '                strSQL1 += vbCrLf + " INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry AND  T0.U_TransEntry = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum) AS InspQty,  ISNULL(T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (SELECT ISNULL(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1 "
    '                strSQL1 += vbCrLf + "INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry AND  T0.U_TransEntry = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum),0) AS PendQty FROM WTR1 T2  "
    '                strSQL1 += vbCrLf + "INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry WHERE T2.DocEntry = '" & DocumentEntry & "' "
    '                If InvExceptionInWhse <> "" Then
    '                    strSQL1 += vbCrLf + " AND (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
    '                Else
    '                    strSQL1 += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
    '                End If
    '                strSQL1 += vbCrLf + "GROUP BY T2.DocEntry, T2.U_Rem,T2.ItemCode, T2.LineNum, T2.ItemCode, T2.WhsCode,T2.Dscription, T2.StockPrice,T2.OcrCode, T2.unitMsr,T4.BaseQty,T4.AltQty"

    '                '      strSQL1 = "select T2.LineNum, T2.ItemCode,T2. Dscription,T2.StockPrice AS Price, T3.BuyUnitMsr, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) TotQty , (select isnull(sum(T1.U_QtyInsp),0) from [@MIPLQC1] T1 " & _
    '                ' " join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and T0. U_TransEntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) InspQty, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (select isnull(sum(T1.U_QtyInsp),0)   from [@MIPLQC1] T1 join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and T0. U_TransEntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) PendQty " & _
    '                '" from WTR1 T2  inner join OITM T3 on T3.ItemCode= T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry ='" & DocumentEntry & "'  AND isnull(T3.U_InspReq,'')='Y' group by T2.DocEntry ,T2.ItemCode ,T2.LineNum, T2.ItemCode,T2. Dscription, T2.StockPrice, T3.BuyUnitMsr ,T4.BaseQty,T4.AltQty"

    '            End If
    '        Case "R"
    '            If objAddOn.HANA Then
    '                strSQL1 = " SELECT ROW_NUMBER() OVER () AS ""LineId"", T2.""LineNum"", T2.""U_Rem"",T2.""ItemCode"", T2.""WhsCode"" as ""FrmWhse"",T2.""Dscription"", T2.""Price"",T2.""OcrCode"", T2.""unitMsr"", IFNULL( T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity""),0) AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) "
    '                strSQL1 += vbCrLf + " FROM ""@MIPLQC1"" T1 INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry"" = T1.""DocEntry"" AND T0.""U_GRNum"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" "
    '                strSQL1 += vbCrLf + " AND T1.""U_BaseLinNum"" = T2.""LineNum"") AS ""InspQty"",  IFNULL(T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) "
    '                strSQL1 += vbCrLf + " FROM ""@MIPLQC1"" T1 INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry"" = T1.""DocEntry"" AND T0.""U_GRNum"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" "
    '                strSQL1 += vbCrLf + " AND T1.""U_BaseLinNum"" = T2.""LineNum""),0) AS ""PendQty"" FROM IGN1 T2  INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "' AND T2.""BaseType"" <>'202'  "
    '                'strSQL1 += vbCrLf + "AND IFNULL(T3.""U_InspReq"",'') = 'Y' "
    '                If InvExceptionInWhse <> "" Then
    '                    strSQL1 += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
    '                Else
    '                    strSQL1 += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
    '                End If
    '                strSQL1 += vbCrLf + "GROUP BY T2.""DocEntry"",  T2.""U_Rem"",T2.""ItemCode"", T2.""LineNum"", T2.""ItemCode"", T2.""WhsCode"",T2.""Dscription"", T2.""Price"",T2.""OcrCode"", T2.""unitMsr"",T4.""BaseQty"",T4.""AltQty"";"

    '            Else
    '                strSQL1 = "SELECT  T2.LineNum,T2.U_Rem, T2.ItemCode, T2.WhsCode as FrmWhse,T2.Dscription, T2.Price,T2.OcrCode, T2.unitMsr,  ISNULL(T4.BaseQty/T4.AltQty * Sum(T2.Quantity),0) AS TotQty, (SELECT ISNULL(SUM(T1.U_QtyInsp), 0)  "
    '                strSQL1 += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry = T1.DocEntry AND T0.U_GRNum = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode  "
    '                strSQL1 += vbCrLf + "AND T1.U_BaseLinNum = T2.LineNum) AS InspQty,  ISNULL(T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (SELECT ISNULL(SUM(T1.U_QtyInsp), 0)  "
    '                strSQL1 += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry = T1.DocEntry AND T0.U_GRNum = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode  "
    '                strSQL1 += vbCrLf + "AND T1.U_BaseLinNum = T2.LineNum),0) AS PendQty FROM IGN1 T2  INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry WHERE T2.DocEntry = '" & DocumentEntry & "' AND T2.BaseType <>'202'  "
    '                ' strSQL1 += vbCrLf + " And ISNULL(T3.U_InspReq,'') = 'Y'"
    '                If InvExceptionInWhse <> "" Then
    '                    strSQL1 += vbCrLf + " And (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
    '                Else
    '                    strSQL1 += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
    '            End If
    '                strSQL1 += vbCrLf + " GROUP BY T2.DocEntry, T2.ItemCode,T2.U_Rem, T2.LineNum, T2.ItemCode, T2.WhsCode,T2.Dscription, T2.Price,T2.OcrCode, T2.unitMsr,T4.BaseQty,T4.AltQty"

    '                '                  strSQL1 = "select T2.LineNum, T2.ItemCode,T2. Dscription,T2.Price, T3.BuyUnitMsr, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) TotQty , (select isnull(sum(T1.U_QtyInsp),0) from [@MIPLQC1] T1 " & _
    '                '          " join [@MIPLQC] T0 on T0.DocEntry=T1.DocEntry and T0. U_GREntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) InspQty, " & _
    '                '" T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (select isnull(sum(T1.U_QtyInsp),0)   from [@MIPLQC1] T1 join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and " & _
    '                ' " T0. U_GREntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) PendQty " & _
    '                '       " from IGN1 T2  inner join OITM T3 on T3.ItemCode= T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry ='" & DocumentEntry & "'  and T2.BaseType<>'202'  AND isnull(T3.U_InspReq,'')='Y' group by T2.DocEntry ,T2.ItemCode ,T2.LineNum, T2.ItemCode,T2. Dscription, T2.price, T3.BuyUnitMsr ,T4.BaseQty,T4.AltQty"


    '            End If

    '    End Select
    '    Return strSQL1
    'End Function

    Function getDetailsQuery(ByVal DocumentEntry As String, ByVal Type As String) As String
        ' should return one of below query
        Dim strSQL1 As String = ""
        Dim InvExceptionInWhse As String = ""
        If objAddOn.HANA Then
            InvExceptionInWhse = objAddOn.objGenFunc.getSingleValue("Select ""U_Whse"" from OADM")
        Else
            InvExceptionInWhse = objAddOn.objGenFunc.getSingleValue("Select U_Whse from OADM")
        End If
        Select Case Type
            Case "G"
                If objAddOn.HANA Then
                    strSQL1 = "SELECT ROW_NUMBER() OVER () AS ""LineId"",T2.""LineNum"", T2.""U_Rem"", T2.""ItemCode"", T2.""WhsCode"" as ""FrmWhse"",T2.""Dscription"", T2.""Price"",T2.""OcrCode"", T2.""unitMsr"",IFNULL(T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity""),0) AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 "
                    strSQL1 += vbCrLf + " INNER JOIN ""@MIPLQC"" T0 ON  T0.""DocEntry""=T1.""DocEntry"" AND T0.""U_GRNEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"" and T0.""Status""='C') AS ""InspQty"", IFNULL( T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 "
                    strSQL1 += vbCrLf + " INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_GRNEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"" and T0.""Status""='C'),0) AS ""PendQty"" FROM PDN1 T2 "
                    strSQL1 += vbCrLf + " INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "' "
                    If InvExceptionInWhse <> "" Then
                        strSQL1 += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
                    Else
                        strSQL1 += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
                    End If
                    strSQL1 += vbCrLf + "GROUP BY T2.""DocEntry"", T2.""ItemCode"", T2.""LineNum"", T2.""ItemCode"", T2.""U_Rem"",T2.""WhsCode"", T2.""Dscription"",T2.""Price"",T2.""OcrCode"", T2.""unitMsr"",T4.""BaseQty"",T4.""AltQty"";"

                Else
                    ' strSQL = "Select LineNum, ItemCode, Dscription,Quantity from PDN1 where DocEntry= '" & objForm.Items.Item("13B").Specific.string & "'"
                    strSQL1 = "SELECT T2.LineNum,T2.U_Rem, T2.ItemCode, T2.WhsCode as FrmWhse,T2.Dscription, T2.Price,T2.OcrCode, T2.unitMsr,ISNULL(Isnull(T4.BaseQty/T4.AltQty,1) * Sum(T2.Quantity),0) AS TotQty, (SELECT ISNULL(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1  "
                    strSQL1 += vbCrLf + "INNER JOIN [@MIPLQC] T0 ON  T0.DocEntry=T1.DocEntry AND T0.U_GRNEntry = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum and T0.Status='C') AS InspQty,  ISNULL(isnull(T4.BaseQty/T4.AltQty,1) * Sum(T2.Quantity) - (SELECT ISNULL(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1  "
                    strSQL1 += vbCrLf + "INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry AND  T0.U_GRNEntry = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum and T0.Status='C' and T0.U_AccStkD is not null),0) AS PendQty FROM PDN1 T2 "
                    strSQL1 += vbCrLf + "INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry "
                    strSQL1 += vbCrLf + "WHERE T2.DocEntry = '" & DocumentEntry & "'"
                    'strSQL1 += vbCrLf + " And ISNULL(T3.U_InspReq,'') = 'Y'"
                    If InvExceptionInWhse <> "" Then
                        strSQL1 += vbCrLf + " AND (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
                    Else
                        strSQL1 += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
                    End If
                    strSQL1 += vbCrLf + "GROUP BY T2.DocEntry, T2.ItemCode,T2.U_Rem, T2.LineNum, T2.ItemCode,T2.WhsCode, T2.Dscription,T2.Price,T2.OcrCode, T2.unitMsr,T4.BaseQty,T4.AltQty"

                    'strSQL1 = "select T2.LineNum, T2.ItemCode,T2. Dscription,T2.Price, T3.BuyUnitMsr,T4.BaseQty/T4.AltQty * Sum(T2.Quantity) TotQty , (select isnull(sum(T1.U_QtyInsp),0) from [@MIPLQC1] T1 join [@MIPLQC] T0 on T0.DocEntry=T1.DocEntry and T0. U_GRNEntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) InspQty ,T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (select isnull(sum(T1.U_QtyInsp),0)   from [@MIPLQC1] T1 join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and T0. U_GRNEntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) PendQty" & _
                    '" from PDN1 T2 inner join OITM T3 on T3.ItemCode= T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry ='" & DocumentEntry & "' AND isnull(T3.U_InspReq,'')='Y' group by T2.DocEntry ,T2.ItemCode ,T2.LineNum, T2.ItemCode,T2. Dscription, T2.price,T3.BuyUnitMsr,T4.BaseQty,T4.AltQty"
                End If
            Case "P"
                If objAddOn.HANA Then
                    strSQL1 = " SELECT ROW_NUMBER() OVER () AS ""LineId"", T2.""LineNum"", T2.""U_Rem"",T2.""ItemCode"", T2.""WhsCode"" as ""FrmWhse"",T2.""Dscription"", T2.""Price"", T2.""OcrCode"",T2.""unitMsr"",  IFNULL(T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity""),0) AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) "
                    strSQL1 += vbCrLf + " FROM ""@MIPLQC1"" T1 INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry"" = T1.""DocEntry"" AND T0.""U_GRNum"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" "
                    strSQL1 += vbCrLf + " AND T1.""U_BaseLinNum"" = T2.""LineNum"" and T0.""Status""='C') AS ""InspQty"", IFNULL( T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) "
                    strSQL1 += vbCrLf + " FROM ""@MIPLQC1"" T1 INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry"" = T1.""DocEntry"" AND T0.""U_GRNum"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" "
                    strSQL1 += vbCrLf + " AND T1.""U_BaseLinNum"" = T2.""LineNum"" and T0.""Status""='C'),0) AS ""PendQty"" FROM IGN1 T2  INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "' AND T2.""BaseType"" = '202' "
                    ' strSQL1 += vbCrLf + "AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
                    If InvExceptionInWhse <> "" Then
                        strSQL1 += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
                    Else
                        strSQL1 += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
                    End If
                    strSQL1 += vbCrLf + "GROUP BY T2.""DocEntry"", T2.""ItemCode"",  T2.""U_Rem"",T2.""LineNum"", T2.""ItemCode"",T2.""WhsCode"", T2.""Dscription"", T2.""Price"", T2.""OcrCode"",T2.""unitMsr"",T4.""BaseQty"",T4.""AltQty"";"

                Else
                    strSQL1 = "SELECT T2.LineNum,T2.U_Rem, T2.ItemCode, T2.WhsCode as FrmWhse,T2.Dscription, T2.Price, T2.OcrCode,T2.unitMsr, ISNULL(isnull(T4.BaseQty/T4.AltQty,1) * Sum(T2.Quantity),0) AS TotQty, (SELECT ISNULL(SUM(T1.U_QtyInsp), 0)  "
                    strSQL1 += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry = T1.DocEntry AND T0.U_GRNum = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode  "
                    strSQL1 += vbCrLf + "AND T1.U_BaseLinNum = T2.LineNum and T0.Status='C') AS InspQty,  ISNULL(isnull(T4.BaseQty/T4.AltQty,1) * Sum(T2.Quantity) - (SELECT ISNULL(SUM(T1.U_QtyInsp), 0)  "
                    strSQL1 += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry = T1.DocEntry AND T0.U_GRNum = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode  "
                    strSQL1 += vbCrLf + "AND T1.U_BaseLinNum = T2.LineNum and T0.Status='C' and T0.U_AccStkD is not null),0) AS PendQty FROM IGN1 T2  INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry"
                    strSQL1 += vbCrLf + "WHERE T2.DocEntry = '" & DocumentEntry & "' AND T2.BaseType = '202' "
                    ' strSQL1 += vbCrLf + "And ISNULL(T3.U_InspReq,'') = 'Y'"
                    If InvExceptionInWhse <> "" Then
                        strSQL1 += vbCrLf + " AND (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
                    Else
                        strSQL1 += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
                    End If
                    strSQL1 += vbCrLf + " GROUP BY T2.DocEntry, T2.ItemCode,T2.U_Rem, T2.LineNum, T2.ItemCode,T2.WhsCode, T2.Dscription, T2.Price, T2.OcrCode,T2.unitMsr,T4.BaseQty,T4.AltQty"

                    '                  strSQL1 = "select T2.LineNum, T2.ItemCode,T2. Dscription,T2.Price, T3.BuyUnitMsr, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) TotQty , (select isnull(sum(T1.U_QtyInsp),0) from [@MIPLQC1] T1 " & _
                    '          " join [@MIPLQC] T0 on T0.DocEntry=T1.DocEntry And T0. U_GREntry = T2.DocEntry And T1.U_ItemCode = T2.ItemCode And T1.U_BaseLinNum =T2.LineNum) InspQty, " & _
                    '" T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (select isnull(sum(T1.U_QtyInsp),0)   from [@MIPLQC1] T1 join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry And " & _
                    ' " T0. U_GREntry = T2.DocEntry And T1.U_ItemCode = T2.ItemCode And T1.U_BaseLinNum =T2.LineNum) PendQty " & _
                    '       " from IGN1 T2  inner join OITM T3 on T3.ItemCode= T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  And T4.UgpEntry = T3.UgpEntry where T2.DocEntry ='" & DocumentEntry & "'  and T2.BaseType='202'  AND isnull(T3.U_InspReq,'')='Y' group by T2.DocEntry ,T2.ItemCode ,T2.LineNum, T2.ItemCode,T2. Dscription, T2.price, T3.BuyUnitMsr ,T4.BaseQty,T4.AltQty"


                End If
            Case "T"
                If objAddOn.HANA Then
                    strSQL1 = "SELECT ROW_NUMBER() OVER () AS ""LineId"", T2.""LineNum"", T2.""U_Rem"",T2.""ItemCode"", T2.""WhsCode"" as ""FrmWhse"",T2.""Dscription"", T2.""StockPrice"" AS ""Price"",T2.""OcrCode"", T2.""unitMsr"",  IFNULL(T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity""),0) AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 "
                    strSQL1 += vbCrLf + " INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_TransEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"" and T0.""Status""='C') AS ""InspQty"", IFNULL( T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 "
                    strSQL1 += vbCrLf + "  INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_TransEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"" and T0.""Status""='C'),0) AS ""PendQty"" FROM WTR1 T2 "
                    strSQL1 += vbCrLf + " INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "'   "
                    If InvExceptionInWhse <> "" Then
                        'strSQL1 += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
                        'strSQL1 += vbCrLf + " AND (IFNULL(T3.""U_InspReq"",'') = 'Y' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
                        strSQL1 += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
                    Else
                        strSQL1 += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
                    End If
                    strSQL1 += vbCrLf + "  GROUP BY T2.""DocEntry"", T2.""ItemCode"", T2.""LineNum"",  T2.""U_Rem"",T2.""ItemCode"", T2.""WhsCode"",T2.""Dscription"", T2.""StockPrice"",T2.""OcrCode"", T2.""unitMsr"",T4.""BaseQty"",T4.""AltQty"";"

                    'strSQL1 = "SELECT ROW_NUMBER() OVER () AS ""LineId"", T2.""LineNum"", T2.""U_Rem"",T2.""ItemCode"", T2.""WhsCode"" as ""FrmWhse"",T2.""Dscription"", T2.""StockPrice"" AS ""Price"", T2.""unitMsr"",  IFNULL(T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity""),0) AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 " &
                    '    " INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_TransEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum"") AS ""InspQty"", IFNULL( T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) FROM ""@MIPLQC1"" T1 " &
                    '    " INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry""=T1.""DocEntry"" AND  T0.""U_TransEntry"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" AND T1.""U_BaseLinNum"" =T2.""LineNum""),0) AS ""PendQty"" FROM WTR1 T2 " &
                    '    "  INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "'  AND IFNULL(T3.""U_InspReq"",'') = 'Y' " &
                    '    " GROUP BY T2.""DocEntry"", T2.""ItemCode"", T2.""LineNum"",  T2.""U_Rem"",T2.""ItemCode"", T2.""WhsCode"",T2.""Dscription"", T2.""StockPrice"", T2.""unitMsr"",T4.""BaseQty"",T4.""AltQty"";"
                Else

                    strSQL1 = "SELECT T2.LineNum,T2.U_Rem, T2.ItemCode, T2.WhsCode as FrmWhse,T2.Dscription, T2.StockPrice AS Price,T2.OcrCode, T2.unitMsr,  ISNULL(Isnull(T4.BaseQty/T4.AltQty,1) * Sum(T2.Quantity),0) AS TotQty, (SELECT ISNULL(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1 "
                    strSQL1 += vbCrLf + " INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry AND  T0.U_TransEntry = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum and T0.Status='C') AS InspQty,  ISNULL(isnull(T4.BaseQty/T4.AltQty,1) * Sum(T2.Quantity) - (SELECT ISNULL(SUM(T1.U_QtyInsp), 0) FROM [@MIPLQC1] T1 "
                    strSQL1 += vbCrLf + "INNER JOIN [@MIPLQC] T0 ON T0.DocEntry=T1.DocEntry AND  T0.U_TransEntry = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode AND T1.U_BaseLinNum =T2.LineNum and T0.Status='C' and T0.U_AccStkD is not null),0) AS PendQty FROM WTR1 T2  "
                    strSQL1 += vbCrLf + "INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry WHERE T2.DocEntry = '" & DocumentEntry & "' "
                    If InvExceptionInWhse <> "" Then
                        strSQL1 += vbCrLf + " AND (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
                    Else
                        strSQL1 += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
                    End If
                    strSQL1 += vbCrLf + "GROUP BY T2.DocEntry, T2.U_Rem,T2.ItemCode, T2.LineNum, T2.ItemCode, T2.WhsCode,T2.Dscription, T2.StockPrice,T2.OcrCode, T2.unitMsr,T4.BaseQty,T4.AltQty"

                    '      strSQL1 = "select T2.LineNum, T2.ItemCode,T2. Dscription,T2.StockPrice AS Price, T3.BuyUnitMsr, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) TotQty , (select isnull(sum(T1.U_QtyInsp),0) from [@MIPLQC1] T1 " & _
                    ' " join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and T0. U_TransEntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) InspQty, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (select isnull(sum(T1.U_QtyInsp),0)   from [@MIPLQC1] T1 join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and T0. U_TransEntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) PendQty " & _
                    '" from WTR1 T2  inner join OITM T3 on T3.ItemCode= T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry ='" & DocumentEntry & "'  AND isnull(T3.U_InspReq,'')='Y' group by T2.DocEntry ,T2.ItemCode ,T2.LineNum, T2.ItemCode,T2. Dscription, T2.StockPrice, T3.BuyUnitMsr ,T4.BaseQty,T4.AltQty"

                End If
            Case "R"
                If objAddOn.HANA Then
                    strSQL1 = " SELECT ROW_NUMBER() OVER () AS ""LineId"", T2.""LineNum"", T2.""U_Rem"",T2.""ItemCode"", T2.""WhsCode"" as ""FrmWhse"",T2.""Dscription"", T2.""Price"",T2.""OcrCode"", T2.""unitMsr"", IFNULL( T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity""),0) AS ""TotQty"", (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) "
                    strSQL1 += vbCrLf + " FROM ""@MIPLQC1"" T1 INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry"" = T1.""DocEntry"" AND T0.""U_GRNum"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" "
                    strSQL1 += vbCrLf + " AND T1.""U_BaseLinNum"" = T2.""LineNum"" and T0.""Status""='C') AS ""InspQty"",  IFNULL(T4.""BaseQty""/T4.""AltQty"" * Sum(T2.""Quantity"") - (SELECT IFNULL(SUM(T1.""U_QtyInsp""), 0) "
                    strSQL1 += vbCrLf + " FROM ""@MIPLQC1"" T1 INNER JOIN ""@MIPLQC"" T0 ON T0.""DocEntry"" = T1.""DocEntry"" AND T0.""U_GRNum"" = cast( T2.""DocEntry"" as varchar) AND T1.""U_ItemCode"" = T2.""ItemCode"" "
                    strSQL1 += vbCrLf + " AND T1.""U_BaseLinNum"" = T2.""LineNum"" and T0.""Status""='C'),0) AS ""PendQty"" FROM IGN1 T2  INNER JOIN ""OITM"" T3 ON T3.""ItemCode""=T2.""ItemCode"" left outer join UGP1 T4 on T4.""UomEntry"" = T2.""UomEntry""  and T4.""UgpEntry"" = T3.""UgpEntry"" WHERE T2.""DocEntry"" = '" & DocumentEntry & "' AND T2.""BaseType"" <>'202'  "
                    'strSQL1 += vbCrLf + "AND IFNULL(T3.""U_InspReq"",'') = 'Y' "
                    If InvExceptionInWhse <> "" Then
                        strSQL1 += vbCrLf + " AND (T2.""WhsCode"" like '%" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '" & InvExceptionInWhse & "%' or T2.""WhsCode"" like '%" & InvExceptionInWhse & "')"
                    Else
                        strSQL1 += vbCrLf + " AND IFNULL(T3.""U_InspReq"",'') = 'Y'"
                    End If
                    strSQL1 += vbCrLf + "GROUP BY T2.""DocEntry"",  T2.""U_Rem"",T2.""ItemCode"", T2.""LineNum"", T2.""ItemCode"", T2.""WhsCode"",T2.""Dscription"", T2.""Price"",T2.""OcrCode"", T2.""unitMsr"",T4.""BaseQty"",T4.""AltQty"";"

                Else
                    strSQL1 = "SELECT  T2.LineNum,T2.U_Rem, T2.ItemCode, T2.WhsCode as FrmWhse,T2.Dscription, T2.Price,T2.OcrCode, T2.unitMsr,  ISNULL(isnull(T4.BaseQty/T4.AltQty,1) * Sum(T2.Quantity),0) AS TotQty, (SELECT ISNULL(SUM(T1.U_QtyInsp), 0)  "
                    strSQL1 += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry = T1.DocEntry AND T0.U_GRNum = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode  "
                    strSQL1 += vbCrLf + "AND T1.U_BaseLinNum = T2.LineNum and T0.Status='C') AS InspQty,  ISNULL(isnull(T4.BaseQty/T4.AltQty,1) * Sum(T2.Quantity) - (SELECT ISNULL(SUM(T1.U_QtyInsp), 0)  "
                    strSQL1 += vbCrLf + "FROM [@MIPLQC1] T1 INNER JOIN [@MIPLQC] T0 ON T0.DocEntry = T1.DocEntry AND T0.U_GRNum = cast( T2.DocEntry as varchar) AND T1.U_ItemCode = T2.ItemCode  "
                    strSQL1 += vbCrLf + "AND T1.U_BaseLinNum = T2.LineNum and T0.Status='C' and T0.U_AccStkD is not null),0) AS PendQty FROM IGN1 T2  INNER JOIN OITM T3 ON T3.ItemCode=T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry WHERE T2.DocEntry = '" & DocumentEntry & "' AND T2.BaseType <>'202'  "
                    ' strSQL1 += vbCrLf + " And ISNULL(T3.U_InspReq,'') = 'Y'"
                    If InvExceptionInWhse <> "" Then
                        strSQL1 += vbCrLf + " And (ISNULL(T3.U_InspReq,'') = 'Y' or T2.WhsCode like '%" & InvExceptionInWhse & "%' or T2.WhsCode like '" & InvExceptionInWhse & "%' or T2.WhsCode like '%" & InvExceptionInWhse & "')"
                    Else
                        strSQL1 += vbCrLf + "   And ISNULL(T3.U_InspReq,'') = 'Y' "
                    End If
                    strSQL1 += vbCrLf + " GROUP BY T2.DocEntry, T2.ItemCode,T2.U_Rem, T2.LineNum, T2.ItemCode, T2.WhsCode,T2.Dscription, T2.Price,T2.OcrCode, T2.unitMsr,T4.BaseQty,T4.AltQty"

                    '                  strSQL1 = "select T2.LineNum, T2.ItemCode,T2. Dscription,T2.Price, T3.BuyUnitMsr, T4.BaseQty/T4.AltQty * Sum(T2.Quantity) TotQty , (select isnull(sum(T1.U_QtyInsp),0) from [@MIPLQC1] T1 " & _
                    '          " join [@MIPLQC] T0 on T0.DocEntry=T1.DocEntry and T0. U_GREntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) InspQty, " & _
                    '" T4.BaseQty/T4.AltQty * Sum(T2.Quantity) - (select isnull(sum(T1.U_QtyInsp),0)   from [@MIPLQC1] T1 join [@MIPLQC] T0 on  T0.DocEntry=T1.DocEntry and " & _
                    ' " T0. U_GREntry = T2.DocEntry and T1.U_ItemCode = T2.ItemCode and T1.U_BaseLinNum =T2.LineNum) PendQty " & _
                    '       " from IGN1 T2  inner join OITM T3 on T3.ItemCode= T2.ItemCode left outer join UGP1 T4 on T4.UomEntry = T2.UomEntry  and T4.UgpEntry = T3.UgpEntry where T2.DocEntry ='" & DocumentEntry & "'  and T2.BaseType<>'202'  AND isnull(T3.U_InspReq,'')='Y' group by T2.DocEntry ,T2.ItemCode ,T2.LineNum, T2.ItemCode,T2. Dscription, T2.price, T3.BuyUnitMsr ,T4.BaseQty,T4.AltQty"


                End If

        End Select
        Return strSQL1
    End Function

    Function getDocumentEntry(ByVal FormUID As String, ByVal Type As String, ByVal DocEntry As String) As String
        Dim DocumentEntry As String = ""
        If DocEntry = "" Then Return ""
        Select Case Type
            Case "G"

                If objAddOn.HANA Then
                    strSQL = "Select Top 1 T0.""DocEntry"", T1.""WhsCode"", T0.""CardName"" from OPDN T0 join PDN1 T1 on T0.""DocEntry"" = T1.""DocEntry"" where T0.""InvntSttus""='O' AND T0.""DocEntry"" ='" & DocEntry & "' order by ""DocEntry"" Desc"
                Else
                    strSQL = "Select Top 1 T0.DocEntry, T1.WhsCode, T0.CardName from OPDN T0 join PDN1 T1 on T0.DocEntry = T1.DocEntry where T0.InvntSttus='O' AND T0.DocEntry ='" & DocEntry & "' order by DocEntry Desc"
                End If
                objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordSet.DoQuery(strSQL)
                If Not objRecordSet.EoF Then
                    DocumentEntry = CStr(objRecordSet.Fields.Item("DocEntry").Value)
                    'objForm.Items.Item("13B").Specific.string = DocumentEntry
                    objForm.Items.Item("25").Specific.string = CStr(objRecordSet.Fields.Item("WhsCode").Value)
                    objForm.Items.Item("27").Specific.string = CStr(objRecordSet.Fields.Item("CardName").Value)
                End If
                objRecordSet = Nothing
            Case "P"
                If objAddOn.HANA Then
                    strSQL = "Select Top 1 T0.""DocEntry"", T1.""WhsCode"", T0.""CardName"", T0.""DocNum"" from OIGN T0 join IGN1 T1 on T0.""DocEntry"" = T1.""DocEntry"" where T0.""InvntSttus""='O' AND T0.""DocEntry"" ='" & DocEntry & "' order by T0.""DocEntry"" Desc"
                Else
                    strSQL = "Select Top 1 T0.DocEntry, T1.WhsCode, T0.CardName, T0.DocNum from OIGN T0 join IGN1 T1 on T0.DocEntry = T1.DocEntry where T0.InvntSttus='O' AND T0.DocEntry ='" & DocEntry & "' order by T0.DocEntry Desc"
                End If
                'If objAddOn.HANA Then
                '    strSQL = "Select Top 1 ""DocEntry"", ""Warehouse"" from OWOR where ""DocEntry"" ='" & DocEntry & "' order by ""DocEntry"" Desc"
                'Else
                '    strSQL = "Select Top 1 DocEntry, Warehouse from OWOR where DocEntry ='" & DocEntry & "' order by DocEntry Desc"
                'End If

                objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordSet.DoQuery(strSQL)
                If Not objRecordSet.EoF Then
                    DocumentEntry = CStr(objRecordSet.Fields.Item("DocEntry").Value)
                    'objForm.Items.Item("15B").Specific.string = DocumentEntry
                    objForm.Items.Item("25").Specific.string = CStr(objRecordSet.Fields.Item("WhsCode").Value)
                End If
                objRecordSet = Nothing
            Case "T"
                If objAddOn.HANA Then
                    strSQL = "Select Top 1 T0.""DocEntry"", T1.""WhsCode"", T0.""CardName"" from OWTR T0 join WTR1 T1 on T0.""DocEntry"" = T1.""DocEntry"" where T0.""DocEntry"" ='" & DocEntry & "' order by ""DocEntry"" Desc"
                Else
                    strSQL = "Select Top 1 T0.DocEntry, T1.WhsCode, T0.CardName from OWTR T0 join WTR1 T1 on T0.DocEntry = T1.DocEntry where T0.DocEntry ='" & DocEntry & "' order by DocEntry Desc"
                End If

                objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordSet.DoQuery(strSQL)
                If Not objRecordSet.EoF Then
                    DocumentEntry = CStr(objRecordSet.Fields.Item("DocEntry").Value)
                    'objForm.Items.Item("23B").Specific.string = DocumentEntry
                    objForm.Items.Item("25").Specific.string = CStr(objRecordSet.Fields.Item("WhsCode").Value)
                    objForm.Items.Item("27").Specific.string = CStr(objRecordSet.Fields.Item("CardName").Value)
                End If
                objRecordSet = Nothing
            Case "R"
                If objAddOn.HANA Then
                    strSQL = "Select Top 1 T0.""DocEntry"", T1.""WhsCode"", (Select ""CardName"" from OCRD where ""CardCode""= T1.""U_CardCode"") as ""CardName"", T0.""DocNum"" from OIGN T0 join IGN1 T1 on T0.""DocEntry"" = T1.""DocEntry"" where T0.""InvntSttus""='O' AND T0.""DocEntry"" ='" & DocEntry & "' order by T0.""DocEntry"" Desc"
                Else
                    strSQL = "Select Top 1 T0.DocEntry, T1.WhsCode, case when T0.U_BpMaster is null then (Select CardName from OCRD where CardCode= T1.U_CardCode) else T0.U_BpMaster end as CardName, T0.DocNum from OIGN T0 join IGN1 T1 on T0.DocEntry = T1.DocEntry where T0.InvntSttus='O' AND T0.DocEntry ='" & DocEntry & "' order by T0.DocEntry Desc"
                End If

                objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordSet.DoQuery(strSQL)
                If Not objRecordSet.EoF Then
                    DocumentEntry = CStr(objRecordSet.Fields.Item("DocEntry").Value)
                    'objForm.Items.Item("51B").Specific.string = objRecordSet.Fields.Item("DocNum").Value
                    objForm.Items.Item("25").Specific.string = CStr(objRecordSet.Fields.Item("WhsCode").Value)
                    objForm.Items.Item("27").Specific.string = CStr(objRecordSet.Fields.Item("CardName").Value)
                End If
                objRecordSet = Nothing

        End Select
        Dim objedit As SAPbouiCOM.EditText
        objedit = objForm.Items.Item("27").Specific
        Dim Fieldsize As Size = TextRenderer.MeasureText(objedit.Value, New Font("Arial", 12.0F))
        If Fieldsize.Width <= 135 Then
            objedit.Item.Width = 135
        Else
            objedit.Item.Width = Fieldsize.Width
        End If
        Return DocumentEntry
    End Function

    Function StockPosting_GRN(ByVal FormUID As String) As Boolean
        Dim oStockTransfer As SAPbobsCOM.StockTransfer
        Dim BinWhse As String = ""
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("20").Specific
        If objMatrix.RowCount = 0 Then
            Return False
        End If
        Try

            Dim SumAccQty, SumRejQty, SumRewQty As Integer
            SumAccQty = SumRejQty = SumRewQty = 0

            For intloop = 1 To objMatrix.RowCount
                SumAccQty += objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string
                SumRejQty += objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string
                SumRewQty += objMatrix.Columns.Item("8").Cells.Item(intloop).Specific.string
                If BinWhse = "" Then
                    BinWhse = BinEnabled(objMatrix.Columns.Item("6A").Cells.Item(intloop).Specific.string)
                End If
            Next intloop


            InWhse = objForm.Items.Item("25").Specific.string
            If SumAccQty > 0 And BinWhse = "" Then
                oStockTransfer = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                Dim DocNum = objForm.Items.Item("4").Specific.string
                oStockTransfer.Comments = "Accepted Stock Posted From QC DocNum -> " & DocNum
                oStockTransfer.Reference2 = DocNum
                oStockTransfer.DocDate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("6").Specific.String) '.ToString("yyyyMMdd") 'Date.Now
                oStockTransfer.FromWarehouse = InWhse
                Dim ItemCode = ""
                For intloop = 1 To objMatrix.RowCount
                    AccQty = IIf(objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string))
                    If AccQty > 0 Then
                        oStockTransfer.ToWarehouse = objMatrix.Columns.Item("6A").Cells.Item(intloop).Specific.string
                        ItemCode = objMatrix.Columns.Item("1").Cells.Item(intloop).Specific.string

                        If objAddOn.HANA Then
                            strSQL = "SELECT IFNULL(""InvntItem"", 'N') AS ""InvtItem"" FROM OITM WHERE ""ItemCode"" = '" & ItemCode & "';"
                        Else
                            strSQL = "Select isnull(InvntItem,'N') InvtItem from OITM where ItemCode = '" & ItemCode & "'"
                        End If
                        Dim ChkInvt = objAddOn.objGenFunc.getSingleValue(strSQL)
                        If ChkInvt = "N" Then Return True
                        oStockTransfer.Lines.ItemCode = ItemCode
                        oStockTransfer.Lines.Quantity = AccQty
                        oStockTransfer.Lines.FromWarehouseCode = InWhse
                        oStockTransfer.Lines.WarehouseCode = objMatrix.Columns.Item("6A").Cells.Item(intloop).Specific.string
                        oStockTransfer.Lines.BinAllocations.BinAbsEntry = 3
                        oStockTransfer.Lines.BinAllocations.Quantity = AccQty
                        oStockTransfer.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = 0
                        oStockTransfer.Lines.BinAllocations.Add()
                        If objAddOn.HANA Then
                            strSQL = "Select ""ManBtchNum""  from OITM Where ""ItemCode""='" & ItemCode & "';"
                        Else
                            strSQL = "Select ManBtchNum  from OITM Where ItemCode='" & ItemCode & "'"
                        End If

                        Dim ManBtchNum = objAddOn.objGenFunc.getSingleValue(strSQL)

                        If ManBtchNum.Trim <> "N" Then

                            If objAddOn.HANA Then
                                strSQL = "select * from OBTQ where ""ItemCode"" ='" & ItemCode & "' and ""WhsCode"" ='" & InWhse & "' And ""Quantity"" > 0 order by ""SysNumber"""
                            Else
                                strSQL = "select * from OBTQ where ItemCode ='" & ItemCode & "' and WhsCode ='" & InWhse & "' And Quantity > 0 order by SysNumber"

                            End If
                            Dim objRecordSet As SAPbobsCOM.Recordset = objAddOn.objGenFunc.DoQuery(strSQL)

                            Dim count As Double = CDbl(AccQty)

                            For k As Integer = 0 To objRecordSet.RecordCount - 1
                                If objAddOn.HANA Then
                                    strSQL = "select ""DistNumber"" from OBTN where ""ItemCode"" ='" & ItemCode & "' and ""SysNumber"" ='" & objRecordSet.Fields.Item("SysNumber").Value & "'"
                                Else
                                    strSQL = "select DistNumber from OBTN where ItemCode ='" & ItemCode & "' and SysNumber ='" & objRecordSet.Fields.Item("SysNumber").Value & "'"

                                End If
                                oStockTransfer.Lines.BatchNumbers.BatchNumber = objAddOn.objGenFunc.getSingleValue(strSQL)
                                If CDbl(objRecordSet.Fields.Item("Quantity").Value) >= count Then
                                    ' Dim ss = objRecordSet.Fields.Item("SysNumber").Value
                                    oStockTransfer.Lines.BatchNumbers.Quantity = count

                                    oStockTransfer.Lines.BatchNumbers.Add()
                                    Exit For
                                Else

                                    oStockTransfer.Lines.BatchNumbers.Quantity = objRecordSet.Fields.Item("Quantity").Value
                                    oStockTransfer.Lines.BatchNumbers.Add()
                                    count = count - CDbl(objRecordSet.Fields.Item("Quantity").Value)

                                End If
                                objRecordSet.MoveNext()
                            Next

                        End If

                        oStockTransfer.Lines.Add()
                    End If ' AccQty>0
                Next intloop ' accepted lines loop end
                Dim ErrCode = oStockTransfer.Add()
                If ErrCode <> 0 Then
                    objAddOn.objApplication.SetStatusBarMessage(" QC Acc Qty Posting Error : " & objAddOn.objCompany.GetLastErrorDescription)

                    Return False
                Else
                    QCHeader.SetValue("U_AccStk", 0, objAddOn.objCompany.GetNewObjectKey)
                    objAddOn.objApplication.SetStatusBarMessage("QC Accepted Quantity Posted", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                End If
            End If
            '--------------------------------- Rejected Qty Stock Transfer------------------------------------------

            If SumRejQty > 0 Then
                oStockTransfer = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

                Dim DocNum = objForm.Items.Item("4").Specific.string
                oStockTransfer.Comments = "Rejected Stock Posted From QC DocNum -> " & DocNum
                oStockTransfer.Reference2 = DocNum
                oStockTransfer.DocDate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("6").Specific.String) 'Date.Now
                oStockTransfer.FromWarehouse = InWhse


                Dim ItemCode = ""
                Dim ToRejWorkCenter = ""
                Dim ToRewWorkCenter = ""
                For intloop = 1 To objMatrix.RowCount
                    RejQty = IIf(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string))
                    If RejQty > 0 Then
                        oStockTransfer.ToWarehouse = objMatrix.Columns.Item("7A").Cells.Item(intloop).Specific.string
                        ItemCode = objMatrix.Columns.Item("1").Cells.Item(intloop).Specific.string
                        'Dim oFlag As Boolean = False
                        If objAddOn.HANA Then
                            strSQL = "SELECT IFNULL(""InvntItem"", 'N') AS ""InvtItem"" FROM OITM WHERE ""ItemCode"" = '" & ItemCode & "';"
                        Else
                            strSQL = "Select isnull(InvntItem,'N') InvtItem from OITM where ItemCode = '" & ItemCode & "'"
                        End If
                        Dim ChkInvt = objAddOn.objGenFunc.getSingleValue(strSQL)
                        If ChkInvt = "N" Then Return True
                        oStockTransfer.Lines.ItemCode = ItemCode
                        oStockTransfer.Lines.Quantity = RejQty
                        oStockTransfer.Lines.FromWarehouseCode = InWhse
                        oStockTransfer.Lines.WarehouseCode = objMatrix.Columns.Item("7A").Cells.Item(intloop).Specific.string

                        'Batch Number Allocation
                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        If objAddOn.HANA Then
                            strSQL = "Select ""ManBtchNum""  from OITM Where ""ItemCode""='" & ItemCode & "';"
                        Else
                            strSQL = "Select ManBtchNum  from OITM Where ItemCode='" & ItemCode & "'"
                        End If
                        Dim ManBtchNum = objAddOn.objGenFunc.getSingleValue(strSQL)
                        If ManBtchNum.Trim <> "N" Then
                            If objAddOn.HANA Then
                                strSQL = "select * from OBTQ where ""ItemCode"" ='" & ItemCode & "' and ""WhsCode"" ='" & InWhse & "' And ""Quantity"" > 0 order by ""SysNumber"""
                            Else
                                strSQL = "select * from OBTQ where ItemCode ='" & ItemCode & "' and WhsCode ='" & InWhse & "' And Quantity > 0 order by SysNumber"

                            End If
                            Dim objRecordSet As SAPbobsCOM.Recordset = objAddOn.objGenFunc.DoQuery(strSQL)

                            Dim count As Double = CDbl(RejQty)

                            For k As Integer = 0 To objRecordSet.RecordCount - 1
                                If objAddOn.HANA Then
                                    strSQL = "select ""DistNumber"" from OBTN where ""ItemCode"" ='" & ItemCode & "' and ""SysNumber"" ='" & objRecordSet.Fields.Item("SysNumber").Value & "'"
                                Else
                                    strSQL = "select DistNumber from OBTN where ItemCode ='" & ItemCode & "' and SysNumber ='" & objRecordSet.Fields.Item("SysNumber").Value & "'"

                                End If
                                oStockTransfer.Lines.BatchNumbers.BatchNumber = objAddOn.objGenFunc.getSingleValue(strSQL)
                                If CDbl(objRecordSet.Fields.Item("Quantity").Value) >= count Then
                                    ' Dim ss = objRecordSet.Fields.Item("SysNumber").Value
                                    oStockTransfer.Lines.BatchNumbers.Quantity = count
                                    oStockTransfer.Lines.BatchNumbers.Add()
                                    Exit For
                                Else

                                    oStockTransfer.Lines.BatchNumbers.Quantity = objRecordSet.Fields.Item("Quantity").Value
                                    oStockTransfer.Lines.BatchNumbers.Add()
                                    count = count - CDbl(objRecordSet.Fields.Item("Quantity").Value)

                                End If
                                objRecordSet.MoveNext()
                            Next

                        End If


                        oStockTransfer.Lines.Add()
                    End If ' RejQty>0
                Next intloop ' Rejected lines loop end
                Dim ErrCode = oStockTransfer.Add()
                If ErrCode <> 0 Then
                    objAddOn.objApplication.SetStatusBarMessage(" QC Rej Qty Posting Error : " & objAddOn.objCompany.GetLastErrorDescription)
                    Return False
                Else
                    QCHeader.SetValue("U_RejStk", 0, objAddOn.objCompany.GetNewObjectKey)
                    objAddOn.objApplication.SetStatusBarMessage(" QC Rejected Qty Posted ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                End If
            End If

            '------------------------Rework Qty Stock Transfer ------------------------------------------
            If SumRewQty > 0 Then
                oStockTransfer = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

                Dim DocNum = objForm.Items.Item("4").Specific.string
                oStockTransfer.Comments = "Rework stock Posted From QC DocNum -> " & DocNum
                oStockTransfer.Reference2 = DocNum
                oStockTransfer.DocDate = objAddOn.objGenFunc.GetDateTimeValue(objForm.Items.Item("6").Specific.String) ' Date.Now
                oStockTransfer.FromWarehouse = InWhse


                Dim ItemCode = ""
                Dim ToRejWorkCenter = ""
                Dim ToRewWorkCenter = ""
                For intloop = 1 To objMatrix.RowCount
                    RewQty = IIf(objMatrix.Columns.Item("8").Cells.Item(intloop).Specific.string.Trim = "", 0, CInt(objMatrix.Columns.Item("8").Cells.Item(intloop).Specific.string))
                    If RewQty > 0 Then
                        oStockTransfer.ToWarehouse = objMatrix.Columns.Item("8A").Cells.Item(intloop).Specific.string
                        ItemCode = objMatrix.Columns.Item("1").Cells.Item(intloop).Specific.string
                        'Dim oFlag As Boolean = False
                        If objAddOn.HANA Then
                            strSQL = "SELECT IFNULL(""InvntItem"", 'N') AS ""InvtItem"" FROM OITM WHERE ""ItemCode"" = '" & ItemCode & "';"
                        Else
                            strSQL = "Select isnull(InvntItem,'N') InvtItem from OITM where ItemCode = '" & ItemCode & "'"
                        End If
                        Dim ChkInvt = objAddOn.objGenFunc.getSingleValue(strSQL)
                        If ChkInvt = "N" Then Return True
                        oStockTransfer.Lines.ItemCode = ItemCode
                        oStockTransfer.Lines.Quantity = RewQty
                        oStockTransfer.Lines.FromWarehouseCode = InWhse
                        oStockTransfer.Lines.WarehouseCode = objMatrix.Columns.Item("8A").Cells.Item(intloop).Specific.string


                        'Batch Number Allocation
                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        If objAddOn.HANA Then
                            strSQL = "Select ""ManBtchNum""  from OITM Where ""ItemCode""='" & ItemCode & "';"
                        Else
                            strSQL = "Select ManBtchNum  from OITM Where ItemCode='" & ItemCode & "'"
                        End If
                        Dim ManBtchNum = objAddOn.objGenFunc.getSingleValue(strSQL)
                        If ManBtchNum.Trim <> "N" Then
                            If objAddOn.HANA Then
                                strSQL = "select * from OBTQ where ""ItemCode"" ='" & ItemCode & "' and ""WhsCode"" ='" & InWhse & "' And ""Quantity"" > 0 order by ""SysNumber"""
                            Else
                                strSQL = "select * from OBTQ where ItemCode ='" & ItemCode & "' and WhsCode ='" & InWhse & "' And Quantity > 0 order by SysNumber"

                            End If
                            Dim objRecordSet As SAPbobsCOM.Recordset = objAddOn.objGenFunc.DoQuery(strSQL)

                            Dim count As Double = CDbl(RewQty)

                            For k As Integer = 0 To objRecordSet.RecordCount - 1
                                If objAddOn.HANA Then
                                    strSQL = "select ""DistNumber"" from OBTN where ""ItemCode"" ='" & ItemCode & "' and ""SysNumber"" ='" & objRecordSet.Fields.Item("SysNumber").Value & "'"
                                Else
                                    strSQL = "select DistNumber from OBTN where ItemCode ='" & ItemCode & "' and SysNumber ='" & objRecordSet.Fields.Item("SysNumber").Value & "'"

                                End If
                                oStockTransfer.Lines.BatchNumbers.BatchNumber = objAddOn.objGenFunc.getSingleValue(strSQL)
                                If CDbl(objRecordSet.Fields.Item("Quantity").Value) >= count Then
                                    ' Dim ss = objRecordSet.Fields.Item("SysNumber").Value
                                    oStockTransfer.Lines.BatchNumbers.Quantity = count
                                    oStockTransfer.Lines.BatchNumbers.Add()
                                    Exit For
                                Else

                                    oStockTransfer.Lines.BatchNumbers.Quantity = objRecordSet.Fields.Item("Quantity").Value
                                    oStockTransfer.Lines.BatchNumbers.Add()
                                    count = count - CDbl(objRecordSet.Fields.Item("Quantity").Value)

                                End If
                                objRecordSet.MoveNext()
                            Next
                        End If
                        oStockTransfer.Lines.Add()
                    End If ' RewQty>0
                Next intloop ' Rework lines loop end
                Dim ErrCode = oStockTransfer.Add()
                If ErrCode <> 0 Then
                    objAddOn.objApplication.SetStatusBarMessage(" QC Rework Qty Posting Error : " & objAddOn.objCompany.GetLastErrorDescription)
                    Return False
                Else
                    QCHeader.SetValue("U_RewStk", 0, objAddOn.objCompany.GetNewObjectKey)
                    objAddOn.objApplication.SetStatusBarMessage("QC Rework Qty Posted", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                End If
            End If

            '----------------------------------End of Rework quantity------------------
            QCHeader.SetValue("U_StkPost", 0, "Y")
            Return True

        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(" GRN Stock Posting Method Failed " & ex.Message)
            Return False
        End Try
    End Function

    Sub StockTransfer_BinLocation(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("20").Specific
        Dim SumAccQty, SumRejQty, SumRewQty As Double
        Dim BinWhse As String = ""
        Dim QCEntry As String = ""
        If objForm.Items.Item("34").Specific.String = "" Then
            'QCEntry = getQCEntry(FormUID)
            If QCEntry = "" Then
                SumAccQty = SumRejQty = SumRewQty = 0

                For intloop = 1 To objMatrix.RowCount
                    SumAccQty += objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string
                    SumRejQty += objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string
                    SumRewQty += objMatrix.Columns.Item("8").Cells.Item(intloop).Specific.string
                    If BinWhse = "" Then
                        BinWhse = BinEnabled(objMatrix.Columns.Item("6A").Cells.Item(intloop).Specific.string)
                    End If
                Next intloop

                If (SumAccQty > 0 Or SumRejQty > 0 Or SumRewQty > 0) And objForm.Items.Item("34").Specific.String = "" Then
                    objAddOn.objApplication.Menus.Item("3080").Activate()
                End If
            Else
                objForm.Items.Item("32").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objForm.Items.Item("34").Specific.String = QCEntry
            End If
        End If

    End Sub

    Function BinEnabled(ByVal Whse As String) As String
        If objAddOn.HANA Then
            strSQL = "SELECT IFNULL(""BinActivat"", 'N') AS ""BinActivat"" FROM OWHS WHERE ""WhsCode"" = '" & Whse & "';"
        Else
            strSQL = "Select isnull(BinActivat,'N') BinActivat from OWHS where WhsCode = '" & Whse & "'"
        End If
        Dim ChkInvt = objAddOn.objGenFunc.getSingleValue(strSQL)
        If ChkInvt = "Y" Then Return Whse
        Return ""

    End Function

    Function getQCEntry(ByVal FormUID As String) As String
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        QCHeader = objForm.DataSources.DBDataSources.Item("@MIPLQC")
        If objAddOn.HANA Then
            strSQL = "SELECT Top 1 ""DocEntry"" FROM OWTR WHERE ""U_QCEntry"" = '" & objForm.Items.Item("6E").Specific.string & "' ORDER BY ""DocEntry"" DESC;"
        Else
            strSQL = "Select top 1 DocEntry  from OWTR where U_QCEntry = '" & objForm.Items.Item("6E").Specific.string & "' order by DocEntry DESC"
        End If
        Return objAddOn.objGenFunc.getSingleValue(strSQL)
    End Function

    Private Sub BatchUpdate()
        Dim objDoc As SAPbobsCOM.Documents
        objDoc = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
        Dim batch As SAPbobsCOM.InventoryPostingBatchNumber

        '    '' MsgBox(objDoc.Lines.ItemCode)
        If objDoc.GetByKey(27) Then
            objDoc.Lines.SetCurrentLine(0)

            objDoc.Lines.ItemCode = objDoc.Lines.ItemCode
            objDoc.Lines.Quantity = objDoc.Lines.Quantity
            objDoc.Lines.WarehouseCode = objDoc.Lines.WarehouseCode
            objDoc.Lines.BatchNumbers.BaseLineNumber = 0
            objDoc.Lines.BatchNumbers.InternalSerialNumber = "REL1234"
            objDoc.Lines.BatchNumbers.ManufacturerSerialNumber = "REL1234"
            objDoc.Lines.BatchNumbers.Quantity = 10
            objDoc.Lines.BatchNumbers.BatchNumber = "REL1234"
            objDoc.Lines.BatchNumbers.Add()

            objDoc.SaveXML("C:\GRPO.xml")
            If objDoc.Update <> 0 Then
                MsgBox(objAddOn.objCompany.GetLastErrorDescription)
            Else
                MsgBox(objAddOn.objCompany.GetLastErrorDescription)
                MsgBox("Updated")
            End If
        End If
    End Sub

    Public Function Auto_StockTransfer(ByVal FormUID As String) As Boolean
        Try
            Dim Batch, Serial, DocEntry, DocNum As String, TranDocEntry As String = ""
            Dim objstocktransfer As SAPbobsCOM.StockTransfer
            Dim objcombo As SAPbouiCOM.ComboBox
            Dim objrs, objcc As SAPbobsCOM.Recordset
            objcombo = objForm.Items.Item("8").Specific
            Dim Quantity As Double
            Dim QCDocNum As Long
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Dim series, BranchCode As String
            'series = objGM.GetSeries("67", CDate(dtst1.Tables(Header).Rows(0)("DocDate")).ToString("yyyy-MM-dd"), Branch)
            Try

                objAddOn.objApplication.StatusBar.SetText("Stock Transfer Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objstocktransfer = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                'If Not objAddOn.objCompany.InTransaction Then objAddOn.objCompany.StartTransaction()
                Dim objedit As SAPbouiCOM.EditText
                objedit = objForm.Items.Item("6").Specific
                Dim DocDate As Date = Date.ParseExact(objedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                QCHeader = objForm.DataSources.DBDataSources.Item("@MIPLQC")
                QCDocNum = objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("21").Specific.Selected.value, Formtype)
                objstocktransfer.DocDate = DocDate 'CDate(dtst1.Tables(Header).Rows(0)("DocDate")).ToString("yyyy-MM-dd")
                ' objstocktransfer.Series = "" 'series
                'objstocktransfer.Comments = "" 'dtst1.Tables(Header).Rows(0)("Comments").ToString
                objstocktransfer.JournalMemo = "Auto Generated " & Now.ToString ' dtst1.Tables(Header).Rows(0)("MYGOAL_KEY").ToString
                objstocktransfer.Comments = "QC DocNum-> " & CStr(QCDocNum) 'QCHeader.GetValue("DocNum", 0) ' & objForm.Items.Item("4").Specific.string
                'Branch = objAddOn.objGenFunc.getSingleValue("select ""MltpBrnchs"" from OADM")
                'If Branch = "Y" Then
                '    BranchCode = objAddOn.objGenFunc.getSingleValue("select ""BPLId"" from OBPL where ifnull(""MainBPL"",'Y')='Y'")
                '    'objstocktransfer.BPLID. = BranchCode
                'End If
                If objcombo.Selected.Value = "G" Then
                    objstocktransfer.UserFields.Fields.Item("U_GRPONum").Value = objForm.Items.Item("13B").Specific.string
                    objstocktransfer.UserFields.Fields.Item("U_GRNEntry").Value = objForm.Items.Item("13").Specific.string
                ElseIf objcombo.Selected.Value = "P" Then
                    objstocktransfer.UserFields.Fields.Item("U_PORNum").Value = objForm.Items.Item("15B").Specific.string
                    objstocktransfer.UserFields.Fields.Item("U_ProdEntry").Value = objForm.Items.Item("15").Specific.string
                    objstocktransfer.UserFields.Fields.Item("U_REntry").Value = objForm.Items.Item("51B").Specific.string
                    objstocktransfer.UserFields.Fields.Item("U_GREntry").Value = objForm.Items.Item("51").Specific.string
                ElseIf objcombo.Selected.Value = "T" Then
                    objstocktransfer.UserFields.Fields.Item("U_StkEntry").Value = objForm.Items.Item("23B").Specific.string
                    objstocktransfer.UserFields.Fields.Item("U_StkNum").Value = objForm.Items.Item("23").Specific.string
                ElseIf objcombo.Selected.Value = "R" Then
                    objstocktransfer.UserFields.Fields.Item("U_REntry").Value = objForm.Items.Item("51B").Specific.string
                    objstocktransfer.UserFields.Fields.Item("U_GREntry").Value = objForm.Items.Item("51").Specific.string
                End If
                objstocktransfer.UserFields.Fields.Item("U_QCEntry").Value = QCHeader.GetValue("DocEntry", 0) ' objForm.Items.Item("6E").Specific.string '
                objstocktransfer.UserFields.Fields.Item("U_QCNum").Value = CStr(QCDocNum) ' QCHeader.GetValue("DocNum", 0) 'objForm.Items.Item("4").Specific.string  '
                objMatrix = objForm.Items.Item("20").Specific
                objstocktransfer.FromWarehouse = objMatrix.Columns.Item("3B").Cells.Item(1).Specific.string ' objForm.Items.Item("25").Specific.string
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If objMatrix.Columns.Item("6A").Cells.Item(i).Specific.string <> "" Then
                        objstocktransfer.ToWarehouse = objMatrix.Columns.Item("6A").Cells.Item(i).Specific.String
                        Exit For
                    ElseIf objMatrix.Columns.Item("7A").Cells.Item(i).Specific.string <> "" Then
                        objstocktransfer.ToWarehouse = objMatrix.Columns.Item("7A").Cells.Item(i).Specific.string
                        Exit For
                    ElseIf objMatrix.Columns.Item("8A").Cells.Item(i).Specific.string <> "" Then
                        objstocktransfer.ToWarehouse = objMatrix.Columns.Item("8A").Cells.Item(i).Specific.string
                        Exit For
                    End If
                Next

                If objcombo.Selected.Value = "G" Then
                    TranDocEntry = objForm.Items.Item("13B").Specific.string
                ElseIf objcombo.Selected.Value = "P" Then
                    TranDocEntry = objForm.Items.Item("51B").Specific.string
                ElseIf objcombo.Selected.Value = "T" Then
                    TranDocEntry = objForm.Items.Item("23B").Specific.string
                ElseIf objcombo.Selected.Value = "R" Then
                    TranDocEntry = objForm.Items.Item("51B").Specific.string
                End If
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If CDbl(objMatrix.Columns.Item("6").Cells.Item(i).Specific.string) > 0 Or CDbl(objMatrix.Columns.Item("7").Cells.Item(i).Specific.string) > 0 Or CDbl(objMatrix.Columns.Item("8").Cells.Item(i).Specific.string) > 0 Then
                        'Dim UOMQty As String = 0
                        If objAddOn.HANA Then
                            Serial = objAddOn.objGenFunc.getSingleValue("select ""ManSerNum"" from OITM WHERE ""ItemCode""='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.string & "'")
                            Batch = objAddOn.objGenFunc.getSingleValue("select ""ManBtchNum"" from OITM WHERE ""ItemCode""='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.string & "'")
                            ' UOMQty = objAddOn.objGenFunc.getSingleValue("SELECT ""NumInBuy"" from OITM where ""ItemCode""='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.string & "'")
                            'UOMQty = GetUOMQty(FormUID, TranDocEntry, objcombo.Selected.Value, objMatrix.Columns.Item("1").Cells.Item(i).Specific.string)
                        Else
                            Serial = objAddOn.objGenFunc.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.string & "'")
                            Batch = objAddOn.objGenFunc.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.string & "'")
                            'UOMQty = objAddOn.objGenFunc.getSingleValue("SELECT NumInBuy from OITM where ItemCode='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.string & "'")
                            'UOMQty = GetUOMQty(FormUID, TranDocEntry, objcombo.Selected.Value, objMatrix.Columns.Item("1").Cells.Item(i).Specific.string)
                        End If
                        ' GRPODocEntry = objAddOn.objGenFunc.getSingleValue("Select ""DocEntry"" from OPDN where ""DocNum""='" & objForm.Items.Item("13").Specific.string & "' ")

                        If Batch = "Y" And Serial = "N" Then
                            Dim BQty As Double = 0, TotBatchQty As Double = 0, LastBQty As Double = 0
                            Dim BatchNum As String = ""
                            objrs = GetBatch_Serial("N", TranDocEntry, objMatrix.Columns.Item("1").Cells.Item(i).Specific.string, objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string, objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string)
                            objcc = GetCostCenter()
                            If CDbl(objMatrix.Columns.Item("6").Cells.Item(i).Specific.string) > 0 Then
                                Quantity = CDbl(objMatrix.Columns.Item("6").Cells.Item(i).Specific.string)
                                objstocktransfer.Lines.ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.string
                                objstocktransfer.Lines.Quantity = Quantity '* CDbl(UOMQty)
                                objstocktransfer.Lines.FromWarehouseCode = objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string 'objForm.Items.Item("25").Specific.string
                                objstocktransfer.Lines.WarehouseCode = objMatrix.Columns.Item("6A").Cells.Item(i).Specific.string

                                If objcc.Fields.Item(0).Value <> "" Then objstocktransfer.Lines.DistributionRule = objcc.Fields.Item(0).Value
                                If objcc.Fields.Item(1).Value <> "" Then objstocktransfer.Lines.DistributionRule2 = objcc.Fields.Item(1).Value
                                If objcc.Fields.Item(2).Value <> "" Then objstocktransfer.Lines.DistributionRule3 = objcc.Fields.Item(2).Value
                                If objcc.Fields.Item(3).Value <> "" Then objstocktransfer.Lines.DistributionRule4 = objcc.Fields.Item(3).Value
                                If objcc.Fields.Item(4).Value <> "" Then objstocktransfer.Lines.DistributionRule5 = objcc.Fields.Item(4).Value

                                objstocktransfer.Lines.UserFields.Fields.Item("U_Rem").Value = "QC"
                                objstocktransfer.Lines.UserFields.Fields.Item("U_BaseLine").Value = objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string
                                'objrs = GetBatch_Serial("N", TranDocEntry, objMatrix.Columns.Item("1").Cells.Item(i).Specific.string, objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string, objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string)
                                BQty = Quantity '* CDbl(UOMQty)  ' 3
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
                                            BatchNum = CStr(objrs.Fields.Item("BatchSerial").Value)
                                            LastBQty = CDbl(objrs.Fields.Item("Qty").Value) - PendQty
                                            Exit For
                                        End If
                                    Next
                                    objstocktransfer.Lines.Add()
                                End If
                            End If
                            If CDbl(objMatrix.Columns.Item("7").Cells.Item(i).Specific.string) > 0 Then
                                Quantity = CDbl(objMatrix.Columns.Item("7").Cells.Item(i).Specific.string)
                                objstocktransfer.Lines.ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.string
                                objstocktransfer.Lines.Quantity = Quantity '* CDbl(UOMQty)
                                objstocktransfer.Lines.FromWarehouseCode = objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string 'objForm.Items.Item("25").Specific.string
                                objstocktransfer.Lines.WarehouseCode = objMatrix.Columns.Item("7A").Cells.Item(i).Specific.string

                                If objcc.Fields.Item(0).Value <> "" Then objstocktransfer.Lines.DistributionRule = objcc.Fields.Item(0).Value
                                If objcc.Fields.Item(1).Value <> "" Then objstocktransfer.Lines.DistributionRule2 = objcc.Fields.Item(1).Value
                                If objcc.Fields.Item(2).Value <> "" Then objstocktransfer.Lines.DistributionRule3 = objcc.Fields.Item(2).Value
                                If objcc.Fields.Item(3).Value <> "" Then objstocktransfer.Lines.DistributionRule4 = objcc.Fields.Item(3).Value
                                If objcc.Fields.Item(4).Value <> "" Then objstocktransfer.Lines.DistributionRule5 = objcc.Fields.Item(4).Value
                                objstocktransfer.Lines.UserFields.Fields.Item("U_Rem").Value = "QC"
                                objstocktransfer.Lines.UserFields.Fields.Item("U_BaseLine").Value = objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string
                                'objrs = GetBatch_Serial("N", TranDocEntry, objMatrix.Columns.Item("1").Cells.Item(i).Specific.string, objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string, objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string)
                                BQty = Quantity '* CDbl(UOMQty)
                                TotBatchQty = 0
                                If objrs.RecordCount > 0 Then
                                    If CStr(objrs.Fields.Item("BatchSerial").Value) = BatchNum Then
                                        If CDbl(LastBQty) > 0 Then
                                            If (BQty - TotBatchQty) - CDbl(LastBQty) > 0 Then
                                                PendQty = CDbl(LastBQty)
                                            Else
                                                PendQty = BQty
                                            End If
                                        Else
                                            objrs.MoveNext()
                                        End If
                                    End If
                                    For j As Integer = 0 To objrs.RecordCount - 1
                                        If BatchNum <> CStr(objrs.Fields.Item("BatchSerial").Value) Then
                                            If (BQty - TotBatchQty) - CDbl(objrs.Fields.Item("Qty").Value) > 0 Then
                                                PendQty = CDbl(objrs.Fields.Item("Qty").Value)
                                            Else
                                                PendQty = BQty - TotBatchQty
                                            End If
                                        End If
                                        objstocktransfer.Lines.BatchNumbers.BatchNumber = CStr(objrs.Fields.Item("BatchSerial").Value)
                                        objstocktransfer.Lines.BatchNumbers.Quantity = PendQty
                                        objstocktransfer.Lines.BatchNumbers.Add()
                                        TotBatchQty += PendQty  '2
                                        If BQty - TotBatchQty > 0 Then
                                            objrs.MoveNext()
                                        Else
                                            If CStr(objrs.Fields.Item("BatchSerial").Value) = BatchNum Then
                                                LastBQty = CDbl(objrs.Fields.Item("Qty").Value) - (CDbl(objrs.Fields.Item("Qty").Value) - (CDbl(LastBQty) - PendQty))
                                            Else
                                                LastBQty = CDbl(objrs.Fields.Item("Qty").Value) - PendQty
                                            End If
                                            BatchNum = CStr(objrs.Fields.Item("BatchSerial").Value)
                                            Exit For
                                        End If
                                    Next
                                    objstocktransfer.Lines.Add()
                                End If
                            End If
                            If CDbl(objMatrix.Columns.Item("8").Cells.Item(i).Specific.string) > 0 Then
                                Quantity = CDbl(objMatrix.Columns.Item("8").Cells.Item(i).Specific.string)
                                objstocktransfer.Lines.ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.string
                                objstocktransfer.Lines.Quantity = Quantity '* CDbl(UOMQty)
                                objstocktransfer.Lines.FromWarehouseCode = objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string ' objForm.Items.Item("25").Specific.string
                                objstocktransfer.Lines.WarehouseCode = objMatrix.Columns.Item("8A").Cells.Item(i).Specific.string
                                If objcc.Fields.Item(0).Value <> "" Then objstocktransfer.Lines.DistributionRule = objcc.Fields.Item(0).Value
                                If objcc.Fields.Item(1).Value <> "" Then objstocktransfer.Lines.DistributionRule2 = objcc.Fields.Item(1).Value
                                If objcc.Fields.Item(2).Value <> "" Then objstocktransfer.Lines.DistributionRule3 = objcc.Fields.Item(2).Value
                                If objcc.Fields.Item(3).Value <> "" Then objstocktransfer.Lines.DistributionRule4 = objcc.Fields.Item(3).Value
                                If objcc.Fields.Item(4).Value <> "" Then objstocktransfer.Lines.DistributionRule5 = objcc.Fields.Item(4).Value
                                objstocktransfer.Lines.UserFields.Fields.Item("U_Rem").Value = "QC"
                                objstocktransfer.Lines.UserFields.Fields.Item("U_BaseLine").Value = objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string
                                'objrs = GetBatch_Serial("N", TranDocEntry, objMatrix.Columns.Item("1").Cells.Item(i).Specific.string, objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string, objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string)
                                BQty = Quantity '* CDbl(UOMQty)
                                TotBatchQty = 0
                                If objrs.RecordCount > 0 Then
                                    If CStr(objrs.Fields.Item("BatchSerial").Value) = BatchNum Then
                                        If CDbl(LastBQty) > 0 Then
                                            If (BQty - TotBatchQty) - CDbl(LastBQty) > 0 Then
                                                PendQty = CDbl(LastBQty)
                                            Else
                                                PendQty = BQty
                                            End If
                                        Else
                                            objrs.MoveNext()
                                        End If
                                    End If
                                    For j As Integer = 0 To objrs.RecordCount - 1
                                        If BatchNum <> CStr(objrs.Fields.Item("BatchSerial").Value) Then
                                            If (BQty - TotBatchQty) - CDbl(objrs.Fields.Item("Qty").Value) > 0 Then
                                                PendQty = CDbl(objrs.Fields.Item("Qty").Value)
                                            Else
                                                PendQty = BQty - TotBatchQty
                                            End If
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
                                    objstocktransfer.Lines.Add()
                                End If
                            End If
                        ElseIf Batch = "N" And Serial = "Y" Then
                            Dim SerialNum As String = ""
                            Dim SQty As Double = 0, TotSerialQty As Double = 0
                            objrs = GetBatch_Serial("Y", TranDocEntry, objMatrix.Columns.Item("1").Cells.Item(i).Specific.string, objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string, objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string)
                            objcc = GetCostCenter()

                            If CDbl(objMatrix.Columns.Item("6").Cells.Item(i).Specific.string) > 0 Then
                                Quantity = CDbl(objMatrix.Columns.Item("6").Cells.Item(i).Specific.string)
                                objstocktransfer.Lines.ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.string
                                objstocktransfer.Lines.Quantity = Quantity '* CDbl(UOMQty)
                                objstocktransfer.Lines.FromWarehouseCode = objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string 'objForm.Items.Item("25").Specific.string
                                objstocktransfer.Lines.WarehouseCode = objMatrix.Columns.Item("6A").Cells.Item(i).Specific.string
                                If objcc.Fields.Item(0).Value <> "" Then objstocktransfer.Lines.DistributionRule = objcc.Fields.Item(0).Value
                                If objcc.Fields.Item(1).Value <> "" Then objstocktransfer.Lines.DistributionRule2 = objcc.Fields.Item(1).Value
                                If objcc.Fields.Item(2).Value <> "" Then objstocktransfer.Lines.DistributionRule3 = objcc.Fields.Item(2).Value
                                If objcc.Fields.Item(3).Value <> "" Then objstocktransfer.Lines.DistributionRule4 = objcc.Fields.Item(3).Value
                                If objcc.Fields.Item(4).Value <> "" Then objstocktransfer.Lines.DistributionRule5 = objcc.Fields.Item(4).Value
                                objstocktransfer.Lines.UserFields.Fields.Item("U_Rem").Value = "QC"
                                objstocktransfer.Lines.UserFields.Fields.Item("U_BaseLine").Value = objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string
                                'objrs = GetBatch_Serial("Y", TranDocEntry, objMatrix.Columns.Item("1").Cells.Item(i).Specific.string, objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string, objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string)
                                SQty = Quantity '* CDbl(UOMQty)
                                If objrs.RecordCount > 0 Then
                                    For j As Integer = 0 To objrs.RecordCount - 1
                                        objstocktransfer.Lines.SerialNumbers.InternalSerialNumber = CStr(objrs.Fields.Item("BatchSerial").Value)
                                        objstocktransfer.Lines.SerialNumbers.Quantity = CDbl(1)
                                        objstocktransfer.Lines.SerialNumbers.Add()
                                        TotSerialQty += CDbl(1)  '2
                                        If SQty - TotSerialQty > 0 Then
                                            objrs.MoveNext()
                                        Else
                                            SerialNum = CStr(objrs.Fields.Item("BatchSerial").Value)
                                            Exit For
                                        End If
                                    Next
                                    objstocktransfer.Lines.Add()
                                End If
                            End If
                            If CDbl(objMatrix.Columns.Item("7").Cells.Item(i).Specific.string) > 0 Then
                                Quantity = CDbl(objMatrix.Columns.Item("7").Cells.Item(i).Specific.string)
                                objstocktransfer.Lines.ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.string
                                objstocktransfer.Lines.Quantity = Quantity '* CDbl(UOMQty)
                                objstocktransfer.Lines.FromWarehouseCode = objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string 'objForm.Items.Item("25").Specific.string
                                objstocktransfer.Lines.WarehouseCode = objMatrix.Columns.Item("7A").Cells.Item(i).Specific.string
                                If objcc.Fields.Item(0).Value <> "" Then objstocktransfer.Lines.DistributionRule = objcc.Fields.Item(0).Value
                                If objcc.Fields.Item(1).Value <> "" Then objstocktransfer.Lines.DistributionRule2 = objcc.Fields.Item(1).Value
                                If objcc.Fields.Item(2).Value <> "" Then objstocktransfer.Lines.DistributionRule3 = objcc.Fields.Item(2).Value
                                If objcc.Fields.Item(3).Value <> "" Then objstocktransfer.Lines.DistributionRule4 = objcc.Fields.Item(3).Value
                                If objcc.Fields.Item(4).Value <> "" Then objstocktransfer.Lines.DistributionRule5 = objcc.Fields.Item(4).Value
                                objstocktransfer.Lines.UserFields.Fields.Item("U_Rem").Value = "QC"
                                objstocktransfer.Lines.UserFields.Fields.Item("U_BaseLine").Value = objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string
                                'objrs = GetBatch_Serial("Y", TranDocEntry, objMatrix.Columns.Item("1").Cells.Item(i).Specific.string, objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string, objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string)
                                SQty = Quantity '* CDbl(UOMQty)
                                TotSerialQty = 0
                                If objrs.RecordCount > 0 Then
                                    For j As Integer = 0 To objrs.RecordCount - 1
                                        If CStr(objrs.Fields.Item("BatchSerial").Value) = SerialNum Then
                                            objrs.MoveNext()
                                        End If
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
                                    objstocktransfer.Lines.Add()
                                End If
                            End If
                            If CDbl(objMatrix.Columns.Item("8").Cells.Item(i).Specific.string) > 0 Then
                                Quantity = CDbl(objMatrix.Columns.Item("8").Cells.Item(i).Specific.string)
                                objstocktransfer.Lines.ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.string
                                objstocktransfer.Lines.Quantity = Quantity '* CDbl(UOMQty)
                                objstocktransfer.Lines.FromWarehouseCode = objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string ' objForm.Items.Item("25").Specific.string
                                objstocktransfer.Lines.WarehouseCode = objMatrix.Columns.Item("8A").Cells.Item(i).Specific.string
                                If objcc.Fields.Item(0).Value <> "" Then objstocktransfer.Lines.DistributionRule = objcc.Fields.Item(0).Value
                                If objcc.Fields.Item(1).Value <> "" Then objstocktransfer.Lines.DistributionRule2 = objcc.Fields.Item(1).Value
                                If objcc.Fields.Item(2).Value <> "" Then objstocktransfer.Lines.DistributionRule3 = objcc.Fields.Item(2).Value
                                If objcc.Fields.Item(3).Value <> "" Then objstocktransfer.Lines.DistributionRule4 = objcc.Fields.Item(3).Value
                                If objcc.Fields.Item(4).Value <> "" Then objstocktransfer.Lines.DistributionRule5 = objcc.Fields.Item(4).Value
                                objstocktransfer.Lines.UserFields.Fields.Item("U_Rem").Value = "QC"
                                objstocktransfer.Lines.UserFields.Fields.Item("U_BaseLine").Value = objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string
                                'objrs = GetBatch_Serial("Y", TranDocEntry, objMatrix.Columns.Item("1").Cells.Item(i).Specific.string, objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string, objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string)
                                SQty = Quantity '* CDbl(UOMQty)
                                TotSerialQty = 0
                                If objrs.RecordCount > 0 Then
                                    For j As Integer = 0 To objrs.RecordCount - 1
                                        If CStr(objrs.Fields.Item("BatchSerial").Value) = SerialNum Then
                                            objrs.MoveNext()
                                        End If
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
                                    objstocktransfer.Lines.Add()
                                End If
                            End If
                        ElseIf Batch = "N" And Serial = "N" Then
                            objcc = GetCostCenter()
                            If CDbl(objMatrix.Columns.Item("6").Cells.Item(i).Specific.string) > 0 Then
                                Quantity = CDbl(objMatrix.Columns.Item("6").Cells.Item(i).Specific.string)
                                objstocktransfer.Lines.ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.string
                                objstocktransfer.Lines.Quantity = Quantity '* CDbl(UOMQty)
                                objstocktransfer.Lines.FromWarehouseCode = objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string 'objForm.Items.Item("25").Specific.string
                                objstocktransfer.Lines.WarehouseCode = objMatrix.Columns.Item("6A").Cells.Item(i).Specific.string
                                If objcc.Fields.Item(0).Value <> "" Then objstocktransfer.Lines.DistributionRule = objcc.Fields.Item(0).Value
                                If objcc.Fields.Item(1).Value <> "" Then objstocktransfer.Lines.DistributionRule2 = objcc.Fields.Item(1).Value
                                If objcc.Fields.Item(2).Value <> "" Then objstocktransfer.Lines.DistributionRule3 = objcc.Fields.Item(2).Value
                                If objcc.Fields.Item(3).Value <> "" Then objstocktransfer.Lines.DistributionRule4 = objcc.Fields.Item(3).Value
                                If objcc.Fields.Item(4).Value <> "" Then objstocktransfer.Lines.DistributionRule5 = objcc.Fields.Item(4).Value
                                objstocktransfer.Lines.UserFields.Fields.Item("U_Rem").Value = "QC"
                                objstocktransfer.Lines.UserFields.Fields.Item("U_BaseLine").Value = objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string
                                objstocktransfer.Lines.Add()
                            End If
                            If CDbl(objMatrix.Columns.Item("7").Cells.Item(i).Specific.string) > 0 Then
                                Quantity = CDbl(objMatrix.Columns.Item("7").Cells.Item(i).Specific.string)
                                objstocktransfer.Lines.ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.string
                                objstocktransfer.Lines.Quantity = Quantity '* CDbl(UOMQty)
                                objstocktransfer.Lines.FromWarehouseCode = objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string ' objForm.Items.Item("25").Specific.string
                                objstocktransfer.Lines.WarehouseCode = objMatrix.Columns.Item("7A").Cells.Item(i).Specific.string
                                If objcc.Fields.Item(0).Value <> "" Then objstocktransfer.Lines.DistributionRule = objcc.Fields.Item(0).Value
                                If objcc.Fields.Item(1).Value <> "" Then objstocktransfer.Lines.DistributionRule2 = objcc.Fields.Item(1).Value
                                If objcc.Fields.Item(2).Value <> "" Then objstocktransfer.Lines.DistributionRule3 = objcc.Fields.Item(2).Value
                                If objcc.Fields.Item(3).Value <> "" Then objstocktransfer.Lines.DistributionRule4 = objcc.Fields.Item(3).Value
                                If objcc.Fields.Item(4).Value <> "" Then objstocktransfer.Lines.DistributionRule5 = objcc.Fields.Item(4).Value
                                objstocktransfer.Lines.UserFields.Fields.Item("U_Rem").Value = "QC"
                                objstocktransfer.Lines.UserFields.Fields.Item("U_BaseLine").Value = objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string
                                objstocktransfer.Lines.Add()
                            End If
                            If CDbl(objMatrix.Columns.Item("8").Cells.Item(i).Specific.string) > 0 Then
                                Quantity = CDbl(objMatrix.Columns.Item("8").Cells.Item(i).Specific.string)
                                objstocktransfer.Lines.ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.string
                                objstocktransfer.Lines.Quantity = Quantity '* CDbl(UOMQty)
                                objstocktransfer.Lines.FromWarehouseCode = objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string ' objForm.Items.Item("25").Specific.string
                                objstocktransfer.Lines.WarehouseCode = objMatrix.Columns.Item("8A").Cells.Item(i).Specific.string
                                If objcc.Fields.Item(0).Value <> "" Then objstocktransfer.Lines.DistributionRule = objcc.Fields.Item(0).Value
                                If objcc.Fields.Item(1).Value <> "" Then objstocktransfer.Lines.DistributionRule2 = objcc.Fields.Item(1).Value
                                If objcc.Fields.Item(2).Value <> "" Then objstocktransfer.Lines.DistributionRule3 = objcc.Fields.Item(2).Value
                                If objcc.Fields.Item(3).Value <> "" Then objstocktransfer.Lines.DistributionRule4 = objcc.Fields.Item(3).Value
                                If objcc.Fields.Item(4).Value <> "" Then objstocktransfer.Lines.DistributionRule5 = objcc.Fields.Item(4).Value
                                objstocktransfer.Lines.UserFields.Fields.Item("U_Rem").Value = "QC"
                                objstocktransfer.Lines.UserFields.Fields.Item("U_BaseLine").Value = objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string
                                objstocktransfer.Lines.Add()
                            End If
                        End If
                    End If
                Next

                If objstocktransfer.Add() <> 0 Then
                    'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objAddOn.objApplication.SetStatusBarMessage("Stock Transfer: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    objAddOn.objApplication.MessageBox("Stock Transfer: " & objAddOn.objCompany.GetLastErrorDescription & "-" & objAddOn.objCompany.GetLastErrorCode,, "OK")
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objstocktransfer)
                    Return False
                Else
                    'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    DocEntry = objAddOn.objCompany.GetNewObjectKey()
                    objForm.Items.Item("34A").Specific.string = DocEntry
                    If objAddOn.HANA Then
                        DocNum = objAddOn.objGenFunc.getSingleValue("Select ""DocNum"" from OWTR where ""DocEntry""=" & DocEntry & "")
                        objForm.Items.Item("34").Specific.string = DocNum
                        'strSQL = "Update ""@MIPLQC"" set ""U_AccStk""='" & DocNum & "',""U_AccStkD""='" & DocEntry & "',""Status""='C' where ""DocEntry""='" & QCHeader.GetValue("DocEntry", 0) & "' "
                        'objrs.DoQuery(strSQL)
                    Else
                        DocNum = objAddOn.objGenFunc.getSingleValue("Select DocNum from OWTR where DocEntry=" & DocEntry & "")
                        objForm.Items.Item("34").Specific.string = DocNum
                        'strSQL = "Update [@MIPLQC] set U_AccStk='" & DocNum & "',U_AccStkD='" & DocEntry & "',Status='C' where DocEntry='" & QCHeader.GetValue("DocEntry", 0) & "' "
                        'objrs.DoQuery(strSQL)
                    End If
                    'objAddOn.objApplication.ActivateMenuItem("1304")
                    objAddOn.objApplication.StatusBar.SetText("Stock Transfer Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objstocktransfer)
                    Return True
                End If

            Catch ex As Exception
                'objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                objAddOn.objApplication.MessageBox(ex.Message, , "OK")
                Return False
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

    Private Function GetBatch_Serial(ByVal BatchSerial As String, ByVal DocEntry As String, ByVal ItemCode As String, ByVal Line As String, ByVal WhsCode As String) As SAPbobsCOM.Recordset
        Dim objrs As SAPbobsCOM.Recordset
        objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim objCombo As SAPbouiCOM.ComboBox
        Dim THeader As String = "", TLine As String = "", BaseType As String = "", UDFField As String = "", GetEntry As String = ""
        objCombo = objForm.Items.Item("8").Specific
        If objCombo.Selected.Value = "G" Then
            THeader = "OPDN"
            TLine = "PDN1"
            BaseType = "20"
            UDFField = "U_GRPONum"
        ElseIf objCombo.Selected.Value = "P" Or objCombo.Selected.Value = "R" Then
            THeader = "OIGN"
            TLine = "IGN1"
            BaseType = "59"
            UDFField = "U_REntry"
        ElseIf objCombo.Selected.Value = "T" Then
            THeader = "OWTR"
            TLine = "WTR1"
            BaseType = "67"
            UDFField = "U_StkEntry"
        Else
            Exit Function
        End If
        objrs = GetTransactions(objCombo.Selected.Value)
        If objrs.RecordCount > 0 Then
            'Dim GetValues() As String = {""}
            'For Rec As Integer = 0 To objrs.RecordCount - 1
            '    GetValues(Rec) = objrs.Fields.Item(0).Value.ToString
            '    objrs.MoveNext()
            'Next
            'Dim DocEntryList = (From gv In GetValues Select New String(gv)).ToList()
            GetEntry = objrs.Fields.Item(0).Value.ToString 'String.Join(",", DocEntryList)
        End If
        'FinalBatch = String.Join(",", TrackBatch)
        If objAddOn.HANA Then
            If BatchSerial = "N" Then
                'strSQL = "Select * from ("
                'strSQL += vbCrLf + "SELECT distinct I1.""BatchNum"" ""BatchSerial"",T1.""DocEntry"",T1.""ItemCode"",T4.""WhsCode"",T4.""Quantity"" as ""Qty"", I1.""Quantity"",T4.""Status"",T1.""LineNum"" from " & THeader & " T0 left join " & TLine & " T1 on T0.""DocEntry""=T1.""DocEntry"""
                'strSQL += vbCrLf + "left outer join IBT1 I1 on T1.""ItemCode""=I1.""ItemCode""   and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                'strSQL += vbCrLf + "left outer join OIBT T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""BatchNum""=T4.""BatchNum"" and I1.""WhsCode"" = T4.""WhsCode"""
                'strSQL += vbCrLf + ")A Where  A.""DocEntry"" = '" & DocEntry & "' and A.""ItemCode""='" & ItemCode & "' and A.""BatchSerial"" <>'' and A.""Status""=0 and A.""LineNum""='" & Line & "'  and A.""Qty"">0  and A.""WhsCode""='" & WhsCode & "'"
                strSQL = "SELECT A.""BatchNum"" as ""BatchSerial"",  SUM(A.""Quantity"") as ""Qty"" FROM ("
                strSQL += vbCrLf + "select T.""BatchNum"",  T.""Quantity"" from ibt1 T inner join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                strSQL += vbCrLf + "inner join " & TLine & " T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                strSQL += vbCrLf + "inner join " & THeader & " T3 on T2.""DocEntry""=T3.""DocEntry"""
                strSQL += vbCrLf + "where T.""BaseType""='" & BaseType & "' and T.""Direction""=0 and T3.""DocEntry""='" & DocEntry & "' and T.""ItemCode""='" & ItemCode & "' and T2.""LineNum""='" & Line & "' "
                'If FinalBatch <> "" Then
                '    strSQL += vbCrLf + "  and T.""BatchNum"" not in (" & FinalBatch & ")"
                'End If

                strSQL += vbCrLf + "UNION ALL"
                strSQL += vbCrLf + "select T.""BatchNum"", -T.""Quantity"" from ibt1 T inner join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                strSQL += vbCrLf + "inner join WTR1 T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                strSQL += vbCrLf + "inner join OWTR T3 on T2.""DocEntry""=T3.""DocEntry"""
                strSQL += vbCrLf + "where T.""BaseType""='67' and T.""Direction""=1  and T.""ItemCode""='" & ItemCode & "' "
                If GetEntry <> "" Then
                    strSQL += vbCrLf + "and T3.""" & UDFField & """='" & DocEntry & "' and  T2.""U_BaseLine""='" & Line & "'"
                Else
                    strSQL += vbCrLf + "and IFNULL(T3.""" & UDFField & """,'') is Null"
                End If
                strSQL += vbCrLf + ")A GROUP BY A.""BatchNum"" having SUM(A.""Quantity"") >0 "

            ElseIf BatchSerial = "Y" Then
                strSQL = "Select * from ("
                strSQL += vbCrLf + "SELECT distinct T4.""IntrSerial"" ""BatchSerial"",T1.""DocEntry"",T1.""ItemCode"", T4.""WhsCode"",T4.""Quantity"",T4.""Status"",T1.""LineNum"" from " & THeader & "  T0 inner join " & TLine & " T1 on T0.""DocEntry""=T1.""DocEntry"""
                strSQL += vbCrLf + "left outer join SRI1 I1 on T1.""ItemCode""=I1.""ItemCode""   and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                strSQL += vbCrLf + "left outer join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"" ) A "
                strSQL += vbCrLf + " Where A.""DocEntry"" = '" & DocEntry & "' and A.""ItemCode""='" & ItemCode & "' and A.""BatchSerial"" <>'' and A.""Status""=0 and A.""LineNum""='" & Line & "' and A.""WhsCode""='" & WhsCode & "'"
            Else
                Return Nothing
            End If
        Else
            If BatchSerial = "N" Then
                'strSQL = "Select * from ("
                'strSQL += vbCrLf + "SELECT distinct I1.BatchNum BatchSerial,T1.DocEntry,T1.ItemCode,T4.Quantity as Qty, T4.WhsCode,I1.Quantity,T4.Status,T1.LineNum from " & THeader & " T0 left join " & TLine & " T1 on T0.DocEntry=T1.DocEntry"
                'strSQL += vbCrLf + "left outer join IBT1 I1 on T1.ItemCode=I1.ItemCode   and (T1.DocEntry=I1.BaseEntry and T1.ObjType=I1.BaseType) and T1.LineNum=I1.BaseLinNum"
                'strSQL += vbCrLf + "left outer join OIBT T4 on T4.ItemCode=I1.ItemCode and I1.BatchNum=T4.BatchNum and I1.WhsCode = T4.WhsCode"
                'strSQL += vbCrLf + ")A Where  A. DocEntry  = '" & DocEntry & "' and A. ItemCode ='" & ItemCode & "' and A. BatchSerial  <>'' and A. Status =0 and A. LineNum ='" & Line & "'  and A. Qty >0 and A.WhsCode='" & WhsCode & "'"

                strSQL = "SELECT A.BatchNum BatchSerial,  SUM(A.Quantity) as Qty FROM ("
                strSQL += vbCrLf + "select T.BatchNum,  T.Quantity from ibt1 T inner join oibt T1 on T.ItemCode=T1.ItemCode and T.BatchNum=T1.BatchNum and T.WhsCode=T1.WhsCode"
                strSQL += vbCrLf + "inner join " & TLine & " T2 on T2.DocEntry=T.BaseEntry and T2.ItemCode=T.ItemCode and T2.LineNum=T.BaseLinNum"
                strSQL += vbCrLf + "inner join " & THeader & " T3 on T2.DocEntry=T3.DocEntry"
                strSQL += vbCrLf + "where T.BaseType='" & BaseType & "' and T.Direction=0 and T3.DocEntry='" & DocEntry & "' and T.ItemCode='" & ItemCode & "' and T2.LineNum='" & Line & "' "
                'If FinalBatch <> "" Then
                '    strSQL += vbCrLf + "  and T.BatchNum not in (" & FinalBatch & ")"
                'End If
                strSQL += vbCrLf + "UNION ALL"
                strSQL += vbCrLf + "select T.BatchNum, -T.Quantity from ibt1 T inner join oibt T1 on T.ItemCode=T1.ItemCode and T.BatchNum=T1.BatchNum and T.WhsCode=T1.WhsCode"
                strSQL += vbCrLf + "inner join WTR1 T2 on T2.DocEntry=T.BaseEntry and T2.ItemCode=T.ItemCode and T2.LineNum=T.BaseLinNum"
                strSQL += vbCrLf + "inner join OWTR T3 on T2.DocEntry=T3.DocEntry"
                strSQL += vbCrLf + "where T.BaseType='67' and T.Direction=1  and T.ItemCode='" & ItemCode & "' "
                If GetEntry <> "" Then
                    strSQL += vbCrLf + "and T3." & UDFField & "='" & DocEntry & "' and  T2.U_BaseLine='" & Line & "'"
                Else
                    strSQL += vbCrLf + "and ISNULL(T3." & UDFField & ",'') is Null"
                End If
                strSQL += vbCrLf + ")A GROUP BY A.BatchNum having SUM(A.Quantity) >0 "
            ElseIf BatchSerial = "Y" Then
                strSQL = "Select * from ("
                strSQL += vbCrLf + "SELECT distinct T4.IntrSerial BatchSerial,T1.DocEntry,T1.ItemCode, T4.Quantity,T4.WhsCode,T4.Status,T1.LineNum from " & THeader & "  T0 inner join " & TLine & " T1 on T0.DocEntry=T1.DocEntry"
                strSQL += vbCrLf + "left outer join SRI1 I1 on T1.ItemCode=I1.ItemCode   and (T1.DocEntry=I1.BaseEntry and T1.ObjType=I1.BaseType) and T1.LineNum=I1.BaseLinNum"
                strSQL += vbCrLf + "left outer join OSRI T4 on T4.ItemCode=I1.ItemCode and I1.SysSerial=T4.SysSerial and I1.WhsCode = T4.WhsCode ) A "
                strSQL += vbCrLf + " Where A. DocEntry  = '" & DocEntry & "' and A. ItemCode ='" & ItemCode & "' and A. BatchSerial  <>'' and A. Status =0 and A. LineNum ='" & Line & "' and A.WhsCode='" & WhsCode & "'"
            Else
                Return Nothing
            End If
        End If
        objrs.DoQuery(strSQL)
        Return objrs
    End Function

    Private Function Validate_Batch_Serial() As Boolean
        Try
            Dim objBatch_serial As SAPbobsCOM.Recordset
            objCombo = objForm.Items.Item("8").Specific
            Dim ErrCount As Integer = 0
            Dim TranDocEntry As String = "", Serial As String = "", Batch As String = ""
            If objCombo.Selected.Value = "G" Then
                TranDocEntry = objForm.Items.Item("13B").Specific.string
            ElseIf objCombo.Selected.Value = "P" Then
                TranDocEntry = objForm.Items.Item("51B").Specific.string
            ElseIf objCombo.Selected.Value = "T" Then
                TranDocEntry = objForm.Items.Item("23B").Specific.string
            ElseIf objCombo.Selected.Value = "R" Then
                TranDocEntry = objForm.Items.Item("51B").Specific.string
            End If
            For i As Integer = 1 To objMatrix.VisualRowCount
                If objAddOn.HANA Then
                    Serial = objAddOn.objGenFunc.getSingleValue("select ""ManSerNum"" from OITM WHERE ""ItemCode""='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.string & "'")
                    Batch = objAddOn.objGenFunc.getSingleValue("select ""ManBtchNum"" from OITM WHERE ""ItemCode""='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.string & "'")
                Else
                    Serial = objAddOn.objGenFunc.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.string & "'")
                    Batch = objAddOn.objGenFunc.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.string & "'")
                End If

                If Batch = "Y" And Serial = "N" Then
                    objBatch_serial = GetBatch_Serial("N", TranDocEntry, objMatrix.Columns.Item("1").Cells.Item(i).Specific.string, objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string, objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string)
                    If objBatch_serial.RecordCount = 0 Then
                        ErrCount += 1
                        objAddOn.objApplication.StatusBar.SetText("Batch or Serial not available. ItemCode " & objMatrix.Columns.Item("1").Cells.Item(i).Specific.string & " Please remove the line " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                ElseIf Batch = "N" And Serial = "Y" Then
                    objBatch_serial = GetBatch_Serial("Y", TranDocEntry, objMatrix.Columns.Item("1").Cells.Item(i).Specific.string, objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string, objMatrix.Columns.Item("3B").Cells.Item(i).Specific.string)
                    If objBatch_serial.RecordCount = 0 Then
                        ErrCount += 1
                        objAddOn.objApplication.StatusBar.SetText("Batch or Serial not available. ItemCode " & objMatrix.Columns.Item("1").Cells.Item(i).Specific.string & " Please remove the line " & i, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                End If
            Next
            If ErrCount > 0 Then
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End Try
    End Function

    Private Function GetTransactions(ByVal TranValue As String)
        Try
            Dim objrs As SAPbobsCOM.Recordset
            Dim str_sql As String
            If TranValue = "G" Then
                If objAddOn.HANA Then
                    str_sql = "select distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from OWTR T0 left join ""@MIPLQC"" T1 on T0.""U_GRPONum"" = T1.""U_GRNEntry"" where  T0.""U_GRPONum"" ='" & objForm.Items.Item("13B").Specific.String & "';"
                Else
                    str_sql = "select distinct T0.DocEntry,T0.DocNum,T0.DocDate from OWTR T0 left join [@MIPLQC] T1 on T0.U_GRPONum = T1.U_GRNEntry where  T0.U_GRPONum ='" & objForm.Items.Item("13B").Specific.String & "'"
                End If
            ElseIf TranValue = "P" Then
                If objAddOn.HANA Then
                    str_sql = "select distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from OWTR T0 left join ""@MIPLQC"" T1 on T0.""U_PORNum"" = T1.""U_ProdEntry""  and T1.""U_GRNum""=T0.""U_REntry"" where  T0.""U_PORNum""='" & objForm.Items.Item("15B").Specific.String & "';"
                Else
                    str_sql = "select distinct T0. DocEntry ,T0. DocNum ,T0. DocDate  from OWTR T0 left join  [@MIPLQC]  T1 on T0. U_PORNum  = T1. U_ProdEntry   and T1. U_GRNum =T0. U_REntry  where  T0. U_PORNum ='" & objForm.Items.Item("15B").Specific.String & "'"
                End If

            ElseIf TranValue = "T" Then
                If objAddOn.HANA Then
                    str_sql = "select distinct T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"" from OWTR T0 left join ""@MIPLQC"" T1 on T0.""U_StkEntry"" = T1.""U_TransEntry"" where  T0.""U_StkEntry"" ='" & objForm.Items.Item("23B").Specific.String & "';"
                Else
                    str_sql = "select distinct T0.DocEntry,T0.DocNum,T0.DocDate from OWTR T0 left join [@MIPLQC] T1 on T0.U_StkEntry = T1.U_TransEntry where  T0.U_StkEntry ='" & objForm.Items.Item("23B").Specific.String & "'"
                End If
            ElseIf TranValue = "R" Then
                If objAddOn.HANA Then
                    str_sql = "select distinct T0.""DocEntry"",T0.""U_GRPONum"",T0.""DocNum"",T0.""DocDate"" from OWTR T0 left join ""@MIPLQC"" T1 on T0.""U_REntry"" = T1.""U_GRNum"" where  T0.""U_REntry"" ='" & objForm.Items.Item("51B").Specific.String & "';"
                Else
                    str_sql = "select distinct T0.DocEntry,T0.U_GRPONum,T0.DocNum,T0.DocDate from OWTR T0 left join [@MIPLQC] T1 on T0.U_REntry = T1.U_GRNum where  T0.U_REntry ='" & objForm.Items.Item("51B").Specific.String & "'"
                End If
            Else
                Return Nothing
            End If
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(str_sql)
            Return objrs
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function GetCostCenter() As SAPbobsCOM.Recordset
        Try
            Dim objrs As SAPbobsCOM.Recordset
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim THeader As String = "", TLine As String = "", TranDocEntry As String
            objCombo = objForm.Items.Item("8").Specific
            If objCombo.Selected.Value = "G" Then
                THeader = "OPDN"
                TLine = "PDN1"
                TranDocEntry = objForm.Items.Item("13B").Specific.string
            ElseIf objCombo.Selected.Value = "P" Or objCombo.Selected.Value = "R" Then
                THeader = "OIGN"
                TLine = "IGN1"
                TranDocEntry = objForm.Items.Item("51B").Specific.string
            ElseIf objCombo.Selected.Value = "T" Then
                THeader = "OWTR"
                TLine = "WTR1"
                TranDocEntry = objForm.Items.Item("23B").Specific.string
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

    Private Function Create_GoodsIssue(ByVal FormUID As String) As Boolean
        Try
            Dim objGoodsIssue As SAPbobsCOM.Documents
            Dim objrs, objcc As SAPbobsCOM.Recordset
            Dim series, Branch, BranchEnabled, Serial, Batchs, Docentry, DocNum As String
            Dim Quantity As Double
            Dim QCDocNum As Long
            Dim Flag As Boolean = False
            If DisposableGoodsIssue = "N" Then Return True
            For i As Integer = 1 To objMatrix.VisualRowCount
                If objMatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" And CDbl(objMatrix.Columns.Item("9B").Cells.Item(i).Specific.String) > 0 Then
                    Flag = True : Exit For
                End If
            Next
            objCombo = objForm.Items.Item("8").Specific
            If Flag = False Or objCombo.Selected.Value = "T" Then Return True
            objGoodsIssue = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objMatrix = objForm.Items.Item("20").Specific
            objAddOn.objApplication.StatusBar.SetText("Disposable Goods Issue Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGoodsIssue.DocDate = Now.Date 'DocDate
            ' objGoodsIssue.DocDueDate = DocDueDate
            objGoodsIssue.TaxDate = Now.Date
            If objAddOn.HANA Then
                BranchEnabled = objAddOn.objGenFunc.getSingleValue("select ""MltpBrnchs"" from OADM")
            Else
                BranchEnabled = objAddOn.objGenFunc.getSingleValue("select MltpBrnchs from OADM")
            End If
            If BranchEnabled = "Y" Then
                If objAddOn.HANA Then
                    series = objAddOn.objGenFunc.getSingleValue("select ""Series"" From NNM1 where ""ObjectCode""='60' and ""Indicator""=(select Top 1 ""Indicator""  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                        " and ""BPLId""=(select Top 1 ""BPLid"" from OWHS where ""WhsCode""='" & objMatrix.Columns.Item("3B").Cells.Item(1).Specific.String & "') ")
                    Branch = objAddOn.objGenFunc.getSingleValue("select Top 1 ""BPLid"" from OWHS where ""WhsCode""='" & objMatrix.Columns.Item("3B").Cells.Item(1).Specific.String & "'")
                Else
                    series = objAddOn.objGenFunc.getSingleValue("select Series From NNM1 where ObjectCode='60' and Indicator=(select Top 1 Indicator  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between F_RefDate and T_RefDate) " &
                                                                                        " and BPLId=(select Top 1 BPLid from OWHS where WhsCode='" & objMatrix.Columns.Item("3B").Cells.Item(1).Specific.String & "') ")
                    Branch = objAddOn.objGenFunc.getSingleValue("select Top 1 BPLid from OWHS where WhsCode='" & objMatrix.Columns.Item("3B").Cells.Item(1).Specific.String & "'")
                End If
                If series <> "" Then
                    objGoodsIssue.Series = series
                End If
                objGoodsIssue.BPL_IDAssignedToInvoice = Branch
            End If
            'objGoodsIssue.Reference1 = objForm.Items.Item("txtvrefno").Specific.string
            QCHeader = objForm.DataSources.DBDataSources.Item("@MIPLQC")
            QCDocNum = objForm.BusinessObject.GetNextSerialNumber(objForm.Items.Item("21").Specific.Selected.value, Formtype)
            objGoodsIssue.Comments = "Auto Posted-Disposable from QC Addon ->" & Now.ToString 'objform.Items.Item("txtremark").Specific.string 
            objGoodsIssue.UserFields.Fields.Item("U_QCNum").Value = CStr(QCDocNum)
            objGoodsIssue.UserFields.Fields.Item("U_QCEntry").Value = QCHeader.GetValue("DocEntry", 0)
            If objCombo.Selected.Value = "G" Then
                objGoodsIssue.UserFields.Fields.Item("U_GRPONum").Value = objForm.Items.Item("13B").Specific.string
                objGoodsIssue.UserFields.Fields.Item("U_GRNEntry").Value = objForm.Items.Item("13").Specific.string
            ElseIf objCombo.Selected.Value = "P" Then
                objGoodsIssue.UserFields.Fields.Item("U_PORNum").Value = objForm.Items.Item("15B").Specific.string
                objGoodsIssue.UserFields.Fields.Item("U_ProdEntry").Value = objForm.Items.Item("15").Specific.string
                objGoodsIssue.UserFields.Fields.Item("U_REntry").Value = objForm.Items.Item("51B").Specific.string
                objGoodsIssue.UserFields.Fields.Item("U_GREntry").Value = objForm.Items.Item("51").Specific.string
                'ElseIf objCombo.Selected.Value = "T" Then
                '    objGoodsIssue.UserFields.Fields.Item("U_StkEntry").Value = objForm.Items.Item("23B").Specific.string
                '    objGoodsIssue.UserFields.Fields.Item("U_StkNum").Value = objForm.Items.Item("23").Specific.string
            ElseIf objCombo.Selected.Value = "R" Then
                objGoodsIssue.UserFields.Fields.Item("U_REntry").Value = objForm.Items.Item("51B").Specific.string
                objGoodsIssue.UserFields.Fields.Item("U_GREntry").Value = objForm.Items.Item("51").Specific.string
            End If
            For i As Integer = 1 To objMatrix.VisualRowCount
                If objMatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" And CDbl(objMatrix.Columns.Item("9B").Cells.Item(i).Specific.String) > 0 Then
                    Quantity = CDbl(objMatrix.Columns.Item("9B").Cells.Item(i).Specific.String)
                    objGoodsIssue.Lines.ItemCode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.String
                    objGoodsIssue.Lines.Quantity = Quantity
                    objGoodsIssue.Lines.WarehouseCode = objMatrix.Columns.Item("3B").Cells.Item(i).Specific.String
                    objcc = GetCostCenter()
                    If objcc.Fields.Item(0).Value <> "" Then objGoodsIssue.Lines.CostingCode = objcc.Fields.Item(0).Value
                    If objcc.Fields.Item(1).Value <> "" Then objGoodsIssue.Lines.CostingCode2 = objcc.Fields.Item(1).Value
                    If objcc.Fields.Item(2).Value <> "" Then objGoodsIssue.Lines.CostingCode3 = objcc.Fields.Item(2).Value
                    If objcc.Fields.Item(3).Value <> "" Then objGoodsIssue.Lines.CostingCode4 = objcc.Fields.Item(3).Value
                    If objcc.Fields.Item(4).Value <> "" Then objGoodsIssue.Lines.CostingCode5 = objcc.Fields.Item(4).Value
                    objGoodsIssue.Lines.UserFields.Fields.Item("U_BaseLine").Value = objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string
                    Dim BatchQty As Double = 0, BatchQty1 As Double = 0
                    If objAddOn.HANA Then
                        Serial = objAddOn.objGenFunc.getSingleValue("select ""ManSerNum"" from OITM WHERE ""ItemCode""='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.String & "'")
                        Batchs = objAddOn.objGenFunc.getSingleValue("select ""ManBtchNum"" from OITM WHERE ""ItemCode""='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.String & "'")
                    Else
                        Serial = objAddOn.objGenFunc.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.String & "'")
                        Batchs = objAddOn.objGenFunc.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.String & "'")
                    End If
                    If Batchs = "Y" And Serial = "N" Then
                        Dim InvEntry, GIEntry As String
                        strSQL = "Select ""DocEntry"" from OWTR where ""U_GRPONum""='" & objForm.Items.Item("13B").Specific.String & "' or ""U_REntry""='" & objForm.Items.Item("51B").Specific.String & "' "
                        objrs.DoQuery(strSQL)
                        For IEntry As Integer = 0 To objrs.RecordCount - 1
                            If IEntry = 0 Then
                                InvEntry = "'" & objrs.Fields.Item("DocEntry").Value & "'"
                            Else
                                InvEntry = InvEntry & ",'" & objrs.Fields.Item("DocEntry").Value & "'"
                            End If
                            objrs.MoveNext()
                        Next
                        strSQL = "Select ""DocEntry"" from OIGE where ""U_GRPONum""='" & objForm.Items.Item("13B").Specific.String & "' or ""U_REntry""='" & objForm.Items.Item("51B").Specific.String & "' "
                        objrs.DoQuery(strSQL)
                        For IEntry As Integer = 0 To objrs.RecordCount - 1
                            If IEntry = 0 Then
                                GIEntry = "'" & objrs.Fields.Item("DocEntry").Value & "'"
                            Else
                                GIEntry = GIEntry & ",'" & objrs.Fields.Item("DocEntry").Value & "'"
                            End If
                            objrs.MoveNext()
                        Next
                        Dim THeader, TLine, BaseType, Entry As String
                        If objCombo.Selected.Value = "G" Then
                            THeader = "OPDN"
                            TLine = "PDN1"
                            BaseType = "20"
                            Entry = objForm.Items.Item("13B").Specific.String
                        ElseIf objCombo.Selected.Value = "P" Or objCombo.Selected.Value = "R" Then
                            THeader = "OIGN"
                            TLine = "IGN1"
                            BaseType = "59"
                            Entry = objForm.Items.Item("51B").Specific.String
                            'ElseIf objCombo.Selected.Value = "T" Then
                            '    THeader = "OWTR"
                            '    TLine = "WTR1"
                            '    BaseType = "67"
                            '    UDFField = "U_StkEntry"
                        Else
                            Exit Function
                        End If
                        If objAddOn.HANA Then

                            strSQL = "SELECT A.""BatchNum"" as ""BatchSerial"",  SUM(A.""Quantity"") as ""Qty"" FROM ("
                            strSQL += vbCrLf + "select T.""BatchNum"",  T.""Quantity"" from ibt1 T inner join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                            strSQL += vbCrLf + "inner join " & TLine & " T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                            strSQL += vbCrLf + "inner join " & THeader & " T3 on T2.""DocEntry""=T3.""DocEntry"""
                            strSQL += vbCrLf + "where T.""BaseType""='" & BaseType & "' and T.""Direction""=0 and T3.""DocEntry""='" & Entry & "' and T.""ItemCode""='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.String & "' and T2.""LineNum""='" & objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string & "'"
                            strSQL += vbCrLf + "UNION ALL"
                            strSQL += vbCrLf + "select T.""BatchNum"", -T.""Quantity"" from ibt1 T inner join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                            strSQL += vbCrLf + "inner join WTR1 T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                            strSQL += vbCrLf + "inner join OWTR T3 on T2.""DocEntry""=T3.""DocEntry"""
                            strSQL += vbCrLf + "where T.""BaseType""='67' and T.""Direction""=1 and T.""ItemCode""='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.String & "' and T2.""U_BaseLine""='" & objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string & "' "
                            If InvEntry <> "" Then
                                strSQL += vbCrLf + "and T3.""DocEntry"" in (" & InvEntry & ")"
                            Else
                                strSQL += vbCrLf + "and IFNULL(T3.""DocEntry"",'') is Null"
                            End If
                            strSQL += vbCrLf + "UNION ALL"
                            strSQL += vbCrLf + "select T.""BatchNum"", -T.""Quantity"" from ibt1 T inner join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                            strSQL += vbCrLf + "inner join IGE1 T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                            strSQL += vbCrLf + "inner join OIGE T3 on T2.""DocEntry""=T3.""DocEntry"""
                            strSQL += vbCrLf + "where T.""BaseType""='60' and T.""Direction""=1 and T.""ItemCode""='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.String & "' and T2.""U_BaseLine""='" & objMatrix.Columns.Item("0A").Cells.Item(i).Specific.string & "' "
                            If GIEntry <> "" Then
                                strSQL += vbCrLf + "and T3.""DocEntry"" in (" & GIEntry & ")"
                            Else
                                strSQL += vbCrLf + "and IFNULL(T3.""DocEntry"",'') is Null"
                            End If
                            strSQL += vbCrLf + ")A "
                            strSQL += vbCrLf + "GROUP BY A.""BatchNum"" having SUM(A.""Quantity"") >0"

                        Else


                        End If

                        objrs.DoQuery(strSQL)
                        Dim BQty As Double = 0, PendQty As Double = 0, TotBatchQty As Double = 0
                        BQty = Quantity '2
                        If objrs.RecordCount > 0 Then
                            For j As Integer = 0 To objrs.RecordCount - 1
                                If (BQty - TotBatchQty) - CDbl(objrs.Fields.Item("Qty").Value) > 0 Then
                                    PendQty = CDbl(objrs.Fields.Item("Qty").Value)
                                Else
                                    PendQty = BQty - TotBatchQty
                                End If
                                objGoodsIssue.Lines.BatchNumbers.BatchNumber = CStr(objrs.Fields.Item("BatchSerial").Value)
                                objGoodsIssue.Lines.BatchNumbers.Quantity = PendQty
                                objGoodsIssue.Lines.BatchNumbers.Add()
                                TotBatchQty += PendQty
                                If BQty - TotBatchQty > 0 Then
                                    objrs.MoveNext()
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                    ElseIf Batchs = "N" And Serial = "Y" Then
                        If objAddOn.HANA Then
                            strSQL = "Select * from ("
                            strSQL += vbCrLf + "SELECT distinct T1.""DocEntry"",T4.""IntrSerial"" ""BatchSerial"",T4.""WhsCode"",T1.""ItemCode"", T4.""Quantity"",T4.""Status"" from OPDN T0 inner join PDN1 T1 on T0.""DocEntry""=T1.""DocEntry"""
                            strSQL += vbCrLf + "left outer join SRI1 I1 on T1.""ItemCode""=I1.""ItemCode""   and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                            strSQL += vbCrLf + "left outer join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"" ) A "
                            strSQL += vbCrLf + " Where A.""DocEntry""='" & objForm.Items.Item("13B").Specific.String & "' and A.""DocEntry"" in (2)  and A.""ItemCode""='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.String & "' and  A.""WhsCode""='" & objMatrix.Columns.Item("3B").Cells.Item(i).Specific.String & "' and A.""BatchSerial"" <>'' and A.""Status""=0"
                        Else
                            strSQL = "Select * from ("
                            strSQL += vbCrLf + "SELECT distinct T1.DocEntry,T4.IntrSerial BatchSerial,T4.WhsCode,T1.ItemCode, T4.Quantity,T4.Status from OPDN T0 inner join PDN1 T1 on T0.DocEntry=T1.DocEntry"
                            strSQL += vbCrLf + "left outer join SRI1 I1 on T1.ItemCode=I1.ItemCode   and (T1.DocEntry=I1.BaseEntry and T1.ObjType=I1.BaseType) and T1.LineNum=I1.BaseLinNum"
                            strSQL += vbCrLf + "left outer join OSRI T4 on T4.ItemCode=I1.ItemCode and I1.SysSerial=T4.SysSerial and I1.WhsCode = T4.WhsCode ) A "
                            strSQL += vbCrLf + " Where A.DocEntry='" & objForm.Items.Item("13B").Specific.String & "' and A.DocEntry in (2)  and A.ItemCode='" & objMatrix.Columns.Item("1").Cells.Item(i).Specific.String & "' and  A.WhsCode='" & objMatrix.Columns.Item("3B").Cells.Item(i).Specific.String & "' and A.BatchSerial <>'' and A.Status=0"
                        End If
                        objrs.DoQuery(strSQL)
                        Dim SQty As Double = 0, TotSerialQty As Double = 0
                        SQty = Quantity
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
                    Else
                    End If
                    objGoodsIssue.Lines.Add()
                End If
            Next i

            Dim Retval As Integer
            Retval = objGoodsIssue.Add()
            If Retval <> 0 Then
                'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objAddOn.objApplication.StatusBar.SetText("Goods Issue : " & objAddOn.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objGoodsIssue)
                Return False
            Else
                'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                Docentry = objAddOn.objCompany.GetNewObjectKey()
                objForm.Items.Item("34C").Specific.string = Docentry
                If objAddOn.HANA Then
                    DocNum = objAddOn.objGenFunc.getSingleValue("Select ""DocNum"" from OIGE where ""DocEntry""=" & Docentry & "")
                    objForm.Items.Item("34B").Specific.string = DocNum
                    'strSQL = "Update ""@MIPLQC"" set ""U_AccStk""='" & DocNum & "',""U_AccStkD""='" & DocEntry & "',""Status""='C' where ""DocEntry""='" & QCHeader.GetValue("DocEntry", 0) & "' "
                    'objrs.DoQuery(strSQL)
                Else
                    DocNum = objAddOn.objGenFunc.getSingleValue("Select DocNum from OIGE where DocEntry=" & Docentry & "")
                    objForm.Items.Item("34B").Specific.string = DocNum
                    'strSQL = "Update [@MIPLQC] set U_AccStk='" & DocNum & "',U_AccStkD='" & DocEntry & "',Status='C' where DocEntry='" & QCHeader.GetValue("DocEntry", 0) & "' "
                    'objrs.DoQuery(strSQL)
                End If
                objAddOn.objApplication.StatusBar.SetText("Disposable Goods Issue Created Successfully!... ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objGoodsIssue)
                GC.Collect()
                Return True
            End If


        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage("Disposable_GoodsIssue: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    Private Sub SendAlertMessageToUsers(ByVal UserCode As String, ByVal QCEntry As String, ByVal TranEntry As String, ByVal ObjTran As String, ByVal Message As String)
        Dim oMessage As Message
        Dim pMessageDataColumns As MessageDataColumns
        Dim pMessageDataColumn As MessageDataColumn
        Dim oLines As MessageDataLines
        Dim oLine As MessageDataLine
        Dim oRecipientCollection As RecipientCollection
        Try
            Dim Row As Integer = 0
            Dim oMessageService As MessagesService
            oMessageService = objAddOn.objCompany.GetCompanyService.GetBusinessService(ServiceTypes.MessagesService)
            oMessage = oMessageService.GetDataInterface(MessagesServiceDataInterfaces.msdiMessage)
            oMessage.Subject = "Quality Check Status"
            oRecipientCollection = oMessage.RecipientCollection
            oRecipientCollection.Add()
            oMessage.Text = Message ' "QC Completed for the transaction"

            Dim UserName As String = objAddOn.objGenFunc.getSingleValue("select ""USER_CODE"" from OUSR where ""USERID""='" & UserCode & "'")
            For i As Integer = 0 To oRecipientCollection.Count - 1
                oRecipientCollection.Item(i).SendInternal = BoYesNoEnum.tYES
                oRecipientCollection.Item(i).UserCode = UserName '"srmprof1"
                pMessageDataColumns = oMessage.MessageDataColumns
                pMessageDataColumn = pMessageDataColumns.Add()
                pMessageDataColumn.ColumnName = "QC Document Entry"
                pMessageDataColumn.Link = BoYesNoEnum.tYES
                oLines = pMessageDataColumn.MessageDataLines()
                oLine = oLines.Add()
                'set the line value
                oLine.Value = QCEntry
                oLine.Object = "MIPLQC"  'Object Type
                oLine.ObjectKey = QCEntry

                pMessageDataColumn = pMessageDataColumns.Add()
                pMessageDataColumn.ColumnName = "Transaction Entry"
                pMessageDataColumn.Link = BoYesNoEnum.tYES
                oLines = pMessageDataColumn.MessageDataLines()
                oLine = oLines.Add()
                oLine.Value = TranEntry
                oLine.Object = ObjTran 'Object Type
                oLine.ObjectKey = TranEntry

                'pMessageDataColumn = pMessageDataColumns.Add()
                'pMessageDataColumn.ColumnName = "UserName"
                'oLines = pMessageDataColumn.MessageDataLines()
                'oLine = oLines.Add()
                'oLine.Value = objAddOn.objCompany.UserName
                ''send the message
                oMessageService.SendMessage(oMessage)
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class