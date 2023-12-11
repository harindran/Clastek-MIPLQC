Public Class ClsSubPO
    Public Const Formtype = "SUBCTPO"
    Dim objSubPOform As SAPbouiCOM.Form
    Dim ObjQCForm As SAPbouiCOM.Form

    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        objSubPOform = objAddOn.objApplication.Forms.Item(FormUID)
        If pVal.BeforeAction Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                   ' CreateButton(FormUID)
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "lnk_qca" Then
                        Try
                            If objSubPOform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                objAddOn.objApplication.Menus.Item("MIQCACT").Activate()
                                ObjQCForm = objAddOn.objApplication.Forms.ActiveForm
                                'ObjQCForm = objAddOn.objApplication.Forms.GetForm("MIQCACT", 1)
                                ObjQCForm.Freeze(True)
                                ObjQCForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                ObjQCForm.Items.Item("15").Enabled = True
                                ObjQCForm.Items.Item("15").Specific.String = objSubPOform.Items.Item("t_qcanum").Specific.String
                                ObjQCForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                ObjQCForm.Freeze(False)
                            End If
                        Catch ex As Exception
                            ObjQCForm.Freeze(False)
                            ObjQCForm = Nothing
                        End Try
                    End If
            End Select
        Else
            Try
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        'If pVal.ItemUID = "1" Then

                        'End If
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If objSubPOform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.ItemUID <> "" Then
                                If objSubPOform.Items.Item("t_qcanum").Specific.String <> "" Then
                                    objSubPOform.Items.Item("txtbitem").Enabled = False
                                    objSubPOform.Items.Item("t_qcanum").Enabled = False
                                    objSubPOform.Items.Item("SQty").Enabled = False
                                End If
                            End If
                        End If

                End Select
            Catch ex As Exception

            End Try
        End If
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objSubPOform = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
            If BusinessObjectInfo.BeforeAction = True Then
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                End Select
            Else
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        Try
                            If objSubPOform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                                If objSubPOform.Items.Item("t_qcanum").Specific.String = "" Then Exit Sub
                                Dim GREntry As String = ""
                                Dim objmatrix, objQCAMatrix As SAPbouiCOM.Matrix
                                Dim objSelect As SAPbouiCOM.CheckBox
                                Dim objrs As SAPbobsCOM.Recordset
                                Try
                                    If objSubPOform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objSubPOform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then Exit Sub
                                    'If objSubPOform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    objmatrix = objSubPOform.Items.Item("MtxinputN").Specific
                                        ObjQCForm = objAddOn.objApplication.Forms.GetForm("MIQCACT", 0)
                                        objQCAMatrix = ObjQCForm.Items.Item("20").Specific
                                        objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        If objSubPOform.Items.Item("t_qcanum").Specific.String <> "" Then
                                        GREntry = objSubPOform.DataSources.DBDataSources.Item("@MIPL_OPOR").GetValue("DocEntry", 0)
                                    End If
                                        For j = 1 To objQCAMatrix.VisualRowCount
                                            objSelect = objQCAMatrix.Columns.Item("0A").Cells.Item(j).Specific
                                            If objSelect.Checked = True Then
                                                objQCAMatrix.Columns.Item("11").Cells.Item(j).Specific.String = GREntry
                                            End If
                                        Next
                                    'End If
                                    objrs = Nothing
                                    GC.Collect()
                                Catch ex As Exception
                                    objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                End Try
                            End If
                        Catch ex As Exception
                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        Try
                            'objSubPOform.ActiveItem = "deldate"
                            If objSubPOform.Items.Item("t_qcanum").Specific.String <> "" Then
                                objSubPOform.Items.Item("t_qcanum").Enabled = False
                                objSubPOform.Update()
                            End If
                        Catch ex As Exception
                        End Try
                End Select
            End If

        Catch ex As Exception
            'objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub CreateButton(ByVal FormUID As String)
        Try
            Dim objButton As SAPbouiCOM.StaticText
            Dim objItem As SAPbouiCOM.Item
            objSubPOform = objAddOn.objApplication.Forms.Item(FormUID)
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

    End Sub

    Sub RightClickEvent(ByRef EventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            objSubPOform = objAddOn.objApplication.Forms.Item(objSubPOform.UniqueID)
            'objMatrix = objSubPOform.Items.Item("20").Specific
            If EventInfo.BeforeAction Then
                Select Case EventInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK
                        Dim oMenuItem As SAPbouiCOM.MenuItem
                        If EventInfo.ItemUID = "" Then
                            objSubPOform.EnableMenu("5907", True)
                            oMenuItem = objAddOn.objApplication.Menus.Item("5907")
                            oMenuItem.Enabled = True
                            'objSubPOform.EnableMenu("5907", True)
                        Else
                            'objSubPOform.EnableMenu("5907", False)
                        End If
                        'Select Case EventInfo.ItemUID
                        '    Case "20"
                        '        If EventInfo.ColUID = "0" Or EventInfo.ColUID = "1" And objSubPOform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        '            objSubPOform.EnableMenu("1293", True)
                        '        Else
                        '            objSubPOform.EnableMenu("1293", False)
                        '        End If
                        'End Select
                End Select
            Else

            End If

        Catch ex As Exception
            'objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
End Class
