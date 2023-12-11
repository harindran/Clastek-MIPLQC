Public Class ClsARInvoice
    Public Const Formtype = "133"
    Dim objARform As SAPbouiCOM.Form
    Dim ObjQCForm As SAPbouiCOM.Form

    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        objARform = objAddOn.objApplication.Forms.Item(FormUID)
        If pVal.BeforeAction Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If objARform.Items.Item("t_qcanum").Specific.String = "" Then Exit Sub
                    If pVal.ItemUID = "1" Then
                        Dim objMatrix As SAPbouiCOM.Matrix
                        objMatrix = objARform.Items.Item("38").Specific
                        For i As Integer = 1 To objMatrix.VisualRowCount
                            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.String <> "" Then
                                If objMatrix.Columns.Item("1").Cells.Item(i).Specific.String <> objMatrix.Columns.Item("U_Item").Cells.Item(i).Specific.String Or objMatrix.Columns.Item("24").Cells.Item(i).Specific.String <> objMatrix.Columns.Item("U_Whse").Cells.Item(i).Specific.String Or CDbl(objMatrix.Columns.Item("11").Cells.Item(i).Specific.String) <> CDbl(objMatrix.Columns.Item("U_Qty").Cells.Item(i).Specific.String) Then
                                    objAddOn.objApplication.StatusBar.SetText("Data mismatch on line " & i & " Please re-generate...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                            End If
                        Next
                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    CreateButton(FormUID)
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "lnk_qca" Then
                        Try
                            If objARform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                objAddOn.objApplication.Menus.Item("MIQCACT").Activate()
                                ObjQCForm = objAddOn.objApplication.Forms.ActiveForm
                                'ObjQCForm = objAddOn.objApplication.Forms.GetForm("MIQCACT", 1)
                                ObjQCForm.Freeze(True)
                                ObjQCForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                ObjQCForm.Items.Item("15").Enabled = True
                                ObjQCForm.Items.Item("15").Specific.String = objARform.Items.Item("t_qcanum").Specific.String
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
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                End Select
            Catch ex As Exception

            End Try
        End If
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objARform = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
            If BusinessObjectInfo.BeforeAction = True Then
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                End Select
            Else
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        Try
                            If objARform.Items.Item("t_qcanum").Specific.String = "" Then Exit Sub
                            Dim GREntry As String = ""
                            Dim objmatrix, objQCAMatrix As SAPbouiCOM.Matrix
                            Dim objSelect As SAPbouiCOM.CheckBox
                            Dim objrs As SAPbobsCOM.Recordset
                            Try
                                If objARform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objARform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then Exit Sub
                                If objARform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    objmatrix = objARform.Items.Item("38").Specific
                                    ObjQCForm = objAddOn.objApplication.Forms.GetForm("MIQCACT", 1)
                                    objQCAMatrix = ObjQCForm.Items.Item("20").Specific
                                    objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    If objARform.Items.Item("t_qcanum").Specific.String <> "" Then
                                        GREntry = objARform.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0)
                                    End If
                                    For j = 1 To objQCAMatrix.VisualRowCount
                                        objSelect = objQCAMatrix.Columns.Item("0A").Cells.Item(j).Specific
                                        If objSelect.Checked = True Then
                                            objQCAMatrix.Columns.Item("11").Cells.Item(j).Specific.String = GREntry
                                        End If
                                    Next
                                End If
                                objrs = Nothing
                                GC.Collect()
                            Catch ex As Exception
                                objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            End Try
                        Catch ex As Exception
                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                        If objARform.Items.Item("t_qcanum").Specific.String <> "" Then
                            objARform.Items.Item("t_qcanum").Enabled = False
                        End If
                End Select
            End If

        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub CreateButton(ByVal FormUID As String)
        Try
            Dim objButton As SAPbouiCOM.StaticText
            Dim objItem As SAPbouiCOM.Item
            objARform = objAddOn.objApplication.Forms.Item(FormUID)
            objItem = objARform.Items.Add("l_qcanum", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            objItem.Left = objARform.Items.Item("230").Left '+ objARform.Items.Item("230").Width + 10
            objItem.Width = 80
            objItem.Top = objARform.Items.Item("230").Top + objARform.Items.Item("230").Height + 2
            objItem.Height = 14 'objARform.Items.Item("2").Height
            objButton = objItem.Specific
            objButton.Caption = "QCA Num"

            Dim objedit As SAPbouiCOM.EditText
            objItem = objARform.Items.Add("t_qcanum", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            objItem.Left = objARform.Items.Item("222").Left '+ objARform.Items.Item("l_qcanum").Width + 5
            objItem.Width = 70
            objItem.Top = objARform.Items.Item("l_qcanum").Top
            objItem.Height = 14 'objARform.Items.Item("l_qcanum").Height
            objItem.LinkTo = "l_qcanum"
            objedit = objItem.Specific
            objedit.Item.Enabled = False
            objedit.DataBind.SetBound(True, "OINV", "U_QCANum")

            Dim objlink As SAPbouiCOM.LinkedButton
            objItem = objARform.Items.Add("lnk_qca", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            objItem.Left = objARform.Items.Item("l_qcanum").Left + objARform.Items.Item("l_qcanum").Width + 15
            objItem.Width = 12
            objItem.Top = objARform.Items.Item("l_qcanum").Top + 3
            objItem.Height = 10 'objARform.Items.Item("l_qcanum").Height
            'objItem.LinkTo = "t_qcanum"
            objlink = objItem.Specific
            objlink.LinkedObjectType = "-1"
            objlink.Item.LinkTo = "t_qcanum"

            'objAddOn.objApplication.SetStatusBarMessage("Button Created", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Catch ex As Exception
        End Try

    End Sub

End Class
