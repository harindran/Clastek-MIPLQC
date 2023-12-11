Public Class clsInventoryTransfer
    Dim InventoryForm As SAPbouiCOM.Form
    Dim ObjQCForm As SAPbouiCOM.Form
    Public Const Formtype = "940"
    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        InventoryForm = objAddOn.objApplication.Forms.Item(FormUID)
        If pVal.BeforeAction Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "1" And InventoryForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Try
                            'ObjQCForm.Items.Item("32").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            'ObjQCForm.Items.Item("34").Specific.value = "0"

                            'objAddOn.objCompany.GetNewObjectKey() 'InventoryForm.DataSources.DBDataSources.Item("OWTR").GetValue("DocEntry", 0)
                        Catch ex As Exception

                        End Try

                    End If
            End Select
        Else
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    If pVal.ActionSuccess Then
                        loadScreen(FormUID)
                    End If
            End Select
        End If
    End Sub

    Public Sub loadScreen(ByVal FormUID As String)
        Try
            Dim StkTrnfMatrix As SAPbouiCOM.Matrix
            'Dim objformUDF As SAPbouiCOM.Form
            Dim objQCMatrix As SAPbouiCOM.Matrix
            InventoryForm = objAddOn.objApplication.Forms.Item(FormUID)
            InventoryForm.Visible = True
            'InventoryForm.UDFFormUID
            ObjQCForm = objAddOn.objApplication.Forms.GetForm("MIPLQC", 0)
            Dim QCHeader As SAPbouiCOM.DBDataSource = ObjQCForm.DataSources.DBDataSources.Item("@MIPLQC")
            InventoryForm.Items.Item("14").Specific.String = ObjQCForm.Items.Item("6").Specific.String
            ' InventoryForm.Items.Item("1320000098").Specific.String = ObjQCForm.Items.Item("50").Specific.String
            StkTrnfMatrix = InventoryForm.Items.Item("23").Specific '18 - from warehouse; 1470000101 - to warehouse
            'InventoryForm.Items.Item("18").Specific.String = CStr(QCHeader.GetValue("U_InWhse", 0))
            Try
                If Not objAddOn.objApplication.Menus.Item("6913").Checked = True Then
                    objAddOn.objApplication.Menus.Item("6913").Activate()
                    'objformUDF = objAddOn.objApplication.Forms.GetForm("-940", 0)
                End If
            Catch ex As Exception

            End Try
            '-------------------------------------------------------------------------------
            Try

                'InventoryForm.DataSources.DBDataSources.Item("OWTR").SetValue("U_GRNEntry", 0, QCHeader.GetValue("U_GRNEntry", 0))
                'InventoryForm.DataSources.DBDataSources.Item("OWTR").SetValue("U_ProdEntry", 0, QCHeader.GetValue("U_ProdEntry", 0))
                'InventoryForm.DataSources.DBDataSources.Item("OWTR").SetValue("U_GREntry", 0, QCHeader.GetValue("U_GREntry", 0))

                'InventoryUDFForm.Items.Item("U_GRNEntry").Specific.String = QCHeader.GetValue("U_GRNEntry", 0)
                'InventoryUDFForm.Items.Item("U_ProdEntry").Specific.String = QCHeader.GetValue("U_ProdEntry", 0)
                'InventoryUDFForm.Items.Item("U_GREntry").Specific.String = QCHeader.GetValue("U_GREntry", 0)
                InventoryForm.Items.Item("U_QCEntry").Specific.String = QCHeader.GetValue("DocEntry", 0)
            Catch ex As Exception

                Dim objText As SAPbouiCOM.EditText
                Dim objItem As SAPbouiCOM.Item
                Dim objLabel As SAPbouiCOM.StaticText
                objItem = InventoryForm.Items.Add("U_QCEntry", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                objItem.Left = InventoryForm.Items.Item("36").Left
                objItem.Top = InventoryForm.Items.Item("36").Top + InventoryForm.Items.Item("36").Height + 5
                objItem.Width = InventoryForm.Items.Item("36").Width
                objItem.Height = InventoryForm.Items.Item("36").Height
                objText = objItem.Specific
                objText.Item.Enabled = False
                objText.DataBind.SetBound(True, "OWTR", "U_QCNum")

                objItem = InventoryForm.Items.Add("QCLabel", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                objItem.Left = InventoryForm.Items.Item("37").Left
                objItem.Top = InventoryForm.Items.Item("37").Top + InventoryForm.Items.Item("37").Height + 5
                objItem.Width = InventoryForm.Items.Item("37").Width
                objItem.Height = InventoryForm.Items.Item("37").Height
                objLabel = objItem.Specific
                objLabel.Caption = "QC Num"
                objItem.LinkTo = "U_QCEntry"

                InventoryForm.Items.Item("U_QCEntry").Specific.String = QCHeader.GetValue("DocNum", 0) ' objForm.Items.Item("6E").Specific.string '
                'InventoryForm.Items.Item("U_QCNum").Specific.String = QCHeader.GetValue("DocNum", 0)

                'InventoryForm.Items.Item("U_QCEntry").Specific.String = QCHeader.GetValue("DocEntry", 0)
            'InventoryUDFForm.Items.Item("U_GRNEntry").Specific.String = QCHeader.GetValue("U_GRNEntry", 0)
            'InventoryUDFForm.Items.Item("U_ProdEntry").Specific.String = QCHeader.GetValue("U_ProdEntry", 0)
            'InventoryUDFForm.Items.Item("U_GREntry").Specific.String = QCHeader.GetValue("U_GREntry", 0)
            ''-----------------------------------------------------------------------------------


            'objItem = InventoryForm.Items.Add("U_GREntry", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'objItem.Left = InventoryForm.Items.Item("36").Left
            'objItem.Top = InventoryForm.Items.Item("36").Top + InventoryForm.Items.Item("36").Height + 5
            'objItem.Width = InventoryForm.Items.Item("36").Width
            'objItem.Height = InventoryForm.Items.Item("36").Height
            'objText = objItem.Specific
            'objText.DataBind.SetBound(True, "OWTR", "U_GRPOEntry")

            'objItem = InventoryForm.Items.Add("GRPOLabel", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'objItem.Left = InventoryForm.Items.Item("37").Left
            'objItem.Top = InventoryForm.Items.Item("37").Top + InventoryForm.Items.Item("37").Height + 5
            'objItem.Width = InventoryForm.Items.Item("37").Width
            'objItem.Height = InventoryForm.Items.Item("37").Height
            'objLabel = objItem.Specific
            'objLabel.Caption = "GRPO Entry"
            'objItem.LinkTo = "U_GREntry"

            'InventoryForm.Items.Item("U_GREntry").Specific.String = QCHeader.GetValue("DocEntry", 0)

            'objItem = InventoryForm.Items.Add("U_PREntry", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            'objItem.Left = InventoryForm.Items.Item("36").Left
            'objItem.Top = InventoryForm.Items.Item("36").Top + InventoryForm.Items.Item("36").Height + 5
            'objItem.Width = InventoryForm.Items.Item("36").Width
            'objItem.Height = InventoryForm.Items.Item("36").Height
            'objText = objItem.Specific
            'objText.DataBind.SetBound(True, "OWTR", "U_ProdEntry")

            'objItem = InventoryForm.Items.Add("ProdLabel", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            'objItem.Left = InventoryForm.Items.Item("37").Left
            'objItem.Top = InventoryForm.Items.Item("37").Top + InventoryForm.Items.Item("37").Height + 5
            'objItem.Width = InventoryForm.Items.Item("37").Width
            'objItem.Height = InventoryForm.Items.Item("37").Height
            'objLabel = objItem.Specific
            'objLabel.Caption = "Prod Entry"
            'objItem.LinkTo = "U_PREntry"

            'InventoryForm.Items.Item("U_PREntry").Specific.String = QCHeader.GetValue("DocEntry", 0)

            '-----------------------------------------------------------------------------------
        End Try
        '-------------------------------------------------------------------------------
        objQCMatrix = ObjQCForm.Items.Item("20").Specific
        Dim ToWhseSet As Boolean = False
        Dim QCLine As SAPbouiCOM.DBDataSource = ObjQCForm.DataSources.DBDataSources.Item("@MIPLQC1")

        For j As Integer = 1 To objQCMatrix.RowCount
            objQCMatrix.GetLineData(j)
            If CInt(QCLine.GetValue("U_AccQty", j - 1)) > 0 Then

                'If ToWhseSet = False Then
                '    InventoryForm.Items.Item("1470000101").Specific.String = CStr(QCLine.GetValue("U_AccWhse", j - 1))
                '    ToWhseSet = True
                'End If

                StkTrnfMatrix.Columns.Item("1").Cells.Item(StkTrnfMatrix.RowCount).Specific.String = Trim(CStr(QCLine.GetValue("U_ItemCode", j - 1)))
                StkTrnfMatrix.Columns.Item("1470001039").Cells.Item(StkTrnfMatrix.RowCount - 1).Specific.String = Trim(CStr(QCLine.GetValue("U_FrmWhse", j - 1)))
                StkTrnfMatrix.Columns.Item("5").Cells.Item(StkTrnfMatrix.RowCount - 1).Specific.String = Trim(CStr(QCLine.GetValue("U_AccWhse", j - 1)))
                StkTrnfMatrix.Columns.Item("10").Cells.Item(StkTrnfMatrix.RowCount - 1).Specific.String = CStr(QCLine.GetValue("U_AccQty", j - 1))
                ' StkTrnfMatrix.Columns.Item("10").Cells.Item(StkTrnfMatrix.RowCount - 1).Specific.String = GetConvertedQty(FormUID, Trim(CStr(QCHeader.GetValue("U_GRNEntry", 0))), Trim(CStr(QCHeader.GetValue("U_Type", 0))), Trim(CStr(QCLine.GetValue("U_ItemCode", j - 1))), QCLine.GetValue("U_AccQty", j - 1))
            End If
            If CInt(QCLine.GetValue("U_RejQty", j - 1)) > 0 Then

                'If ToWhseSet = False Then
                '    InventoryForm.Items.Item("1470000101").Specific.String = CStr(QCLine.GetValue("U_RejWhse", j - 1))
                '    ToWhseSet = True
                'End If
                StkTrnfMatrix.Columns.Item("1").Cells.Item(StkTrnfMatrix.RowCount).Specific.String = Trim(CStr(QCLine.GetValue("U_ItemCode", j - 1)))
                StkTrnfMatrix.Columns.Item("1470001039").Cells.Item(StkTrnfMatrix.RowCount - 1).Specific.String = Trim(CStr(QCLine.GetValue("U_FrmWhse", j - 1)))
                StkTrnfMatrix.Columns.Item("5").Cells.Item(StkTrnfMatrix.RowCount - 1).Specific.String = Trim(CStr(QCLine.GetValue("U_RejWhse", j - 1)))
                StkTrnfMatrix.Columns.Item("10").Cells.Item(StkTrnfMatrix.RowCount - 1).Specific.String = CStr(QCLine.GetValue("U_RejQty", j - 1))
                'StkTrnfMatrix.Columns.Item("10").Cells.Item(StkTrnfMatrix.RowCount - 1).Specific.String = GetConvertedQty(FormUID, Trim(CStr(QCHeader.GetValue("U_GRNEntry", 0))), Trim(CStr(QCHeader.GetValue("U_Type", 0))), Trim(CStr(QCLine.GetValue("U_ItemCode", j - 1))), QCLine.GetValue("U_RejQty", j - 1))
            End If
            If CInt(QCLine.GetValue("U_RewQty", j - 1)) > 0 Then

                'If ToWhseSet = False Then
                '    InventoryForm.Items.Item("1470000101").Specific.String = CStr(QCLine.GetValue("U_RewWhse", j - 1))
                '    ToWhseSet = True
                'End If
                StkTrnfMatrix.Columns.Item("1").Cells.Item(StkTrnfMatrix.RowCount).Specific.String = Trim(CStr(QCLine.GetValue("U_ItemCode", j - 1)))
                StkTrnfMatrix.Columns.Item("1470001039").Cells.Item(StkTrnfMatrix.RowCount - 1).Specific.String = Trim(CStr(QCLine.GetValue("U_FrmWhse", j - 1)))
                StkTrnfMatrix.Columns.Item("5").Cells.Item(StkTrnfMatrix.RowCount - 1).Specific.String = Trim(CStr(QCLine.GetValue("U_RewWhse", j - 1)))
                StkTrnfMatrix.Columns.Item("10").Cells.Item(StkTrnfMatrix.RowCount - 1).Specific.String = CStr(QCLine.GetValue("U_RewQty", j - 1))
                ' StkTrnfMatrix.Columns.Item("10").Cells.Item(StkTrnfMatrix.RowCount - 1).Specific.String = GetConvertedQty(FormUID, Trim(CStr(QCHeader.GetValue("U_GRNEntry", 0))), Trim(CStr(QCHeader.GetValue("U_Type", 0))), Trim(CStr(QCLine.GetValue("U_ItemCode", j - 1))), QCLine.GetValue("U_RewQty", j - 1))
            End If
        Next
        '  Dim InventoryUDFForm As SAPbouiCOM.Form
        ' InventoryUDFForm = objAddOn.objApplication.Forms.GetFormByTypeAndCount("-940", 0)

        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            InventoryForm = objAddOn.objApplication.Forms.Item(BusinessObjectInfo.FormUID)
            ObjQCForm = objAddOn.objApplication.Forms.GetForm("MIPLQC", 0)
            If BusinessObjectInfo.BeforeAction = True Then
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                End Select
            Else
                Select Case BusinessObjectInfo.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                        If objAddOn.HANA Then
                            ObjQCForm.Items.Item("34").Specific.string = InventoryForm.DataSources.DBDataSources.Item("OWTR").GetValue("DocEntry", 0)
                        Else
                            ObjQCForm.Items.Item("34").Specific.string = InventoryForm.DataSources.DBDataSources.Item("OWTR").GetValue("DocNum", 0)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                End Select
            End If

        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
End Class
