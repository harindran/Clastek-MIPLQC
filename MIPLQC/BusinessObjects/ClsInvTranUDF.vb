Public Class ClsInvTranUDF
    Dim InvUDFForm As SAPbouiCOM.Form
    Public Const Formtype = "-940"
    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        InvUDFForm = objAddOn.objApplication.Forms.Item(FormUID)
        If pVal.BeforeAction Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "1" And InvUDFForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Try
                        Catch ex As Exception
                        End Try
                    End If
            End Select
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    FieldEnabled(FormUID)
            End Select
        End If
    End Sub
    
    Public Sub TranTrackInUDF(ByVal FormUID As String)
        Try
            InvUDFForm = objAddOn.objApplication.Forms.Item(FormUID)
            Dim ObjQCForm As SAPbouiCOM.Form
            ObjQCForm = objAddOn.objApplication.Forms.GetForm("MIPLQC", 0)
            Dim QCHeader As SAPbouiCOM.DBDataSource = ObjQCForm.DataSources.DBDataSources.Item("@MIPLQC")
            UDFEnable(FormUID, "U_QCEntry")
            UDFEnable(FormUID, "U_QCNum")
            InvUDFForm.Items.Item("U_QCNum").Specific.String = QCHeader.GetValue("DocNum", 0)
            InvUDFForm.Items.Item("U_QCEntry").Specific.String = QCHeader.GetValue("DocEntry", 0)
            InvUDFForm.Items.Item("U_QCNum").Enabled = False
            'InvUDFForm.Items.Item("U_QCEntry").Enabled = False
            If QCHeader.GetValue("U_Type", 0) = "G" Then
                UDFEnable(FormUID, "U_GRPONum")
                UDFEnable(FormUID, "U_GRNEntry")
                InvUDFForm.Items.Item("U_GRPONum").Specific.String = ObjQCForm.Items.Item("13B").Specific.string
                InvUDFForm.Items.Item("U_GRNEntry").Specific.String = ObjQCForm.Items.Item("13").Specific.string
                InvUDFForm.Items.Item("U_GRPONum").Enabled = False
                'InvUDFForm.Items.Item("U_GRNEntry").Enabled = False
            ElseIf QCHeader.GetValue("U_Type", 0) = "P" Then
                UDFEnable(FormUID, "U_PORNum")
                UDFEnable(FormUID, "U_ProdEntry")
                UDFEnable(FormUID, "U_REntry")
                UDFEnable(FormUID, "U_GREntry")
                InvUDFForm.Items.Item("U_PORNum").Specific.String = ObjQCForm.Items.Item("15B").Specific.string
                InvUDFForm.Items.Item("U_ProdEntry").Specific.String = ObjQCForm.Items.Item("15").Specific.string
                InvUDFForm.Items.Item("U_REntry").Specific.String = ObjQCForm.Items.Item("51B").Specific.string
                InvUDFForm.Items.Item("U_GREntry").Specific.String = ObjQCForm.Items.Item("51").Specific.string
                InvUDFForm.Items.Item("U_PORNum").Enabled = False
                'InvUDFForm.Items.Item("U_ProdEntry").Enabled = False
                InvUDFForm.Items.Item("U_REntry").Enabled = False
                'InvUDFForm.Items.Item("U_GREntry").Enabled = False
            ElseIf QCHeader.GetValue("U_Type", 0) = "T" Then
                UDFEnable(FormUID, "U_StkEntry")
                UDFEnable(FormUID, "U_StkNum")
                InvUDFForm.Items.Item("U_StkEntry").Specific.String = ObjQCForm.Items.Item("23B").Specific.string
                InvUDFForm.Items.Item("U_StkNum").Specific.String = ObjQCForm.Items.Item("23").Specific.string
                InvUDFForm.Items.Item("U_StkEntry").Enabled = False
                'InvUDFForm.Items.Item("U_StkNum").Enabled = False
            ElseIf QCHeader.GetValue("U_Type", 0) = "R" Then
                UDFEnable(FormUID, "U_REntry")
                UDFEnable(FormUID, "U_GREntry")
                InvUDFForm.Items.Item("U_REntry").Specific.String = ObjQCForm.Items.Item("51B").Specific.string
                InvUDFForm.Items.Item("U_GREntry").Specific.String = ObjQCForm.Items.Item("51").Specific.string
                InvUDFForm.Items.Item("U_REntry").Enabled = False
                'InvUDFForm.Items.Item("U_GREntry").Enabled = False
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub FieldEnabled(ByVal FormUID As String)
        Try
            InvUDFForm = objAddOn.objApplication.Forms.Item(FormUID)
            InvUDFForm.Items.Item("U_GRPONum").Enabled = False
            InvUDFForm.Items.Item("U_GRNEntry").Enabled = False
            InvUDFForm.Items.Item("U_PORNum").Enabled = False
            InvUDFForm.Items.Item("U_ProdEntry").Enabled = False
            InvUDFForm.Items.Item("U_REntry").Enabled = False
            InvUDFForm.Items.Item("U_GREntry").Enabled = False
            InvUDFForm.Items.Item("U_StkEntry").Enabled = False
            InvUDFForm.Items.Item("U_StkNum").Enabled = False
            InvUDFForm.Items.Item("U_REntry").Enabled = False
            InvUDFForm.Items.Item("U_GREntry").Enabled = False
            InvUDFForm.Items.Item("U_QCEntry").Enabled = False
            InvUDFForm.Items.Item("U_QCNum").Enabled = False
        Catch ex As Exception

        End Try
    End Sub

    Private Sub UDFEnable(ByVal FormUID As String, ByVal ItemUID As String)
        Try
            InvUDFForm = objAddOn.objApplication.Forms.Item(FormUID)
            'InvUDFForm.ActiveItem = "U_Dummy"
            If InvUDFForm.Items.Item(ItemUID).Enabled = False Then
                InvUDFForm.Items.Item(ItemUID).Enabled = True
            End If
        Catch ex As Exception
            'objAddOn.objApplication.SetStatusBarMessage(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try
    End Sub

End Class
