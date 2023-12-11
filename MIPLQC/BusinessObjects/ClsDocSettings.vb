Public Class ClsDocSettings
    Public Const Formtype = "228"
    Dim objForm As SAPbouiCOM.Form
    Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        If pVal.BeforeAction Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    FieldCreationUI(FormUID)
            End Select
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True Then

                    End If
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True Then

                    End If

            End Select
        End If
    End Sub

    Private Sub FieldCreationUI(FormUID As String)
        Dim objLabel As SAPbouiCOM.StaticText
        Dim objedit As SAPbouiCOM.EditText
        Dim oItem As SAPbouiCOM.Item
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            oItem = objForm.Items.Add("lblInv", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = objForm.Items.Item("20").Left + objForm.Items.Item("20").Width + 5
            oItem.Width = 200
            oItem.Height = objForm.Items.Item("20").Height
            oItem.Top = objForm.Items.Item("20").Top
            objLabel = oItem.Specific
            objLabel.Caption = "InvTran Exception Whse"
            objLabel.Item.FromPane = 1
            objLabel.Item.ToPane = 1

            oItem = objForm.Items.Add("txtInv", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = objForm.Items.Item("lblInv").Left + objForm.Items.Item("lblInv").Width + 5
            oItem.Width = 50
            oItem.Height = objForm.Items.Item("lblInv").Height
            oItem.Top = objForm.Items.Item("lblInv").Top
            objedit = oItem.Specific
            objedit.DataBind.SetBound(True, "OADM", "U_Whse")
            objedit.Item.FromPane = 1
            objedit.Item.ToPane = 1
        Catch ex As Exception

        End Try
    End Sub
End Class
