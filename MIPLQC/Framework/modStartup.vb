Imports SAPbouiCOM.Framework
Module modStartup
    Public objAddOn As clsAddOn

    'Public Sub Main()
    '    Try
    '        objAddOn = New clsAddOn
    '        objAddOn.Intialize()
    '        System.Windows.Forms.Application.Run()
    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '    End Try
    'End Sub
    Sub Main(ByVal args() As String)
        Try
            'Application & Company Connection                
            objAddOn = New clsAddOn
            objAddOn.Intialize(args)

        Catch ex As Exception
            MsgBox("Error in Module : " & ex.Message.ToString)
        End Try
    End Sub
End Module
