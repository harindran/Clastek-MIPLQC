
Namespace Mukesh.SBOLib
    Public Class clsRecordSetPool
        Private Const POOL_SIZE As Integer = 1
        Private Shared intNextRecordSet
        Private Shared objRSPool(POOL_SIZE) As enuObjStatus
        Private Shared objRS(POOL_SIZE) As SAPbobsCOM.Recordset

        Sub New(ByVal Company As SAPbobsCOM.Company)
            Dim intLoop As Integer
            For intLoop = 1 To POOL_SIZE
                objRS(intLoop) = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRSPool(intLoop) = enuObjStatus.Free
            Next
        End Sub

        Public Shared Function GetRecordSet() As SAPbobsCOM.Recordset
            Dim intLoop As Integer

            For intLoop = 1 To POOL_SIZE
                If objRSPool(intLoop) = enuObjStatus.Free Then
                    objRSPool(intLoop) = enuObjStatus.Occupaied
                    Return objRS(intLoop)
                End If
            Next
            Throw (New Exception("Recordset Pool Is Full. Increase Pool Size"))
            Return Nothing
        End Function

        Public Shared Sub ReleaseRecordSet(ByVal RS As SAPbobsCOM.Recordset)
            Dim intLoop As Integer

            For intLoop = 1 To POOL_SIZE
                If RS Is objRS(intLoop) Then
                    objRSPool(intLoop) = enuObjStatus.Free
                End If
            Next
        End Sub

        Private Enum enuObjStatus
            Occupaied
            Free
        End Enum
    End Class
End Namespace