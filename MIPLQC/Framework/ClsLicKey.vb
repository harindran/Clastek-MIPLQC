Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Public Class ClsLicKey
    Private temp As String = ""

    Public Sub New()
    End Sub

    Public Sub SetNewDate()
        Dim newDate As DateTime = DateTime.Now.AddDays(31)
        temp = newDate.ToLongDateString()
        StoreDate(temp)
    End Sub
    Public Sub Expired()
        Dim d As String = ""

        Using key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Add-ons")
            d = CStr(key.GetValue("Date"))
        End Using

        Dim now As DateTime = DateTime.Parse(d)
        Dim day As Integer = (now.Subtract(DateTime.Now)).Days

        If day > 30 Then
        ElseIf 0 < day AndAlso day <= 30 Then
            Dim daysLeft As String = String.Format("{0} days more to expire", now.Subtract(DateTime.Now).Days)
            MessageBox.Show(daysLeft)
        ElseIf day <= 0 Then
        End If
    End Sub

    Private Sub StoreDate(ByVal value As String)
        Try

            Using key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey("SOFTWARE\Add-ons")
                key.SetValue("Date", value, Microsoft.Win32.RegistryValueKind.String)
            End Using
            Expired()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
