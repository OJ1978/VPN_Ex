Imports System.Net.NetworkInformation

Public Class frmVPN

    Private Sub CheckConnection()

        Dim niVPN As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces

        Dim blnExist As Boolean = niVPN.AsEnumerable().Any(Function(x) x.Name = "VPN Name")

        If blnExist Then

            MessageBox.Show("VPN Exists")

        Else

            MessageBox.Show("VPN Does Not Exist")

        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        CheckConnection()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim vp As New clsVPN

        vp.Test()

        If vp.Connected Then

            MessageBox.Show("Connected")

        Else

            MessageBox.Show("Not Connected")

        End If

    End Sub

End Class
