Imports System.Net.NetworkInformation

Public Class clsVPN

    Public Delegate Sub delPing()
    Public Delegate Sub delConnect()
    Public Delegate Sub delIdle()
    Public Delegate Sub delDisconnect()
    Public Delegate Sub delStatus(blnConnected As Boolean)

    Public Event Ping As delPing
    Public Event Con As delConnect
    Public Event Discon As delDisconnect
    Public Event Idle As delIdle
    Public Event StatusChanged As delStatus

    Private strRASPhone As String = "C:\WINDOWS\system32\rasphone.exe"
    Private strVPNCon As String = ""
    Private strIPAddress As String = ""

    Private blnConnected As Boolean = False

    Protected Sub OnStatusChanged(blnConnected As Boolean)

        RaiseEvent StatusChanged(blnConnected)

    End Sub

    Protected Sub OnDisconnect()

        RaiseEvent Discon()

    End Sub

    Protected Sub OnPing()

        RaiseEvent Ping()

    End Sub

    Protected Sub OnIdle()

        RaiseEvent Idle()

    End Sub

    Protected Sub OnConnect()

        RaiseEvent Con()

    End Sub

    Public ReadOnly Property Connected() As Boolean

        Get

            Return blnConnected

        End Get

    End Property


    Public Property ConName() As String

        Get

            Return strVPNCon

        End Get

        Set(strValue As String)

            strVPNCon = strValue

        End Set

    End Property

    Public Function Test() As Boolean

        Dim blnSucceed As Boolean = False

        OnPing()

        Dim p As New Ping()

        If p.Send(strIPAddress).Status = IPStatus.Success Then

            blnSucceed = True

        Else

            blnSucceed = False

        End If

        p = Nothing

        If blnSucceed <> blnConnected Then

            blnConnected = blnSucceed

            OnStatusChanged(blnConnected)

        End If

        OnIdle()

        Return blnSucceed

    End Function

    Private Function Connect() As Boolean

        Dim blnSucceed As Boolean = False

        OnConnect()

        Process.Start(strRASPhone, Convert.ToString(" -d ") & strVPNCon)

        Application.DoEvents()

        System.Threading.Thread.Sleep(5000)

        Application.DoEvents()

        blnSucceed = True

        OnIdle()

        Return blnSucceed

    End Function

    Private Function Disconnect() As Boolean

        Dim blnSucceed As Boolean = False

        OnDisconnect()

        Process.Start(strRASPhone, Convert.ToString(" -h ") & strVPNCon)

        Application.DoEvents()

        System.Threading.Thread.Sleep(8000)

        Application.DoEvents()

        blnSucceed = True

        OnIdle()

        Return blnSucceed

    End Function


    Public Sub New()

    End Sub

End Class
