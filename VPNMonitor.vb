Imports System.Net.NetworkInformation
Imports System.Configuration
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Menu

Public Class VPNMonitor
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ctxNotify As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuOn As System.Windows.Forms.MenuItem
    Friend WithEvents TmrCheckVPN As Timer
    Friend WithEvents mnuOff As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSep As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ctxNotify = New System.Windows.Forms.ContextMenu()
        Me.mnuOn = New System.Windows.Forms.MenuItem()
        Me.mnuOff = New System.Windows.Forms.MenuItem()
        Me.mnuExit = New System.Windows.Forms.MenuItem()
        Me.mnuSep = New MenuItem("-")
        Me.TmrCheckVPN = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ctxNotify
        Me.NotifyIcon1.Text = "VPN Monitor"
        Me.NotifyIcon1.Visible = True
        '
        'ctxNotify
        '
        Me.ctxNotify.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuOn, Me.mnuOff, Me.mnuSep, Me.mnuExit})
        '
        'mnuOn
        '
        Me.mnuOn.Index = 0
        Me.mnuOn.Text = "O&n"
        '
        'mnuOff
        '
        Me.mnuOff.Index = 1
        Me.mnuOff.Text = "O&ff"
        '
        'mnuSep
        '
        Me.mnuSep.Index = 2
        '
        'mnuExit
        '
        Me.mnuExit.Index = 3
        Me.mnuExit.Text = "&Exit"
        '
        'TmrCheckVPN
        '
        Me.TmrCheckVPN.Enabled = True
        Me.TmrCheckVPN.Interval = 5000
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 77)
        Me.Name = "VPN Monitor"
        Me.ShowInTaskbar = False
        Me.FormBorderStyle = FormBorderStyle.SizableToolWindow
        Me.Text = "VPN Monitor"
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private m_icoOn As Icon
    Private m_icoOff As Icon
    Private connectionName As String

    Private VPN_Name As String = ConfigurationManager.AppSettings("VPN_Name")

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        m_icoOn = My.Resources.green
        m_icoOff = My.Resources.red

        Check_VPN()
    End Sub

    Private Sub mnuOn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOn.Click
        Me.NotifyIcon1.Icon = m_icoOn
        Me.Icon = m_icoOn
        Turn_VPN("On")
    End Sub

    Private Sub mnuOff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOff.Click
        Me.NotifyIcon1.Icon = m_icoOff
        Me.Icon = m_icoOff
        Turn_VPN("Off")
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Application.Exit()
    End Sub

    Private Function Check_VPN() As Boolean

        Dim nics As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces
        If nics.Length < 0 Or nics Is Nothing Then
            MsgBox("NO NETWORK CONNECTIONS")
            Application.Exit()
        Else
            For Each netadapter As NetworkInterface In nics
                If netadapter.Name = VPN_Name Then
                    connectionName = netadapter.Name
                    Me.mnuOff.Enabled = True
                    Me.mnuOn.Enabled = False
                    Me.NotifyIcon1.Icon = m_icoOn
                    Me.Icon = m_icoOn
                    Return True
                End If
            Next
        End If

        Me.mnuOff.Enabled = False
        Me.mnuOn.Enabled = True
        Me.NotifyIcon1.Icon = m_icoOff
        Me.Icon = m_icoOff
        Return False
    End Function

    Private Sub Turn_VPN(ByVal OnOff As String)

        TmrCheckVPN.Stop()
        TmrCheckVPN.Start()

        Dim argVPN As String
        argVPN = Chr(34) + connectionName + Chr(34)

        If OnOff = "Off" Then
            argVPN = argVPN + " /disconnect"
        End If

        Dim ProcessProperties As New ProcessStartInfo
        ProcessProperties.FileName = "c:\windows\system32\rasdial.exe"
        ProcessProperties.Arguments = argVPN
        ProcessProperties.WindowStyle = ProcessWindowStyle.Hidden
        Dim myProcess As Process = Process.Start(ProcessProperties)

    End Sub

    Private Sub TmrCheckVPN_Tick(sender As Object, e As EventArgs) Handles TmrCheckVPN.Tick
        Check_VPN()
    End Sub

End Class
