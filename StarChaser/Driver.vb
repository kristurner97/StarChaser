'tabs=4
' --------------------------------------------------------------------------------
'
' ASCOM Dome driver for StarChaser
'
' Description:	The Star Chaser dome control software was built by Kris Turner as
'               part of a project commissioned by Keele Observatory in order to
'               improve the functionality of their 24" Reflector telescope.
'
' Implements:	ASCOM Dome interface version: 1.0
' Author:		Kris Turner. For support contact <kristurner1997@gmail.com>
'
' Edit Log:
'
' Date			Who	Vers	Description
' -----------	---	-----	-------------------------------------------------------
' 07-05-2016	KST	1.0.0	Initial edit, from Dome template
' 07-05-2016	KST	1.0.0	Connected Property Programmed
' ---------------------------------------------------------------------------------
'

#Const Device = "Dome"

Imports ASCOM
Imports ASCOM.Astrometry
Imports ASCOM.Astrometry.AstroUtils
Imports ASCOM.DeviceInterface
Imports ASCOM.Utilities

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text

<Guid("e726dd49-0135-4a45-86be-e0aedb358ffd")>
<ClassInterface(ClassInterfaceType.None)>
Public Class Dome

    ' The Guid attribute sets the CLSID for ASCOM.StarChaser.Dome
    ' The ClassInterface/None addribute prevents an empty interface called
    ' _StarChaser from being created and used as the [default] interface

    ' TODO Replace the not implemented exceptions with code to implement the function or
    ' throw the appropriate ASCOM exception.
    '
    Implements IDomeV2

    '
    ' Driver ID and descriptive string that shows in the Chooser
    '
    Friend Shared driverID As String = "ASCOM.StarChaser.Dome"
    Private Shared driverDescription As String = "StarChaser Dome"

    Friend Shared comPortProfileName As String = "COM Port" 'Constants used for Profile persistence
    Friend Shared traceStateProfileName As String = "Trace Level"
    Friend Shared comPortDefault As String = "COM1"
    Friend Shared traceStateDefault As String = "False"

    Friend Shared comPort As String ' Variables to hold the currrent device configuration
    Friend Shared traceState As Boolean

    Private connectedState As Boolean ' Private variable to hold the connected state
    Private utilities As Util ' Private variable to hold an ASCOM Utilities object
    Private astroUtilities As AstroUtils ' Private variable to hold an AstroUtils object to provide the Range method
    Private TL As TraceLogger ' Private variable to hold the trace logger object (creates a diagnostic log file with information that you specify)

    Private DomeSerial As IO.Ports.SerialPort = Nothing

    '
    ' Constructor - Must be public for COM registration!
    '
    Public Sub New()

        ReadProfile() ' Read device configuration from the ASCOM Profile store
        TL = New TraceLogger("", "StarChaser")
        TL.Enabled = traceState
        TL.LogMessage("Dome", "Starting initialisation")

        connectedState = False ' Initialise connected to false
        utilities = New Util() ' Initialise util object
        astroUtilities = New AstroUtils 'Initialise new astro utiliites object

        'TODO: Implement your additional construction here

        TL.LogMessage("Dome", "Completed initialisation")
    End Sub

    '
    ' PUBLIC COM INTERFACE IDomeV2 IMPLEMENTATION
    '

#Region "Common properties and methods"
    ''' <summary>
    ''' Displays the Setup Dialog form.
    ''' If the user clicks the OK button to dismiss the form, then
    ''' the new settings are saved, otherwise the old values are reloaded.
    ''' THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
    ''' </summary>
    Public Sub SetupDialog() Implements IDomeV2.SetupDialog
        ' consider only showing the setup dialog if not connected
        ' or call a different dialog if connected
        If IsConnected Then
            System.Windows.Forms.MessageBox.Show("Already connected, just press OK")
        End If

        Using F As SetupDialogForm = New SetupDialogForm()
            Dim result As System.Windows.Forms.DialogResult = F.ShowDialog()
            If result = DialogResult.OK Then
                WriteProfile() ' Persist device configuration values to the ASCOM Profile store
            End If
        End Using
    End Sub

    Public ReadOnly Property SupportedActions() As ArrayList Implements IDomeV2.SupportedActions
        Get
            TL.LogMessage("SupportedActions Get", "Returning empty arraylist")
            Return New ArrayList()
        End Get
    End Property

    Public Function Action(ByVal ActionName As String, ByVal ActionParameters As String) As String Implements IDomeV2.Action
        Throw New ActionNotImplementedException("Action " & ActionName & " is not supported by this driver")
    End Function

    Public Sub CommandBlind(ByVal Command As String, Optional ByVal Raw As Boolean = False) Implements IDomeV2.CommandBlind
        CheckConnected("CommandBlind")
        ' Call CommandString and return as soon as it finishes
        Me.CommandString(Command, Raw)
        ' or
        Throw New MethodNotImplementedException("CommandBlind")
    End Sub

    Public Function CommandBool(ByVal Command As String, Optional ByVal Raw As Boolean = False) As Boolean _
        Implements IDomeV2.CommandBool
        CheckConnected("CommandBool")
        Dim ret As String = CommandString(Command, Raw)
        ' TODO decode the return string and return true or false
        ' or
        Throw New MethodNotImplementedException("CommandBool")
    End Function

    Public Function CommandString(ByVal Command As String, Optional ByVal Raw As Boolean = False) As String _
        Implements IDomeV2.CommandString
        CheckConnected("CommandString")
        ' it's a good idea to put all the low level communication with the device here,
        ' then all communication calls this function
        ' you need something to ensure that only one command is in progress at a time
        Throw New MethodNotImplementedException("CommandString")
    End Function

    Public Property Connected() As Boolean Implements IDomeV2.Connected
        Get
            TL.LogMessage("Connected Get", IsConnected.ToString())
            Return IsConnected
        End Get
        Set(value As Boolean)
            TL.LogMessage("Connected Set", value.ToString())
            If value = IsConnected Then
                Return
            End If

            If value Then
                Try
                    DomeSerial = My.Computer.Ports.OpenSerialPort(comPort)
                    DomeSerial.ReadTimeout = 10000
                    connectedState = True

                    DomeSerial.Write("ConfirmConnect#")

                    Dim StartTime As TimeSpan = Today.TimeOfDay

                    Do
                        Dim ArduinoMessage As String = DomeSerial.ReadTo("#")
                        If ArduinoMessage Is Nothing Then
                            If StartTime.TotalMilliseconds - Today.TimeOfDay.TotalMilliseconds > 5000 Then
                                Throw New TimeoutException
                            End If
                        ElseIf ArduinoMessage = "StarChaser Confirmed" Then
                            TL.LogMessage("Connected Set", "Connecting to port " + comPort)
                            connectedState = True
                        End If

                    Loop
                Catch ex As TimeoutException
                    TL.LogMessage("Connected Set", "Time Out Error. Check device is connected.")
                End Try
            Else

                Try
                    DomeSerial.Write("BeginDisconnect#")
                    Dim StartTime As TimeSpan = Today.TimeOfDay

                    Do
                        Dim ArduinoMessage As String = DomeSerial.ReadTo("#")

                        If ArduinoMessage Is Nothing Then
                            If StartTime.TotalMilliseconds - Today.TimeOfDay.TotalMilliseconds > 10000 Then
                                Throw New TimeoutException
                            End If
                        ElseIf ArduinoMessage = "DisconnectConfirmed" Then
                            DomeSerial.Close()
                            connectedState = False
                            TL.LogMessage("Connected Set", "Disconnecting from port " + comPort)
                        End If

                    Loop
                Catch ex As TimeoutException
                    TL.LogMessage("Connected Set", "Time Out Error on Disconnect.")
                End Try

            End If
        End Set
    End Property 'COMPLETED 07/05/2016 KST

    Public ReadOnly Property Description As String Implements IDomeV2.Description
        Get
            ' this pattern seems to be needed to allow a public property to return a private field
            Dim d As String = driverDescription
            TL.LogMessage("Description Get", d)
            Return d
        End Get
    End Property

    Public ReadOnly Property DriverInfo As String Implements IDomeV2.DriverInfo
        Get
            Dim m_version As Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            ' TODO customise this driver description
            Dim s_driverInfo As String = "Information about the driver itself. Version: " + m_version.Major.ToString() + "." + m_version.Minor.ToString()
            TL.LogMessage("DriverInfo Get", s_driverInfo)
            Return s_driverInfo
        End Get
    End Property

    Public ReadOnly Property DriverVersion() As String Implements IDomeV2.DriverVersion
        Get
            ' Get our own assembly and report its version number
            TL.LogMessage("DriverVersion Get", Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString(2))
            Return Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString(2)
        End Get
    End Property

    Public ReadOnly Property InterfaceVersion() As Short Implements IDomeV2.InterfaceVersion
        Get
            TL.LogMessage("InterfaceVersion Get", "2")
            Return 2
        End Get
    End Property

    Public ReadOnly Property Name As String Implements IDomeV2.Name
        Get
            Dim s_name As String = "Star Chaser"
            TL.LogMessage("Name Get", s_name)
            Return s_name
        End Get
    End Property

    Public Sub Dispose() Implements IDomeV2.Dispose
        ' Clean up the tracelogger and util objects
        TL.Enabled = False
        TL.Dispose()
        TL = Nothing
        utilities.Dispose()
        utilities = Nothing
        astroUtilities.Dispose()
        astroUtilities = Nothing
    End Sub

#End Region

#Region "IDome Implementation"

    Private domeShutterState As Boolean = False ' Variable to hold the open/closed status of the shutter, true = Open

    ''' <summary>
    ''' This method instructs the Arduino to abort any slewing commands. It does not receive a response from Arduino.
    ''' Modified 07/05/2016 KST
    ''' </summary>

    Public Sub AbortSlew() Implements IDomeV2.AbortSlew

        DomeSerial.Write("AbortSlew#")

        TL.LogMessage("AbortSlew", "Completed")
    End Sub

    ''' <summary>
    ''' ALTITUDE
    ''' At this time StarChaser does not have the capabilities of reporting shutter altitude. Returns zero.
    ''' Modified 07/05/2016 KST
    ''' </summary>

    Public ReadOnly Property Altitude() As Double Implements IDomeV2.Altitude
        Get
            TL.LogMessage("Altitude Get", "Not implemented")
            Return 0
        End Get
    End Property

    Public ReadOnly Property CanSetAltitude() As Boolean Implements IDomeV2.CanSetAltitude
        Get
            TL.LogMessage("CanSetAltitude Get", False.ToString())
            Return False
        End Get
    End Property

    ''' <summary>
    ''' At Home
    ''' StarChaser will poll the reed switch (which marks home position). If it is at home it will return "Home#" otherwise "NotHome#".
    ''' Modified 07/05/2016 KST
    ''' </summary

    Public ReadOnly Property AtHome() As Boolean Implements IDomeV2.AtHome
        Get

            DomeSerial.Write("AtHome#")

            Try

                Dim StartTime As TimeSpan = Today.TimeOfDay

                Do
                    Dim ArduinoMessage As String = DomeSerial.ReadTo("#")
                    If ArduinoMessage Is Nothing Then
                        If StartTime.TotalMilliseconds - Today.TimeOfDay.TotalMilliseconds > 1000 Then
                            Throw New TimeoutException
                            Exit Do
                        End If
                    ElseIf ArduinoMessage = "Home" Then
                        Return True
                        Exit Do
                    ElseIf ArduinoMessage = "NotHome" Then
                        Return False
                        Exit Do
                    End If
                Loop
            Catch ex As TimeoutException
                TL.LogMessage("Communication", "Time Out Error. Check device is connected.")
            End Try

        End Get
    End Property

    Public ReadOnly Property CanFindHome() As Boolean Implements IDomeV2.CanFindHome
        Get
            TL.LogMessage("CanFindHome Get", True.ToString())
            Return True
        End Get
    End Property

    ''' <summary>
    ''' At Park
    ''' At this time StarChaser does not have the capabilities of Parking dome rotation. Returns AtPark False.
    ''' Modified 07/05/2016 KST
    ''' </summary>

    Public ReadOnly Property AtPark() As Boolean Implements IDomeV2.AtPark
        Get
            TL.LogMessage("AtPark", "Not implemented")
            Return False
        End Get
    End Property

    Public ReadOnly Property CanPark() As Boolean Implements IDomeV2.CanPark
        Get
            TL.LogMessage("CanPark Get", False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSetPark() As Boolean Implements IDomeV2.CanSetPark
        Get
            TL.LogMessage("CanSetPark Get", False.ToString())
            Return False
        End Get
    End Property

    ''' <summary>
    ''' Azimuth
    ''' Ask the Arduino for it's current location.
    ''' Modified 07/05/2016 KST
    ''' </summary>

    Public ReadOnly Property Azimuth() As Double Implements IDomeV2.Azimuth
        Get

            DomeSerial.Write("AzGet#")

            Try

                Dim StartTime As TimeSpan = Today.TimeOfDay

                Do
                    Dim ArduinoMessage As String = DomeSerial.ReadTo("#")
                    If ArduinoMessage Is Nothing Then
                        If StartTime.TotalMilliseconds - Today.TimeOfDay.TotalMilliseconds > 2000 Then
                            Throw New TimeoutException
                            Exit Do
                        End If
                    ElseIf ArduinoMessage.Contains("Az[") Then
                        Dim Az As String = ArduinoMessage.Remove(0, 3)
                        Az = Az.Remove(Az.IndexOf("]"))
                        Return Double.Parse(Az)
                        Exit Do
                    End If
                Loop
            Catch ex As TimeoutException
                TL.LogMessage("Communication", "Time Out Error. Check device is connected.")
            End Try


        End Get
    End Property

    Public ReadOnly Property CanSetAzimuth() As Boolean Implements IDomeV2.CanSetAzimuth
        Get
            TL.LogMessage("CanSetAzimuth Get", False.ToString())
            Return False
        End Get
    End Property




    ''' <summary>
    ''' At this time StarChaser does not have the capabilities of automatic shutter control.
    ''' Modified 07/05/2016 KST
    ''' </summary>
    Public ReadOnly Property CanSetShutter() As Boolean Implements IDomeV2.CanSetShutter
        Get
            TL.LogMessage("CanSetShutter Get", False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSlave() As Boolean Implements IDomeV2.CanSlave
        Get
            TL.LogMessage("CanSlave Get", False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSyncAzimuth() As Boolean Implements IDomeV2.CanSyncAzimuth
        Get
            TL.LogMessage("CanSyncAzimuth Get", False.ToString())
            Return False
        End Get
    End Property

    Public Sub CloseShutter() Implements IDomeV2.CloseShutter
        TL.LogMessage("CloseShutter", "Shutter has been closed")
        domeShutterState = False
    End Sub

    Public Sub FindHome() Implements IDomeV2.FindHome
        TL.LogMessage("FindHome", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("FindHome")
    End Sub

    Public Sub OpenShutter() Implements IDomeV2.OpenShutter
        TL.LogMessage("OpenShutter", "Shutter has been opened")
        domeShutterState = True
    End Sub

    Public Sub Park() Implements IDomeV2.Park
        TL.LogMessage("Park", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("Park")
    End Sub

    Public Sub SetPark() Implements IDomeV2.SetPark
        TL.LogMessage("SetPark", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SetPark")
    End Sub

    Public ReadOnly Property ShutterStatus() As ShutterState Implements IDomeV2.ShutterStatus
        Get
            TL.LogMessage("CanSyncAzimuth Get", False.ToString())
            If (domeShutterState) Then
                TL.LogMessage("ShutterStatus", ShutterState.shutterOpen.ToString())
                Return ShutterState.shutterOpen
            Else
                TL.LogMessage("ShutterStatus", ShutterState.shutterClosed.ToString())
                Return ShutterState.shutterClosed
            End If
        End Get
    End Property

    Public Property Slaved() As Boolean Implements IDomeV2.Slaved
        Get
            TL.LogMessage("Slaved Get", False.ToString())
            Return False
        End Get
        Set(value As Boolean)
            TL.LogMessage("Slaved Set", "not implemented")
            Throw New ASCOM.PropertyNotImplementedException("Slaved", True)
        End Set
    End Property

    Public Sub SlewToAltitude(Altitude As Double) Implements IDomeV2.SlewToAltitude
        TL.LogMessage("SlewToAltitude", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SlewToAltitude")
    End Sub

    Public Sub SlewToAzimuth(Azimuth As Double) Implements IDomeV2.SlewToAzimuth
        TL.LogMessage("SlewToAzimuth", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SlewToAzimuth")
    End Sub

    Public ReadOnly Property Slewing() As Boolean Implements IDomeV2.Slewing
        Get
            TL.LogMessage("Slewing Get", False.ToString())
            Return False
        End Get
    End Property

    Public Sub SyncToAzimuth(Azimuth As Double) Implements IDomeV2.SyncToAzimuth
        TL.LogMessage("SyncToAzimuth", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SyncToAzimuth")
    End Sub

#End Region

#Region "Private properties and methods"
    ' here are some useful properties and methods that can be used as required
    ' to help with

#Region "ASCOM Registration"

    Private Shared Sub RegUnregASCOM(ByVal bRegister As Boolean)

        Using P As New Profile() With {.DeviceType = "Dome"}
            If bRegister Then
                P.Register(driverID, driverDescription)
            Else
                P.Unregister(driverID)
            End If
        End Using

    End Sub

    <ComRegisterFunction()>
    Public Shared Sub RegisterASCOM(ByVal T As Type)

        RegUnregASCOM(True)

    End Sub

    <ComUnregisterFunction()>
    Public Shared Sub UnregisterASCOM(ByVal T As Type)

        RegUnregASCOM(False)

    End Sub

#End Region

    ''' <summary>
    ''' Returns true if there is a valid connection to the driver hardware
    ''' </summary>
    ''' IsConnected Property completed 07/05/2016 KST
    Private ReadOnly Property IsConnected As Boolean
        Get
            Return connectedState
        End Get
    End Property

    ''' <summary>
    ''' Use this function to throw an exception if we aren't connected to the hardware
    ''' </summary>
    ''' <param name="message"></param>
    Private Sub CheckConnected(ByVal message As String)
        If Not IsConnected Then
            Throw New NotConnectedException(message)
        End If
    End Sub

    ''' <summary>
    ''' Read the device configuration from the ASCOM Profile store
    ''' </summary>
    Friend Sub ReadProfile()
        Using driverProfile As New Profile()
            driverProfile.DeviceType = "Dome"
            traceState = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, String.Empty, traceStateDefault))
            comPort = driverProfile.GetValue(driverID, comPortProfileName, String.Empty, comPortDefault)
        End Using
    End Sub

    ''' <summary>
    ''' Write the device configuration to the  ASCOM  Profile store
    ''' </summary>
    Friend Sub WriteProfile()
        Using driverProfile As New Profile()
            driverProfile.DeviceType = "Dome"
            driverProfile.WriteValue(driverID, traceStateProfileName, traceState.ToString())
            driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString())
        End Using

    End Sub

#End Region

End Class
