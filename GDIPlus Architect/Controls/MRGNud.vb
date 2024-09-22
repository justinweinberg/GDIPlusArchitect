''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : MRGNud
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A custom numeric updown control.  Created to supplement behavior of the original 
''' up down control.
''' <remarks>There were a couple things 
''' about the default NUD that created issues that needed to be addressed in order 
''' to get the functionality I wanted.  This control addresses those issues.
''' </remarks>
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class MRGNud
    Inherits System.Windows.Forms.UserControl


#Region "Local Fields"



    '''<summary>Last good value.  When the user looses focus with an invalid value, 
    ''' the nud resets to this value</summary>
    Private _LastKnownValue As Decimal = 0

    '''<summary>The maximum allowed value.</summary>
    Private _MaxValue As Decimal = 1000

    '''<summary>The minimum allowed value.</summary>
    Private _MinValue As Decimal = 0

    '''<summary>The amount to increment on an up or down operation</summary>
    Private _Increment As Decimal = 1


    '''<summary>The number of allowed decimal places</summary>
    Private _DecimalPlaces As Int32 = 0

    '''<summary>Whether to allow blank entries (true) or not (false)</summary>
    Private _AllowBlank As Boolean = False

#End Region


#Region "Constructors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of an MRGNud control
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
#End Region


#Region " Windows Form Designer generated code "


    'UserControl overrides dispose to clean up the component list.
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
    Private WithEvents txtNud As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        txtNud = New TextBox
        Me.SuspendLayout()
        txtNud.Parent = Me
        '
        'txtNud
        '
        txtNud.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        txtNud.Location = New Point(0, 0)
        txtNud.TabIndex = 0
        txtNud.Text = ""
        '
        'MRGNud
        '
        Me.Size = New Size(176, 32)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the text inside of the NUD
    ''' </summary>
    ''' <value>A string containing the value of the NUD</value>
    ''' -----------------------------------------------------------------------------
    Public Overrides Property Text() As String
        Get
            Return txtNud.Text
        End Get
        Set(ByVal Value As String)
            If ProposedValueValid(Value) Then
                assignText(Value)
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the minimum valid value of the NUD.
    ''' </summary>
    ''' <value>A decimal that is the lower bound on values in the NUD</value>
    ''' -----------------------------------------------------------------------------
    Public Property MinValue() As Decimal
        Get
            Return _MinValue
        End Get
        Set(ByVal Value As Decimal)
            If _MinValue < MaxValue Then
                _MinValue = Value
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the number of acceptable decimal places
    ''' </summary>
    ''' <value>An int32 holding the maximum decimals allowed in the NUD</value>
    ''' -----------------------------------------------------------------------------
    Public Property DecimalPlaces() As Int32
        Get
            Return _DecimalPlaces
        End Get
        Set(ByVal Value As Int32)
            If Value > 0 AndAlso Value < 10 Then
                _DecimalPlaces = Value
            End If
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets a value indicating whether a blank value is allowed in the NUD
    ''' </summary>
    ''' <value>Boolean indicating whether the NUD should allow blanks or not.</value>
    ''' -----------------------------------------------------------------------------
    Public Property AllowBlank() As Boolean
        Get
            Return _AllowBlank
        End Get
        Set(ByVal Value As Boolean)
            _AllowBlank = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the maximum acceptable value in the NUD
    ''' </summary>
    ''' <value>Decimal representing the upper bound on values in the NUD</value>
    ''' -----------------------------------------------------------------------------
    Public Property MaxValue() As Decimal
        Get
            Return _MaxValue
        End Get
        Set(ByVal Value As Decimal)
            If _MaxValue > MinValue Then
                _MaxValue = Value
            End If
        End Set
    End Property

#End Region



#Region "Event Handlers and Overrides"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles loosing focus on the textbox to make sure a valid value is contained 
    ''' within the controls.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub txtNud_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNud.LostFocus
        If Not _AllowBlank AndAlso Trim(txtNud.Text).Length = 0 Then
            assignText(_LastKnownValue.ToString)
        End If

        handleChange(txtNud.Text)
        MyBase.OnLostFocus(e)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Overrides default wheel behavior to increment the NUD values.
    ''' </summary>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseWheel(ByVal e As System.Windows.Forms.MouseEventArgs)
        If e.Delta > 0 Then
            increment(True)
        Else
            increment(False)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles clicks on the text box, selecting the entire text of the nud the first 
    ''' time the click is registered.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub txtNud_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNud.Click
        If txtNud.Text.Length > 0 Then
            txtNud.SelectionStart = 0
            txtNud.SelectionLength = txtNud.Text.Length
        End If
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes in the text box entry, calling handleChange to ensure that 
    ''' only valid values are entered.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub txtNud_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNud.TextChanged
        handleChange(txtNud.Text)
        MyBase.OnTextChanged(e)
    End Sub

#End Region



#Region "Implementation Members"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs the value assignment to the control based upon settings in the 
    ''' NUD.  A couple rules are checked here and enforced as the value is assigned and 
    ''' then the last good value is saved in _lastKnownValue.
    ''' </summary>
    ''' <param name="s"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub assignText(ByVal s As String)
        If s = "." Then
            s = "0" & s
        End If

        If IsNumeric(s) Then
            _LastKnownValue = CDec(s)
        End If

        Dim curPos As Int32 = txtNud.SelectionStart

        If txtNud.SelectionLength > 0 Then
            curPos = txtNud.Text.Length
        End If

        txtNud.Text = s

        If curPos > txtNud.Text.Length Then
            txtNud.SelectionStart = txtNud.Text.Length

        End If
        If curPos <= txtNud.Text.Length Then
            txtNud.SelectionStart = curPos
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Checks a proposed value to make sure it conforms to the settings on the NUD. 
    ''' Returns a boolean indicating if the value is acceptable or not.
    ''' </summary>
    ''' <param name="sVal">The proposed value as a string</param>
    ''' <returns>A boolean indicating if the proposed value is valid or not.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function ProposedValueValid(ByVal sVal As String) As Boolean

        If sVal.Length = 0 Then
            Return True
        End If

        If IsNumeric(sVal) Then
            Dim val As Decimal = CDec(sVal)
            If val <= MaxValue AndAlso val >= MinValue Then
                'verify precision
                If sVal.IndexOf(".") > 0 Then

                    Dim str() As String = sVal.Split("."c)
                    Dim right As String = str(1)
                    If right.Length > _DecimalPlaces Then
                        Return False
                    Else
                        Return True
                    End If
                Else
                    Return True
                End If
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the nud control's values and either sets the value to 
    ''' the new value, or if invalid, sets the value to the last good value.
    ''' </summary>
    ''' <param name="sProposedValue">The proposed value as a string.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub handleChange(ByVal sProposedValue As String)
        If ProposedValueValid(sProposedValue) Then
            assignText(sProposedValue)
        Else
            assignText(_LastKnownValue.ToString)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles an increment request from one of the NUD buttons.
    ''' </summary>
    ''' <param name="bPositive">Whether the NUD should increment its value 
    ''' up (true) or down (false).</param>
    ''' -----------------------------------------------------------------------------
    Private Sub increment(ByVal bPositive As Boolean)
        Dim val As Decimal


        If IsNumeric(txtNud.Text) Then
            val = CDec(txtNud.Text)
        Else
            val = _LastKnownValue
        End If

        If bPositive Then
            val += _Increment
        Else
            val -= _Increment
        End If


        handleChange(val.ToString)


        If txtNud.Text.Length > 0 Then
            txtNud.SelectionStart = 0
            txtNud.SelectionLength = txtNud.Text.Length
        End If
    End Sub
#End Region

End Class
