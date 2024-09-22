Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : GDIPlusToolTips
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for managing all of the tooltips used in GDI+ Architect.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class ToolTipManager
    Inherits System.ComponentModel.Component



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the ToolTipManager.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        _ToolTip = New ToolTip
    End Sub

#Region "Local Declarations"


    '''<summary>A simple tooltip control that is used whenever we need 
    ''' a tool tip in GDI+ Architect</summary>
    Private _ToolTip As ToolTip

    '''<summary>Holds an XML document that contains all of the tool tips 
    ''' for the application organized by container and control.</summary>
    Private _xml As Xml.XmlDocument



#End Region



#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the XML document of tool tips from the embedded resource or from the _xml 
    ''' member if available.
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------
    Private ReadOnly Property ToolTipXML() As System.Xml.XmlDocument
        Get
            If Not _xml Is Nothing Then
                Return _xml
            Else
                Dim xmlToolStream As System.IO.Stream

                Try
                    _xml = New Xml.XmlDocument

                    xmlToolStream = System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream("GDIPlusArchitect.Tooltips.xml")

                    _xml.Load(xmlToolStream)
                Catch ex As Exception

                Finally
                    If Not xmlToolStream Is Nothing Then
                        xmlToolStream.Close()
                    End If
                End Try
            End If

            Return _xml
        End Get
    End Property

#End Region


#Region "Implementation Members"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a boolean indicating whether a given control has a tool tip or not.
    ''' </summary>
    ''' <param name="sSection">The section name to look under for the tool tip in the 
    ''' xml document</param>
    ''' <param name="ctl">The control to check for a tooltip for.</param>
    ''' <returns>A boolean indicating if a tooltip exists for the document</returns>
    ''' -----------------------------------------------------------------------------
    Private Function hasTipForControl(ByVal sSection As String, ByVal ctl As Control) As Boolean
        Dim pNode As Xml.XmlNode = ToolTipXML.ChildNodes(0).SelectSingleNode("toolset[@name= '" & sSection & "']/tip[@name='" & ctl.Name & "']")
        If pNode Is Nothing Then
            Return False
        Else
            Return True
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the tooltip text for a specific control within a specific section of the 
    ''' xml document.
    ''' </summary>
    ''' <param name="sSection">The section name to look under for the tool tip in the 
    ''' xml document</param>
    ''' <param name="ctl">The control to check for a tooltip for.</param>
    ''' <returns></returns>
    ''' -----------------------------------------------------------------------------
    Private Function getText(ByVal sSection As String, ByVal ctl As Control) As String
        Dim pNode As Xml.XmlNode = ToolTipXML.ChildNodes(0).SelectSingleNode("toolset[@name= '" & sSection & "']/tip[@name='" & ctl.Name & "']")
        Return pNode.InnerText.Replace("%", vbLf)

    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Assigns the appropriate tooltip in the xml source to a specific tooltip control 
    ''' </summary>
    ''' <param name="tip">The tool tip to assign to</param>
    ''' <param name="sSection">The section to look for the tool tip in.</param>
    ''' <param name="ctls">The control to assign the tip to.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub AssignTips(ByVal tip As ToolTip, ByVal sSection As String, ByVal ctls As Control.ControlCollection)
        For Each ctl As Control In ctls
            If ctl.HasChildren Then
                AssignTips(tip, sSection, ctl.Controls)
            End If

            If hasTipForControl(sSection, ctl) Then
                If TypeOf ctl Is NumericUpDown Then
                    tip.SetToolTip(ctl.Controls(0), getText(sSection, ctl))
                    tip.SetToolTip(ctl.Controls(1), getText(sSection, ctl))
                Else
                    tip.SetToolTip(ctl, getText(sSection, ctl))
                End If
            End If

        Next
    End Sub

#End Region

#Region "Public Members"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populates a pop up form (one that does not exist for the lifetime of the application)
    ''' with its appropriate tool tips
    ''' </summary>
    ''' <param name="tip">The tool tip control to assign the tips to.</param>
    ''' <param name="Section">The section to look for tooltips in the xml document for</param>
    ''' <param name="frm">The Windows form that needs tool tip information.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub PopulatePopupTip(ByVal tip As ToolTip, ByVal Section As String, ByVal frm As System.Windows.Forms.Form)

        If App.Options.ShowToolTips Then
            If tip Is Nothing Then
                tip = New ToolTip
            End If
            AssignTips(tip, Section, frm.Controls)
        End If

    End Sub



 


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Assigns the appropriate tooltip in the xml source to a specific tooltip control 
    ''' </summary>
    ''' <param name="tip">The tool tip to assign to</param>
    ''' <param name="sSection">The section to look for the tool tip in.</param>
    ''' <param name="ctls">The control to assign the tip to.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub AssignPanelTip(ByVal sSection As String, ByVal ctls As Control.ControlCollection)
        For Each ctl As Control In ctls
            If ctl.HasChildren Then
                AssignTips(_ToolTip, sSection, ctl.Controls)
            End If

            If hasTipForControl(sSection, ctl) Then
                If TypeOf ctl Is NumericUpDown Then
                    _ToolTip.SetToolTip(ctl.Controls(0), getText(sSection, ctl))
                    _ToolTip.SetToolTip(ctl.Controls(1), getText(sSection, ctl))
                Else
                    _ToolTip.SetToolTip(ctl, getText(sSection, ctl))
                End If
            End If

        Next
    End Sub

#End Region

#Region "Disposal and Cleanup"
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not _ToolTip Is Nothing Then
            _ToolTip.Dispose()
        End If

        MyBase.Dispose(disposing)
    End Sub
#End Region


End Class
