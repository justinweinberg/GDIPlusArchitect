Imports System.Design
Imports System.Windows.Forms.Design

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : OpenFolderDialog
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a mechanism for selecting a folder
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class OpenFolderDialog
    Inherits FolderNameEditor

#Region "Local Declarations"
    '''<summary>A folder browser control </summary>
    Private _FolderBrowser As FolderBrowser
#End Region


#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the OpenFolderDialog 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        _FolderBrowser = New FolderBrowser
    End Sub

#End Region

#Region "Public Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Requests a folder from the user using the folder browser and returns the 
    ''' path returned from the folder browser.
    ''' </summary>
    ''' <returns></returns>
    ''' -----------------------------------------------------------------------------
    Public Function GetFolder() As String

        Dim iresp As DialogResult = _FolderBrowser.ShowDialog()
        Return _FolderBrowser.DirectoryPath
    End Function

#End Region
End Class
