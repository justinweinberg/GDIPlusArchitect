

''' -----------------------------------------------------------------------------
''' <summary>
''' The types of colors supported (named or argb style)
''' </summary>
''' -----------------------------------------------------------------------------
Public Enum ColorFormat
    NamedColor = 0
    ARGBColor
End Enum
 
''' -----------------------------------------------------------------------------
''' <summary>
''' The tools that GDI+ Architect supports.
''' </summary>
''' -----------------------------------------------------------------------------
Public Enum EnumTools As Integer
    ePointer = 0
    eFill
    eDropper
    eSquare
    eCircle
    eText
    eLine
    ePen
    ePlacing
    eField
    eMagnify
    eHand

End Enum

Friend Class Constants


    Friend Const URL_BUYSOURCE As String = "www.mrgsoft.com/buy/"
    'File name to save GDI+ Architect options to.
    Friend Const CONST_OPTIONS_FILE_NAME As String = "gdiarchoptions.xml"
    'File name to save docking preferences to
    Friend Const CONST_FILE_DOCKING As String = "gdiarchdocking.dat"

    'URL to magic's site
    Friend Const URL_DOTNETMAGIC As String = "www.dotnetmagic.com"


End Class
