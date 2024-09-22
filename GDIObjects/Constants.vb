
''' -----------------------------------------------------------------------------
''' <summary>
''' The Two types of code documents GDI+ Architect produces.
''' </summary>
''' <remarks>
''' PrintDocument is for  multiple pages of printed information.  
''' It creates a class that inherits from  System.Drawing.Printing.PrintDocument and 
''' contains code to draw documents page by page.
''' 
''' The graphics class produces code to create a custom graphics class document.  
''' A graphics class document is a single class to which a graphics context can be 
''' passed to draw upon a surface.
''' </remarks>
''' -----------------------------------------------------------------------------
Public Enum EnumDocumentTypes
    ''' <summary>The PrintDocument type of document.  Generates a PrintDocument which 
    ''' can print over multiple pages and inherits from the Framework's 
    ''' System.Drawing.Printing.PrintDocument class</summary>
    ePrintDocument
    ''' <summary>Document style class.  Generates a class that can draw itself to 
    ''' a System.Drawing.Graphics context.</summary>
    eClass
End Enum

''' -----------------------------------------------------------------------------
''' <summary>
''' Enumeration of handles for objects on the design surface.  Handles correspond to the 
''' rectangular areas on the edge of objects that can be manipulated, or on paths, to the 
''' points that can be moved.  In the latter case, ePointHandle represents this type of 
''' handle.
''' </summary>
''' <remarks>
''' The values of this enumeration is significant for arrays that manipulate sets of handles.
''' The order starts from the Left most point and rotates around clockwise back on itself.
''' </remarks>
''' -----------------------------------------------------------------------------
Public Enum EnumDragHandles As Integer
    ''' <summary> No handle has been selected</summary>
    eNone = -1
    ''' <summary> The left side handle</summary>
    eLeft = 0
    ''' <summary> The top left corner handle</summary>
    eTopLeft = 1
    ''' <summary> The top side handle</summary>
    eTop = 2
    ''' <summary> The top right corner handle</summary>
    eTopRight = 3
    ''' <summary> The right side handle</summary>
    eRight = 4
    ''' <summary> The bottom right corner handle</summary>
    eBottomRight = 5
    ''' <summary> The bottom side handle</summary>
    eBottom = 6
    ''' <summary> The bottom left corner handle</summary>
    eBottomleft = 7
    ''' <summary> An arbitrary point handle has been selected.  This is used in graphics 
    ''' paths which are manipulated by the points in the paths instead of by its bounding 
    ''' corner handles</summary>    
    ePointHandle = 100

End Enum


''' -----------------------------------------------------------------------------
''' <summary>
''' The types of code that GDI+ Architect produces (VB and C#).
''' </summary>
''' -----------------------------------------------------------------------------
Public Enum EnumCodeTypes As Integer
    ''' <summary>Code will be generated in Visual Basic</summary>
    eVB
    ''' <summary>Code will be generated in C#</summary>
    eCSharp
End Enum

''' -----------------------------------------------------------------------------
''' <summary>
''' EnumDrawMode is used to enumerate the reason objects are being rendered, and allows 
''' objects that need to paint differently to different types of surfaces to customize 
''' their painting.  For example, text objects draw borders at design time, but not at 
''' print preview time. 
''' </summary>
''' -----------------------------------------------------------------------------
Public Enum EnumDrawMode As Integer
    ''' <summary> Normal is drawing to the GDI+ Architect surface (a screen)</summary>
    eNormal
    ''' <summary> Printing is drawing to a PrintDocument.</summary>
    ePrinting
    ''' <summary> Graphics Export is drawing to a Bitmap or other GDI+ graphic item being 
    ''' exported</summary>
    eGraphicExport
End Enum

''' -----------------------------------------------------------------------------
''' <summary>
''' Enumeration of valid alignment modes.  When the user elects to align
''' an object they can align "normal" (to other selected objects), to the document's 
''' margins or to the entire canvas.
''' </summary>
''' -----------------------------------------------------------------------------
Public Enum EnumAlignMode As Integer
    ''' <summary> Normal indicates alignment occurs to other selected objects</summary>
    eNormal
    ''' <summary> eMargins indicates alignment occurs to margins </summary>
    eMargins
    ''' <summary> eCanvas indicates alignment occurs to the edges of the canvas </summary>
    eCanvas
End Enum

''' -----------------------------------------------------------------------------
''' <summary>
''' The way a resource (texture image or image) will be referenced in the generated code.
''' </summary>
''' <remarks>
''' In the code generated by GDI+ Architect, images can be referenced by one of three 
''' methods.  The user may indicate they plan to embed an image, refer to the image by an 
''' absolute path, or refer to the image by a relative path.  This Enumeration is used 
''' to denote which of the three options the user has selected.
''' </remarks>
''' -----------------------------------------------------------------------------
Public Enum EnumLinkType As Integer
    ''' <summary>The final resource will be embedded in the solution</summary>
    Embedded
    ''' <summary>The final resource will be specified as an absolute path to the resource</summary>
    AbsolutePath
    ''' <summary>The final resource will be specified as
    '''  a relative path from the BIN directory (runtime path)</summary>
    RelativePath
End Enum


''' -----------------------------------------------------------------------------
''' <summary>
''' The valid scopes for objects in generated code.
''' </summary>
''' -----------------------------------------------------------------------------
Public Enum EnumScope As Integer
    ''' <summary> Private field scope</summary>
    [Private] = 0
    ''' <summary>Protected field scope</summary>
    [Protected] = 1
    ''' <summary>Friend (Internal) scope</summary>
    [FriendFamilyInternal] = 2
    ''' <summary>Protected Friend scope (available in project and to inheritors)</summary>
    [ProtectedFriendFamilyInternal] = 3
    ''' <summary>Public scope</summary>
    [Public] = 4
End Enum



