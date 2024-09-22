Imports System.CodeDom


''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIObjCol
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Maintains an ordered collection of GDIObjects and provides helper methods to work with 
''' sets of GDIObjects.  The order of the collection determines Z-Order when rendering to 
''' surfaces.
''' </summary>
''' <remarks>The GDIObjCol maintains a collection GDIObjects for both GDIPages and 
''' the SelectedSet object which holds the sets of selected objects.
''' </remarks>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDIObjCol
    Inherits CollectionBase

#Region "Local fields"



 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Used for clipboard operations.  Sets the dataformat for the GDIObjCol.  This is 
    ''' later used to identify whether a resource is pastable into the project 
    ''' or not.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private Shared _DataFormat As System.Windows.Forms.DataFormats.Format = System.Windows.Forms.DataFormats.GetFormat(GetType(GDIObjCol).FullName)

#End Region

#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a GDIObjCol.  Sets up the dataformat provider.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        _DataFormat = System.Windows.Forms.DataFormats.GetFormat(GetType(GDIObjCol).FullName)
    End Sub



#End Region
 

#Region "Drawing related methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws the collection as selected (with appropriate handles, highlight, etc).
    ''' </summary>
    ''' <param name="g">Drawing context to draw the set to</param>
    ''' <param name="fscale">Current surface zoom factor.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub DrawSelected(ByVal g As Graphics, ByVal fscale As Single)
        Dim gCon As Drawing2D.GraphicsContainer

        With g
            gCon = .BeginContainer
            .ScaleTransform(fscale, fscale)
        End With



        'Ask each item to draw itself to the graphic container
        If Me.HasItems() Then
            For Each gObj As GDIObject In innerlist
                gObj.DrawSelectedObject(g, fscale)
            Next

        End If

        'reset the graphics object to its original value
        g.EndContainer(gCon)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Highlights a specific object.  This arguably doesn't belong in the collection's 
    ''' implementation since it uses no local members nor has much to do with a collection
    ''' of objects.
    ''' </summary>
    ''' <param name="g">The graphics context being draw to.</param>
    ''' <param name="obj">The object to highlight</param>
    ''' <param name="fscale">The current scale factor.</param>
    ''' <param name="highlightColor">The highlight color to highlight with.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub highlightObject(ByVal g As Graphics, ByVal obj As GDIObject, ByVal fscale As Single, ByVal highlightColor As Color)
        Session.GraphicsManager.BeginScaleMode(g, fscale)

        obj.HighlightObject(g, fscale, highlightColor)

        Session.GraphicsManager.EndScaleMode(g)

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws a set of objects to the specific graphics surface.
    ''' </summary>
    ''' <param name="g">The graphics context to draw to.</param>
    ''' <param name="fScale">the current scale factor of the surface.</param>
    ''' <param name="mode">The draw mode. See EnumDrawMode for more information.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub DrawObjects(ByVal g As Graphics, ByVal fScale As Single, ByVal mode As GDIObjects.EnumDrawMode)
        Dim gobj As GDIObject

        Session.GraphicsManager.BeginScaleMode(g, fScale)
        'Ask each item to draw itself to the graphic container
        If Me.HasItems() Then
            For i As Int32 = 0 To InnerList.Count - 1
                gobj = CType(InnerList(i), GDIObject)
                gobj.Draw(g, mode)
            Next
        End If


        Session.GraphicsManager.EndScaleMode(g)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Prints the set of objects to a PrintDocument.
    ''' </summary>
    ''' <param name="g">Graphics context to print the objects to.</param>
    ''' <param name="doc">parent GDIDocument.  Smoothing and TextHint are retrieved 
    ''' from this.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub PrintObjects(ByVal g As Graphics)
        Trace.WriteLineIf(Session.Tracing, "Invoking object collection print")

        Dim drawObj As GDIObject

        If Me.HasItems() Then
            For i As Int32 = 0 To Me.InnerList.Count - 1
                drawObj = CType(Me.InnerList(i), GDIObject)
                drawObj.Draw(g, GDIObjects.EnumDrawMode.ePrinting)
            Next
        End If
    End Sub

#End Region


#Region "Alignment and Position methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Nudges all of the items in the set by an x and/or a y increment.
    ''' </summary>
    ''' <param name="x">The x amount to nudge by</param>
    ''' <param name="y">The y amount to nudge by</param>
    ''' <remarks>X and Y can be negative.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Friend Sub Nudge(ByVal x As Int32, ByVal y As Int32)
        Trace.WriteLineIf(Session.Tracing, "nudge:" & " " & x & ", " & y)

        For Each selobj As GDIObject In Me.InnerList
            selobj.Bounds = New Rectangle(selobj.Bounds.X + x, selobj.Bounds.Y + y, selobj.Bounds.Width, selobj.Bounds.Height)
        Next

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the bottom most point in the set
    ''' </summary>
    ''' <returns>The bottom most point in the set</returns>
    ''' -----------------------------------------------------------------------------
    Friend Function extentBottom() As Int32
        Dim lastY As Int32 = Int32.MinValue

        For Each gdiobj As GDIObject In List
            If gdiobj.Rotation > 0 Then
                Dim bottomMostPt As Point = bottomMostPoint(gdiobj.RotatedBoundPoints)
                If bottomMostPt.Y > lastY Then
                    lastY = bottomMostPt.Y
                End If
            Else
                If gdiobj.Bounds.Y + gdiobj.Bounds.Height > lastY Then
                    lastY = gdiobj.Bounds.Y + gdiobj.Bounds.Height
                End If
            End If
        Next

        Return lastY

    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the top most point in the set
    ''' </summary>
    ''' <returns>The top most point in the set</returns>
    ''' -----------------------------------------------------------------------------
    Friend Function extentTop() As Int32
        Dim lastY As Int32 = Int32.MaxValue

        For Each gdiobj As GDIObject In List
            If gdiobj.Rotation > 0 Then
                Dim topMostPt As Point = topMostPoint(gdiobj.RotatedBoundPoints)
                If topMostPt.Y < lastY Then
                    lastY = topMostPt.Y
                End If
            Else
                If gdiobj.Bounds.Y < lastY Then
                    lastY = gdiobj.Bounds.Y
                End If
            End If
        Next


        Return lastY

    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the right most point in the set
    ''' </summary>
    ''' <returns>The right most point in the set</returns>
    ''' ----------------------------------------------------------------------------- 
    Friend Function extentRight() As Int32
        Dim lastX As Int32 = Int32.MinValue

        For Each gdiobj As GDIObject In List
            If gdiobj.Rotation > 0 Then
                Dim rightMostpt As Point = rightMostPoint(gdiobj.RotatedBoundPoints)
                If rightMostpt.X > lastX Then
                    lastX = rightMostpt.X

                End If
            Else
                If gdiobj.Bounds.X + gdiobj.Bounds.Width > lastX Then
                    lastX = gdiobj.Bounds.X + gdiobj.Bounds.Width

                End If
            End If
        Next


        Return lastX

    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the left most point in the set
    ''' </summary>
    ''' <returns>The left most point in the set</returns>
    ''' -----------------------------------------------------------------------------  
    Friend Function extentLeft() As Int32
        Dim lastX As Int32 = 100000
        Dim outObj As GDIObject
        For Each gdiobj As GDIObject In List
            If gdiobj.Rotation > 0 Then
                Dim leftMostpt As Point = leftMostPoint(gdiobj.RotatedBoundPoints)
                If leftMostpt.X < lastX Then
                    lastX = leftMostpt.X

                End If
            Else
                If gdiobj.Bounds.X < lastX Then
                    lastX = gdiobj.Bounds.X

                End If
            End If
        Next


        Return lastX

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sorts all of the items in the set by their Y values, returning a new sorted set.
    ''' </summary>
    ''' <returns>A set of GDIObjects sorted by their Y positions</returns>
    ''' <remarks>Within a GDIObjCol, objects are sorted by their Z Order values.  Prior to 
    ''' distributing heights, thee need to be ordered by their Y positions.  This method returns 
    ''' the objects in such a way that height distribution can be performed.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Function insertionSortY() As ArrayList
        Dim i As Int32 = -1

        Dim tempArr As New ArrayList(Me.InnerList)

        For i = 0 To tempArr.Count - 1
            Dim Temp As GDIObject = DirectCast(tempArr(i), GDIObject)

            Dim j As Int32

            For j = i - 1 To 0 Step -1
                If Temp.Bounds.Y >= DirectCast(tempArr(j), GDIObject).Bounds.Y Then
                    Exit For
                End If

                tempArr(j + 1) = tempArr(j)
            Next

            tempArr(j + 1) = Temp

        Next

        Return tempArr
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sorts all of the items in the set by their X values, returning a new sorted set.
    ''' </summary>
    ''' <returns>A set of GDIObjects sorted by their X positions</returns>
    ''' <remarks>Within a GDIObjCol, objects are sorted by their Z Order values.  Prior to 
    ''' distributing widths, they nede to be ordered by their X positions.  This method returns 
    ''' the objects in such a way that width distribution can be performed.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Function InsertionSortX() As ArrayList
        Dim i As Int32
        Dim j As Int32

        Dim tempArr As New ArrayList(Me.InnerList)

        For i = 0 To tempArr.Count - 1
            Dim Temp As GDIObject = DirectCast(tempArr(i), GDIObject)

            For j = i - 1 To 0 Step -1
                If Temp.Bounds.X >= DirectCast(tempArr(j), GDIObject).Bounds.X Then
                    Exit For
                End If

                tempArr(j + 1) = tempArr(j)
            Next

            tempArr(j + 1) = Temp

        Next

        Return tempArr
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the right most point in array of points.
    ''' </summary>
    ''' <param name="ptArr">Array to find the right most point in.</param>
    ''' <returns>The point with the greatest X value</returns>
    ''' -----------------------------------------------------------------------------
    Private Function rightMostPoint(ByVal ptArr As Point()) As Point
        Dim lastX As Int32 = Int32.MinValue
        Dim outpt As Point
        For Each pt As Point In ptArr
            If pt.X > lastX Then
                lastX = pt.X
                outpt = pt
            End If
        Next

        Return outpt
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the bottom most point in array of points.
    ''' </summary>
    ''' <param name="ptArr">Array to find the bottom most point in.</param>
    ''' <returns>The point in the set that has the greatest Y value</returns>
    ''' -----------------------------------------------------------------------------
    Private Function bottomMostPoint(ByVal ptArr As Point()) As Point
        Dim lasty As Int32 = Int32.MinValue
        Dim outpt As Point
        For Each pt As Point In ptArr
            If pt.Y > lasty Then
                lasty = pt.Y
                outpt = pt
            End If
        Next

        Return outpt
    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the top most point in array of points.
    ''' </summary>
    ''' <param name="ptArr">Array to find the top most point in.</param>
    ''' <returns>The point with the smallest Y value in the array.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function topMostPoint(ByVal ptArr As Point()) As Point
        Dim lasty As Int32 = Int32.MaxValue
        Dim outpt As Point
        For Each pt As Point In ptArr
            If pt.Y < lasty Then
                lasty = pt.Y
                outpt = pt
            End If
        Next

        Return outpt
    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the left most point in array of points.
    ''' </summary>
    ''' <param name="ptArr">Array to find the left most point in.</param>
    ''' <returns>The point with the smallest X value in the array.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function leftMostPoint(ByVal ptArr As Point()) As Point
        Dim lastX As Int32 = Int32.MaxValue
        Dim outpt As Point
        For Each pt As Point In ptArr
            If pt.X < lastX Then
                lastX = pt.X
                outpt = pt
            End If
        Next

        Return outpt
    End Function




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Distributes the heights of selected objects along a specific height bounds.
    ''' </summary>
    ''' <param name="bounds">Rectangle from which the height and Y is extracted to distribute 
    ''' bounds.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub distributeHeights(ByVal bounds As Rectangle)
        Trace.WriteLineIf(Session.Tracing, "distributeHeights")

        If list.Count > 1 Then


            Dim ysortedList As ArrayList = insertionSortY()

            Dim totalHeight As Int32 = 0

            Dim iTotalSpace As Int32 = 0
            Dim iEvenSpace As Int32

            For Each gdiobj As GDIObject In list
                If gdiobj.Rotation > 0 Then
                    Dim ptTopmost As Point = topMostPoint(gdiobj.RotatedBoundPoints)
                    Dim ptBottommost As Point = bottomMostPoint(gdiobj.RotatedBoundPoints)
                    totalHeight += ptBottommost.Y - ptTopmost.Y
                Else
                    totalHeight += gdiobj.Bounds.Height
                End If

            Next


            Dim iAreaCovered As Int32 = bounds.Height
            iTotalSpace = iAreaCovered - totalHeight

            iEvenSpace = iTotalSpace \ (list.Count - 1)

            Dim iLastY As Int32 = bounds.Y

            For i As Int32 = 0 To ysortedList.Count - 1
                Dim gdiobj As GDIObject = DirectCast(ysortedList(i), GDIObject)


                If gdiobj.Rotation > 0 Then
                    Dim curY As Int32 = topMostPoint(gdiobj.RotatedBoundPoints).Y

                    gdiobj.Bounds = New Rectangle(gdiobj.Bounds.x, gdiobj.Bounds.Y - (curY - iLastY), gdiobj.Bounds.Width, gdiobj.Bounds.Height)

                    Dim pts() As Point = gdiobj.RotatedBoundPoints
                    iLastY += iEvenSpace + bottomMostPoint(pts).Y - topMostPoint(pts).Y
                Else
                    gdiobj.Bounds = New Rectangle(gdiobj.Bounds.x, iLastY, gdiobj.Bounds.Width, gdiobj.Bounds.Height)
                    iLastY += iEvenSpace + gdiobj.Bounds.height
                End If

            Next



            ysortedList = Nothing
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Distributes heights within the set in relation to each other.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Sub distributeHeights()
        Trace.WriteLineIf(Session.Tracing, "distributeHeights")

        If list.Count > 2 Then

            Dim ysortedList As ArrayList = insertionSortY()

            Dim totalHeight As Int32 = 0

            Dim topObj As GDIObject = DirectCast(ysortedList(0), GDIObject)
            Dim bottomObj As GDIObject = DirectCast(ysortedList(ysortedList.Count - 1), GDIObject)

            Dim iAreaCovered As Int32 = 0
            Dim iTotalSpace As Int32 = 0
            Dim iEvenSpace As Int32

            For Each gdiobj As GDIObject In list
                If gdiobj.Rotation > 0 Then
                    Dim ptTopmost As Point = topMostPoint(gdiobj.RotatedBoundPoints)
                    Dim ptBottommost As Point = bottomMostPoint(gdiobj.RotatedBoundPoints)
                    totalHeight += ptBottommost.Y - ptTopmost.Y
                Else
                    totalHeight += gdiobj.Bounds.Height
                End If

            Next

            Dim bottomBounds As Rectangle
            Dim topBounds As Rectangle

            If bottomObj.Rotation > 0 Then
                Dim mtx As New Drawing2D.Matrix
                Dim pth As New Drawing2D.GraphicsPath
                pth.AddRectangle(bottomObj.Bounds)
                mtx.RotateAt(bottomObj.Rotation, bottomObj.RotationPoint)
                pth.Transform(mtx)
                bottomBounds = Rectangle.Round(pth.GetBounds())
                pth.Dispose()
            Else
                bottomBounds = bottomObj.Bounds
            End If

            If topObj.Rotation > 0 Then
                Dim mtx As New Drawing2D.Matrix
                Dim pth As New Drawing2D.GraphicsPath
                pth.AddRectangle(topObj.Bounds)
                mtx.RotateAt(topObj.Rotation, topObj.RotationPoint)
                pth.Transform(mtx)
                topBounds = Rectangle.Round(pth.GetBounds())
                pth.Dispose()
            Else
                topBounds = topObj.Bounds
            End If

            iAreaCovered = (bottomBounds.Y + bottomBounds.Height) - topBounds.Y
            iTotalSpace = iAreaCovered - totalHeight

            iEvenSpace = iTotalSpace \ (list.Count - 1)

            Dim iLastY As Int32 = topBounds.Y

            For i As Int32 = 0 To ysortedList.Count - 2
                Dim gdiobj As GDIObject = DirectCast(ysortedList(i), GDIObject)


                If gdiobj.Rotation > 0 Then
                    Dim curY As Int32 = topMostPoint(gdiobj.RotatedBoundPoints).Y

                    gdiobj.Bounds = New Rectangle(gdiobj.Bounds.x, gdiobj.Bounds.Y - (curY - iLastY), gdiobj.Bounds.Width, gdiobj.Bounds.Height)

                    Dim pts() As Point = gdiobj.RotatedBoundPoints
                    iLastY += iEvenSpace + bottomMostPoint(pts).Y - topMostPoint(pts).Y
                Else
                    gdiobj.Bounds = New Rectangle(gdiobj.Bounds.x, iLastY, gdiobj.Bounds.Width, gdiobj.Bounds.Height)
                    iLastY += iEvenSpace + gdiobj.Bounds.height
                End If

            Next


            ysortedList = Nothing
        End If


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Distributes the widths of items in the set in relation to a specific width
    ''' </summary>
    ''' <param name="bounds">A rectangle from which the width and X values are extracted 
    ''' in order to distribute widths.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub distributeWidths(ByVal bounds As Rectangle)
        Trace.WriteLineIf(Session.Tracing, "distributeWidths")

        If list.Count > 1 Then

            Dim xsortedlist As ArrayList = InsertionSortX()

            Dim totalWidth As Int32 = 0

            Dim iTotalSpace As Int32 = 0
            Dim iEvenSpace As Int32

            For Each gdiobj As GDIObject In list
                If gdiobj.Rotation > 0 Then
                    Dim ptLeftmost As Point = leftMostPoint(gdiobj.RotatedBoundPoints)
                    Dim ptRightMost As Point = rightMostPoint(gdiobj.RotatedBoundPoints)
                    totalWidth += ptRightMost.X - ptLeftmost.X
                Else
                    totalWidth += gdiobj.Bounds.Width
                End If
            Next


            Dim iAreaCovered As Int32 = bounds.Width
            iTotalSpace = iAreaCovered - totalWidth

            iEvenSpace = iTotalSpace \ (list.Count - 1)

            Dim iLastX As Int32 = bounds.X

            For i As Int32 = 0 To xsortedlist.Count - 1
                Dim gdiobj As GDIObject = DirectCast(xsortedlist(i), GDIObject)


                If gdiobj.Rotation > 0 Then
                    Dim curX As Int32 = leftMostPoint(gdiobj.RotatedBoundPoints).X

                    gdiobj.Bounds = New Rectangle(gdiobj.Bounds.x - (curX - iLastX), gdiobj.Bounds.y, gdiobj.Bounds.Width, gdiobj.Bounds.Height)

                    Dim pts() As Point = gdiobj.RotatedBoundPoints
                    iLastX += iEvenSpace + rightMostPoint(pts).X - leftMostPoint(pts).X
                Else
                    gdiobj.Bounds = New Rectangle(iLastX, gdiobj.Bounds.y, gdiobj.Bounds.Width, gdiobj.Bounds.Height)
                    iLastX += iEvenSpace + gdiobj.Bounds.Width
                End If
            Next


            xsortedlist = Nothing
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Distributes widths within a set in relation to the set's members.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Sub distributeWidths()
        Trace.WriteLineIf(Session.Tracing, "distributeWidths")

        If list.Count > 2 Then

            Dim xsortedList As ArrayList = InsertionSortX()

            Dim totalWidth As Int32 = 0

            Dim leftObj As GDIObject = DirectCast(xsortedList(0), GDIObject)
            Dim rightObj As GDIObject = DirectCast(xsortedList(xsortedList.Count - 1), GDIObject)

            Dim iAreaCovered As Int32 = 0
            Dim iTotalSpace As Int32 = 0
            Dim iEvenSpace As Int32

            For Each gdiobj As GDIObject In list
                If gdiobj.Rotation > 0 Then
                    Dim ptLeftmost As Point = leftMostPoint(gdiobj.RotatedBoundPoints)
                    Dim ptRightMost As Point = rightMostPoint(gdiobj.RotatedBoundPoints)
                    totalWidth += ptRightMost.X - ptLeftmost.X
                Else
                    totalWidth += gdiobj.Bounds.Width
                End If

            Next

            Dim leftBounds As Rectangle
            Dim rightBounds As Rectangle

            If leftObj.Rotation > 0 Then
                Dim mtx As New Drawing2D.Matrix
                Dim pth As New Drawing2D.GraphicsPath
                pth.AddRectangle(leftObj.Bounds)
                mtx.RotateAt(leftObj.Rotation, leftObj.RotationPoint)
                pth.Transform(mtx)
                leftBounds = Rectangle.Round(pth.GetBounds())
                pth.Dispose()
            Else
                leftBounds = leftObj.Bounds
            End If

            If rightObj.Rotation > 0 Then
                Dim mtx As New Drawing2D.Matrix
                Dim pth As New Drawing2D.GraphicsPath
                pth.AddRectangle(rightObj.Bounds)
                mtx.RotateAt(rightObj.Rotation, rightObj.RotationPoint)
                pth.Transform(mtx)
                rightBounds = Rectangle.Round(pth.GetBounds())
                pth.Dispose()
            Else
                rightBounds = rightObj.Bounds
            End If

            iAreaCovered = (rightBounds.X + rightBounds.Width) - leftBounds.X
            iTotalSpace = iAreaCovered - totalWidth

            iEvenSpace = iTotalSpace \ (list.Count - 1)

            Dim iLastX As Int32 = leftBounds.X

            For i As Int32 = 0 To xsortedList.Count - 2
                Dim gdiobj As GDIObject = DirectCast(xsortedList(i), GDIObject)


                If gdiobj.Rotation > 0 Then
                    Dim curX As Int32 = leftMostPoint(gdiobj.RotatedBoundPoints).X

                    gdiobj.Bounds = New Rectangle(gdiobj.Bounds.x - (curX - iLastX), gdiobj.Bounds.y, gdiobj.Bounds.Width, gdiobj.Bounds.Height)

                    Dim pts() As Point = gdiobj.RotatedBoundPoints
                    iLastX += iEvenSpace + rightMostPoint(pts).X - leftMostPoint(pts).X
                Else
                    gdiobj.Bounds = New Rectangle(iLastX, gdiobj.Bounds.y, gdiobj.Bounds.Width, gdiobj.Bounds.Height)
                    iLastX += iEvenSpace + gdiobj.Bounds.Width
                End If

            Next


            xsortedList = Nothing
        End If



    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the set of objects to a specific left position
    ''' </summary>
    ''' <param name="left">The left position to align to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub alignLeft(ByVal left As Int32)
        Trace.WriteLineIf(Session.Tracing, "alignLeft")

        For Each selobj As GDIObject In Me.InnerList

            If selobj.Rotation > 0 Then
                Dim rect As Rectangle = selobj.Bounds
                Dim ptLeftMost As Point = leftMostPoint(selobj.RotatedBoundPoints)
                Dim diff As Int32 = left - ptLeftMost.X

                rect.Offset(diff, 0)
                selobj.Bounds = rect
            Else
                selobj.Bounds = New Rectangle(left, selobj.Bounds.Y, selobj.Bounds.Width, selobj.Bounds.Height)

            End If
        Next
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the set of objects to the leftmost object in the set.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Sub alignLeft()
        Trace.WriteLineIf(Session.Tracing, "alignLeft")

        Dim left As Int32 = Me.extentLeft()

        alignLeft(left)



    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the set of objects to a specific right most position.
    ''' </summary>
    ''' <param name="right">The right value to align to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub alignRight(ByVal right As Int32)
        Trace.WriteLineIf(Session.Tracing, "alignRight")

        For Each selobj As GDIObject In Me.InnerList

            If selobj.Rotation > 0 Then
                Dim rect As Rectangle = selobj.Bounds
                Dim ptrightMost As Point = rightMostPoint(selobj.RotatedBoundPoints)
                Dim diff As Int32 = right - ptrightMost.X

                rect.Offset(diff, 0)
                selobj.Bounds = rect
            Else
                selobj.Bounds = New Rectangle(right - selobj.Bounds.Width, selobj.Bounds.Y, selobj.Bounds.Width, selobj.Bounds.Height)

            End If
        Next


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the set of objects to the right most point in the set
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Sub alignRight()
        Trace.WriteLineIf(Session.Tracing, "alignRight")

        Dim right As Int32 = Me.extentRight

        alignRight(right)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the set of objects to a specific bottom most value.
    ''' </summary>
    ''' <param name="bottom">The bottom most value to align to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub alignBottom(ByVal bottom As Int32)
        Trace.WriteLineIf(Session.Tracing, "alignBottom")


        For Each selobj As GDIObject In Me.InnerList

            If selobj.Rotation > 0 Then
                Dim rect As Rectangle = selobj.Bounds
                Dim ptBottommost As Point = bottomMostPoint(selobj.RotatedBoundPoints)
                Dim diff As Int32 = bottom - ptBottommost.Y

                rect.Offset(0, diff)
                selobj.Bounds = rect
            Else
                selobj.Bounds = New Rectangle(selobj.Bounds.X, bottom - selobj.Bounds.Height, selobj.Bounds.Width, selobj.Bounds.Height)

            End If
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the set of objects to the bottom most point in the set.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Sub alignBottom()
        Trace.WriteLineIf(Session.Tracing, "alignBottom")

        Dim bottom As Int32 = extentBottom()

        alignBottom(bottom)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the set of objects topmost to a specific fixed value
    ''' </summary>
    ''' <param name="top">The fixed value to align objects to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub alignTop(ByVal top As Int32)

        For Each selobj As GDIObject In Me.InnerList

            If selobj.Rotation > 0 Then
                Dim rect As Rectangle = selobj.Bounds
                Dim ptTopmost As Point = topMostPoint(selobj.RotatedBoundPoints)
                Dim diff As Int32 = top - ptTopmost.Y

                rect.Offset(0, diff)
                selobj.Bounds = rect
            Else
                selobj.Bounds = New Rectangle(selobj.Bounds.X, top, selobj.Bounds.Width, selobj.Bounds.Height)

            End If
        Next

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns all objects to the topmost point in the selected set.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Sub alignTop()
        Dim top As Int32 = extentTop()

        alignTop(top)

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the objects in the set's height to a specific value
    ''' </summary>
    ''' <param name="height">The height value to set the object's height to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub makeHeightEqual(ByVal height As Int32)
        Trace.WriteLineIf(Session.Tracing, "makeHeightEqual")

        For Each selobj As GDIObject In Me.InnerList
            selobj.Bounds = New Rectangle(selobj.Bounds.X, selobj.Bounds.Y, selobj.Bounds.Width, height)
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the height of the objects in the set the same as a specified key object.
    ''' </summary>
    ''' <param name="keyobj">The object upon which to base the other object's heights.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub makeHeightEqual(ByVal keyobj As GDIObject)
        Trace.WriteLineIf(Session.Tracing, "makeHeightEqual: " & keyobj.Name)

        For Each selobj As GDIObject In Me.InnerList
            selobj.Bounds = New Rectangle(selobj.Bounds.X, selobj.Bounds.Y, selobj.Bounds.Width, keyobj.Bounds.Height)
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Makes the width of the set equal to a specific width value.
    ''' </summary>
    ''' <param name="width">The width value to set the width of the objects to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub makeWidthEqual(ByVal width As Int32)
        Trace.WriteLineIf(Session.Tracing, "makeWidthEqual")

        For Each selobj As GDIObject In Me.InnerList
            selobj.Bounds = New Rectangle(selobj.Bounds.X, selobj.Bounds.Y, width, selobj.Bounds.Height)
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Makes the items in a set have equal width to a specific key object.
    ''' </summary>
    ''' <param name="keyobj">The key object to set widths to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub makeWidthEqual(ByVal keyobj As GDIObject)
        Trace.WriteLineIf(Session.Tracing, "makeWidthEqual" & keyobj.Name)

        For Each selobj As GDIObject In Me.InnerList
            selobj.Bounds = New Rectangle(selobj.Bounds.X, selobj.Bounds.Y, _
            keyobj.Bounds.Width, selobj.Bounds.Height)
        Next
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Makes the items in a set have equal height and width to a specified size structure.
    ''' </summary>
    ''' <param name="size">Size structure (height/width elements) to size to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub makeHeightWidthEqual(ByVal size As Size)
        Trace.WriteLineIf(Session.Tracing, "makeHeightWidthEqual")

        makeWidthEqual(size.Width)
        makeHeightEqual(size.Height)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Makes the items in a set have equal height and width to a specified GDIObject.
    ''' </summary>
    ''' <param name="keyobj">The object which to set equal height and width to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub makeHeightWidthEqual(ByVal keyobj As GDIObject)
        Trace.WriteLineIf(Session.Tracing, "makeHeightWidthEqual: " & keyobj.Name)

        makeWidthEqual(keyobj)
        makeHeightEqual(keyobj)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the set of objects along a vertical center.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Sub alignHorizCenter()
        Trace.WriteLineIf(Session.Tracing, "alignHorizCenter")

        Dim iCenterTotal As Int32


        For Each selobj As GDIObject In Me.InnerList
            iCenterTotal += selobj.Bounds.Y + (selobj.Bounds.Height \ 2)
        Next

        iCenterTotal = iCenterTotal \ Me.InnerList.Count

        For Each selobj As GDIObject In Me.InnerList
            selobj.Bounds = New Rectangle(selobj.Bounds.X, iCenterTotal - (selobj.Bounds.Height \ 2), selobj.Bounds.Width, selobj.Bounds.Height)

        Next
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the set along a specified horizontal center
    ''' </summary>
    ''' <param name="horizCenter">A specific value to align upon</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub alignHorizCenter(ByVal horizCenter As Int32)
        Trace.WriteLineIf(Session.Tracing, "alignHorizCenter")


        For Each selobj As GDIObject In Me.InnerList
            selobj.Bounds = New Rectangle(selobj.Bounds.X, horizCenter - (selobj.Bounds.Height \ 2), selobj.Bounds.Width, selobj.Bounds.Height)

        Next
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the set along a specified vertical center 
    ''' </summary>
    ''' <param name="verticalcenter">A specific value to align upon</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub alignVertCenter(ByVal verticalcenter As Int32)
        Trace.WriteLineIf(Session.Tracing, "alignVertCenter")

        For Each selobj As GDIObject In Me.InnerList
            selobj.Bounds = New Rectangle(verticalcenter - (selobj.Bounds.Width \ 2), selobj.Bounds.Y, selobj.Bounds.Width, selobj.Bounds.Height)
        Next
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the set along the vertical center of the set.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Sub alignVertCenter()
        Trace.WriteLineIf(Session.Tracing, "alignVertCenter")

        Dim iCenterTotal As Int32


        For Each selobj As GDIObject In Me.InnerList
            iCenterTotal += selobj.Bounds.x + (selobj.Bounds.Width \ 2)
        Next

        iCenterTotal = iCenterTotal \ Me.InnerList.Count

        For Each selobj As GDIObject In Me.InnerList
            selobj.Bounds = New Rectangle(iCenterTotal - (selobj.Bounds.Width \ 2), selobj.Bounds.Y, selobj.Bounds.Width, selobj.Bounds.Height)
        Next
    End Sub

#End Region


#Region "Hit Test Related members"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the subset of objects within the collection that intersect with a specific 
    ''' rectangle.  The method name is somewhat of a misnomer.
    ''' </summary>
    ''' <param name="boundrect">The rectangle to test against for intersection</param>
    ''' <returns>A subset of objects bound by the rectangle.</returns>
    ''' -----------------------------------------------------------------------------

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the subset of objects within the collection that intersect with a specific 
    ''' rectangle.  The method name is somewhat of a misnomer.
    ''' </summary>
    ''' <param name="boundrect">The rectangle to test against for intersection</param>
    ''' <returns>A subset of objects bound by the rectangle.</returns>
    ''' -----------------------------------------------------------------------------
    Friend Function getBoundObjects(ByVal boundrect As Rectangle) As GDIObjCol

        Dim tempCol As New GDIObjCol

        For Each obj As GDIObject In list
            If obj.HitTest(boundrect) Then
                tempCol.Add(obj)
            End If

        Next

        Return tempCol
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Hit tests all items in the object collection against a specific point, returning 
    ''' the first hit object, which is the topmost in the zorder.
    ''' </summary>
    ''' <param name="pt">The point to test against </param>
    ''' <param name="fscale">The scale of the drawing surface.</param>
    ''' <returns>The topmost object hit, or Nothing (null) if no objects are hit.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function FindObjectAtPoint(ByVal pt As PointF, ByVal fscale As Single) As GDIObject

        For i As Int32 = Me.InnerList.Count - 1 To 0 Step -1
            Dim gobj As GDIObject = Me.Item(i)

            'return first successful hit test
            If gobj.HitTest(pt, fscale) Then
                Return gobj
            End If
        Next


        'no hit item found
        Return Nothing

    End Function

#End Region

#Region "Zorder related code"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sends a subset of this collection's items to the back (behind other items in the set).
    ''' </summary>
    ''' <param name="gdiObjCol">The subset of items to send backward</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub sendBackward(ByVal gdiObjCol As GDIObjCol)
        Trace.WriteLineIf(Session.Tracing, "sendBackward")

        For i As Int32 = 0 To gdiObjCol.Count - 1
            Dim curObj As GDIObject = gdiObjCol(i)

            If list.Contains(curObj) Then
                Dim iPos As Int32 = list.IndexOf(curObj)
                If iPos > 0 Then
                    list.RemoveAt(iPos)

                    list.Insert(iPos - 1, curObj)
                End If
            End If
        Next

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Brings a subset of the collections items forward (before other items in the set)
    ''' </summary>
    ''' <param name="gdiObjCol">The collection of items to bring forward.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub bringForward(ByVal gdiObjCol As GDIObjCol)
        Trace.WriteLineIf(Session.Tracing, "bringForward")

        For i As Int32 = gdiObjCol.Count - 1 To 0 Step -1
            Dim curObj As GDIObject = gdiObjCol(i)

            If list.Contains(curObj) Then
                Dim iPos As Int32 = list.IndexOf(curObj)
                If iPos < list.Count - 1 Then
                    list.RemoveAt(iPos)

                    list.Insert(iPos + 1, curObj)
                End If
            End If
        Next

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sends a set of selected items to the back (behind all other items in the set)
    ''' </summary>
    ''' <param name="gdiObjCol">The set of items to send to the back.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub sendToback(ByVal gdiObjCol As GDIObjCol)
        Trace.WriteLineIf(Session.Tracing, "sendToback")

        For Each obj As GDIObject In gdiObjCol
            List.Remove(obj)
            list.Insert(0, obj)
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Brings a set of selected items to the front (before all other items in the set)
    ''' </summary>
    ''' <param name="gdiObjCol">The set of items to send to the back.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub bringToFront(ByVal gdiObjCol As GDIObjCol)
        Trace.WriteLineIf(Session.Tracing, "bringToFront")

        For Each gdiobject As GDIObject In gdiObjCol
            List.Remove(gdiobject)
            list.Insert(list.Count, gdiobject)
        Next
    End Sub

#End Region


#Region "Miscellaneous members"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deserializes a set of items.
    ''' </summary>
    ''' <param name="doc">The parent document this set belongs to.</param>
    ''' <param name="pg">The page this set belongs to.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub deserialize(ByVal doc As GDIDocument, ByVal pg As GDIPage)
        Trace.WriteLineIf(Session.Tracing, "deserialize, GDIObjCol")

        For Each obj As GDIObject In Me.InnerList
            obj.deserialize(doc, pg)
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a data format provider used for clipboard operations on GDIObjCol objects.
    ''' </summary>
    ''' <value>A Format consumable by clip board operations to identify the GDIObjCol.</value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property Format() As System.Windows.Forms.DataFormats.Format
        Get
            Return _DataFormat
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a Boolean indicating if there are any items in this particular set.
    ''' </summary>
    ''' <returns>A Boolean indicating if any items exist in the set.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function HasItems() As Boolean
        Return Not InnerList Is Nothing AndAlso InnerList.Count > 0
    End Function

#End Region



#Region "Collection implementation"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets up the indexer on the GDIObjCol to return GDIObjects instead of objects.
    ''' </summary>
    ''' <param name="index">Positional index of the object to retrieve.</param>
    ''' <value>Returns the GDIObject at the position of index</value>
    ''' -----------------------------------------------------------------------------
    Default Public Property Item(ByVal index As Integer) As GDIObject
        Get
            Return DirectCast(innerlist(index), GDIObject)
        End Get
        Set(ByVal Value As GDIObject)
            innerlist(index) = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a GDIObject to the collection
    ''' </summary>
    ''' <param name="value">The object to add</param>
    ''' -----------------------------------------------------------------------------
    Public Sub Add(ByVal value As GDIObject)
        innerList.Add(value)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a range of objects to the gdiobjcol
    ''' </summary>
    ''' <param name="value">An array of GDIObjects to add</param>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub AddRange(ByVal value As GDIObject())
        For i As Int32 = 0 To value.Length - 1
            Me.Add(value(i))
        Next

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a range of objects to the collection from another GDIObjCol
    ''' </summary>
    ''' <param name="value">The GDIObjCol to extract objects from</param>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub AddRange(ByVal value As GDIObjCol)
        For i As Int32 = 0 To value.Count - 1
            Me.Add(value(i))
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a Boolean indicating if a particular GDIObject exists in this collection 
    ''' instance.
    ''' </summary>
    ''' <param name="value">The GDIObject to test for existence in the collection</param>
    ''' <returns>A Boolean indicating if the GDIObject exists in the collection or not.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function Contains(ByVal value As GDIObject) As Boolean
        Return innerlist.Contains(value)
    End Function




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Removes a specific GDIObject from the underlying collection
    ''' </summary>
    ''' <param name="value">A GDIObject to remove from the list</param>
    ''' -----------------------------------------------------------------------------
    Public Sub Remove(ByVal value As GDIObject)
        If innerlist.Contains(value) Then
            innerlist.Remove(value)
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Removes a set of objects from the collection
    ''' </summary>
    ''' <param name="gdiObjects">The objects within another GDIObjCol 
    ''' to remove from the collection</param>
    ''' -----------------------------------------------------------------------------
    Public Sub Remove(ByVal gdiObjects As GDIObjCol)
        For i As Int32 = gdiObjects.Count - 1 To 0 Step -1
            innerlist.Remove(gdiObjects.Item(i))
        Next

    End Sub

#End Region

End Class
