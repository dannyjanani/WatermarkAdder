This program takes a directory of DWG files, creates a file called params.out with the names of all the files. It also requests for a drawing containing a watermark to add as a block to a directory of drawings. The path to the watermark drawing is stored in a file called watermark.out. Once it has the information, it opens a drawing in AutoCAD which has VBA code embedded into it. The VBA code then splits the params.out file using a delimeter and uses the data appropriately.


CODE FROM VBA EMBEDDED DRAWING "WatermarkAdder.dwg":
_________________________________________________

' Declares a few constants such as the parameter file, abort execution files, and the delimeterConst paramPath As String = "params.out"
Const waterMarkPath As String = "watermark.out"
Const disableExecutionPath As String = "disable.ex"
Const abortPath As String = "abort.ex"
Const delimeter As String = "[%]"
______________________________________________________________________________________________________________________________________________________
' This sets up all the files in corresponding variables and calls the addWatermark function to do the conversion.
' We start off by reading the watermark.out file and open the watermark drawing. We then call DetermineWmCoordinates function which
' tells us the size of the watermark or basically any drawing in modelspace.
' Once the size is determined, we now know how much to scale the watermark onto the desired drawing and we close the watermark.
' We then read the params.out file and while the file still has data and the abort.ex file is not in the directory,
' we make sure the dataline is not an empty string and then put the file path into a variable called directPath.
' We update watermark in the loop every time since the AutoCAD InsertBlock function changes the path for some reason.
' We open the drawing and call the addWatermark function which does all the heavy work.
' After it completes all drawings it closes the file and goes to next.
' NOTE: You will get an error if you try to run this 2 times in a row. In order to avoid you must completely quit AutoCAD or open a new instance of it like the DWGtoPDFConverter.
' What you can do is change ThisDrawing.Close to ThisDrawing.SendCommand "_QUIT " below where it says "QUIT AUTOCAD HERE"
Public Sub doSomething(ByVal currentDirPath As String)
    Dim FileNum As Integer
    Dim DataLine As String
    Dim originalWatermark As String
    Dim dwg As AcadDocument
    Dim errorMsg As String: errorMsg = ""
    
    FileNum = FreeFile()
    Open (currentDirPath & "\" & waterMarkPath) For Input As #FileNum
    Line Input #FileNum, DataLine
    originalWatermark = DataLine
    Close #FileNum
        
    Set dwg = Application.Documents.Open(originalWatermark)
    
    Dim minX As Double
    Dim minY As Double
    Dim maxX As Double
    Dim maxY As Double
    
    DetermineWmCoordinates dwg, minX, maxX, minY, maxY
    
    Dim waterMarkHeight As Double
    Dim waterMarkWidth As Double
    
    waterMarkHeight = maxY - minY
    waterMarkWidth = maxX - minX
    'MsgBox "Width: " & waterMarkWidth & ", Height: " & waterMarkHeight
    dwg.Save
    dwg.Close
    Open (currentDirPath & "\" & paramPath) For Input As #FileNum
    Do While Not EOF(FileNum)
        On Error GoTo ReportError
        If Dir(curDir & "\" & abortPath) = "" Then
            Line Input #FileNum, DataLine
            If DataLine <> "" Then
                'Dim splitParam() As String: splitParam = Split(DataLine, delimeter)
                Dim directPath As String: directPath = DataLine
                Dim watermark As String: watermark = originalWatermark
                Dim backPlot As Integer: backPlot = CInt(ThisDrawing.GetVariable("BACKGROUNDPLOT"))
                ThisDrawing.SetVariable "BACKGROUNDPLOT", 0
                Set dwg = Application.Documents.Open(directPath)
                Application.ActiveDocument = dwg
                addWatermark watermark, dwg, waterMarkHeight, waterMarkWidth
                dwg.Save
                dwg.Close False
                ThisDrawing.SetVariable "BACKGROUNDPLOT", backPlot
            Else
                MsgBox ("The DWG is not found!")
            End If
        Else
            MsgBox ("Project successfully aborted")
            DeleteFile abortPath
            ThisDrawing.Close
            Exit Sub
        End If
NextDrawing:
    Loop
    Close #FileNum
    MsgBox ("Execution Succesfully Completed!")
    ThisDrawing.Save
    ThisDrawing.Close
    ' QUIT AUTOCAD HERE
    Exit Sub
ReportError:
    MsgBox Err.Description
    Err.Clear
    GoTo NextDrawing
End Sub
______________________________________________________________________________________________________________________________________________________
' This function makes sure the disable.ex is not in the directory (if you want to edit code create this empty file in that directory)
' If it isn't it will call doSomething function and if it is it will send the vbaide function which opens the code.
Private Sub UserForm_Activate()
    Me.Hide
    Dim curDir As String: curDir = Replace(LCase(ThisDrawing.Path), "watermarkadder.dwg", "")
    If Dir(curDir & "\" & disableExecutionPath) = "" Then
        doSomething (curDir)
    Else
        ThisDrawing.SendCommand "vbaide "
    End If
End Sub
______________________________________________________________________________________________________________________________________________________
' Here is all the heavy work. We first add a layer, then check the coordinates of the drawing we are adding the watermark to.
' Some are in layout and some are in modelspace. Function DetermineCoordinates checks the coordinates in layout
' while function DetermineWmCoordinates checks the coordinates in modelspace. These are passed by reference and updated in the function.
' We then scale the insertion point based on the width and height of the page and insert on both layout and modelspace.
' Finally, we reset the active layer back to what it was and purge all.
Public Function addWatermark(watermark As String, ByRef dwg As AcadDocument, wmheight As Double, wmwidth As Double)
    dwg.PurgeAll
    Dim strLayerName As String
    Dim objLayer As AcadLayer
    Dim insertionPnt(0 To 2) As Double
    Dim myIns As AcadBlockReference
    Dim paperWidth As Double
    Dim paperHeight As Double
    Dim xscale As Double
    Dim yscale As Double
    Dim minX As Double
    Dim minY As Double
    Dim maxX As Double
    Dim maxY As Double
    
    'strLayerName = "Watermark"
    'Set objLayer = dwg.Layers.Add(strLayerName)
    Set objLayer = dwg.ActiveLayer
    'dwg.ActiveLayer = objLayer
    'objLayer.Plottable = True
    'objLayer.color = acWhite

    'dwg.ActiveLayout.GetPaperSize paperWidth, paperHeight
    
    DetermineCoordinates dwg, minX, maxX, minY, maxY
    
    paperWidth = maxX - minX
    paperHeight = maxY - minY
    
    'MsgBox paperWidth & ", " & paperHeight
    xscale = 0.7 * (paperWidth / wmwidth)
    yscale = 0.7 * (paperHeight / wmheight)
    
    insertionPnt(0) = (paperWidth * 0.15): insertionPnt(1) = (paperHeight * 0.15): insertionPnt(2) = 0
    'MsgBox "insertion width: " & insertionPnt(0) & ", insertion height: " & insertionPnt(1)
    Set myIns = dwg.Layouts(0).Block.InsertBlock(insertionPnt, watermark, xscale, yscale, 1, 0)
    
    DetermineWmCoordinates dwg, minX, maxX, minY, maxY
    
    paperWidth = maxX - minX
    paperHeight = maxY - minY
    
    'MsgBox paperWidth & ", " & paperHeight
    xscale = 0.7 * (paperWidth / wmwidth)
    yscale = 0.7 * (paperHeight / wmheight)
    
    insertionPnt(0) = (paperWidth * 0.15) + minX: insertionPnt(1) = (paperHeight * 0.15) + minY: insertionPnt(2) = 0
    Set myIns = dwg.ModelSpace.InsertBlock(insertionPnt, watermark, xscale, yscale, 1, 0)
    
    AutoCAD.Application.ZoomExtents
    dwg.ActiveLayer = objLayer
    dwg.PurgeAll
End Function
______________________________________________________________________________________________________________________________________________________
' Checks if a specific file passed into it exists
Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function
______________________________________________________________________________________________________________________________________________________
' Gets the coordinates of a drawing from the Layout 1 space.
' NOTE: the function of AutoCAD for some reason returns the height then the width for the getPaperSize function
Public Sub DetermineCoordinates(ByRef dwg As AcadDocument, ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double)
    Dim i As Integer: i = 0
    
    Dim pWidth As Double: pWidth = 0
    Dim pHeight As Double: pHeight = 0
    Dim paperWidth As Double: paperWidth = 0
    Dim paperHeight As Double: paperHeight = 0
    Dim temp As Double: temp = 0
    
    dwg.Layouts(0).GetPaperSize pHeight, pWidth
    'If dwg.Layouts(0).PlotRotation = ac90degrees Then
    '    temp = pHeight
    '    pHeight = pWidth
    '    pWidth = temp
    'End If
    
    paperWidth = pWidth / 25.4
    paperHeight = pHeight / 25.4
    
    minX = 0
    maxX = paperWidth
    minY = 0
    maxY = paperHeight
End Sub
______________________________________________________________________________________________________________________________________________________
' Gets the coordinates of a drawing from the modelspace.
' Filter 67 is the one to use to get the objects and bounding box.
Public Sub DetermineWmCoordinates(ByRef dwg As AcadDocument, ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double)
    Dim i As Integer: i = 0
    
    Dim ss As AcadSelectionSet
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim GroupCode As Variant
    Dim DataValue As Variant
    
    FilterType(0) = 67
    FilterData(0) = 0
    
    GroupCode = FilterType
    DataValue = FilterData
    
    Set ss = ThisDrawing.SelectionSets.Add("SS1")
    ss.Select acSelectionSetAll, , , GroupCode, FilterData
    
    Dim pointA As Variant
    Dim pointB As Variant
    
    Dim LLx As Double: LLx = 67
    Dim LLy As Double: LLy = 67
    Dim URx As Double: URx = 67
    Dim URy As Double: URy = 67
    Dim aceCount As Integer: aceCount = 0
    
    For i = 0 To dwg.ModelSpace.Count - 1
        If TypeOf dwg.ModelSpace.Item(i) Is AcadEntity Then
           Dim acE As AcadEntity
            Set acE = dwg.ModelSpace.Item(i)
            acE.GetBoundingBox pointA, pointB
            If LLx = 67 Or pointA(0) < LLx Then
                LLx = pointA(0)
            End If
            If LLy = 67 Or pointA(1) < LLy Then
                LLy = pointA(1)
            End If
            If URx = 67 Or pointB(0) > URx Then
                URx = pointB(0)
            End If
            If URy = 67 Or pointB(1) > URy Then
                URy = pointB(1)
            End If
            aceCount = aceCount + 1
        End If
    Next
    minX = LLx
    maxX = URx
    minY = LLy
    maxY = URy
    
    ss.Delete
End Sub
______________________________________________________________________________________________________________________________________________________
' This function deletes a file passed into it
Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub
______________________________________________________________________________________________________________________________________________________