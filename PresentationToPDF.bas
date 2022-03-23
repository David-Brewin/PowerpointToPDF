Attribute VB_Name = "PresentationToPDF"
' Converts presentation into a PDF file, with a separate page for each animation (mouse-click)
' Author: David Brewin (dtbrewin@icloud.com)
' Based on code published by Neil Mitchell (see http://neilmitchell.blogspot.com/2007/11/powerpoint-pdf-part-2.html)
Option Explicit
Public Enum SlideNumberingOption
    NoNumbers
    BottomLeft
    BottomCentre
    BottomRight
    TopLeft
    TopCentre
    TopRight
    CancelMacro
End Enum
Function DebugMode() As Boolean
    DebugMode = False
End Function
Sub PrintToPDF()
    Dim response As Integer
    Dim numberingOption As SlideNumberingOption
    
    If ActivePresentation.Slides.Count = 0 Then
        response = MsgBox(Prompt:="No slides detected in " & ActivePresentation.Name, Title:="Presentation to PDF")
        Exit Sub
    End If
    
    If ActivePresentation.Tags("ptpInProgress") = "True" Then
        Call PrintAndRestore(oPresentation:=ActivePresentation, RestoreOnly:=True)
        Exit Sub
    End If
    
    numberingOption = GetNumberingOption()
    If numberingOption = CancelMacro Then
        Exit Sub
    End If
    
    ActivePresentation.Tags.Add "ptpInProgress", "True"
    
    Call CreateTemporarySlides(oPresentation:=ActivePresentation, numberingOption:=numberingOption)
    Call PrintAndRestore(oPresentation:=ActivePresentation, RestoreOnly:=False)

End Sub
' https://excelmacromastery.com/
Function Exists(coll As Collection, ByVal key As String) As Boolean
    
    On Error GoTo EH
    IsObject (coll.Item(key))
    Exists = True
EH:

End Function
Function GetAnimationsByShape(oSlide As Slide) As Collection
    Dim oShapeAnimationsCollection As New Collection
    Dim oEffect As Effect
    Dim animationNumber As Integer
    
    animationNumber = 0
    For Each oEffect In oSlide.TimeLine.MainSequence
        If DebugMode() Then oEffect.Shape.Select
        If oEffect.Timing.TriggerType = msoAnimTriggerOnPageClick _
                Or oEffect.Timing.TriggerType = msoAnimTriggerAfterPrevious _
                Or oEffect.Timing.TriggerType = msoAnimTriggerWithPrevious Then
                
            If oEffect.Timing.TriggerType = msoAnimTriggerOnPageClick Then
                animationNumber = animationNumber + 1
            End If
            
            Call AddEffectToCollection(oEffect, animationNumber, oShapeAnimationsCollection)
        End If
    Next oEffect
    
    Set GetAnimationsByShape = oShapeAnimationsCollection

End Function
Function GetBestContrast(colorToContrastWith As Long) As Long
    Dim blue As Long
    Dim green As Long
    Dim red As Long
    
    red = colorToContrastWith Mod 256
    green = Int(colorToContrastWith / 256) Mod 256
    blue = Int(colorToContrastWith / 65536)
    
    If (red + green + blue) < 385 Then
    ' textColor is closer to black than white, choose white
    GetBestContrast = RGB(255, 255, 255)
    Else
        ' textColor is closer to white than black, choose black
        GetBestContrast = RGB(0, 0, 0)
    End If
End Function
Function GetEffectParagraph(oEffect As Effect) As Integer

    ' Attempting to access Effect.Paragraph when Effect applies to shape as a
    ' whole throws an error. There must be a better way of coding this!
    GetEffectParagraph = 0
    On Error Resume Next
    GetEffectParagraph = oEffect.Paragraph
    On Error GoTo 0

End Function
Function GetNumberingOption() As SlideNumberingOption
    Dim response As String
    Do
        response = InputBox(Prompt:="Type one of " & vbCrLf _
            & "    'TL' to print at Top Left" & vbCrLf _
            & "    'TC' to print at Top Centre" & vbCrLf _
            & "    'TR' to print at Top Right" & vbCrLf _
            & "    'BR' to print at Bottom Right" & vbCrLf _
            & "    'BC' to print at Bottom Centre" & vbCrLf _
            & "    'BL' to print at Bottom Left" & vbCrLf _
            & "or type 'N' to omit slide numbers", _
            Title:="Slide Numbering")
        response = UCase(response)
        Select Case response
            Case "BL"
                GetNumberingOption = BottomLeft: Exit Do
            Case "BC"
                GetNumberingOption = BottomCentre: Exit Do
            Case "BR"
                GetNumberingOption = BottomRight: Exit Do
            Case "TL"
                GetNumberingOption = TopLeft: Exit Do
            Case "TC"
                GetNumberingOption = TopCentre: Exit Do
            Case "TR"
                GetNumberingOption = TopRight: Exit Do
            Case "N"
                GetNumberingOption = NoNumbers: Exit Do
            Case ""
                GetNumberingOption = CancelMacro: Exit Do
            Case Else
                response = MsgBox(Prompt:="'" & response & "' is not a valid option", Buttons:=vbOK + vbCritical)
        End Select
    Loop
    
End Function
Function IsItemVisible(animationNumber As Integer, oAnimationCollection As Collection) As Boolean
    Dim oAnimation As Animation
    
    IsItemVisible = (oAnimationCollection.Count = 0) ' Default value
    For Each oAnimation In oAnimationCollection
        If animationNumber = 0 Then
            ' if first animation makes shape visible, then it must be invisible at start and vice versa
            IsItemVisible = Not (oAnimation.Effect.Exit = msoFalse)
        End If
        
        If oAnimation.SlideNumber > animationNumber Then
            Exit For
        End If
        
        IsItemVisible = (oAnimation.Effect.Exit = msoFalse)
    Next oAnimation
    
End Function
Sub AddEffectToCollection( _
        oEffect As Effect, _
        animationNumber As Integer, _
        oShapeAnimationsCollection As Collection)
    
    Dim oAnimation As Animation
    Dim oShapeAnimations As ShapeAnimations
    Dim paragraphNumber As Integer
    Dim oShapeTextRange As TextRange2
    Dim oParagraphAnimations As ParagraphAnimations
    Dim cKey As String

    Set oAnimation = New Animation
    oAnimation.SlideNumber = animationNumber
    Set oAnimation.Effect = oEffect

    cKey = CStr(oEffect.Shape.Id)
    If Exists(oShapeAnimationsCollection, cKey) Then
        Set oShapeAnimations = oShapeAnimationsCollection(cKey)
    Else
        Set oShapeAnimations = New ShapeAnimations
        Set oShapeAnimations.AnimationCollection = New Collection
        Set oShapeAnimations.ParagraphAnimationsCollection = New Collection
        oShapeAnimationsCollection.Add Item:=oShapeAnimations, key:=cKey
    End If
        
    paragraphNumber = GetEffectParagraph(oEffect)
    
    If paragraphNumber = 0 Then
        ' Effect applies to shape as a whole
        oShapeAnimations.AnimationCollection.Add Item:=oAnimation
    Else
        If DebugMode() Then
            Set oShapeTextRange = oEffect.Shape.TextFrame2.TextRange
            oShapeTextRange.Paragraphs(paragraphNumber).Select
        End If
        
        cKey = CStr(paragraphNumber)
        If Exists(oShapeAnimations.ParagraphAnimationsCollection, cKey) Then
            Set oParagraphAnimations = oShapeAnimations.ParagraphAnimationsCollection(cKey)
        Else
            Set oParagraphAnimations = New ParagraphAnimations
            oParagraphAnimations.Paragraph = paragraphNumber
            Set oParagraphAnimations.AnimationCollection = New Collection
            oShapeAnimations.ParagraphAnimationsCollection.Add Item:=oParagraphAnimations, key:=cKey
        End If

        oParagraphAnimations.AnimationCollection.Add Item:=oAnimation
    End If

End Sub
Sub AddSlideNumber( _
        newSlide As Slide, _
        originalSlideNumber As Integer, _
        animationNumber As Integer, _
        numberingOption As SlideNumberingOption)
    
    Dim oNewShape As Shape
    Dim slideBackgroundRGB As Long
    Dim margin As Long
    Dim boxWidth As Long, boxHeight As Long
    Dim slideWidth As Long, slideHeight As Long
    Dim ppAlignment As PpParagraphAlignment
    Dim xPos As Long, yPos As Long
    
    If numberingOption = NoNumbers Then
        Exit Sub
    End If
    
    margin = 5
    boxWidth = 80
    boxHeight = 40
    slideWidth = ActivePresentation.PageSetup.slideWidth
    slideHeight = ActivePresentation.PageSetup.slideHeight
    Select Case numberingOption
        Case BottomLeft
            ppAlignment = ppAlignLeft: xPos = margin: yPos = (slideHeight - boxHeight - margin)
        Case BottomCentre
            ppAlignment = ppAlignCenter: xPos = (slideWidth - boxWidth) / 2: yPos = (slideHeight - boxHeight - margin)
        Case BottomRight
            ppAlignment = ppAlignRight: xPos = (slideWidth - boxWidth - margin): yPos = (slideHeight - boxHeight - margin)
        Case TopLeft
            ppAlignment = ppAlignLeft: xPos = margin: yPos = margin
        Case TopCentre
            ppAlignment = ppAlignCenter: xPos = (slideWidth - boxWidth) / 2: yPos = margin
        Case TopRight
            ppAlignment = ppAlignRight: xPos = (slideWidth - boxWidth - margin): yPos = margin
    End Select
    
    Set oNewShape = newSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, xPos, yPos, boxWidth, boxHeight)
    slideBackgroundRGB = newSlide.Background.Fill.ForeColor.RGB
    
    With oNewShape.TextFrame2
        .AutoSize = msoAutoSizeTextToFitShape
        .TextRange.ParagraphFormat.Alignment = ppAlignment
        .TextRange.Font.Name = "Calibri"
        .TextRange.Font.Italic = msoTrue
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = GetBestContrast(slideBackgroundRGB)
    End With

    oNewShape.TextFrame2.TextRange.InsertAfter "" & originalSlideNumber & "-" & animationNumber
End Sub
Sub CamouflageUnanimatedParagraphs(oShape As Shape, animationNumber As Integer, oShapeAnimations As ShapeAnimations)
    Dim oShapeTextRange As TextRange2
    Dim paragraphIndex As Long
    Dim oParagraph As TextRange2
    Dim oParagraphAnimations As ParagraphAnimations
    Dim oRun As TextRange2
    
    Set oShapeTextRange = oShape.TextFrame2.TextRange
    
    For paragraphIndex = 1 To oShapeTextRange.Paragraphs.Count
        Set oParagraph = oShapeTextRange.Paragraphs(paragraphIndex)
        If DebugMode() Then oParagraph.Select
        If Exists(oShapeAnimations.ParagraphAnimationsCollection, CStr(paragraphIndex)) Then
            Set oParagraphAnimations = oShapeAnimations.ParagraphAnimationsCollection(CStr(paragraphIndex))
            
            If Not IsItemVisible(animationNumber, oParagraphAnimations.AnimationCollection) Then
                oParagraph.Font.Fill.ForeColor.RGB = oShape.Fill.ForeColor.RGB
                oParagraph.Font.Fill.BackColor.RGB = oShape.Fill.ForeColor.RGB
                
                For Each oRun In oParagraph.Runs
                    If oRun.Font.Highlight.ObjectThemeColor = msoNotThemeColor Then
                        oRun.Font.Highlight = oShape.Fill.ForeColor
                    End If
                Next oRun
            End If
            
        End If
    Next paragraphIndex
                                        
End Sub
Sub CreateNewSlide( _
        originalSlide As Slide, animationNumber As Integer, _
        oShapeAnimationsCollection As Collection, _
        numberingOption As SlideNumberingOption)
    Dim newSlide As Slide
    Dim shapeIndex As Long
    Dim oShape As Shape
    Dim cKey As String

    Set newSlide = originalSlide.Duplicate(1)

    With newSlide
        .Name = "AutoGenerated: " & newSlide.slideId
        .Tags.Add "ptpSlideType", "Generated"
        .SlideShowTransition.Hidden = msoFalse
        .MoveTo ActivePresentation.Slides.Count
        If DebugMode() Then .Select

        ' Get rid of all animation effects before deletion
        While .TimeLine.MainSequence.Count > 0
            .TimeLine.MainSequence.Item(1).Delete
        Wend
    End With
    
    ' Perform shape deletions (descending order to avoid issues with reindexing)
    For shapeIndex = newSlide.Shapes.Count To 1 Step -1
        Set oShape = newSlide.Shapes(shapeIndex)
        If DebugMode() Then
            On Error Resume Next
            oShape.Select
            On Error GoTo 0
        End If
        cKey = CStr(oShape.Id)
        If Exists(oShapeAnimationsCollection, CStr(oShape.Id)) Then
            Call DeleteUnanimatedItems(oShape, animationNumber, oShapeAnimationsCollection(cKey))
        End If
        
    Next shapeIndex
    
    Call AddSlideNumber(newSlide, originalSlide.SlideNumber, animationNumber, numberingOption)

End Sub
Sub CreateTemporarySlides(oPresentation As Presentation, numberingOption As SlideNumberingOption)
    Dim oSlide As Slide
    Dim originalSlideCollection As New Collection
    Dim oShapeAnimationsCollection As Collection
    Dim animationNumber As Integer
    Dim oEffect As Effect
    
    For Each oSlide In oPresentation.Slides
        With oSlide
            If .SlideShowTransition.Hidden = msoTrue Then
                .Tags.Add "ptpSlideType", "Hidden"
            Else
                .Tags.Add "ptpSlideType", "Visible"
            End If
            .SlideShowTransition.Hidden = msoTrue
        End With
        
        originalSlideCollection.Add oSlide
    Next oSlide
    
    For Each oSlide In originalSlideCollection
        If DebugMode() Then oSlide.Select
        Set oShapeAnimationsCollection = GetAnimationsByShape(oSlide)
    
        animationNumber = 0
        ' Create slide for initial slide state
        Call CreateNewSlide(oSlide, animationNumber, oShapeAnimationsCollection, numberingOption)
        
        For Each oEffect In oSlide.TimeLine.MainSequence
    
            If oEffect.Timing.TriggerType = msoAnimTriggerOnPageClick Then
                animationNumber = animationNumber + 1
                ' Create slide for this animation step
                Call CreateNewSlide(oSlide, animationNumber, oShapeAnimationsCollection, numberingOption)
            End If
    
        Next oEffect
    Next oSlide

End Sub
Sub DeletePlaceholder(oShape As Shape)
    ' When a placeholder with text is deleted, an "empty" placeholder, e.g.
    ' "Click to add subtitle" is added to the end of the Shapes collection.
    ' To mimic the animation, both have to be deleted.
    Dim deleteEmpty As Boolean
    Dim oShapes As Shapes
    
    If oShape.HasTextFrame Then
        If oShape.TextFrame2.HasText Then
            Set oShapes = oShape.Parent.Shapes
            deleteEmpty = True
        End If
    End If
    
    oShape.Delete
    If deleteEmpty Then
        oShapes(oShapes.Count).Delete
    End If
End Sub
Sub DeleteUnanimatedItems(oShape As Shape, animationNumber As Integer, oShapeAnimations As ShapeAnimations)
    
    If IsItemVisible(animationNumber, oShapeAnimations.AnimationCollection) Then
        If oShapeAnimations.ParagraphAnimationsCollection.Count > 0 Then
            Call CamouflageUnanimatedParagraphs(oShape, animationNumber, oShapeAnimations)
        End If
    Else
        If oShape.Type = msoPlaceholder Then
            Call DeletePlaceholder(oShape)
        Else
            oShape.Delete
        End If
    End If
    
End Sub
Sub PrintAndRestore(oPresentation As Presentation, RestoreOnly As Boolean)
    Dim appendString As String
    Dim response As Integer
    
    If RestoreOnly Then
        appendString = "?"
    Else
        appendString = " after printing?"
    End If
    
    response = MsgBox(Prompt:="Restore " & oPresentation.Name & " to normal state" & appendString, _
                 Buttons:=vbYesNo + vbQuestion + vbDefaultButton2, _
                 Title:="Presentation to PDF")
    
    If Not RestoreOnly Then
        Call PrintTemporarySlides(oPresentation)
    End If
    
    If response = vbNo Then
        Exit Sub
    End If
    
    Call RestorePresentation(oPresentation)
    ActivePresentation.Tags.Delete ("ptpInProgress")
End Sub
Sub PrintTemporarySlides(oPresentation As Presentation)
    Dim response As Integer

    On Error Resume Next
    With oPresentation
        .PrintOptions.PrintHiddenSlides = False
        .PrintOptions.ActivePrinter = "Microsoft Print to PDF"
        .PrintOut
    End With
    
    If Err <> 0 Then
        On Error GoTo 0
        response = MsgBox(Prompt:="Printing to '" _
            & oPresentation.PrintOptions.ActivePrinter _
            & "'failed with error " + CStr(Err), _
            Title:="Presentation to PDF")
        Exit Sub
    End If
    On Error GoTo 0

End Sub
Sub RestorePresentation(oPresentation As Presentation)
    Dim oSlidesCollection As Slides
    Dim oSlide As Slide
    Dim ptpSlideType As String
    
    Set oSlidesCollection = oPresentation.Slides
    Do
        Set oSlide = oSlidesCollection(oSlidesCollection.Count)
        ptpSlideType = oSlide.Tags("ptpSlideType")
        If ptpSlideType = "Generated" Then
            oSlide.Delete
        Else
            Exit Do
        End If
    Loop While oSlidesCollection.Count > 0
    
    For Each oSlide In oSlidesCollection
        ptpSlideType = oSlide.Tags("ptpSlideType")
        If ptpSlideType = "Hidden" Then
            oSlide.SlideShowTransition.Hidden = msoTrue
        Else
            oSlide.SlideShowTransition.Hidden = msoFalse
        End If
        oSlide.Tags.Delete ("ptpSlideType")
    Next oSlide

End Sub
