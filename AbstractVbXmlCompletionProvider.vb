' Written By Mohammad Hamdy Ghanem

Imports System.Collections.Immutable
Imports System.Threading
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.Completion
Imports System.Xml.Schema
' Imports System.Text.Json

Namespace Microsoft.CodeAnalysis.VisualBasic.Completion.Providers


    Friend MustInherit Class AbstractVbXmlCompletionProvider
        Inherits CommonCompletionProvider


        Shared HtmlSchema As Xml.Schema.XmlSchema

        Private Shared ReadOnly SelfClosing As String() = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"}

        Protected Const Quote As String = ChrW(34)
        Protected Const LambdaExpression As String = "vbxmlLambda"
        Protected Const VbXmlRootName As String = "vbxml"

        Protected Const _BeforeCaretText As String = NameOf(_BeforeCaretText)
        Protected Const _AfterCaretText As String = NameOf(_AfterCaretText)
        Protected Const _BeforeCaretTextOnSpace As String = NameOf(_BeforeCaretTextOnSpace)
        Protected Const _AfterCaretTextOnSpace As String = NameOf(_AfterCaretTextOnSpace)
        Protected Const _DeletePreviousStr As String = NameOf(_DeletePreviousStr)

        Private Shared ReadOnly _s_tagMap As New Dictionary(
               Of String, CompletionParts)

        Shared Sub New()
            LoadHtmlSchema()
            PopulateTagMap()
        End Sub

        Private Shared Sub LoadHtmlSchema()

            Dim html5XsdFile = Environment.ExpandEnvironmentVariables("%userprofile%\documents\html_5.xsd")
            Dim commonHTML5TypesXsdFile = Environment.ExpandEnvironmentVariables("%userprofile%\documents\commonHTML5Types.xsd")
            If Not IO.File.Exists(html5XsdFile) Then
                ' I copy the xsd file to another location like My Documents folder. 
                ' In VB projects this code works fine with the xsd in its location in VS installation dir, 
                ' but in this code inside Roslyn running in VS itself, the xsd doesn't work for unkown reason to me
                ' as the shema turns out to be Nothing!

                Try
                    Dim x = Process.GetCurrentProcess().MainModule.FileName
                    x = IO.Path.GetDirectoryName(x) ' \IDE
                    x = IO.Path.GetDirectoryName(x) '\Common7
                    Dim file = IO.Path.Combine(x, "Packages\Schemas\html\html_5.xsd")
                    IO.File.Copy(file, html5XsdFile)
                    file = IO.Path.Combine(x, "Packages\Schemas\html\commonHTML5Types.xsd")
                    IO.File.Copy(file, commonHTML5TypesXsdFile)
                Catch

                End Try
            End If

            Dim schemaSet As New XmlSchemaSet()
            AddHandler schemaSet.ValidationEventHandler, AddressOf ValidationCallback
            Dim ns1 = "http://schemas.microsoft.com/intellisense/html5"
            Dim ns2 = "http://schemas.microsoft.com/Visual-Studio-Intellisense"
            schemaSet.Add(ns1, html5XsdFile)
            schemaSet.Add(ns2, commonHTML5TypesXsdFile)
            schemaSet.Compile()

            Dim schema2 As XmlSchema = Nothing
            For Each schema As XmlSchema In schemaSet.Schemas()
                If schema.TargetNamespace = ns1 Then
                    HtmlSchema = schema
                ElseIf schema.TargetNamespace = ns2 Then
                    schema2 = schema
                End If
            Next

            Dim import As New XmlSchemaImport()
            import.Namespace = ns2
            import.Schema = schema2
            HtmlSchema.Includes.Add(import)
            schemaSet.Reprocess(HtmlSchema)
            schemaSet.Compile()
        End Sub

        Private Shared Sub PopulateTagMap()
            Dim tagName As String = ""
            For Each tag As XmlSchemaElement In HtmlSchema.Elements.Values
                tagName = If(tag.Name, tag.QualifiedName.Name)

                Dim tagClose = If(SelfClosing.Contains(tagName), "/>", $"></{tagName}>")

                ' Auto add required attributes after the tag
                Dim textB4 = ""
                Dim txtAfter = ""
                Dim scmaType = If(tag.ElementSchemaType, tag.SchemaType)
                Dim elm = TryCast(scmaType, XmlSchemaComplexType)
                ' If tagName = "img" Then Stop
                If elm IsNot Nothing Then
                    Dim attrs = GetAttributes(elm.Attributes, OnlyRequiredAttrs:=True)
                    If attrs.Count > 0 Then
                        textB4 = $" {attrs(0).Name}={Quote}"
                        txtAfter = Quote & " " & String.Join($"={Quote}{Quote} ", (attrs.Select(Function(a) a.Name)).ToArray(), 1, attrs.Count - 1)
                        txtAfter = txtAfter.TrimEnd
                        If txtAfter <> Quote Then txtAfter &= $"={Quote}{Quote}"
                    End If
                End If


                _s_tagMap.Add(tagName, New CompletionParts(
                              $"<{tagName}", textB4, txtAfter, tagClose))
            Next

            _s_tagMap.Add(LambdaExpression,
                          New CompletionParts(
                               "= (Function()" & vbCrLf,
                               "Return <",
                               " />" & vbCrLf & "End Function)()" & vbCrLf,
                               "%>")
                          )
        End Sub

        Private Shared Function GetXmlComplexElement(tagName As String) As XmlSchemaComplexType
            Dim n = New Xml.XmlQualifiedName(tagName, HtmlSchema.TargetNamespace)
            Dim tag = CType(HtmlSchema.Elements.Item(n), XmlSchemaElement)
            If tag Is Nothing Then Return Nothing

            Dim scmaType = If(tag.ElementSchemaType, tag.SchemaType)
            Return TryCast(scmaType, XmlSchemaComplexType)
        End Function


        Private Shared Function GetAttribute(
                           attrName As String,
                           attributes As XmlSchemaObjectCollection
                     ) As XmlSchemaAttribute

            For Each child As XmlSchemaAnnotated In attributes
                If TypeOf child Is XmlSchemaAttribute Then
                    Dim attr = CType(child, XmlSchemaAttribute)
                    If attr.Name = attrName Then Return attr
                ElseIf TypeOf child Is XmlSchemaAttributeGroup Then
                    Dim a = GetAttribute(attrName, CType(child, XmlSchemaAttributeGroup).Attributes)
                    If a IsNot Nothing Then Return a
                ElseIf TypeOf child Is XmlSchemaAttributeGroupRef Then
                    Dim ref = CType(child, XmlSchemaAttributeGroupRef)
                    Dim grp = TryCast(HtmlSchema.AttributeGroups.Item(ref.RefName), XmlSchemaAttributeGroup)
                    If grp IsNot Nothing Then
                        Dim a = GetAttribute(attrName, grp.Attributes)
                        If a IsNot Nothing Then Return a
                    End If
                End If
            Next

            Return Nothing
        End Function


        Private Shared Function GetAttributes(
                            tagName As String,
                            Optional OnlyRequiredAttrs As Boolean = False,
                            Optional reqAttrs As List(Of String) = Nothing
                      ) As List(Of AttrInfo)

            Dim elm = GetXmlComplexElement(tagName)

            If elm Is Nothing Then Return New List(Of AttrInfo)

            Return GetAttributes(elm.Attributes, OnlyRequiredAttrs)
        End Function

        Private Shared Function GetAttributes(
                           attributes As XmlSchemaObjectCollection,
                           Optional OnlyRequiredAttrs As Boolean = False,
                           Optional reqAttrs As List(Of AttrInfo) = Nothing
                     ) As List(Of AttrInfo)

            reqAttrs = If(reqAttrs, New List(Of AttrInfo))

            For Each child As XmlSchemaAnnotated In attributes
                If TypeOf child Is XmlSchemaAttribute Then
                    Dim attr = CType(child, XmlSchemaAttribute)

                    If OnlyRequiredAttrs = False OrElse attr.Use = XmlSchemaUse.Required Then
                        Dim EnumValues = TryCast(attr.AttributeSchemaType?.Content, XmlSchemaSimpleTypeRestriction)?.Facets
                        reqAttrs.Add(New AttrInfo(attr.Name, EnumValues IsNot Nothing AndAlso EnumValues.Count > 0))
                    End If

                ElseIf Not OnlyRequiredAttrs Then
                    If TypeOf child Is XmlSchemaAttributeGroup Then
                        GetAttributes(CType(child, XmlSchemaAttributeGroup).Attributes, OnlyRequiredAttrs, reqAttrs)

                    ElseIf TypeOf child Is XmlSchemaAttributeGroupRef Then
                        Dim ref = CType(child, XmlSchemaAttributeGroupRef)
                        Dim grp = TryCast(HtmlSchema.AttributeGroups.Item(ref.RefName), XmlSchemaAttributeGroup)
                        If grp IsNot Nothing Then GetAttributes(grp.Attributes, OnlyRequiredAttrs, reqAttrs)
                    End If
                End If
            Next

            Return reqAttrs
        End Function

        Shared Sub ValidationCallback(ByVal sender As Object, ByVal args As System.Xml.Schema.ValidationEventArgs)
            If args.Severity = System.Xml.Schema.XmlSeverityType.Warning Then
                Console.Write("WARNING: ")
            ElseIf args.Severity = System.Xml.Schema.XmlSeverityType.Error Then
                Console.Write("ERROR: ")
            End If
            Console.WriteLine(args.Message)
        End Sub


        Private Shared ReadOnly _s_listTypeValues As ImmutableArray(Of String) = ImmutableArray.Create(
            "bullet", "number", "table"
            )

        Private ReadOnly _defaultRules As CompletionItemRules

        Protected Sub New(_defaultRules As CompletionItemRules)
            If _defaultRules Is Nothing Then Throw New ArgumentNullException(NameOf(_defaultRules))
            Me._defaultRules = _defaultRules

        End Sub

        Public Overrides Async Function ProvideCompletionsAsync(context As CompletionContext) As Task

            If HtmlSchema Is Nothing Then Return

            ' Dim jsonString = JsonSerializer.Serialize(context)


            Dim items = Await GetItemsWorkerAsync(
                        context.Document,
                        context.Position,
                        context.Trigger,
                        context.CancellationToken
                    ).ConfigureAwait(False)

            If items IsNot Nothing Then
                context.AddItems(items)
            End If
        End Function

        Protected MustOverride Function GetItemsWorkerAsync(_document As Document, _position As Integer, _trigger As CompletionTrigger, _cancellationToken As CancellationToken) As Task(Of IEnumerable(Of CompletionItem))

        Protected Function GetItem(name As String, AddCloseTag As Boolean) As CompletionItem
            Dim value As CompletionParts
            If _s_tagMap.TryGetValue(name, value) Then
                Dim parts = GetAttrAndClosingTag(value, AddCloseTag)
                Return CreateCompletionItem(name, parts.beforeCaretText, parts.afterCaretText, "")
            Else
                Return CreateCompletionItem(name)
            End If
        End Function

        Function GetAttrAndClosingTag(value As CompletionParts, AddCloseTag As Boolean) As (beforeCaretText As String, afterCaretText As String)
            If AddCloseTag Then
                Return (value.tagOpen & value.textBeforeCaret,
                value.textAfterCaret & value.tagClose)
            Else
                Return (value.tagOpen, "")
            End If
        End Function

        Protected Function GetAttributeItems(tagName As String, StartsWith As String, nextChar As String, existingAttributes As ISet(Of String)) As IEnumerable(Of CompletionItem)
            Dim txtB4, txtAfter As String

            If nextChar = "=" Then
                txtB4 = ""
                txtAfter = ""
            ElseIf nextChar = Quote Then
                txtB4 = "="
                txtAfter = ""
            Else
                txtB4 = $"={Quote}"
                txtAfter = Quote
            End If

            Dim AttrItems As New List(Of CompletionItem)
            Dim Attrs = GetAttributes(tagName)
            Dim groupedItems As New List(Of String)

            Dim ShowDataAttr = False

            If StartsWith.StartsWith("on") OrElse StartsWith.StartsWith("On") Then
                StartsWith = StartsWith.Substring(0, 2)
                ShowDataAttr = True
            ElseIf StartsWith.Contains("-") Then
                ShowDataAttr = True
            Else
                StartsWith = ""
            End If

            For Each attr In Attrs
                Dim attrName = attr.Name

                If Not existingAttributes.Contains(attrName) Then

                    If Not ShowDataAttr Then
                        Dim i = attrName.IndexOf("-")
                        If i > 0 Then ' data- & aria-

                            Dim x = attrName.Substring(0, i + 1) & "..."
                            If Not groupedItems.Contains(x) Then
                                groupedItems.Add(x)
                                AttrItems.Add(CreateCompletionItem(x, x.TrimEnd("."c), "", ""))
                            End If
                            Continue For
                        End If

                        If attrName.Length > 2 Then ' onEVENT
                            If attrName.StartsWith("on") Then
                                Dim x = "on..."
                                If Not groupedItems.Contains(x) Then
                                    groupedItems.Add(x)
                                    AttrItems.Add(CreateCompletionItem(x, "on", "", ""))
                                End If
                                Continue For
                            ElseIf attrName.StartsWith("On") Then
                                Dim x = "On..."
                                If Not groupedItems.Contains(x) Then
                                    groupedItems.Add(x)
                                    AttrItems.Add(CreateCompletionItem(x, "On", "", ""))
                                End If
                                Continue For
                            End If
                        End If
                    End If

                    If StartsWith = "" OrElse attrName.StartsWith(StartsWith) Then
                        AttrItems.Add(CreateCompletionItem(
                            attrName,
                            attrName & If(txtB4.Length < 2, txtB4, If(attr.HasEnum, "=", txtB4)),
                            If(attr.HasEnum, "", txtAfter),
                            If(StartsWith.EndsWith("-"), StartsWith, "")
                       ))
                    End If
                End If
            Next

            Return AttrItems
        End Function

        Private Function GetCommentItem() As CompletionItem
            Const prefix As String = "!--"
            Const suffix As String = "-->"
            Return CreateCompletionItem(prefix, "<" & prefix, suffix, "")
        End Function

        Private Function GetCDataItem() As CompletionItem
            Const prefix As String = "![CDATA["
            Const suffix As String = "]]>"
            Return CreateCompletionItem(prefix, "<" & prefix, suffix, "")
        End Function

        Protected Function GetAttributeValueItems(tagName As String, attributeName As String, addOpeningQuote As Boolean, addClosingQuote As Boolean, deletePreviousStr As String) As IEnumerable(Of CompletionItem)
            Dim values As New List(Of CompletionItem)
            Dim elm = GetXmlComplexElement(tagName)
            Dim attr = GetAttribute(attributeName, elm.Attributes)
            If attr Is Nothing Then Return values

            Dim EnumValues = CType(attr.AttributeSchemaType?.Content, XmlSchemaSimpleTypeRestriction)?.Facets
            If EnumValues Is Nothing Then Return values

            For i = 0 To EnumValues.Count - 1
                Dim value = CType(EnumValues(i), XmlSchemaEnumerationFacet).Value
                values.Add(CreateCompletionItem(
                               value,
                               If(addOpeningQuote, Quote, "") & value & If(addClosingQuote, Quote, ""),
                               "",
                              deletePreviousStr)
                         )
            Next

            Return values
        End Function

        Protected Function GetChildItems(tagName As String, AddCloseTag As Boolean) As IEnumerable(Of CompletionItem)
            Dim tags As New List(Of CompletionItem)

            If tagName = VbXmlRootName Then
                ' All elements are allowed
                For Each entry In _s_tagMap

                    Dim parts = GetAttrAndClosingTag(entry.Value, AddCloseTag)

                    tags.Add(CreateCompletionItem(
                        entry.Key,
                        parts.beforeCaretText,
                        parts.afterCaretText,
                        "")
                    )
                Next
                Return tags
            End If

            Dim n = New Xml.XmlQualifiedName(tagName, HtmlSchema.TargetNamespace)
            Dim tag = CType(HtmlSchema.Elements.Item(n), XmlSchemaElement)
            If tag Is Nothing Then Return tags


            Dim scmaType = If(tag.SchemaType, tag.ElementSchemaType)
            If scmaType Is Nothing OrElse TypeOf scmaType Is XmlSchemaSimpleType _
                  Then Return tags

            Dim particle = CType(scmaType, XmlSchemaComplexType).ContentTypeParticle
            If particle Is Nothing Then Return tags

            Dim group = TryCast(particle, XmlSchemaGroupBase)
            If group Is Nothing Then Return tags

            For Each child In group.Items
                Dim elm = TryCast(child, XmlSchemaElement)
                If elm IsNot Nothing Then tags.Add(GetItem(If(elm.Name, elm.QualifiedName.Name), AddCloseTag))
            Next

            Return tags
        End Function

        Public Overrides Async Function GetChangeAsync(
                          _document As Document,
                          item As CompletionItem,
                          Optional commitChar As Char? = Nothing,
                          Optional _cancellationToken As CancellationToken = Nothing
                    ) As Task(Of CompletionChange)

            Dim beforeCaretText = "", afterCaretText = ""
            item.Properties.TryGetValue(_BeforeCaretText, beforeCaretText)
            item.Properties.TryGetValue(_AfterCaretText, afterCaretText)

            Dim SpaceCond = item.Properties.ContainsKey(_BeforeCaretTextOnSpace) AndAlso
                item.Properties.ContainsKey(_AfterCaretTextOnSpace) AndAlso
                (beforeCaretText <> "" OrElse afterCaretText <> "")

            Dim includesCommitCharacter = Not (commitChar.HasValue AndAlso commitChar.Value = " "c AndAlso SpaceCond)


            Dim text = Await _document.GetTextAsync(_cancellationToken).ConfigureAwait(False)

            Dim itemSpan = item.Span
            Dim curPos = itemSpan.Start
            Dim prevPos = curPos - 1
            Dim prevChar = If(prevPos > -1 AndAlso text.Length > 0, text(prevPos), "")

            Dim delCount = 0
            Dim deletePreviousStr = ""
            item.Properties.TryGetValue(_DeletePreviousStr, deletePreviousStr)

            If prevChar = "<"c AndAlso beforeCaretText <> "" AndAlso
                          beforeCaretText(0) = "<"c Then
                delCount = 1
            ElseIf deletePreviousStr <> "" Then
                Dim x = text.GetSubText(New TextSpan(0, curPos)).ToString()
                If deletePreviousStr = " " Then ' Delete all previous spaces not just one
                    delCount = x.Length - x.TrimEnd().Length
                ElseIf x.EndsWith(deletePreviousStr) Then
                    delCount = deletePreviousStr.Length
                ElseIf x.TrimEnd().EndsWith(deletePreviousStr) Then
                    delCount = x.Length - x.LastIndexOf(deletePreviousStr)
                End If
            End If


            Dim replacementSpan = TextSpan.FromBounds(
                    curPos - delCount,
                    itemSpan.End
                )

            Dim replacementText = beforeCaretText
            Dim newPosition = replacementSpan.Start + beforeCaretText.Length

            If commitChar.HasValue AndAlso Not Char.IsWhiteSpace(commitChar.Value) AndAlso commitChar.Value <> replacementText(replacementText.Length - 1) Then
                ' include the commit character
                replacementText += commitChar.Value

                ' The caret goes after whatever commit character we spit.
                newPosition += 1
            End If

            replacementText += afterCaretText

            Return CompletionChange.Create(
                    New TextChange(replacementSpan, replacementText),
                    newPosition, includesCommitCharacter)
        End Function

        Private Function CreateCompletionItem(displayText As String) As CompletionItem
            Return CreateCompletionItem(
                   displayText:=displayText,
                   beforeCaretText:=displayText,
                   afterCaretText:=String.Empty,
                   "")
        End Function

        Protected Function CreateCompletionItem(
                                               displayText As String,
                                               beforeCaretText As String,
                                               afterCaretText As String,
                                               deletePreviousStr As String,
                                               Optional beforeCaretTextOnSpace As String = Nothing,
                                               Optional afterCaretTextOnSpace As String = Nothing
                                        ) As CompletionItem

            Dim props = ImmutableDictionary(Of String, String).Empty.
            Add(_BeforeCaretText, beforeCaretText).
            Add(_AfterCaretText, afterCaretText).
            Add(_BeforeCaretTextOnSpace, beforeCaretTextOnSpace).
            Add(_AfterCaretTextOnSpace, afterCaretTextOnSpace).
            Add(_DeletePreviousStr, deletePreviousStr)

            Return CommonCompletionItem.Create(
                displayText:=displayText,
                displayTextSuffix:="",
                glyph:=Glyph.Keyword,
                properties:=props,
                rules:=GetCompletionItemRules(displayText))
        End Function

        Private Shared ReadOnly WithoutQuoteRule As CharacterSetModificationRule = CharacterSetModificationRule.Create(CharacterSetModificationKind.Remove, """"c)
        Private Shared ReadOnly WithoutSpaceRule As CharacterSetModificationRule = CharacterSetModificationRule.Create(CharacterSetModificationKind.Remove, " "c)

        Friend Shared ReadOnly FilterRules As ImmutableArray(Of CharacterSetModificationRule) = ImmutableArray.Create(
        CharacterSetModificationRule.Create(CharacterSetModificationKind.Add, "!"c, "-"c, "["c))

        Private Function GetCompletionItemRules(displayText As String) As CompletionItemRules
            Dim commitRules = _defaultRules.CommitCharacterRules

            If displayText.Contains("""") Then
                commitRules = commitRules.Add(WithoutQuoteRule)
            End If

            If displayText.Contains(" ") Then
                commitRules = commitRules.Add(WithoutSpaceRule)
            End If

            Return _defaultRules.WithCommitCharacterRules(commitRules)
        End Function
    End Class

    Friend Structure CompletionParts
        Public tagOpen As String
        Public textBeforeCaret As String
        Public textAfterCaret As String
        Public tagClose As String

        Public Sub New(tagOpen As String, textBeforeCaret As String, textAfterCaret As String, tagClose As String)
            Me.tagOpen = tagOpen
            Me.textBeforeCaret = textBeforeCaret
            Me.textAfterCaret = textAfterCaret
            Me.tagClose = tagClose
        End Sub

        Public Sub Deconstruct(ByRef tagOpen As String, ByRef textBeforeCaret As String, ByRef textAfterCaret As String, ByRef tagClose As String)
            tagOpen = Me.tagOpen
            textBeforeCaret = Me.textBeforeCaret
            textAfterCaret = Me.textAfterCaret
            tagClose = Me.tagClose
        End Sub
    End Structure

    Friend Structure AttrInfo
        Public Property Name As String

        Public Property HasEnum As Boolean

        Public Sub New(name As String, hasEnum As Boolean)
            Me.Name = name
            Me.HasEnum = hasEnum
        End Sub

        Public Shared Widening Operator CType(value As (name As String, hasEnum As Boolean)) As AttrInfo
            Return New AttrInfo(value.name, value.hasEnum)
        End Operator
    End Structure
End Namespace
