#tag Class
Protected Class ReportWriter
Inherits Canvas
	#tag Method, Flags = &h0
		Sub addComment(Comment As String)
		  AppendString("<!--" + Comment + "-->")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AppendString(sText As String)
		  // Append a HTML string
		  
		  #Pragma DisableBackgroundTasks
		  
		  Dim l As Integer
		  l = LenB(sText)
		  mHTML.StringValue(Length, l) = sText
		  Length = Length + l
		  
		  Exception
		    // Increase size of MemoryBlock
		    TotalLength = TotalLength + BufferSize
		    mHTML.Size = TotalLength
		    mHTML.StringValue(Length, l) = sText
		    Length = Length + l
		    
		    
		    
		    
		    
		    
		    
		    
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub clearColumns(paramarray columns as Dictionary)
		  ' empty arrays
		  
		  Dim column as Dictionary
		  
		  For Each column in columns
		    
		    ' clear
		    column.Clear
		    
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub createGraphics()
		  // Initialize HTML head tag, document title and HighChart stylesheets and scripts
		  
		  
		  ReportName = "Graphiques.html"
		  
		  // This control uses a MemoryBlock as a temporary buffer
		  // while creating long strings by appending to already existing
		  // strings. It can be a lot quicker than s = s + "more text" if
		  // you have very long strings.
		  
		  // This is both the size of the original MemoryBlock
		  // but also how much it increases by each time if
		  // more space is needed.
		  
		  BufferSize = 100000 // bigger than really needed
		  
		  mHTML = New MemoryBlock(BufferSize)
		  TotalLength = BufferSize
		  Length = 0
		  
		  AppendString("<!DOCTYPE HTML>")
		  AppendString("<html>")
		  AppendString("<head>")
		  AppendString("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">")
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub createReport(Title As String, Optional pReportName As String, optional Landscape As Boolean = False)
		  // Initialize HTML head tag, document title and stylesheet
		  
		  If len(pReportName) = 0 Then
		    ReportName = "rapport.html"
		  Else
		    ReportName = pReportName + ".html"
		  End If
		  
		  // This control uses a MemoryBlock as a temporary buffer
		  // while creating long strings by appending to already existing
		  // strings. It can be a lot quicker than s = s + "more text" if
		  // you have very long strings.
		  
		  // This is both the size of the original MemoryBlock
		  // but also how much it increases by each time if
		  // more space is needed.
		  
		  BufferSize = 100000 // bigger than really needed
		  
		  mHTML = New MemoryBlock(BufferSize)
		  TotalLength = BufferSize
		  Length = 0
		  
		  AppendString("<!doctype html>")
		  AppendString("<html lang=""fr"">")
		  AppendString("<head>")
		  AppendString("<meta charset=""utf-8"">")
		  AppendString("<meta name=""viewport"" content=""width=device-width, initial-scale=1, shrink-to-fit=no"">")
		  AppendString("<title>" + title + "</title>")
		  
		  AppendString("<link rel=""stylesheet"" href=""https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css"" integrity=""sha384-JcKb8q3iqJ61gNV9KGb8thSsNjpSL0n8PARn9HuZOnIxN0hoP+VmmDGMN5t9UJ0Z"" crossorigin=""anonymous"">")
		  AppendString("<script src=""https://code.jquery.com/jquery-3.5.1.slim.min.js"" integrity=""sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj"" crossorigin=""anonymous""></script>")
		  AppendString("<script src=""https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js"" integrity=""sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl"" crossorigin=""anonymous""></script>")
		  If Landscape Then
		    AppendString("<style>@media print{@page {size: landscape}}</style>")
		  End If
		  AppendString("</head>")
		  AppendString("<body>")
		  AppendString("<div class=""container"">")
		  AppendString("<h5 class=""text-center"">" + Title + "</h5>")
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub finishReport()
		  // Insert closing HTML tags
		  
		  AppendString("</div>")
		  AppendString("</body>")
		  AppendString("</html>")
		  sHTML = GetString()
		  
		  e = SpecialFolder.Temporary.Child(ReportName)
		  
		  If e.Exists Then
		    e.Delete
		  End If
		  
		  Dim t As TextOutputStream
		  t = TextOutputStream.Append(e)
		  t.Write(sHTML)
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetString() As String
		  #Pragma DisableBackgroundTasks
		  Return mHTML.StringValue(0, Length)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ShowPDF(FileName As FolderItem) As Integer
		  // Créer un fichier PDF
		  
		  Dim sh As Shell
		  Dim cmdString As String
		  Dim paramString As String
		  
		  #If TargetWin32 Then
		    Dim f As FolderItem = App.ExecutableFile.Parent.Child("Adds").Child("wkhtmltopdf.exe")
		    cmdString = f.NativePath
		  #Else
		    cmdString= "/usr/local/bin/wkhtmltopdf"
		  #EndIf
		  
		  paramString = e.NativePath + " " + FileName.NativePath
		  sh = New Shell
		  sh.Execute(cmdString, paramString)
		  
		  Return sh.ErrorCode
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ShowReport()
		  // Afficher la page web dans le fureteur de l'utilisateur
		  
		  ShowURL("file://" + e.NativePath)
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub tableCreate(columns() As Dictionary)
		  // Create a table
		  // Parameters: columns
		  
		  Dim column As Dictionary
		  
		  AppendString("<table class=""table table-striped table-bordered"" style=""width:100%"">")
		  AppendString("<thead><tr>")
		  For Each column In Columns
		    If column.HasKey("colspan") Then
		      If column.HasKey("class") Then
		        AppendString("<th class =""" + column.Value("class") + """ colspan=""" + column.Value("colspan") + """ scope=""col"">" + column.Value("text") + "</th>")
		      Else
		        AppendString("<th colspan=""" + column.Value("colspan") + """ scope=""col"">" + column.Value("text") + "</th>")
		      End If
		      
		    Else
		      If column.HasKey("class") Then
		        AppendString("<th class =""" + column.Value("class") + """ scope=""col"">" + column.Value("text") + "</th>")
		      Else
		        AppendString("<th scope=""col"">" + column.Value("text") + "</th>")
		      End If
		      
		    End If
		  Next
		  
		  AppendString("</tr></thead>")
		  AppendString("<tbody>")
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub tableCreate(paramarray columns as Dictionary)
		  ' create a table
		  
		  Dim data(-1) as Dictionary
		  Dim column as Dictionary
		  
		  For Each column in columns
		    
		    ' add to array
		    data.Append(column)
		    
		  Next
		  
		  ' write table
		  tableCreate(data)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub tableEnd()
		  AppendString("</tbody>")
		  AppendString("</table>")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub tableHeader(columns() As Dictionary)
		  // Create a table
		  // Parameters: columns
		  
		  Dim column As Dictionary
		  
		  For Each column In Columns
		    
		    If column.HasKey("colspan") Then
		      If column.HasKey("class") Then
		        AppendString("<th class =""" + column.Value("class")+ """ colspan=""" + column.Value("colspan") + """ scope=""col"">" + column.Value("text") + "</th>")
		      Else
		        AppendString("<th colspan=""" + column.Value("colspan") + """ scope=""col"">" + column.Value("text") + "</th>")
		      End If
		      
		    Else
		      
		      If column.HasKey("class") Then
		        AppendString("<th class =""" + column.Value("class") + """ scope=""col"">" + column.Value("text") + "</th>")
		      Else
		        AppendString("<th scope=""col"">" + column.Value("text") + "</th>")
		      End If
		    End If
		    
		  Next
		  AppendString("</tr>")
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub tableHeader(paramarray columns as Dictionary)
		  ' create a table
		  
		  Dim data(-1) as Dictionary
		  Dim column as Dictionary
		  
		  For Each column in columns
		    
		    ' add to array
		    data.Append(column)
		    
		  Next
		  
		  ' write table
		  tableHeader(data)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub tableHeaderEnd()
		  AppendString("</thead>")
		  AppendString("<tbody>")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub tableStart()
		  // Multiline table header
		  AppendString("<table class=""table table-striped table-bordered"" style=""width:100%"">")
		  AppendString("<thead><tr>")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub tableWrite(columns() As Dictionary)
		  // Print a table row
		  // Parameter list:
		  // Text = The paragraph to print
		  // align = left, right, center or justify
		  
		  Dim column As Dictionary
		  
		  AppendString ("<tr>")
		  For Each column In Columns
		    If column.HasKey("class") Then
		      AppendString("<td class=""" + column.Value("class") + """>" + column.Value("text") + "</td>")
		    Else
		      AppendString("<td>" + column.Value("text") + "</td>")
		    End If
		  Next
		  AppendString ("</tr>")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub tableWrite(paramarray columns as Dictionary)
		  ' create a table
		  
		  Dim data(-1) as Dictionary
		  Dim column as Dictionary
		  
		  For Each column in columns
		    
		    ' add to array
		    data.Append(column)
		    
		  Next
		  
		  ' write table
		  tableWrite(data)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub write2Cols(Col1 As Dictionary, Col2 As Dictionary)
		  // Define 2 equal-width columns
		  
		  AppendString("<div class=""row"">")
		  
		  If Col1.HasKey("class") Then
		    AppendString("<p class=""col-6 " + Col1.Value("class") + """>" + Col1.Value("text") + "</p>")
		  Else
		    AppendString("<p class=""col-6"">" + Col1.Value("text") + "</p>")
		  End If
		  
		  If Col2.HasKey("class") Then
		    AppendString("<p class=""col-6 " + Col2.Value("class") + """>" + Col2.Value("text") + "</p>")
		  Else
		    AppendString("<p class=""col-6"">" + Col2.Value("text") + "</p>")
		  End If
		  
		  AppendString("</div>")
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub write2Cols(Title1 As String, Title2 As String)
		  // Define 2 equal-width columns
		  
		  AppendString("<div class=""row"">" _
		  + "<p class=""col-6 text-left row-compress"">" _
		  + Title1 _
		  + "</p>" _
		  + "<p class=""col-6 text-left row-compress"">" _
		  + Title2 _
		  + "</p>" _
		  + "</div>")
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub write3Cols(Col1 As Dictionary, Col2 As Dictionary, Col3 As Dictionary)
		  // Define 3 equal-width columns
		  
		  AppendString("<div class=""row"">")
		  
		  If Col1.HasKey("class") Then
		    AppendString("<p class=""col-4 " + Col1.Value("class") + """>" + Col1.Value("text") + "</p>")
		  Else
		    AppendString("<p class=""col-4"">" + Col1.Value("text") + "</p>")
		  End If
		  
		  If Col2.HasKey("class") Then
		    AppendString("<p class=""col-4 " + Col2.Value("class") + """>" + Col2.Value("text") + "</p>")
		  Else
		    AppendString("<p class=""col-4"">" + Col2.Value("text") + "</p>")
		  End If
		  
		  If Col3.HasKey("class") Then
		    AppendString("<p class=""col-4 " + Col3.Value("class") + """>" + Col3.Value("text") + "</p>")
		  Else
		    AppendString("<p class=""col-4"">" + Col3.Value("text") + "</p>")
		  End If
		  
		  AppendString("</div>")
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub write3Cols(Title1 As String, Title2 As String, Title3 as String, Optional RowClass As String)
		  // Define 3 equal-width columns
		  
		  AppendString("<div class=""row"">" _
		  + "<p class=""col-4 text-left row-compress"">" _
		  + Title1 _
		  + "</p>" _
		  + "<p class=""col-4 text-center row-compress"">" _
		  + Title2 _
		  + "</p>" _
		  + "<p class=""col-4 text-center row-compress"">" _
		  + Title3 _
		  + "</p>" _
		  + "</div>")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub write4Cols(Title1 As String, Title2 As String, Title3 as String, Title4 as String)
		  // Define 4 equal-width columns
		  
		  AppendString("<div class=""row"">" _
		  + "<p class=""col-3 text-left row-compress"">" _
		  + Title1 _
		  + "</p>" _
		  + "<p class=""col-3 text-center row-compress"">" _
		  + Title2 _
		  + "</p>" _
		  + "<p class=""col-3 text-center row-compress"">" _
		  + Title3 _
		  + "</p>" _
		  + "<p class=""col-3 text-left row-compress"">" _
		  + Title4 _
		  + "</p>" _
		  + "</div>")
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub writeBreak(Optional Number As Integer = 1)
		  // Insert line breaks
		  
		  For i As Integer = 1 To Number
		    // AppendString("<br>")
		    AppendString("<p>&nbsp;</p>")
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub writeGrid(Col() As Dictionary, Optional RowClass As String)
		  // Define 2 equal-width columns
		  Dim column As Dictionary
		  Dim m_Class As String
		  
		  m_Class = "row " + RowClass
		  
		  AppendString("<div class=""" + m_Class + """>")
		  
		  For Each column In col
		    If column.HasKey("class") Then
		      AppendString("<p class=""col " + column.Value("class") + """>" _
		      + column.Value("text") _
		      + "</p>")
		    Else
		      AppendString("<p class=""col"">" _
		      + column.Value("text") _
		      + "</p>")
		    End If
		  Next
		  
		  AppendString("</div>")
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub writeGrid(ParamArray columns As Dictionary)
		  ' create a table
		  
		  Dim data(-1) As Dictionary
		  Dim column As Dictionary
		  
		  For Each column In columns
		    
		    ' add to array
		    data.Append(column)
		    
		  Next
		  
		  ' write table
		  writeGrid(data)
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub writeHTML(HTML as String)
		  // Write HTML
		  
		  AppendString(HTML)
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub writeLine()
		  // Write a line
		  
		  AppendString("<hr>")
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub writeText(Text As String, Optional args as Dictionary)
		  // Write a Paragraph
		  // Parameters:
		  // Text: Paragraph
		  // align: text-left, text-right, text-center, text-justify, text-nowrap or any other class
		  
		  Dim m_class As String
		  
		  if args.HasKey("class") then
		    m_class = args.Value("class").StringValue
		    AppendString("<p class = """ + m_class + """>" + Text + "</p>")
		  Else
		    AppendString("<p>" + Text + "</p>")
		  End If
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub writeText(Text As String, paramarray args as Pair)
		  Dim d As New Dictionary
		  For Each arg As Pair In args
		    d.Value( arg.Left ) = arg.Right
		  Next
		  
		  writeText( Text, d )
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h21
		Private BufferSize As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private e As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h21
		Private Length As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mHTML As MemoryBlock
	#tag EndProperty

	#tag Property, Flags = &h21
		Private ReportName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		sHTML As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private TotalLength As Integer
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="AcceptFocus"
			Visible=true
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="AcceptTabs"
			Visible=true
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="AutoDeactivate"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Backdrop"
			Visible=true
			Group="Appearance"
			Type="Picture"
			EditorType="Picture"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DoubleBuffer"
			Visible=true
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Enabled"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="EraseBackground"
			Visible=true
			Group="Behavior"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Height"
			Visible=true
			Group="Position"
			InitialValue="100"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="HelpTag"
			Visible=true
			Group="Appearance"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			Type="Integer"
			EditorType="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="InitialParent"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockBottom"
			Visible=true
			Group="Position"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockLeft"
			Visible=true
			Group="Position"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockRight"
			Visible=true
			Group="Position"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockTop"
			Visible=true
			Group="Position"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
			EditorType="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="sHTML"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
			EditorType="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabIndex"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabPanelIndex"
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabStop"
			Visible=true
			Group="Position"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Transparent"
			Visible=true
			Group="Behavior"
			InitialValue="True"
			Type="Boolean"
			EditorType="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="UseFocusRing"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Visible"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Width"
			Visible=true
			Group="Position"
			InitialValue="100"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
