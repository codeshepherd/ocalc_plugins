
REM  *****  BASIC  *****

'---------------------------------------------------------
' READ THIS IS MANDATORY :-) 
'
' This Open Office macro will 
' - read all the .xls files located in the cFolder defined some line below
' - export each sheet of each file as a separate .csv file, named : filename_sheetname.csv
'
' you need to adapt this file to your needs (directory cFolder, cFieldTypes and maybe other...)
'
' you can adapt it easyly to other imput/ouput formats.
'
'---------------------------------------------------------

Sub Export_CSV
' This is the hardcoded pathname to a folder containing Excel files.
cFolder = "/tmp/ocalc/"
pDoc = ThisComponent

' Get the pathname of each file within the folder.
cFile = Dir$( cFolder + "/*.*" )
Do While cFile <> ""
' If it is not a directory...
If cFile <> "."  And  cFile <> ".." Then
' If it has the right suffix...
If LCase( Right( cFile, 4 ) ) = ".xls" Then
' Open the document.
oDoc = StarDesktop.loadComponentFromURL(_
		ConvertToUrl( cFolder + "/" + cFile ),"_blank", 0, Array() )
'=========
' Options for delimiters in CVS
'cFieldDelimiters = Chr(9)
cFieldDelimiters = ";"
	'cTextDelimiter = ""
cTextDelimiter = Chr(34)
	cFieldTypes = "2/2/2/2/2/2/2/9/9/9/9/9/9/9/9/9/9"
	' options....
	'   cFieldDelimiters = ",;" ' for either commas or semicolons
	'   cFieldDelimiters = Chr(9) ' for tab
	'   cTextDelimiter = Chr(34) ' for double quote
	'   cTextDelimiter = Chr(39) ' for single quote
	' Suppose you want your first field to be numeric, then two text fields, and then a date field....
	'   cFieldTypes = "1/2/2/3"
	' Use 1=Num, 2=Text, 3=MM/DD/YY, 4=DD/MM/YY, 5=YY/MM/DD, 9=ignore field (do not import)
	'----------
	' Build up the Filter Options string
	' From the Developer's Guide
	' http://api.openoffice.org/docs/DevelopersGuide/DevelopersGuide.htm
	' See section 8.2.2 under Filter Options
	' http://api.openoffice.org/docs/DevelopersGuide/Spreadsheet/Spreadsheet.htm#1+2+2+3+Filter+Options 
	cFieldDelims = ""
	For i = 1 To Len( cFieldDelimiters )
c = Mid( cFieldDelimiters, i, 1 )
	If Len( cFieldDelims ) > 0 Then
	cFieldDelims = cFieldDelims + "/"
	EndIf
cFieldDelims = cFieldDelims + CStr(Asc( c ))
	Next

	If Len( cTextDelimiter ) > 0 Then
cTextDelim = CStr(Asc( cTextDelimiter ))
	Else
	cTextDelim = "0"
	EndIf

	cFilterOptions = cFieldDelims + "," + cTextDelim + ",0,1," + cFieldTypes

	'=========
	' Prepare new filename
cNewName = Left( cFile, Len( cFile ) - 4 )

	' Save it in OOo format.
	'oDoc.storeToURL( ConvertToUrl( cFolder + "/" + cNewName + ".sxc" ), Array() )

	' Loop and selects sheets to save as csv
	oSheets = oDoc.Sheets()
aSheetNames = oSheets.getElementNames()
	For index=0 to oSheets.getCount() -1
oSheet = oSheets.getByIndex(index)

	' Define prefix or suffix to append to filename
	appendName = aSheetNames(index) 'define prefix/suffix as the name of the sheet
	appendNum = index + 21 ' define prefix/suffix as the number of the sheet                  
	' Choose new filename, with prefix or suffix
	'cNewFileName = appendName + "_" + cNewName 'prefix name
	'cNewFileName = appendNum + "_" + cNewName ' prefix number
	'cNewFileName = cNewName + "_" + appendName ' suffix name
	'cNewFileName = cNewName +  "_" + appendNum ' suffix number
	cNewFileName = appendName

	' Replace spaces with underscores in filenames.
	cNewFileName = Replace(cNewFileName, " ", "_")

	oController = oDoc.GetCurrentController()  'view controller
	oController.SetActiveSheet(oSheet) 'switches view to sheet object

	' Export it using a filter.
	oDoc.StoreToURL( ConvertToUrl( cFolder + "/" + cNewFileName + CFile  + ".csv" ),_
			Array( MakePropertyValue( "FilterName", "Text - txt - csv (StarCalc)" ),_
				MakePropertyValue( "FilterOptions", cFilterOptions ),_
				MakePropertyValue( "SelectionOnly", true ) ) )

	'Insert sheet

	dispatchURL(oDoc,".uno:SelectAll")
	dispatchURL(oDoc,".uno:Copy")

	'cSheet = oSheet
	'pDoc.getSheets.insertByName("newsheet")
	'pDoc.Sheets.insertByName("inserted",1)
	sc = pDoc.getSheets().getCount()
	pDoc.getSheets().insertNewByName(cFile,sc+1)
selectSheetByName(pDoc, cFile)
	dispatchURL(pDoc,".uno:Paste")


	Next index
	' Close the document.
oDoc.dispose()
	EndIf
	EndIf
	cFile = Dir$
	Loop
	End Sub

	Sub dispatchURL(document, aURL)
Dim noProps()
	Dim URL As New com.sun.star.util.URL

frame = document.getCurrentController().getFrame()
	URL.Complete = aURL
	transf = createUnoService("com.sun.star.util.URLTransformer")
transf.parseStrict(URL)

	disp = frame.queryDispatch(URL, "", com.sun.star.frame.FrameSearchFlag.SELF _
			OR com.sun.star.frame.FrameSearchFlag.CHILDREN)
disp.dispatch(URL, noProps())
	End Sub

	Sub selectSheetByName(document, sheetName)
document.getCurrentController.select(document.getSheets().getByName(sheetName))
	End Sub


