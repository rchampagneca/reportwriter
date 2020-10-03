#tag Class
Protected Class App
Inherits Application
	#tag Event
		Sub Open()
		  Dim TempDB As New SQLiteDatabase
		  
		  If Not TempDB.Connect Then
		    MsgBox("Connection failed: " + TempDB.ErrorMessage)
		  Else
		    Dim SQL As String
		    SQL = "CREATE TABLE Team (ID INTEGER, Name TEXT, Coach TEXT, City TEXT, PRIMARY KEY(ID));"
		    TempDB.SQLExecute(SQL)
		    
		    TempDB.SQLExecute("INSERT INTO Team (Name, Coach, City) VALUES ('Team 1', 'Coach 1', 'City 1');")
		    TempDB.SQLExecute("INSERT INTO Team (Name, Coach, City) VALUES ('Team 2', 'Coach 2', 'City 2');")
		    TempDB.SQLExecute("INSERT INTO Team (Name, Coach, City) VALUES ('Team 3', 'Coach 3', 'City 3');")
		    TempDB.SQLExecute("INSERT INTO Team (Name, Coach, City) VALUES ('Team 4', 'Coach 4', 'City 4');")
		    TempDB.SQLExecute("INSERT INTO Team (Name, Coach, City) VALUES ('Team 5', Coach 5', 'City 5');")
		    TempDB.SQLExecute("INSERT INTO Team (Name, Coach, City) VALUES ('Team 6', 'Coach 6', 'City 6');")
		    TempDB.SQLExecute("INSERT INTO Team (Name, Coach, City) VALUES ('Team 7', 'Coach 7', 'City 7');")
		    TempDB.SQLExecute("INSERT INTO Team (Name, Coach, City) VALUES ('Team 8', 'Coach 8', 'City 8');")
		    
		    TempDB.Commit
		  End If
		  
		  inMemoryDB = TempDB
		End Sub
	#tag EndEvent


	#tag Property, Flags = &h0
		inMemoryDB As SQLiteDatabase
	#tag EndProperty


	#tag Constant, Name = kEditClear, Type = String, Dynamic = False, Default = \"&Delete", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"&Delete"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"&Delete"
	#tag EndConstant

	#tag Constant, Name = kFileQuit, Type = String, Dynamic = False, Default = \"&Quit", Scope = Public
		#Tag Instance, Platform = Windows, Language = Default, Definition  = \"E&xit"
	#tag EndConstant

	#tag Constant, Name = kFileQuitShortcut, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Mac OS, Language = Default, Definition  = \"Cmd+Q"
		#Tag Instance, Platform = Linux, Language = Default, Definition  = \"Ctrl+Q"
	#tag EndConstant


	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
