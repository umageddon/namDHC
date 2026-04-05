/*

 Console() Class

Tested On		Autohotkey_L version  1.1.13.00 ANSI
Author 			Nick McCoy (Ronins)
Initial Date			March 17, 2014
Version Release Date	March 31, 2014
Version			1.1

---------------------------------------------------------------------
Functions
---------------------------------------------------------------------
Write(Line)
WriteLine(Line)
ReadLine()
getch()
ReadConsoleInput() -returns Object Event, having members EventType and EventInfo[]
SetConsoleIcon(Path)
SetConsoleTitle(Title)
ClearScreen()
SetColor(BackgroundColor = 0x0, ForegroundColor = 0xF)
GetColor() - returns Object having members BackgroundColor and ForegroundColor
SetConsoleSize(Width, Height) - width and height are in rows and columns
GetConsoleSize() - returns Object with members BufferWidth/Height, Left, Top, Right, Bottom
SetCursorPosition(X, Y) - X and Y are in Column and row
GetCursorPosition()- returns object with X and Y as members
CreateProgress(X, Y, W, H, SmoothMode=0, Front="", Back="") - returns ProgressObjects
SetProgress(&ProgressObject, Value)
FillConsoleOutputCharacter(Character, StartCoordinates, Length)
FillConsoleOutputAttribute(Attribute, StartCoordinates, Length)
CreateConsoleScreenBuffer()
SetConsoleActiveScreenBuffer(hStdOut)
SetStdHandle(nStdHandle, Handle) - nStdHandle = -10 (input), -11 (output)
GetStdHandle(nStdHandle=-11)
GetConsolePID()

-------------------------------------------------------------------------------
References
-------------------------------------------------------------------------------
http://rdoc.info/github/luislavena/win32console/Win32/Console/Constants
http://msdn.microsoft.com/en-us/library/windows/desktop/ms682073(v=vs.85).aspx
http://msdn.microsoft.com/library/078sfkak
http://www.autohotkey.com/board/topic/42308-embedding-a-console-window-in-a-gui/

*/

class console
{	
	Color := Map("Black", 0x0, "DarkBlue", 0x1, "DarkGreen", 0x2, "Turquoise", 0x3, "DarkRed", 0x4, "Purple", 0x5, "Brown", 0x6, "Gray", 0x7, "DarkGray", 0x8, "Blue", 0x9, "Green", 0xA, "Cyan", 0xB, "Red", 0xC, "Magenta", 0xD, "Yellow", 0xE, "White", 0xF)
	VarCapacity := 1024*8 ;1 mb capacity
	Version := "1.1"
	
	
	__New(TargetPID := -1)
	{
		x:= DllCall("AttachConsole", "int", TargetPID)
		x:= DllCall("AllocConsole")
	}
	
	__Delete()
	{
		return DllCall("FreeConsole")
	}
	
	
	write(Line)
	{
		hStdout := DllCall("GetStdHandle", "int", -11)
		return DllCall("WriteConsole", "Ptr", hStdout, "Str", Line, "UInt", StrLen(Line), "UIntP", charsWritten := 0, "Ptr", 0)
	}
	
	writeLine(Line)
	{
		Line := Line "`n"
		hStdout := DllCall("GetStdHandle", "int", -11)
		return DllCall("WriteConsole", "Ptr", hStdout, "Str", Line, "UInt", StrLen(Line), "UIntP", charsWritten := 0, "Ptr", 0)
	}

	readLine()
	{
		hStdin := DllCall("GetStdHandle", "int", -10)
		Buffer := Buffer(this.VarCapacity, 0)
		DllCall("ReadConsole", "Ptr", hStdIn, "Ptr", Buffer.Ptr, "UInt", this.VarCapacity // 2, "UIntP", charsRead := 0, "Ptr", 0)
		Dummy := RegExReplace(StrGet(Buffer), "[\r\n].*")
		DllCall("FlushConsoleInputBuffer", "int", hStdIn)
		return Dummy
	}
	
	getch()
	{
		return DllCall("msvcrt.dll\_getch")
	}
	
	readConsoleInput()
	{
		Event := {EventList: Map(), EventInfo: []}
		Event.EventList[0x0001] := "4|8|10|12|14|16"
		Event.EventList[0x0002] := "4|6|8|12|16"
		
		InputRecord := Buffer(20, 0)
		s := Buffer(4, 0)
		hStdIn := DllCall("GetStdHandle", "int", -10)
		x:= DllCall("ReadConsoleInput", "Ptr", hStdIn, "Ptr", InputRecord.Ptr, "UInt", 1, "UIntP", eventsRead := 0)
		Event.EventType := NumGet(InputRecord, 0, "short")
		Dummy := Event.EventList[Event.EventType]
		Loop Parse, Dummy, "|"
			Event.EventInfo[A_Index] := NumGet(InputRecord, A_LoopField, "short")
		Event.s := NumGet(s, 0, "UInt")
		return Event
	}
	
	SetConsoleIcon(Path)
	{
		hIcon := DllCall("LoadImage", "uint", 0, "str", Path, "uint", 1, "int", 0, "int", 0, "uint", 0x00000010)
		return DllCall("SetConsoleIcon", "int", hIcon)
	}
	
	SetConsoleTitle(Title)
	{
		return DllCall("SetConsoleTitle", "str", Title)
	}
	
	clearScreen()
	{
		;~ return DllCall("msvcrt.dll\system", "str", "cls")
		consoleInfo := this.GetConsoleSize()
		Dummy := this.FillConsoleOutputCharacter(A_Space, {X:0, Y:0}, consoleInfo.BufferWidth*consoleInfo.BufferHeight)
		this.SetCursorPosition(0,0)
		return Dummy
	}
	
	setColor(BackgroundColor := 0x0, ForegroundColor := 0xF)
	{
		hStdout := DllCall("GetStdHandle", "int", -11)
		return DllCall("SetConsoleTextAttribute","int",hStdOut,"int",BackgroundColor<<4|ForeGroundColor)
	}
	
	getColor()
	{
		ConsoleInfo := Buffer(22, 0)
		hStdout := DllCall("GetStdHandle", "int", -11)
		DllCall("GetConsoleScreenBufferInfo","Ptr",hStdOut,"Ptr",ConsoleInfo.Ptr)
		Dummy := NumGet(ConsoleInfo, 8, "word")
		Color := {}
		Color.BackgroundColor := Dummy >> 4
		Color.ForegroundColor := Dummy & 0x0f
		return Color
	}
	
	setConsoleSize(Width, Height)
	{		
		rect := Buffer(8, 0)
		NumPut("Short", 0, rect, 0)
		NumPut("Short", 0, rect, 2)
		NumPut("Short", Width-1, rect, 4)
		NumPut("Short", Height-1, rect, 6)
		
		Coord := Buffer(4, 0)
		NumPut("Short", Width, Coord, 0)
		NumPut("Short", Height, Coord, 2)
		hStdout := DllCall("GetStdHandle", "int", -11)
		a:= DllCall("SetConsoleScreenBufferSize", "Ptr", hStdOut, "Int", NumGet(Coord, 0, "Int"))
		b:= DllCall("SetConsoleWindowInfo", "Ptr", hStdOut, "Int", 1, "Ptr", rect.Ptr)
		return a&&b
	}
	
	getConsoleSize()
	{
		hStdout := DllCall("GetStdHandle", "int", -11)
		ConsoleScreenBufferInfo := Buffer(20, 0)
		a := DllCall("GetConsoleScreenBufferInfo", "Ptr", hStdout, "Ptr", ConsoleScreenBufferInfo.Ptr)
		ConsoleScreenBufferInfoStructure := {}
		ConsoleScreenBufferInfoStructure.BufferWidth := NumGet(ConsoleScreenBufferInfo, 0, "short")
		ConsoleScreenBufferInfoStructure.BufferHeight := NumGet(ConsoleScreenBufferInfo, 2, "short")
		ConsoleScreenBufferInfoStructure.Left := NumGet(ConsoleScreenBufferInfo, 10, "short")
		ConsoleScreenBufferInfoStructure.Top := NumGet(ConsoleScreenBufferInfo, 12, "short")
		ConsoleScreenBufferInfoStructure.Right := NumGet(ConsoleScreenBufferInfo, 14, "short")
		ConsoleScreenBufferInfoStructure.Bottom := NumGet(ConsoleScreenBufferInfo, 16, "short")
		return ConsoleScreenBufferInfoStructure
	}
	
	setCursorPosition(X, Y)
	{
		hStdout := DllCall("GetStdHandle", "int", -11)
		Coord := Buffer(4, 0)
		NumPut("Short", X, Coord, 0)
		NumPut("Short", Y, Coord, 2)
		return DllCall("SetConsoleCursorPosition", "Ptr", hStdOut, "UInt", NumGet(Coord, 0, "UInt"))
	}
	
	getCursorPosition()
	{
		hStdout := DllCall("GetStdHandle", "int", -11)
		
		ConsoleScreenBufferInfo := Buffer(20, 0)
		DllCall("GetConsoleScreenBufferInfo", "Ptr", hStdout, "Ptr", ConsoleScreenBufferInfo.Ptr)
		return {X:NumGet(ConsoleScreenBufferInfo, 4, "short"), Y:NumGet(ConsoleScreenBufferInfo, 6, "short")}
	}
	
	createProgress(X, Y, W, H, SmoothMode:=0, Front:="", Back:="")
	{
		hStdout := DllCall("GetStdHandle", "int", -11)
		
		ConsoleScreenBufferInfo := Buffer(20, 0)
		DllCall("GetConsoleScreenBufferInfo", "Ptr", hStdout, "Ptr", ConsoleScreenBufferInfo.Ptr)
		MaxX := NumGet(ConsoleScreenBufferInfo, 0, "short")
		MaxY := NumGet(ConsoleScreenBufferInfo, 2, "short")
		W := (X+W>MaxX)?MaxX:W
		H := (Y+H>MaxY)?MaxY:H
		
		Old := this.GetCursorPosition()
		Loop H
		{
			this.SetCursorPosition(X, Y+A_Index-1)
			this.Write("[")
			this.SetCursorPosition(X+W, Y+A_Index-1)
			this.Write("]")
		}
		this.SetCursorPosition(Old.X, Old.Y)
		return {X:X, Y:Y, W:W, H:H, Value:0, SmoothMode:SmoothMode, Front:(SmoothMode?"0x40":"="), Back:(SmoothMode?"0x00":A_Space)}
	}
	
	setProgress(&ProgressObject, Value)
	{
		Increment := Round(Value*(ProgressObject.W-1)/100)
		Old := this.GetCursorPosition()
		Method := (ProgressObject.SmoothMode)? "FillConsoleOutputAttribute":"FillConsoleOutputCharacter"
		Loop ProgressObject.H
		{
			this[Method](ProgressObject.Back, {X:ProgressObject.X+1, Y:ProgressObject.Y+A_Index-1}, ProgressObject.W-1)
			this[Method](ProgressObject.Front, {X:ProgressObject.X+1, Y:ProgressObject.Y+A_Index-1}, Increment)
		}
		this.SetCursorPosition(Old.X, Old.Y)
		ProgressObject.Value := Value
	}
	
	fillConsoleOutputCharacter(Character, StartCoordinates, Length)
	{
		Coord := Buffer(4, 0)
		NumPut("Short", StartCoordinates.X, Coord, 0)
		NumPut("Short", StartCoordinates.Y, Coord, 2)
		hStdout := DllCall("GetStdHandle", "int", -11)
		return DllCall("FillConsoleOutputCharacter", "Ptr", hStdOut, "UShort", Ord(Character), "UInt", Length, "UInt", NumGet(Coord, 0, "UInt"), "UIntP", charsWritten := 0)
	}
	
	fillConsoleOutputAttribute(Attribute, StartCoordinates, Length)
	{
		Coord := Buffer(4, 0)
		NumPut("Short", StartCoordinates.X, Coord, 0)
		NumPut("Short", StartCoordinates.Y, Coord, 2)
		hStdout := DllCall("GetStdHandle", "int", -11)
		return DllCall("FillConsoleOutputAttribute", "Ptr", hStdOut, "UShort", Attribute, "UInt", Length, "UInt", NumGet(Coord, 0, "UInt"), "UIntP", charsWritten := 0)
	}
	
	createConsoleScreenBuffer()
	{
		return DllCall("CreateConsoleScreenBuffer", "int", 0x80000000|0x40000000, "int", 0x00000001|0x00000002, "int", 0, "int", 0x00000001, "int", 0)
	}
	
	setConsoleActiveScreenBuffer(hStdOut)
	{
		return DllCall("SetConsoleActiveScreenBuffer", "int", hStdOut)
	}
	
	setStdHandle(nStdHandle, Handle)
	{
		return DllCall("SetStdHandle", "int", nStdHandle, "int", Handle)
	}
	
	getStdHandle(nStdHandle:=-11)
	{
		return DllCall("GetStdHandle", "int", nStdHandle)
	}
	
	getConsolePID()
	{
		ConsoleHWnd := DllCall("GetConsoleWindow")
		return WinGetPID("ahk_id " ConsoleHWnd)
	}
	
	getConsoleHWND()
	{
		return DllCall("GetConsoleWindow")
	}
}
