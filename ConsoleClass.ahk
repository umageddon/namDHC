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
SetProgress(ByRef ProgressObject, Value)
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
	Color := Object("Black", 0x0, "DarkBlue", 0x1, "DarkGreen", 0x2, "Turquoise", 0x3, "DarkRed", 0x4, "Purple", 0x5, "Brown", 0x6, "Gray", 0x7, "DarkGray", 0x8, "Blue", 0x9, "Green", 0xA, "Cyan", 0xB, "Red", 0xC, "Magenta", 0xD, "Yellow", 0xE, "White", 0xF)
	VarCapacity := 1024*8 ;1 mb capacity
	Version := "1.1"
	
	__New(TargetPID = -1)
	{
		DllCall("AttachConsole", int, TargetPID)
		DllCall("AllocConsole")
	}
	
	__Delete()
	{
		return DllCall("FreeConsole")
	}
	
	write(Line)
	{
		hStdout := DllCall("GetStdHandle", "int", -11)
		return, DllCall("WriteConsole", "int", hStdout, "uint", &Line, "uint", StrLen(Line), "uint", 0)
	}
	
	writeLine(Line)
	{
		Line := Line "`n"
		hStdout := DllCall("GetStdHandle", "int", -11)
		return, DllCall("WriteConsole", "int", hStdout, "uint", &Line, "uint", StrLen(Line), "uint", 0)
	}

	readLine()
	{
		hStdin := DllCall("GetStdHandle", "int", -10)
		VarSetCapacity(Buffer, this.VarCapacity)
		DllCall("ReadConsole", "int", hStdIn, "int", &Buffer, "int", this.VarCapacity, "int", 0)
		RegExMatch(StrGet(&Buffer), ".*", Dummy)
		DllCall("FlushConsoleInputBuffer", "int", hStdIn)
		return, Dummy
	}
	
	getch()
	{
		return, DllCall("msvcrt.dll\_getch")
	}
	
	readConsoleInput()
	{
		Event := {}
		Event.EventList[0x0001] := "4|8|10|12|14|16"
		Event.EventList[0x0002] := "4|6|8|12|16"
		
		VarSetCapacity(InputRecord, 20)
		VarSetCapacity(s, 4)
		hStdIn := DllCall("GetStdHandle", "int", -10)
		DllCall("ReadConsoleInput", "int", hStdIn, "int", &InputRecord, "int", 100, "int", &s)
		Event.EventType := NumGet(InputRecord, 0, "short")
		Dummy := Event.EventList[Event.EventType]
		Loop, Parse, Dummy, |
			Event.EventInfo[A_Index] := NumGet(InputRecord, A_LoopField, "short")
		Event.s := NumGet(s)
		return, Event
	}
	
	SetConsoleIcon(Path)
	{
		hIcon := DllCall("LoadImage", "uint", 0, "str", Path, "uint", 1, "int", 0, "int", 0, "uint", 0x00000010)
		return, DllCall("SetConsoleIcon", "int", hIcon)
	}
	
	SetConsoleTitle(Title)
	{
		return, DllCall("SetConsoleTitle", "str", Title)
	}
	
	clearScreen()
	{
		;~ return, DllCall("msvcrt.dll\system", "str", "cls")
		Console := this.GetConsoleSize()
		Dummy := this.FillConsoleOutputCharacter(A_Space, Object("X", 0, "Y", 0) , Console.BufferWidth*Console.BufferHeight)
		this.SetCursorPosition(0,0)
		return, Dummy
	}
	
	setColor(BackgroundColor = 0x0, ForegroundColor = 0xF)
	{
		hStdout := DllCall("GetStdHandle", "int", -11)
		return, DllCall("SetConsoleTextAttribute","int",hStdOut,"int",BackgroundColor<<4|ForeGroundColor)
	}
	
	getColor()
	{
		VarSetCapacity(ConsoleInfo, 22)
		hStdout := DllCall("GetStdHandle", "int", -11)
		DllCall("GetConsoleScreenBufferInfo","int",hStdOut,"int",&ConsoleInfo)
		Dummy := NumGet(ConsoleInfo, 8, "word")
		Color := {}
		Color.BackgroundColor := Dummy >> 4
		Color.ForegroundColor := Dummy & 0x0f
		return, Color
	}
	
	setConsoleSize(Width, Height)
	{		
		VarSetCapacity(rect, 8)
		NumPut(0, rect, 0, "short")
		NumPut(0, rect, 2, "short")
		NumPut(Width-1, rect, 4, "short")
		NumPut(Height-1, rect, 6, "short")
		
		VarSetCapacity(Coord, 4)
		NumPut(Width, Coord, "Short")
		NumPut(Height, Coord, 2, "Short")
		hStdout := DllCall("GetStdHandle", "int", -11)
		a:= DllCall("SetConsoleScreenBufferSize", "int", hStdOut, "int", Numget(Coord, "int"))
		b:= DllCall("SetConsoleWindowInfo", "int", hStdOut, "int", 1, "int", &rect)
		return, a&&b
	}
	
	getConsoleSize()
	{
		hStdout := DllCall("GetStdHandle", "int", -11)
		VarSetCapacity(ConsoleScreenBufferInfo, 20)
		a := DllCall("GetConsoleScreenBufferInfo", "int", hStdout, "int", &ConsoleScreenBufferInfo)
		ConsoleScreenBufferInfoStructure := {}
		ConsoleScreenBufferInfoStructure.BufferWidth := NumGet(ConsoleScreenBufferInfo, 0, "short")
		ConsoleScreenBufferInfoStructure.BufferHeight := NumGet(ConsoleScreenBufferInfo, 2, "short")
		ConsoleScreenBufferInfoStructure.Left := NumGet(ConsoleScreenBufferInfo, 10, "short")
		ConsoleScreenBufferInfoStructure.Top := NumGet(ConsoleScreenBufferInfo, 12, "short")
		ConsoleScreenBufferInfoStructure.Right := NumGet(ConsoleScreenBufferInfo, 14, "short")
		ConsoleScreenBufferInfoStructure.Bottom := NumGet(ConsoleScreenBufferInfo, 16, "short")
		return, ConsoleScreenBufferInfoStructure
	}
	
	setCursorPosition(X, Y)
	{
		hStdout := DllCall("GetStdHandle", "int", -11)
		VarSetCapacity(Coord, 4)
		NumPut(X, Coord, "Short")
		NumPut(Y, Coord, 2, "Short")
		return, DllCall("SetConsoleCursorPosition", "int", hStdOut, "int", NumGet(Coord, "uint"))
	}
	
	getCursorPosition()
	{
		hStdout := DllCall("GetStdHandle", "int", -11)
		
		VarSetCapacity(ConsoleScreenBufferInfo, 20)
		DllCall("GetConsoleScreenBufferInfo", "int", hStdout, "int", &ConsoleScreenBufferInfo)
		return, Object("X", NumGet(ConsoleScreenBufferInfo, 4, "short"), "Y", NumGet(ConsoleScreenBufferInfo, 6, "short"))
	}
	
	createProgress(X, Y, W, H, SmoothMode=0, Front="", Back="")
	{
		hStdout := DllCall("GetStdHandle", "int", -11)
		
		VarSetCapacity(ConsoleScreenBufferInfo, 20)
		DllCall("GetConsoleScreenBufferInfo", "int", hStdout, "int", &ConsoleScreenBufferInfo)
		MaxX := NumGet(ConsoleScreenBufferInfo, 0, "short")
		MaxY := NumGet(ConsoleScreenBufferInfo, 2, "short")
		W := (X+W>MaxX)?MaxX:W
		H := (Y+H>MaxY)?MaxY:H
		
		Old := this.GetCursorPosition()
		Loop, % H
		{
			this.SetCursorPosition(X, Y+A_Index-1)
			this.Write("[")
			this.SetCursorPosition(X+W, Y+A_Index-1)
			this.Write("]")
		}
		this.SetCursorPosition(Old.X, Old.Y)
		return, Object("X", X, "Y", Y, "W", W, "H", H, "Value", 0, "SmoothMode", SmoothMode, "Front", SmoothMode?"0x40":"=", "Back", SmoothMode?"0x00":A_Space)
	}
	
	setProgress(ByRef ProgressObject, Value)
	{
		Increment := Round(Value*(ProgressObject.W-1)/100)
		Old := this.GetCursorPosition()
		Method := (ProgressObject.SmoothMode)? "FillConsoleOutputAttribute":"FillConsoleOutputCharacter"
		Loop, % ProgressObject.H
		{
			this[Method](ProgressObject.Back, Object("X", ProgressObject.X+1, "Y", ProgressObject.Y+A_Index-1), ProgressObject.W-1)
			this[Method](ProgressObject.Front, Object("X", ProgressObject.X+1, "Y", ProgressObject.Y+A_Index-1), Increment)
		}
		this.SetCursorPosition(Old.X, Old.Y)
		ProgressObject.Value := Value
	}
	
	fillConsoleOutputCharacter(Character, StartCoordinates, Length)
	{
		VarSetCapacity(Coord, 4)
		NumPut(StartCoordinates.X, Coord, 0, "short")
		NumPut(StartCoordinates.Y, Coord, 2, "short")
		VarSetCapacity(s, 4)
		hStdout := DllCall("GetStdHandle", "int", -11)
		return, DllCall("FillConsoleOutputCharacter", "int", hStdOut, "int", Asc(Character), "int", Length, "int", NumGet(Coord, "uint"), "int", &s)
	}
	
	fillConsoleOutputAttribute(Attribute, StartCoordinates, Length)
	{
		VarSetCapacity(Coord, 4)
		NumPut(StartCoordinates.X, Coord, 0, "short")
		NumPut(StartCoordinates.Y, Coord, 2, "short")
		VarSetCapacity(s, 4)
		hStdout := DllCall("GetStdHandle", "int", -11)
		return, DllCall("FillConsoleOutputAttribute", "int", hStdOut, "int", Attribute, "int", Length, "int", NumGet(Coord, "uint"), "int", &s)
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
	
	getStdHandle(nStdHandle=-11)
	{
		return DllCall("GetStdHandle", "int", nStdHandle)
	}
	
	getConsolePID()
	{
		ConsoleHWnd := DllCall("GetConsoleWindow")
		WinGet, ConsolePID, PID, ahk_id %ConsoleHWnd%
		return ConsolePID
	}
	
	getConsoleHWND()
	{
		return DllCall("GetConsoleWindow")
	}
}