[MeterAnchor]
Meter=Shape
X=(#Padding#+10)
Y=([MeterDivider:YH])
Shape=Rectangle 0,0,0,0
DynamicVariables=1
[MeterDay1Header]
Meter=String
Text=Tue Aug 12
X=(#Padding#+10)
Y=8R
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#DateColor#
FontWeight=700
AntiAlias=1
StringAlign=Left
DynamicVariables=1
LeftMouseUpAction=[!CommandMeasure "MeasureOpenOutlook" "Run"]
[MeasureOpenEvent837534]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A0900012C6D86A90000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay1Evt837534]
Meter=String
Text=   12:00 PM Lunch
X=(#Padding#+10)
Y=2R
W=(#Width#-2*#Padding#)
FontFace=#FontName#
FontSize=#FontSize#
FontColor=128,128,128,255
AntiAlias=1
ClipString=2
StringAlign=Left
DynamicVariables=1
ToolTipText=10 min
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent837534 Run]
[MeasureOpenEvent9CCA6C]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A090001802AE9550000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay1Evt9CCA6C]
Meter=String
Text=   2:00 PM Canceled: JPCI - ...
X=(#Padding#+10)
Y=2R
W=(#Width#-2*#Padding#)
FontFace=#FontName#
FontSize=#FontSize#
FontColor=128,128,128,255
AntiAlias=1
ClipString=2
StringAlign=Left
DynamicVariables=1
ToolTipText=25 min - Microsoft Teams Meeting
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent9CCA6C Run]
[MeasureOpenEvent9AE9ED]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A090004DCBB6ABA0000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay1Evt9AE9ED]
Meter=String
Text=   6:30 PM test
X=(#Padding#+10)
Y=2R
W=(#Width#-2*#Padding#)
FontFace=#FontName#
FontSize=#FontSize#
FontColor=128,128,128,255
AntiAlias=1
ClipString=2
StringAlign=Left
DynamicVariables=1
ToolTipText=25 min
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent9AE9ED Run]
[MeterDay2Header]
Meter=String
Text=Wed Aug 13
X=(#Padding#+10)
Y=8R
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#DateColor#
FontWeight=700
AntiAlias=1
StringAlign=Left
DynamicVariables=1
LeftMouseUpAction=[!CommandMeasure "MeasureOpenOutlook" "Run"]
[MeasureOpenEvent52EEBC]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A0900045871455F0000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay2Evt52EEBC]
Meter=String
Text=   8:30 AM Canceled: Going B...
X=(#Padding#+10)
Y=2R
W=(#Width#-2*#Padding#)
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#FontColor#
AntiAlias=1
ClipString=2
StringAlign=Left
DynamicVariables=1
ToolTipText=1h 30m - PKT / Teams Live; !Conf Room Thompson East; !Conf Room Thompson Main; !Conf Room Thompson West
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent52EEBC Run]
[MeasureOpenEventF3ADF5]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A09000458BEFE000000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay2EvtF3ADF5]
Meter=String
Text=   10:00 AM Weekly Update
X=(#Padding#+10)
Y=2R
W=(#Width#-2*#Padding#)
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#FontColor#
AntiAlias=1
ClipString=2
StringAlign=Left
DynamicVariables=1
ToolTipText=1 hr - !Conf Room Redbud
LeftMouseUpAction=[!CommandMeasure MeasureOpenEventF3ADF5 Run]
[MeasureOpenEvent472E8E]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A0900012C6D86A90000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay2Evt472E8E]
Meter=String
Text=   12:00 PM Lunch
X=(#Padding#+10)
Y=2R
W=(#Width#-2*#Padding#)
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#FontColor#
AntiAlias=1
ClipString=2
StringAlign=Left
DynamicVariables=1
ToolTipText=10 min
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent472E8E Run]
[MeterDay3Header]
Meter=String
Text=Thu Aug 14
X=(#Padding#+10)
Y=8R
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#DateColor#
FontWeight=700
AntiAlias=1
StringAlign=Left
DynamicVariables=1
LeftMouseUpAction=[!CommandMeasure "MeasureOpenOutlook" "Run"]
[MeasureOpenEvent16F404]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A090001802AE9680000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay3Evt16F404]
Meter=String
Text=   10:30 AM Canceled: Weekly ...
X=(#Padding#+10)
Y=2R
W=(#Width#-2*#Padding#)
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#FontColor#
AntiAlias=1
ClipString=2
StringAlign=Left
DynamicVariables=1
ToolTipText=30 min - Microsoft Teams Meeting
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent16F404 Run]
[MeasureOpenEvent858596]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A0900012C6D86A90000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay3Evt858596]
Meter=String
Text=   12:00 PM Lunch
X=(#Padding#+10)
Y=2R
W=(#Width#-2*#Padding#)
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#FontColor#
AntiAlias=1
ClipString=2
StringAlign=Left
DynamicVariables=1
ToolTipText=10 min
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent858596 Run]
[MeasureOpenEvent8393D4]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A090004D4782FFF0000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay3Evt8393D4]
Meter=String
Text=   3:00 PM SoftChoice x Jasc...
X=(#Padding#+10)
Y=2R
W=(#Width#-2*#Padding#)
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#FontColor#
AntiAlias=1
ClipString=2
StringAlign=Left
DynamicVariables=1
ToolTipText=1 hr - Microsoft Teams Meeting
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent8393D4 Run]
[MeterDay4Header]
Meter=String
Text=Fri Aug 15
X=(#Padding#+10)
Y=8R
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#DateColor#
FontWeight=700
AntiAlias=1
StringAlign=Left
DynamicVariables=1
LeftMouseUpAction=[!CommandMeasure "MeasureOpenOutlook" "Run"]
[MeasureOpenEvent7E6BA9]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A0900012C6D86A90000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay4Evt7E6BA9]
Meter=String
Text=   12:00 PM Lunch
X=(#Padding#+10)
Y=2R
W=(#Width#-2*#Padding#)
FontFace=#FontName#
FontSize=#FontSize#
FontColor=#FontColor#
AntiAlias=1
ClipString=2
StringAlign=Left
DynamicVariables=1
ToolTipText=10 min
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent7E6BA9 Run]
[MeterContentEnd]
Meter=Shape
X=0
Y=6R
Shape=Rectangle 0,0,0,0
DynamicVariables=1

