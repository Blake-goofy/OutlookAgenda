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
ToolTipText=Click to open calendar
[MeasureOpenEvent861133]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A0900012C6D86A90000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay1Evt861133]
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
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent861133 Run]
[MeasureOpenEvent96F86B]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A090001802AE9550000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay1Evt96F86B]
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
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent96F86B Run]
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
ToolTipText=Click to open calendar
[MeasureOpenEvent5386A6]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A0900045871455F0000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay2Evt5386A6]
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
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent5386A6 Run]
[MeasureOpenEvent7F9D90]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A09000458BEFE000000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay2Evt7F9D90]
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
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent7F9D90 Run]
[MeasureOpenEvent9AD415]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A0900012C6D86A90000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay2Evt9AD415]
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
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent9AD415 Run]
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
ToolTipText=Click to open calendar
[MeasureOpenEvent7E27F2]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A090001802AE9680000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay3Evt7E27F2]
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
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent7E27F2 Run]
[MeasureOpenEvent2075CC]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A0900012C6D86A90000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay3Evt2075CC]
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
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent2075CC Run]
[MeasureOpenEvent9B75AC]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A090004D4782FFF0000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay3Evt9B75AC]
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
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent9B75AC Run]
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
ToolTipText=Click to open calendar
[MeasureOpenEvent80A36E]
Measure=Plugin
Plugin=RunCommand
Program=#Pwsh#
Parameter=-NoProfile -ExecutionPolicy Bypass -Command "try { $app = New-Object -ComObject Outlook.Application; $item = $app.Session.GetItemFromID('000000000131F912B0B41849882D6B421E6E05720700869667498D2ACE448746378E969C4A0900000000010D0000869667498D2ACE448746378E969C4A0900012C6D86A90000'); $item.Display(); Start-Sleep -Milliseconds 500; Add-Type -name Win32 -namespace Win32Functions -memberDefinition '[DllImport(`"user32.dll`")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName); [DllImport(`"user32.dll`")] public static extern bool SetForegroundWindow(IntPtr hWnd); [DllImport(`"user32.dll`")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'; $subject = $item.Subject; if ($subject) { $hwnd = [Win32Functions.Win32]::FindWindow($null, $subject); if ($hwnd -ne [System.IntPtr]::Zero) { [Win32Functions.Win32]::ShowWindow($hwnd, 3); [Win32Functions.Win32]::SetForegroundWindow($hwnd) } } } catch { Write-Warning 'Failed to open event' }"
OutputType=ANSI
UpdateDivider=-1
[MeterDay4Evt80A36E]
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
LeftMouseUpAction=[!CommandMeasure MeasureOpenEvent80A36E Run]
[MeterContentEnd]
Meter=Shape
X=0
Y=6R
Shape=Rectangle 0,0,0,0
DynamicVariables=1

