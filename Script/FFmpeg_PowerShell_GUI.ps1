##
# FFMPEG PowerShell GUI Script
#
<#
    .SYNOPSIS
    FFmpegの動画エンコード作業を簡略化するための、グラフィカルユーザーインターフェース(GUI)を提供します。

    .DESCRIPTION
    このPowerShellスクリプトはFFmpegのラッパーとして機能し、動画エンコードをより簡単に行うためのユーザーフレンドリーなGUIを提供します。
    主な機能は以下の通りです:
    - 入力ファイルのドラッグ＆ドロップに対応。入力ファイルがパラメータで指定されない場合、ドロップ用のウィンドウが表示されます。
    - エンコードのステータス、進捗率、残り時間などをリアルタイムで表示するプログレスウィンドウ。
    - エンコード処理の一時停止、再開、中止をGUIから制御可能。
    - デフォルトまたはカスタムのFFmpegパラメータをサポート。
    - パラメータ文字列内の`[INPUT]`や`[OUTPUT]`といったプレースホルダーを自動的に置換。
    - 出力ファイルが既に存在する場合に、上書き、名前の変更、またはキャンセルを選択するプロンプトを表示。
    - エンコード中にシステムがスリープするのを防ぐ機能。
    - 複数のファイルを一度にドロップした場合、ファイルごとに個別のプロセスを起動して処理。
    - 必要なC#ヘルパーアセンブリをオンザフライで自動的にコンパイル。

    .PARAMETER path
    エンコードする入力ファイルのパスを指定します。
    このパラメータを省略した場合、ファイルをドラッグ＆ドロップするためのウィンドウが表示されます。

    .PARAMETER Parameters
    FFmpegに渡すコマンドラインパラメータを単一の文字列として指定します。
    省略された場合、スクリプトに定義されているデフォルトのパラメータセットが使用されます。
    プレースホルダー `[INPUT]` は入力ファイルパスに、`[OUTPUT]` は出力ファイルパスに置換されます。
    また、`[OUTPUT(.ext)]` の形式で出力ファイルの拡張子を指定することもできます。

    .PARAMETER OutputSourceDirectory
    このスイッチを指定すると、出力ファイルが入力ファイルと同じディレクトリに作成されます。
    指定しない場合、スクリプトで定義されたデフォルトの出力ディレクトリに保存されます。

    .PARAMETER StartPaused
    このスイッチを指定すると、エンコードプロセスが一時停止した状態で開始されます。GUIの再開ボタンをクリックしてエンコードを開始する必要があります。

    .PARAMETER ForceCompileAssembly
    スクリプトが使用するC#ヘルパーアセンブリを強制的に再コンパイルします。

    .EXAMPLE
    # 入力ファイルを指定して、デフォルト設定でエンコードを開始する
    .\FFmpeg_PowerShell_GUI.ps1 -path "C:\videos\input.mp4"

    # カスタムのFFmpegパラメータを指定してエンコードする
    .\FFmpeg_PowerShell_GUI.ps1
#>


using namespace System.Text.RegularExpressions;
using namespace ViewModelHelper
using namespace HelperClasses

param(
    [Parameter()]
    [Alias("input")] [string] $path,
    [Parameter()]
    [string] $Parameters,
    [Parameter()]
    [switch] $OutputSourceDirectory = $false,
    [Parameter()]
    [switch] $StartPaused,
    [Parameter()]
    [switch] $ForceCompileAssembly = $false,
    [Parameter()]
    [int] $ProcWeight = 5
)

Set-Location -LiteralPath $PSScriptRoot

# ------------------------------------
# Internal parameter variables
# ------------------------------------
$global:path = $path.trim("`'")
$global:parameters = $Parameters

$OUTPUT_DIRECTORY = "D:\Encode"
$OPTION_DIRECTORY = ""
$OUTPUT_EXTENSION = ".mkv"

$FFMPEG_PARAMETERS = @'
-hide_banner
-ignore_unknown
-y
-i [INPUT]
-analyzeduration 20M -probesize 20M
-c:v libx265
-maxrate 20M
-bufsize 20M
-preset medium
-crf 18
-bf 2
-movflags +faststart
-pix_fmt yuv420p
-c:a aac -ab 256000 -af "channelmap=channel_layout=stereo,aresample=48000:resampler=soxr"
[OUTPUT]
'@

$INPUT_PATTERN = "\[INPUT\]"
$OUTPUT_WITH_EXT_PATTERN = "\[OUTPUT\((.*\..+)\)\]"
$OUTPUT_PATTERN = "\[OUTPUT\]"
$PROCESSOR_COUNT_PATTERN = "\[PROCESSOR_COUNT\]"
$REGEX_OPT = [Text.RegularExpressions.RegexOptions]::IgnoreCase

$TASK_NAME = "$([System.IO.Path]::GetFileName($global:path)) - $($myInvocation.MyCommand.name)"
$AUTO_CLOSE_GUI_WINDOW = $false
$AUTO_PLAY_ENCODED = $false
$OPEN_FOLDER_ENCODED = $false
$PREVENT_SLEEP = $true
$PREVENT_SLEEP_STATE = [UInt32]0x00000002 # $ES_SYSTEM_REQUIRED = [UInt32]0x00000002 : , $ES_DISPLAY_REQUIRED = [UInt32]0x00000002,
$SHOW_CONSOLE_PROGRESSBAR = $false
$ENABLE_ACTIVE_ANIMATION = $true
$FFMPEG_FILE = "ffmpeg.exe"
$FFPROBE_FILE = "ffprobe.exe"
$RUNSPACE_SCRIPT = ".\ffmpeg_stderr.ps1"

$NO_REDIRECT = $false

$OpenFileScript = {
    param([string]$filePath)
    if ((Test-Path -LiteralPath $filePath)) {
        Invoke-Item -LiteralPath $filePath
    }
}

$OpenExplorerScript = {
    param([string]$filePath)
    if ((Test-Path -LiteralPath $filePath)) {
        Start-Process "explorer.exe" ('/select,"{0}"' -f ($filePath))
    }
}

# ------------------------------------
# Helper functions
# ------------------------------------

# Resolve output file path
function ResolveOutputPath {
    param([string]$path, [string]$extension, [string]$folder, [string]$optionFolder)

    if (([String]::IsNullOrWhiteSpace($folder))) {

        # If [$folder] is empty and the extension is different from [$path]'s extenstion,
        # only change the extension.
        if ([System.IO.Path]::GetExtension($path) -ne $extension) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($path)
            $fi = New-Object System.IO.FileInfo($path)
            return ([System.IO.Path]::Combine($fi.DirectoryName, ($baseName + $extension)))
        } elseif ([String]::IsNullOrWhiteSpace($optionFolder))  {
            $folder = [System.IO.Path]::GetFileNameWithoutExtension($path) + "_encode"
        } else {
            $folder = $optionFolder
        }
    }

    [System.IO.DirectoryInfo]$di = $(if ([System.IO.Path]::IsPathRooted($folder)) {
        New-Object System.IO.DirectoryInfo($folder)
    } else {
        New-Object System.IO.DirectoryInfo((Join-Path ([System.IO.Path]::GetDirectoryName($path)) $folder))
    })

    if (-not ($di.Exists)) { $di.Create() }

    $result = Join-Path $di.FullName ([System.IO.Path]::GetFileNameWithoutExtension($path) + $extension)
    return $result
}

# Show file drop window
# Return whether the window was manually closed
function showDropWindow {
    param([string]$caption)

    $DropWindowViewModel = New-Object DropWindow.DropWindowViewModel
    $DropWindowViewModel.Title = $captions

    $dropFilesCommand = New-Object DelegateCommand
    $dropFilesCommand.ExecuteHandler = {
        param($param)
        
        $fileList = $param.Data.GetFileDropList()

        switch ($fileList.Count){
            0 { break }
            1
            {
                $global:path = $fileList[0]
                $FileDropWindow.DialogResult = $true
                break
            }
            Default
            {   
                $global:ProcessManagerMode = $true
                    $proc = Start-Process powershell.exe -PassThru -ArgumentList "-ExecutionPolicy RemoteSigned -File `"$($MyInvocation.ScriptName)`" -path `"$f`" -Parameters `"$($global:parameters)`" -StartPaused"
                foreach ($f in $fileList) {
                    $proc = Start-Process powershell.exe -PassThru -ArgumentList "-ExecutionPolicy RemoteSigned -File `"$($MyInvocation.ScriptName)`" -path `"$f`" -Parameters `"$($global:parameters)`" -StartPaused"
                    Start-Sleep -Milliseconds 200
                    $ProcessList.Add($proc)
                }
                $FileDropWindow.DialogResult = $true
                break
            }
        }
    }
    $DropWindowViewModel.DropFilesCommand = $dropFilesCommand

    $FileDropWindow = New-Object DropWindow.MainWindow($DropWindowViewModel)
    $FileDropWindow.Icon = $iconFrame
    return $FileDropWindow.ShowDialog()
}

# Check [$path] parameter
# Return true if the file exists
function checkFilePath {
    param([string]$filePath)

    if ([String]::IsNullOrEmpty($filePath)) { return $false }

    if (-not(Test-Path -LiteralPath $filePath)){
        Write-Host "`"$filePath`" file does not exist." -ForegroundColor Yellow
        return $false
    }

    if ((Get-Item -LiteralPath $filePath).PSIsContainer) {
        Write-Host "`"$filePath`" is a directory." -ForegroundColor Yellow
        return $false
    }

    return $true
}

# ------------------------------------
# Main execution
# ------------------------------------
# Clear the error
$Error.Clear()

[Console]::WriteLine("Start time: {0}" -f ([DateTime]::Now))


# ------------------------------------
# Check required external files
# for the execution
# check C# .NET Framework source code, ffmpeg.exe exists
# ffprobe.exe is option
# ------------------------------------


# ============================================================================ #
# Check assemblies and compile C# source code.
Function Test-Assembly {
    param($sourcePath, $assemblyPath, $refAssemblies, $classNameToBeVerified, $forceCompileAssembly = $false)
    # Check if [$CS_SOURCE] exist and updated.
    if ((Test-Path $sourcePath) -and (Test-Path $assemblyPath)) {
        if ((Get-ItemProperty $sourcePath).LastWriteTime -gt (Get-ItemProperty $assemblyPath).LastWriteTime) {
            Write-Host ("Since [$sourcePath] has been updated, recompilation is required.")
            $forceCompileAssembly = $true
        }
    }
    # Compile [$CS_SOURCE] if needed.
    Try {
        if ($forceCompileAssembly) {
            throw [System.Management.Automation.RuntimeException] "Request compile assembly.(1)"
        }
        
        if (Test-Path $assemblyPath) {
            $asm = [Reflection.Assembly]::LoadFile((Resolve-Path $assemblyPath))
            $null = $asm.GetType($classNameToBeVerified, $true, $true) # 第2引数: throwOnError=false, 第3引数: ignoreCase=true
        } else {
            throw [System.Management.Automation.RuntimeException] "Request compile assembly.(2)"
        }
    } Catch [System.Management.Automation.RuntimeException] {
        Write-Host "$($Error.Exception.message) Compile [$sourcePath] to [$assemblyPath]." -ForegroundColor Yellow

        for ($i = 0; $i -lt $refAssemblies.Count; $i++) {
            if (Test-Path -LiteralPath $refAssemblies[$i]) {

                $refAssemblies[$i] = Resolve-Path -LiteralPath $refAssemblies[$i]
                Write-Host $refAssemblies[$i]
            }
        }

        $Error.Clear()
        $null = Add-Type -Path $sourcePath -OutputAssembly $assemblyPath -ReferencedAssemblies $refAssemblies -ErrorAction Stop -PassThru
    }
}


# ============================================================================ #
# DevLog
$CS_SOURCE = "assemblies\DevLog.cs"
$CS_ASSEMBLY = "assemblies\DevLog.dll"
$FORCE_COMPILE_ASSEMBLY = $false
$REF_ASSEMBLIES = @('System.Management.Automation', 'System.Runtime.InteropServices')
$CLASS_NAME_TO_BE_VERIFIED = "DevLog"
Test-Assembly -sourcePath $CS_SOURCE -assemblyPath $CS_ASSEMBLY -refAssemblies $REF_ASSEMBLIES -classNameToBeVerified $CLASS_NAME_TO_BE_VERIFIED -forceCompileAssembly $FORCE_COMPILE_ASSEMBLY

# ============================================================================ #
# HelperClasses
$CS_SOURCE = "assemblies\helperclasses.cs"
$CS_ASSEMBLY = "assemblies\helperclasses.dll"
$FORCE_COMPILE_ASSEMBLY = $false
$REF_ASSEMBLIES = @(
    'WindowsBase', 
    'PresentationFramework', 
    'PresentationCore', 
    'System.Xaml', 
    'System.Runtime.InteropServices',
    'assemblies\DevLog.dll'
)
$CLASS_NAME_TO_BE_VERIFIED = "HelperClasses.ConsoleHelper"
Test-Assembly -sourcePath $CS_SOURCE -assemblyPath $CS_ASSEMBLY -refAssemblies $REF_ASSEMBLIES -classNameToBeVerified $CLASS_NAME_TO_BE_VERIFIED -forceCompileAssembly $FORCE_COMPILE_ASSEMBLY

# ============================================================================ #
# ThemeHelper
$CS_SOURCE = "assemblies\themehelper.cs"
$CS_ASSEMBLY = "assemblies\themehelper.dll"
$FORCE_COMPILE_ASSEMBLY = $false
$REF_ASSEMBLIES = @(
    'WindowsBase', 
    'PresentationFramework', 
    'PresentationCore', 
    'System.Xaml', 
    'System.Runtime.InteropServices'
)
$CLASS_NAME_TO_BE_VERIFIED = "ThemeHelper.ThemeBase"
Test-Assembly -sourcePath $CS_SOURCE -assemblyPath $CS_ASSEMBLY -refAssemblies $REF_ASSEMBLIES -classNameToBeVerified $CLASS_NAME_TO_BE_VERIFIED -forceCompileAssembly $FORCE_COMPILE_ASSEMBLY

# ============================================================================ #
# ViewModelHelper
$CS_SOURCE = "assemblies\viewmodelhelper.cs"
$CS_ASSEMBLY = "assemblies\viewmodelhelper.dll"
$FORCE_COMPILE_ASSEMBLY = $false
$REF_ASSEMBLIES = @('WindowsBase', 'PresentationFramework', 'PresentationCore', 'System.Xaml', 'System.Runtime.InteropServices')
$CLASS_NAME_TO_BE_VERIFIED = "ViewModelHelper.ViewModelBase"
Test-Assembly -sourcePath $CS_SOURCE -assemblyPath $CS_ASSEMBLY -refAssemblies $REF_ASSEMBLIES -classNameToBeVerified $CLASS_NAME_TO_BE_VERIFIED -forceCompileAssembly $FORCE_COMPILE_ASSEMBLY

# ============================================================================ #
# ProgressWindow
$CS_SOURCE = "assemblies\progresswindow.cs"
$CS_ASSEMBLY = "assemblies\progresswindow.dll"
$FORCE_COMPILE_ASSEMBLY = $false
$REF_ASSEMBLIES = @(
    'assemblies\helperclasses.dll', 
    "assemblies\themehelper.dll",
    'assemblies\viewmodelhelper.dll', 
    'WindowsBase', 
    'PresentationFramework', 
    'PresentationCore', 
    'System.Xaml', 
    'System.Runtime.InteropServices'
)
$CLASS_NAME_TO_BE_VERIFIED = "ProgressWindow.MainWindow"
Test-Assembly -sourcePath $CS_SOURCE -assemblyPath $CS_ASSEMBLY -refAssemblies $REF_ASSEMBLIES -classNameToBeVerified $CLASS_NAME_TO_BE_VERIFIED -forceCompileAssembly $FORCE_COMPILE_ASSEMBLY

# ============================================================================ #
# dropwindow
$CS_SOURCE = "assemblies\dropwindow.cs"
$CS_ASSEMBLY = "assemblies\dropwindow.dll"
$FORCE_COMPILE_ASSEMBLY = $false
$REF_ASSEMBLIES = @(
    'assemblies\viewmodelhelper.dll',
    'WindowsBase', 
    'PresentationFramework', 
    'PresentationCore', 
    'System.Xaml', 
    'System.Runtime.InteropServices' 
)
$CLASS_NAME_TO_BE_VERIFIED = "DropWindow.MainWindow"
Test-Assembly -sourcePath $CS_SOURCE -assemblyPath $CS_ASSEMBLY -refAssemblies $REF_ASSEMBLIES -classNameToBeVerified $CLASS_NAME_TO_BE_VERIFIED -forceCompileAssembly $FORCE_COMPILE_ASSEMBLY

# ============================================================================ #
# JobHelper
$CS_SOURCE = "assemblies\jobhelper.cs"
$CS_ASSEMBLY = "assemblies\jobhelper.dll"
$FORCE_COMPILE_ASSEMBLY = $false
$REF_ASSEMBLIES = @(
    'System.Runtime.InteropServices'
)
$CLASS_NAME_TO_BE_VERIFIED = "JobHelper.JobCpuWeightController"
Test-Assembly -sourcePath $CS_SOURCE -assemblyPath $CS_ASSEMBLY -refAssemblies $REF_ASSEMBLIES -classNameToBeVerified $CLASS_NAME_TO_BE_VERIFIED -forceCompileAssembly $FORCE_COMPILE_ASSEMBLY


######     From here, use [ConsoleHelper] instead of [Write-Host].     ######

$Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.name
[DevLog]::WriteLine("Debugger is active.")



# Load icon
$icoPath = Join-Path (Get-Location) "icon.ico"
[System.Windows.Media.Imaging.BitmapFrame]$iconFrame = $null
if (Test-Path $icoPath) {
    $uri = [System.Uri]::new($icoPath, [System.UriKind]::Absolute)
    $iconFrame = [System.Windows.Media.Imaging.BitmapFrame]::Create($uri)
    $iconFrame | Format-List PixelWidth, PixelHeight
}


# Load custom fonts
try {
    $interFont = [FontHelper]::LoadFont("Inter", (Resolve-Path -LiteralPath "Inter-VariableFont_opsz,wght.ttf"))
    [ProgressWindow.Theme]::I.LogoFontFamily = $interFont
    #[ProgressWindow.Theme]::I.MonospaceFontFamily =$interFont
} catch {
    [DevLog]::WriteLine($_)
}

try {
    $notoFont = [FontHelper]::LoadFont("Noto Sans JP", (Resolve-Path -LiteralPath "NotoSansJP-VariableFont_wght.ttf"))
    [ProgressWindow.Theme]::I.MainFontFamily = $notoFont
} catch {
    [DevLog]::WriteLine($_)
}




# check ffmpeg.exe
[ConsoleHelper]::WriteLine("######    FFmpeg version    ######", 1, 1)
$checkFfmpegProc = $null
try {
    $checkFfmpegProc = Start-Process $FFMPEG_FILE ('-version') -NoNewWindow -PassThru
} catch {
    [ConsoleHelper]::Error("Error : ffmpeg.exe was not found.", 1, 5)


    exit 2
}
$global:ProcessManagerMode = $false

# Show file drop dialog.

while ((checkFilePath $global:path) -ne $true) {


    $result = showDropWindow $myInvocation.MyCommand.name

    if ($result -ne $true) {
        [ConsoleHelper]::Error("Script execution has been aborted.", 1)
        Start-Sleep 2
        exit 0
    }
    
    if ($global:ProcessManagerMode -eq $true) {
        [ConsoleHelper]::Log("Multiple file mode", 1)
        exit 0
    }
}

# check ffmpeg parameters
if (-not ([String]::IsNullOrWhiteSpace($Parameters))) {
    $FFMPEG_PARAMETERS = $Parameters
}

# create parameter(parameter of ffmpeg) object
$paramsObj = New-Object HelperClasses.CommentOutText($FFMPEG_PARAMETERS)

# check keyword of extension replacer
$matchOutputWithExt = $paramsObj.Match($OUTPUT_WITH_EXT_PATTERN)
if ($matchOutputWithExt.Success) {
    $OUTPUT_EXTENSION = $matchOutputWithExt.Groups[1].Value
}

# Resolve path for the output file
if ([String]::IsNullOrEmpty($global:output)) {
    if ($OutputSourceDirectory) {
        $OUTPUT_DIRECTORY = [System.IO.Path]::GetDirectoryName($global:path)
    }
    $global:output = ResolveOutputPath $global:path $OUTPUT_EXTENSION $OUTPUT_DIRECTORY $OPTION_DIRECTORY
}

if ($null -ne $checkFfmpegProc) {
    $checkFfmpegProc.WaitForExit(1000)
}
[ConsoleHelper]::Log("Input file  : $global:path", 2)
[ConsoleHelper]::Log("Output file : $global:output", 1, 1)

# check overwrite option
$matchOverwrite = [Regex]::Match($paramsObj.GetText(" "), "\s-y\s", $REGEX_OPT)
if ((-not $matchOverwrite.Success) -and (Test-Path -LiteralPath $global:output)) {
    [ConsoleHelper]::WriteLine("Already exists in the destination path.")
    [ConsoleHelper]::WriteLine("Do you want to overwrite it? Y(Yes) / N(No) / S(Suffix) :")
    do {
        $answer = Read-Host
    } while (-not ($answer -match "n|no" -or $answer -match "y|yes" -or $answer -match "s|suffix"))

    if ($answer -match "n|no") {
        exit 0
    } elseif ($answer -match "s|suffix") {
        $i = 0
        do {
            $i += 1
            $fi = New-Object System.IO.FileInfo($global:output)
            $global:output = [System.IO.Path]::Combine($fi.DirectoryName, ([System.IO.Path]::GetFileNameWithoutExtension($fi.Name) + "_$i" + $fi.Extension))
        } while ((Test-Path -LiteralPath $global:output))
        [ConsoleHelper]::Log(" -> Output file : $global:output", 1, 1)
    }
}

# replace ffmpeg parameters
$replacedParams = $paramsObj

$replacedParams.Replace($INPUT_PATTERN, "`"$($global:path)`"", $REGEX_OPT)
if ($matchOutputWithExt.Success) {
    $replacedParams.Replace($OUTPUT_WITH_EXT_PATTERN, "`"$($global:output)`"", $REGEX_OPT)
} else {
    $replacedParams.Replace($OUTPUT_PATTERN, "`"$($global:output)`"", $REGEX_OPT)
}
$replacedParams.Replace($PROCESSOR_COUNT_PATTERN, [System.Environment]::ProcessorCount, $REGEX_OPT)

[ConsoleHelper]::Info("######    Encode parameters    ######", 1)
[ConsoleHelper]::Info($replacedParams.GetText("`n"), 1, 3)

# check [$Error]
if ($Error.Count -gt 0) {
    $Error
    exit 1
}

#-nostdin Quitのために消したが、何のためにつけたのか
# → ffmpeg をバックグラウンドのプロセスで実行する時は -nostdinオプションを追加する。Invalid data found when processing input対策？
$ffmpegProcess = New-Object HelperClasses.ProcessInfo($FFMPEG_FILE, ("-y $($replacedParams.GetText(" "))"), $NO_REDIRECT)

# CPU使用率上限設定
$job = New-Object JobHelper.JobCpuWeightController 500

$ffprobeParams = @(
    '-v error',
    '-select_streams v:0',
    '-show_entries',
    'stream=r_frame_rate,duration',
    '-of default=nw=1',
    ('"{0}"' -f $global:path)
)
$ffprobeProcess = New-Object HelperClasses.ProcessInfo($FFPROBE_FILE, ($ffprobeParams -join ' '))
$ffprobeTask = $ffprobeProcess.Start()

$syncData = [HashTable]::Synchronized(@{
    path = $global:path
    output = $global:output
    framerateNum = [Double]0
    framerateDen = [Double]0
    duration = [Double]0
    totalFrames = [Double]0
    exitCode = 1
    openfile = $OpenFileScript
    openExplorer = $OpenExplorerScript
    showConsoleProgress = $SHOW_CONSOLE_PROGRESSBAR
    termination = $false
    previousState = $null
    job = $job
})

# Create ViewModel of progress window.
$viewModel = New-Object ProgressWindow.ProgressViewModel

$viewModel.CpuLimitChangeCommand = New-Object DelegateCommand {
    param($param)

    $newLimit = [int]($param)
    if ($newLimit -ge 1 -and $newLimit -le 9) {
        
        #プロセスに上限を設定
        $syncData.job.RequestWeight([int]($newLimit))
        #[ConsoleHelper]::WriteLine("CPU Usage $($syncData.job.GetCpuUsage($true))")
    }
}

# Create DelegateCommand for command bindings
$processControlCommand = New-Object DelegateCommand
$processControlCommand.ExecuteHandler = {
    param($param)

    if ([bool]$param) {
        $viewModel.BusyMessage = "On pause."
        $viewModel.Busy = $true
        $viewModel.ProgressState = [ProgressWindow.ProgressState]::Paused
        $ffmpegProcess.Suspend()
    } else {
        $viewModel.ProgressState = [ProgressWindow.ProgressState]::Normal
        $ffmpegProcess.Resume()
    }
}
$viewModel.ProcessControlCommand = $processControlCommand

$processExitCommand = New-Object DelegateCommand
$processExitCommand.ExecuteHandler = {
    param($param)

    [ConsoleHelper]::Log("The process is currently shutting down.", 1)

    $syncData.termination = $true
    $viewModel.ProgressState = [ProgressWindow.ProgressState]::None
    $viewModel.BusyMessage = "Shutting down."
    $viewModel.Busy = $true
    $progressWindow.DoEvents()

    [System.Diagnostics.Process]$proc = $null
    if ($ffmpegProcess.TryGetProcess([ref]$proc)) {

        if ((-not $proc.HasExited) -and $proc.StandardInput -ne $null) {
            $proc.StandardInput.WriteLine("q") # Send q as an input to the ffmpeg process window making it stop.

            if ($ffmpegProcess.IsSuspended) {
                $ffmpegProcess.Resume()
            }
            return
        }
    }

    $progressWindow.Close()
}
$viewModel.ProcessExitCommand = $processExitCommand

$openFolderCommand = New-Object DelegateCommand
$openFolderCommand.ExecuteHandler = {
    $OpenExplorerScript.Invoke($syncData.output)
}
$viewModel.OpenFolderCommand = $openFolderCommand

$openFileCommand = New-Object DelegateCommand
$openFileCommand.ExecuteHandler = {
    $OpenFileScript.Invoke($syncData.output)
}
$viewModel.OpenFileCommand = $openFileCommand

$changeExecutionStateCommand = New-Object DelegateCommand
$changeExecutionStateCommand.ExecuteHandler = {
    param($param)

    if ($viewModel.PreventSleep -eq $true) {
        [ThreadHelper]::PreventSleep($PREVENT_SLEEP_STATE)
        [ConsoleHelper]::Log("Prevent sleep mode enabled.")
    } else {
        [ThreadHelper]::AllowSleep()
        [ConsoleHelper]::Log("Prevent sleep mode disabled.")
    }
}
$viewModel.ChangeExecutionStateCommand = $changeExecutionStateCommand

# Some commands are fired at the value setter, so the property is set after the command preparation
$viewModel.CurrentOperation = "Preparing."
$viewModel.ProgressLabel = "-> $([System.IO.Path]::GetFileName($global:output))"
$viewModel.WindowTitle = $TASK_NAME
$viewModel.AutoClose = $AUTO_CLOSE_GUI_WINDOW
$viewModel.AutoPlay = $AUTO_PLAY_ENCODED
$viewModel.OpenExplorer = $OPEN_FOLDER_ENCODED
$viewModel.PreventSleep = $PREVENT_SLEEP
$viewModel.EnableActiveAnimation = $ENABLE_ACTIVE_ANIMATION

$viewModel.CpuRateLimit = $ProcWeight

try{
    $progressWindow = New-Object ProgressWindow.MainWindow($viewModel)
    $progressWindow.Icon = $iconFrame
} catch {
    $Error
    exit 1
}

# ------------------------------------
# Runspace execution
# ------------------------------------
# Handle stderr output from the ffmpeg process

# Create and setup runspace
$Runspace = [RunSpaceFactory]::CreateRunspace($Host)
$Runspace.ApartmentState = "STA"
$Runspace.Open()
$Runspace.SessionStateProxy.setVariable("ffmpegProcess", $ffmpegProcess)
$Runspace.SessionStateProxy.setVariable("ffprobeProcess", $ffprobeProcess)
$Runspace.SessionStateProxy.setVariable("progressWindow", $progressWindow)
$Runspace.SessionStateProxy.setVariable("syncData", $syncData)
$Runspace.SessionStateProxy.setVariable("viewModel", $viewModel)

$externalScript = Get-Command (Resolve-Path -LiteralPath $RUNSPACE_SCRIPT)
$PowerShell = [PowerShell]::Create()
$PowerShell.AddCommand($externalScript).AddArgument($Host).AddArgument($TASK_NAME).AddArgument($StartPaused)
$PowerShell.Runspace = $Runspace

$IASyncResult = $PowerShell.BeginInvoke()

# Show WPF window
$result = $progressWindow.ShowDialog();

if (($Error.Count -gt 0) -or ($result -ne $true)) {
    $processExitCommand.Execute($null)
}

try {
    if ($IASyncResult.AsyncWaitHandle.WaitOne()) {
        $PowerShell.EndInvoke($IASyncResult)
    }
} finally {
    $PowerShell.Dispose()
}

$syncData.job.Dispose()

if (($syncData.exitCode -eq 0) -and (Test-Path -LiteralPath ($global:output)) -and ($Error.Count -eq 0)) {
    Start-Sleep 1
    exit 0
} else {
    exit 1
}
