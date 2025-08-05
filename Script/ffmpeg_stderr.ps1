using namespace HelperClasses

param(
    $PSHost, 
    $taskName, 
    $StartPaused
)

$timePattern = "time=\D*([\d\.:]+)"
$fpsPattern = "fps=\D*(\d+)"
$framePattern = "frame=\D*(\d+)"
$durationPattern = "Duration:\D*([\d\.:]+)"
$errorPattern = "Error"
$failedPattern = "failed"
$isLastError = $false
$regexOpt = [Text.RegularExpressions.RegexOptions]::IgnoreCase

[ReceivedData]$ffprobeOutput = [ReceivedData]::Empty

# [ProgressRecord] is for console progress bar
$progressRecord = New-Object System.Management.Automation.ProgressRecord(1, $taskName, 'Initialize')
$progressRecord.RecordType = [System.Management.Automation.ProgressRecordType]::Processing
$currentOperation = $syncData.path
$progressRecord.CurrentOperation = $currentOperation

$totalDuration = [TimeSpan]::Zero
$startTime = Get-Date
$ffmpegTask = $ffmpegProcess.Start()

[System.Diagnostics.Process]$proc = $null
if ($ffmpegProcess.TryGetProcess([ref]$proc)) {
    $syncData.job.AssignProcess($proc)
}

# --- 残り時間計算の平滑化用 ---
$averageFps = 0.0
$averagePps = 0.0
# 平滑化係数 (0 < alpha < 1)。小さいほど滑らかになり、過去の値を重視する。
$alpha = 0.05

$viewModel.CurrentOperation = $currentOperation

while (-not $ffmpegTask.Wait(100)) {
    #if ($cTokenSource -ne $null) { $cTokenSource.Dispose() }
    #$cTokenSource = New-Object System.Threading.CancellationTokenSource(1000)
    foreach ($receivedData in ($ffmpegProcess.ReceivedDataQueue.GetConsumingEnumerable(<#$cTokenSource.Token#>)))
    {
        switch ($receivedData.Type)
        {
            ('StdOut') { 
                [HelperClasses.ConsoleHelper]::Log($receivedData.Data)
                break
            }
            ('StdError') {
                $data = $receivedData.Data

                if ($data.Contains('frame=')) {

                    $isLastError = $false
                    if ($syncData.totalFrames -ne 0) {
                        # Calculate progress from frame.
                        $matchFramePettern = [Regex]::Match($data, $framePattern)
                        $matchFpsPettern = [Regex]::Match($data, $fpsPattern)

                        if ($matchFramePettern.Success -and $matchFpsPettern.Success) {

                            $frame = [double]::Parse($matchFramePettern.Groups[1].Value)
                            $currentFps = [Double]::Parse($matchFpsPettern.Groups[1].Value)

                            $percentComplete = ($frame / $syncData.totalFrames) * 100.0
                            
                            if ($currentFps -gt 1.0) { 

                                # 指数移動平均 (Exponential Moving Average) を使ってFPSを平滑化し、予測の安定性を向上させる
                                if ($averageFps -eq 0.0) {
                                    $averageFps = $currentFps
                                } else {
                                    $averagePps = ($currentFps * $alpha) + ($averageFps * (1 - $alpha))
                                }
                                $remainingTime = ($syncData.totalFrames - $frame) / $averageFps

                            } else{ # 1fps未満の場合は完了パーセンテージをもとに計算
                                #pps : percent per second
                                $pps = $percentComplete / (((Get-Date) - $startTime).TotalMilliseconds / 1000.0)
                                if ($pps -gt 0) {
                                    $remainingTime = (100.0 - $percentComplete) / $pps
                                }
                            }                  
                        }

                    } elseif ($totalDuration.Ticks -ne 0) { #Frameが取得できないときのFallback
                        # Calculate progress from time.
                        $match = [Regex]::Match($data, $timePattern)
                        if ($match.Success) {
                            $time = [TimeSpan]::Parse($match.Groups[1].Value)
                            $percentComplete = ($time.Ticks / $totalDuration.Ticks) * 100.0

                            # Calculate estimated time remaining.
                            $currentPps = $percentComplete / (((Get-Date) - $startTime).TotalMilliseconds / 1000.0)
                            $averagePps = if ($averagePps -eq 0.0) {
                                $currentPps
                            } else {
                                ($currentPps * $alpha) + ($averagePps * (1 - $alpha))
                            }

                            if ($pps -gt 0) {
                                $remainingTime = (100.0 - $percentComplete) / $averagePps
                            }
                        }
                    }

                    # [ProgressRecord] is for displaying a progress bar on the console screen.
                    $progressRecord.StatusDescription = $data
                    $progressRecord.PercentComplete = $percentComplete
                    if ($null -ne $remainingTime) {
                        $progressRecord.SecondsRemaining = $remainingTime
                    }
                    if ($syncData.showConsoleProgress) {
                        $PSHost.UI.WriteProgress($progressRecord.ActivityId, $progressRecord)
                        $PSHost.UI.RawUI.WindowTitle = "$($progressRecord.PercentComplete)% $taskName"
                    }

                    # Set progress values in the ViewModel of the GUI window
                    $viewModel.StatusDescription = $data
                    $viewModel.Progress = $percentComplete
                    $viewModel.ProgressRemaining = [TimeSpan]::FromSeconds($remainingTime)
                    $viewModel.WindowTitle = "$($progressRecord.PercentComplete)% $taskName"

                    if ($StartPaused) {
                        [ConsoleHelper]::Info("The process started in paused state")
                        $StartPaused = $false
                        $viewModel.BusyMessage = "On pause."
                        $viewModel.ProcessControlCommand.Execute($true)
                    } elseif((0 -lt $percentComplete) -and ($syncData.termination -eq $false)) {
                        
                        $viewModel.Busy = $false

                        # If progress is 100%, set [RecordType] to Completed.
                        if (100 -gt $percentComplete) {
                        
                            # Set [ProgressState] according to progress
                            $viewModel.ProgressState = [ProgressWindow.ProgressState]::Normal
                        } else {
                            #$viewModel.ProgressState = [ProgressWindow.ProgressState]::Completed

                            # and [ProgressRecord.RecordType]
                            $progressRecord.RecordType = [System.Management.Automation.ProgressRecordType]::Completed
                            $PSHost.UI.WriteProgress($progressRecord.ActivityId, $progressRecord)
                        }
                    }

                } elseif ($data.Contains("Duration:")) {

                    $match = [Regex]::Match($data, $durationPattern)
                    if ($match.Success) {
                        $totalDuration = [TimeSpan]::Parse($match.Groups[1].Value)
                    }

                # Processing other than frame informations
                } else {

                    if (([Regex]::Match($data, $errorPattern, $regexOpt).Success) -and ($isLastError -eq $false)) {

                        $isLastError = $true
                        $message = $(if ($data.Length -gt 45) {
                            ($data.Substring(0, 40) + "...")
                        } else {
                            $data
                        })

                        $viewModel.BusyMessage = "Errors detected : $message"
                        $viewModel.Busy = $true
                        [ConsoleHelper]::Error($data)

                    } elseif (([Regex]::Match($data, $failedPattern, $regexOpt).Success) -and ($isLastError -eq $false)) {

                        $isLastError = $true
                        $message = $(if ($data.Length -gt 45) {
                            ($data.Substring(0, 40) + "...")
                        } else {
                            $data
                        })

                        $viewModel.BusyMessage = $message
                        $viewModel.Busy = $true
                        [ConsoleHelper]::Error($data)

                    } else {
                        [ConsoleHelper]::WriteLine($data)
                    }
                }
                break
            }
        }

        # check ffprobe data
        if (($ffprobeProcess.ReceivedDataQueue.Count -gt 0) -and $ffprobeProcess.ReceivedDataQueue.TryTake([ref]$ffprobeOutput)) {
            if ($ffprobeOutput.Type -eq 'StdOut') {

                $match = [Regex]::Match($ffprobeOutput.Data, "r_frame_rate=(\d+)(/\d+)?")
                if ($match.Success) {
                    if ($match.Groups[1].Success) {
                        $syncData.framerateNum = [Double]::Parse($match.Groups[1].Value)
                    }

                    if ($match.Groups[2].Success) {
                        $syncData.framerateDen = [Double]::Parse($match.Groups[2].Value.Trim('/'))
                    }
                }

                $match = [Regex]::Match($ffprobeOutput.Data, "duration=([\d\.]+)")
                if ($match.Success) {
                    $syncData.duration = [Double]::Parse($match.Groups[1].Value)
                }

                if (($syncData.duration -ne 0) -and ($syncData.framerateDen -ne 0) -and ($syncData.framerateNum -ne 0)) {

                    [ConsoleHelper]::Log("Duration : $($syncData.duration)")
                    [ConsoleHelper]::Log("Frame Rate : $($syncData.framerateNum) / $($syncData.framerateDen)")
                    $syncData.totalFrames = ($syncData.framerateNum / $syncData.framerateDen) * $syncData.duration
                }
            }
        }
    }
}

$syncData.exitCode = $ffmpegTask.GetAwaiter().GetResult()
$ffmpegProcess.Dispose()

[ConsoleHelper]::Log("FFMPEG EXIT CODE : $($syncData.exitCode)")

if (($syncData.exitCode -eq 0) -and ($syncData.termination -eq $false)) {
    
    # --- たまに99.9%でプロセスが終了してしまうようなので対策 ※要調査
    $viewModel.Progress = 100.0
    # -----------------------------------------------------------------
    
    if ($viewModel.AutoPlay) {
        $syncData.openfile.Invoke($syncData.output)
    }
    if ($viewModel.OpenExplorer) {
        $syncData.openExplorer.Invoke($syncData.output)
    }
} else {
    $viewModel.ProgressState = [ProgressWindow.ProgressState]::None
}

if (($viewModel.AutoClose -eq $true) -or ($syncData.termination -eq $true)) {
    Start-Sleep 1
    $closing = $progressWindow.Close()
    if (-not ($closing.Wait(1000))) {
        [ConsoleHelper]::Log("Lost control of the GUI window.", 1)
        $syncData.exitCode = 1003
    }
}

foreach ($e in $Error) {
    if ($e.Exception -isnot [System.OperationCanceledException]) {
        [ConsoleHelper]::Error($e, 1)
        $syncData.exitCode = 1
    }
}