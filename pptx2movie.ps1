<#
.SYNOPSIS
PowerPointプレゼンテーションからノートを読み上げ音声付きの動画を生成します。

.DESCRIPTION
このスクリプトは、指定されたPowerPointファイル(.pptx)を処理します。
音声合成エンジンとして、Azure Text-to-Speech または VOICEVOX を選択できます。
.envファイルに VOICEVOX_SPEAKER_NAME を設定するとVOICEVOXが使用され、設定しない場合はAzureが使用されます。

VOICEVOXを使用する場合、.envで指定された実行ファイルを自動で起動し、話者IDを検索します。
スクリプト終了時にVOICEVOXのプロセスは自動で終了します。

実行ごとにタイムスタンプ付きのフォルダを`assets`内に作成し、そこに中間ファイルを保存します。
最後にすべてのスライド動画を結合し、スクリプトと同じ階層に最終的なMP4ファイルを生成します。

ノート欄で `(voice:...)` のように記述することで、スライドごとに音声（話者）を変更できます。
Azureの場合: `(voice:ja-JP-KeitaNeural)`
VOICEVOXの場合: `(voice:8)` (IDでの直接指定)

.PARAMETER PptxPath
処理対象のPowerPointファイル（.pptx）のパスを指定します。
このパラメータを省略した場合、ファイル選択ダイアログが表示されます。

.EXAMPLE
# ファイル選択ダイアログを開いて処理を開始
.\pptx2movie.ps1

.EXAMPLE
# 特定のファイルを指定して処理
.\pptx2movie.ps1 -PptxPath "C:\path\to\presentation.pptx"

.NOTES
- 実行には `ffmpeg.exe` が必要です。場所は.envで指定できます。
- 実行には `.env` ファイルに有効なサービス設定が必要です。
- PowerPointのCOMオブジェクトを使用するため、実行環境にMicrosoft PowerPointがインストールされている必要があります。
#>

# ==============================================================================
# --- スクリプトのパラメータ定義 ---
# ==============================================================================
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$PptxPath
)

# ==============================================================================
# --- グローバル変数と初期設定 ---
# ==============================================================================

# --- スクリプトの基本パス設定 ---
$scriptRoot = $PSScriptRoot

# --- ベースフォルダ設定 ---
$baseAssetsFolder = Join-Path $scriptRoot "assets"
$global:logRootFolder = Join-Path $scriptRoot "logs"

# --- ログファイル設定 ---
$global:logFilePath = Join-Path $global:logRootFolder "ppt2movie.log"

# --- TTS のデフォルト設定 ---
$defaultAzureVoiceName = "ja-JP-NanamiNeural"
$azureOutputFormat = "riff-24khz-16bit-mono-pcm"

# ==============================================================================
# --- 関数定義 ---
# ==============================================================================

#region ユーティリティ関数

function Write-Log {
    param(
        [Parameter(Mandatory = $true)] [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")] [string]$Level = "INFO"
    )
    $logMessage = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $logMessage
    Add-Content -Path $global:logFilePath -Value $logMessage
}

function Get-PptxPathViaDialog {
    Write-Log "PowerPointファイルのパスが指定されなかったため、選択ダイアログを表示します。"
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Title = "動画を生成するPowerPointファイルを選択してください"
        $openFileDialog.Filter = "PowerPoint プレゼンテーション (*.pptx)|*.pptx"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $openFileDialog.FileName
        }
        return $null
    }
    catch {
        Write-Log "GUIダイアログの表示に失敗しました。`$PptxPath`パラメータでファイルを指定してください。" -Level ERROR
        return $null
    }
}

function Initialize-Environment($config) {
    Write-Log "--- 環境の初期化を開始 ---"

    # ffmpegのパスを解決
    $ffmpegPath = if (-not [string]::IsNullOrWhiteSpace($config.FFMPEG_FOLDER_PATH)) {
        Join-Path $config.FFMPEG_FOLDER_PATH "ffmpeg.exe"
    } else {
        Join-Path $scriptRoot "ffmpeg.exe"
    }
    if (-not (Test-Path $ffmpegPath)) { throw "ffmpeg.exe が見つかりません: $ffmpegPath" }
    $global:ffmpegPath = $ffmpegPath # グローバル変数に設定
    Write-Log "ffmpeg.exe を確認しました: $global:ffmpegPath"

    # ルートのlogs, assetsフォルダがなければ作成
    if (-not (Test-Path $global:logRootFolder)) { New-Item -Path $global:logRootFolder -ItemType Directory | Out-Null }
    if (-not (Test-Path $baseAssetsFolder)) { New-Item -Path $baseAssetsFolder -ItemType Directory | Out-Null }
    
    # ログファイルを初期化
    if (Test-Path $global:logFilePath) { Remove-Item $global:logFilePath }

    # 実行ごとのassetsフォルダを決定
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $runAssetsFolder = Join-Path $baseAssetsFolder $timestamp
    Write-Log "今回の実行用assetsフォルダ: $runAssetsFolder"

    # パス情報をハッシュテーブルに格納
    $paths = @{
        runAssetsFolder   = $runAssetsFolder
        imageOutputFolder = Join-Path $runAssetsFolder "slides_as_images"
        audioOutputFolder = Join-Path $runAssetsFolder "slide_audio"
        textOutputFolder  = Join-Path $runAssetsFolder "slide_texts"
        movieOutputFolder = Join-Path $runAssetsFolder "slide_movies"
        ffmpegLogPath     = Join-Path $runAssetsFolder "ffmpeg_error.log"
        concatListPath    = Join-Path $runAssetsFolder "concat_list.txt"
    }

    # 【修正点】PowerShellのバージョンに応じてffmpegのプロセス起動パラメータを決定
    $global:ffmpegStartProcessArgs = @{
        FilePath    = $global:ffmpegPath
        NoNewWindow = $true
        Wait        = $true
        PassThru    = $true
        RedirectStandardError = $paths.ffmpegLogPath
    }
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $global:ffmpegStartProcessArgs['Append'] = $true
        Write-Log "[DEBUG] PowerShell v6+ を検出。ffmpegのエラーログは追記モードで記録されます。"
    } else {
        Write-Log "[WARN] PowerShell v5 以前を検出。ffmpegのエラーログは実行のたびに上書きされます。"
    }

    # 必要なフォルダを作成
    @($paths.runAssetsFolder, $paths.imageOutputFolder, $paths.audioOutputFolder, $paths.textOutputFolder, $paths.movieOutputFolder) | ForEach-Object {
        if (-not (Test-Path $_)) { New-Item -Path $_ -ItemType Directory | Out-Null }
    }

    Write-Log "--- 環境の初期化が完了 ---"
    return $paths
}

function Get-Configurations {
    Write-Log "--- 設定ファイルの読み込みを開始 ---"
    $envPath = Join-Path $scriptRoot ".env"
    if (-not (Test-Path $envPath)) { throw ".env ファイルが見つかりません。" }

    $config = @{}
    Get-Content $envPath -Encoding UTF8 | ForEach-Object {
        if ($_ -match '^(.*?)=(.*)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            if (($value.StartsWith('"') -and $value.EndsWith('"')) -or ($value.StartsWith('"') -and $value.EndsWith("'"))) {
                $value = $value.Substring(1, $value.Length - 2)
            }
            if (-not [string]::IsNullOrWhiteSpace($key) -and -not $key.StartsWith("#")) {
                $config[$key] = $value
            }
        }
    }

    # 使用するTTSエンジンを決定
    if (-not [string]::IsNullOrWhiteSpace($config.VOICEVOX_SPEAKER_NAME)) {
        $config.TtsEngine = "VOICEVOX"
        if (-not $config.VOICEVOX_URL) { $config.VOICEVOX_URL = "http://127.0.0.1:50021" }
        Write-Log "音声合成エンジン: VOICEVOX"
    } else {
        $config.TtsEngine = "Azure"
        Write-Log "音声合成エンジン: Azure Text-to-Speech"
        if (-not $config.SPEECH_KEY -or $config.SPEECH_KEY -eq "YOUR_KEY") { throw "SPEECH_KEYが.envファイルに設定されていません。" }
        if (-not $config.SPEECH_REGION -or $config.SPEECH_REGION -eq "YOUR_REGION") { throw "SPEECH_REGIONが.envファイルに設定されていません。" }
    }
    
    Write-Log ".envファイルを正常に読み込みました。"
    Write-Log "--- 設定ファイルの読み込みが完了 ---"
    return $config
}

function Start-VoicevoxEngine($config) {
    $exePath = Join-Path $config.VOICEVOX_FOLDER_PATH "VOICEVOX.exe"
    if (-not (Test-Path $exePath)) {
        throw "指定されたVOICEVOXの実行ファイルが見つかりません: $exePath"
    }

    Write-Log "VOICEVOXを起動しています: $exePath"
    $process = Start-Process -FilePath $exePath -PassThru -WindowStyle Minimized
    Write-Log "VOICEVOXのプロセスが開始されました (PID: $($process.Id))。エンジンが応答するまで待機します..."

    # エンジンが起動完了するまで待機 (最大60秒)
    $maxWaitSeconds = 60
    $waitTime = 0
    while ($waitTime -lt $maxWaitSeconds) {
        try {
            Invoke-RestMethod -Method Get -Uri "$($config.VOICEVOX_URL)/version" -TimeoutSec 2 | Out-Null
            Write-Log "  [OK] VOICEVOXエンジンが応答しました。"
            return $process
        } catch {
            Start-Sleep -Seconds 2
            $waitTime += 2
            Write-Log "  ...待機中 ($($waitTime)s)"
        }
    }
    throw "VOICEVOXエンジンが指定時間内に起動しませんでした。"
}

function Get-VoicevoxSpeakerId($config) {
    Write-Log "話者 ' $($config.VOICEVOX_SPEAKER_NAME)' (スタイル: $($config.VOICEVOX_SPEAKER_STYLE)) のIDを検索します..."
    try {
        $client = New-Object System.Net.WebClient
        $client.Encoding = [System.Text.Encoding]::UTF8
        $jsonText = $client.DownloadString("$($config.VOICEVOX_URL)/speakers")
        $speakers = $jsonText | ConvertFrom-Json

        $targetSpeaker = $speakers | Where-Object { $_.name -eq $config.VOICEVOX_SPEAKER_NAME } | Select-Object -First 1
        $targetStyle = $targetSpeaker.styles | Where-Object { $_.name -eq $config.VOICEVOX_SPEAKER_STYLE } | Select-Object -First 1
        if ($targetStyle) {
            Write-Log "  [OK] 話者IDが見つかりました: $($targetStyle.id)"
            return $targetStyle.id
        } else {
            throw "指定された話者名またはスタイルが見つかりません。 Name: $($config.VOICEVOX_SPEAKER_NAME), Style: $($config.VOICEVOX_SPEAKER_STYLE)"
        }
    } catch {
        throw "VOICEVOXから話者情報の取得に失敗しました。: $($_.Exception.Message)"
    }
}

#endregion

#region スライド処理関数

function Export-SlideAsImage($slide, $index, $paths) {
    $imagePath = Join-Path $paths.imageOutputFolder "slide_$($index).png"
    if ($PSCmdlet.ShouldProcess($imagePath, "スライド$($index)を画像としてエクスポート")) {
        $slide.Export($imagePath, "PNG")
        Write-Log "  [OK] スライド $index の画像をエクスポートしました: $imagePath"
    }
}

function Generate-AudioFromNotes($slide, $index, $config, $paths) {
    $audioPath = Join-Path $paths.audioOutputFolder "slide_$($index)_audio.wav"
    
    $notesText = $null
    if ($slide.HasNotesPage) {
        $bodyShape = $slide.NotesPage.Shapes | Where-Object { $_.Type -eq [Microsoft.Office.Core.MsoShapeType]::msoPlaceholder -and $_.PlaceholderFormat.Type -eq [Microsoft.Office.Interop.PowerPoint.PpPlaceholderType]::ppPlaceholderBody }
        if ($bodyShape -and $bodyShape.HasTextFrame -and $bodyShape.TextFrame.HasText) { $notesText = $bodyShape.TextFrame.TextRange.Text.Trim() }
    }

    $cleanNotesText = $notesText
    $voiceOverride = $null
    if ($notesText -match '\(voice:(.*?)\)') {
        $voiceOverride = $matches[1].Trim()
        $cleanNotesText = $notesText -replace '\(voice:.*?\)', ''
        Write-Log "  [INFO] スライド $index で音声が指定されました: $voiceOverride"
    }

    if (-not [string]::IsNullOrWhiteSpace($cleanNotesText)) {
        # --- TTSエンジンに応じて処理を分岐 ---
        if ($config.TtsEngine -eq "VOICEVOX") {
            $speakerId = if ($voiceOverride) { $voiceOverride } else { $config.VOICEVOX_SPEAKER_ID }
            if ($PSCmdlet.ShouldProcess($audioPath, "VOICEVOXで音声合成 (ID: $speakerId)")) {
                try {
                    Write-Log "  [INFO] VOICEVOXで音声を作成します。"
                    Write-Log "  [INFO] $cleanNotesText"
                    $encodedText = [System.Net.WebUtility]::UrlEncode($cleanNotesText)
                    $queryUri = "$($config.VOICEVOX_URL)/audio_query?text=$encodedText&speaker=$speakerId"
                    $audioQuery = Invoke-RestMethod -Method Post -Uri $queryUri -Headers @{ "Content-Type" = "application/json; charset=utf-8" }
                    
                    $synthUri = "$($config.VOICEVOX_URL)/synthesis?speaker=$speakerId"
                    $jsonBody = $audioQuery | ConvertTo-Json -Depth 10
                    Invoke-RestMethod -Method Post -Uri $synthUri -Body $jsonBody -ContentType "application/json" -OutFile $audioPath

                    Write-Log "  [OK] 音声ファイルを保存しました: $audioPath"
                }
                catch {
                    throw "VOICEVOXでの音声合成に失敗しました (スライド $index)。: $($_.Exception.Message)"
                }
            }
        } else { # Azure
            $voiceName = if ($voiceOverride) { $voiceOverride } else { $defaultAzureVoiceName }
            if ($PSCmdlet.ShouldProcess($audioPath, "Azure TTSで音声合成 (Voice: $voiceName)")) {
                Write-Log "  [INFO] Azureで音声を作成します。"
                $ssmlBody = "<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='ja-JP'><voice name='$($voiceName)'>$($cleanNotesText)</voice></speak>"
                $ttsEndpoint = "https://$($config.SPEECH_REGION).tts.speech.microsoft.com/cognitiveservices/v1"
                $headers = @{ "Ocp-Apim-Subscription-Key" = $config.SPEECH_KEY; "Content-Type" = "application/ssml+xml"; "X-Microsoft-OutputFormat" = $azureOutputFormat }
                $utf8Body = [System.Text.Encoding]::UTF8.GetBytes($ssmlBody)
                
                Write-Log "  ... Azureに音声合成をリクエストしています ..."
                Invoke-RestMethod -Method Post -Uri $ttsEndpoint -Headers $headers -Body $utf8Body -OutFile $audioPath
                
                $ssmlFilePath = Join-Path $paths.textOutputFolder "slide_${index}.xml"
                [System.IO.File]::WriteAllText($ssmlFilePath, $ssmlBody, [System.Text.Encoding]::UTF8)
                Write-Log "  [OK] 音声ファイルを保存しました: $audioPath"
            }
        }
    } else {
        if ($PSCmdlet.ShouldProcess($audioPath, "スライド$($index)のノートが空のため無音音声を生成")) {
            Write-Log "  [INFO] スライド $index のノートが空のため、2秒間の無音音声を生成します。"
            $ffmpegArgs = "-f lavfi -i anullsrc=channel_layout=mono:sample_rate=24000 -t 2 -c:a pcm_s16le -y `"$audioPath`""
            # 【修正点】Splattingを使って引数を渡す
            $process = Start-Process @global:ffmpegStartProcessArgs -ArgumentList $ffmpegArgs
            if ($process.ExitCode -ne 0) { throw "無音音声の生成に失敗しました (スライド $index)。詳細は $($paths.ffmpegLogPath) を確認してください。" }
            Write-Log "  [OK] スライド $index の無音音声を生成しました。"
        }
    }
}

function Create-VideoFromAssets($index, $paths) {
    $videoPath = Join-Path $paths.movieOutputFolder "slide_${index}.mp4"
    if ($PSCmdlet.ShouldProcess($videoPath, "スライド$($index)の画像と音声から動画を生成")) {
        Write-Log "  ... スライド $index の動画を生成しています ..."
        $imagePath = Join-Path $paths.imageOutputFolder "slide_$($index).png"
        $audioPath = Join-Path $paths.audioOutputFolder "slide_$($index)_audio.wav"
        $ffmpegArgs = "-loop 1 -i `"$imagePath`" -i `"$audioPath`" -c:v libx264 -tune stillimage -c:a aac -b:a 192k -pix_fmt yuv420p -shortest -y `"$videoPath`""
        # 【修正点】Splattingを使って引数を渡す
        $process = Start-Process @global:ffmpegStartProcessArgs -ArgumentList $ffmpegArgs
        if ($process.ExitCode -ne 0) { throw "動画の生成に失敗しました (スライド $index)。詳細は $($paths.ffmpegLogPath) を確認してください。" }
        Write-Log "  [OK] スライド $index の動画を生成しました: $videoPath"
    }
}

#endregion

#region メイン処理関数

function Process-Slides($presentation, $config, $paths) {
    Write-Log "--- 全スライドのアセット生成処理を開始 ---"
    $slideCount = $presentation.Slides.Count
    Write-Log "プレゼンテーションを開きました。全 $slideCount スライドです。"
    
    for ($i = 1; $i -le $slideCount; $i++) {
        Write-Log "--- スライド $i/$slideCount の処理を開始 ---"
        $slide = $presentation.Slides.Item($i)
        Export-SlideAsImage -slide $slide -index $i -paths $paths
        Generate-AudioFromNotes -slide $slide -index $i -config $config -paths $paths
        Create-VideoFromAssets -index $i -paths $paths
    }
    Write-Log "--- 全スライドのアセット生成処理が完了 ---"
}

function Combine-Videos($presentationName, $paths) {
    Write-Log "--- 動画の結合処理を開始 ---"
    
    $videoFiles = Get-ChildItem -Path $paths.movieOutputFolder -Filter "slide_*.mp4" |
                  Sort-Object { [int]($_.BaseName -replace 'slide_','') } |
                  ForEach-Object { $_.FullName }
    
    if ($videoFiles.Count -eq 0) { throw "結合対象の動画ファイルが $($paths.movieOutputFolder) に見つかりません。" }
    Write-Log "$($videoFiles.Count) 個の動画ファイルを結合します。"

    $finalMoviePath = Join-Path $scriptRoot "$($presentationName -replace '\.pptx$','').mp4"
    $concatContent = $videoFiles | ForEach-Object { "file '$($_ -replace "'","''")'" }
    $encodingNoBOM = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($paths.concatListPath, $concatContent, $encodingNoBOM)

    if ($PSCmdlet.ShouldProcess($finalMoviePath, "全スライドの動画を結合")) {
        $ffmpegArgs = "-f concat -safe 0 -i `"$($paths.concatListPath)`" -c copy -y `"$finalMoviePath`""
        # 【修正点】Splattingを使って引数を渡す
        $process = Start-Process @global:ffmpegStartProcessArgs -ArgumentList $ffmpegArgs
        if ($process.ExitCode -ne 0) { throw "最終的な動画の結合に失敗しました。詳細は $($paths.ffmpegLogPath) を確認してください。" }
        Write-Log "--- すべての処理が正常に完了しました ---" -Level INFO
        Write-Log "完成した動画ファイル: $finalMoviePath" -Level INFO
    }
}

function Cleanup-Resources($powerpoint, $presentation, $voicevoxProcess) {
    if ($presentation) { 
        Write-Log "プレゼンテーションを閉じています..."
        $presentation.Close() 
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }
    if ($powerpoint) { 
        Write-Log "PowerPointアプリケーションを終了しています..."
        $powerpoint.Quit() 
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
    }
    if ($voicevoxProcess) {
        Write-Log "VOICEVOXを終了しています... (PID: $($voicevoxProcess.Id))"
        if (Get-Process -Id $voicevoxProcess.Id -ErrorAction SilentlyContinue) {
            Stop-Process -Id $voicevoxProcess.Id -Force
            Write-Log "  [OK] VOICEVOXのプロセスを終了しました。"
        } else {
            Write-Log "  [INFO] VOICEVOXのプロセスは既に終了していました。"
        }
    }
    Remove-Variable presentation, powerpoint, voicevoxProcess -ErrorAction SilentlyContinue
    Write-Log "リソースを解放しました。"
}

#endregion

# ==============================================================================
# --- スクリプト本編 ---
# ==============================================================================

$pp = $null
$pres = $null
$vvProcess = $null

try {
    $appConfig = Get-Configurations
    $runPaths = Initialize-Environment -config $appConfig
    Write-Log "--- スクリプト実行開始 ---"
    
    if ($appConfig.TtsEngine -eq "VOICEVOX") {
        $vvProcess = Start-VoicevoxEngine -config $appConfig
        # Start-VoicevoxEngineがプロセスを返した場合のみID検索を実行
        if ($vvProcess -or (Get-Process -Name $appConfig.VOICEVOX_EXE_NAME.Split('.')[0] -ErrorAction SilentlyContinue)) {
             $appConfig.VOICEVOX_SPEAKER_ID = Get-VoicevoxSpeakerId -config $appConfig
        }
    }
    
    if (-not $PptxPath) { $PptxPath = Get-PptxPathViaDialog }
    if (-not $PptxPath) { throw "処理対象のPowerPointファイルが指定されませんでした。" }
    
    $PptxPath = Resolve-Path -Path $PptxPath
    Write-Log "対象ファイル: $PptxPath"

    Write-Log "PowerPointアプリケーションを起動しています..."
    $pp = New-Object -ComObject PowerPoint.Application
    $pres = $pp.Presentations.Open($PptxPath, $true, $false, $false)
    
    Process-Slides -presentation $pres -config $appConfig -paths $runPaths
    
    $presentationName = $pres.Name
    Combine-Videos -presentationName $presentationName -paths $runPaths

} catch {
    Write-Log "スクリプトの実行中にエラーが発生しました: $($_.Exception.Message)" -Level ERROR
    Write-Log "スクリプトを中断します。" -Level ERROR
    $errorDetails = $_
    Write-Log "エラー発生元: $($errorDetails.InvocationInfo.ScriptName) 行: $($errorDetails.InvocationInfo.ScriptLineNumber)" -Level DEBUG
    Write-Log "コマンド: $($errorDetails.InvocationInfo.Line.Trim())" -Level DEBUG
    Write-Log "スタックトレース: $($errorDetails.ScriptStackTrace)" -Level DEBUG
} finally {
    Cleanup-Resources -powerpoint $pp -presentation $pres -voicevoxProcess $vvProcess
    Write-Log "--- スクリプト実行終了 ---"
}