# PowerPoint to Movie Converter with AI Voice (pptx2movie)

このPowerShellスクリプトは、PowerPointプレゼンテーション（.pptx）をMP4ビデオファイルに自動変換します。各スライドのノートからテキストを抽出し、VOICEVOXまたはAzure Text-to-Speech（TTS）を使用して音声を生成し、スライド画像と生成された音声を組み合わせてシームレスなビデオを作成します。

## 🚀 主な機能

*   PowerPointプレゼンテーションをMP4ビデオに変換します。
*   ノートの内容をAI音声でナレーションします。
*   `.env`ファイルの設定に基づき、VOICEVOX（ローカル）またはAzure Text-to-Speech（クラウド）をサポートします。
*   VOICEVOXが選択された場合、VOICEVOXエンジンプロセスを自動的に起動・管理します。
*   ノート内の `(voice:...)` タグを使用して、スライドごとに音声（話者）をカスタマイズできます。
*   中間アセット（画像、音声、一時的なビデオ）はタイムスタンプ付きのフォルダに保存され、デバッグが容易です。

## 📋 必要要件

*   **Microsoft PowerPoint:** スクリプトを実行するマシンにインストールされている必要があります（COMオブジェクト操作のため）。
*   **FFmpeg:** `ffmpeg.exe` が利用可能である必要があります。
*   **PowerShell:** バージョン5.1以降（Windows）。
*   **.NET Framework:** ファイル選択ダイアログの表示に必要です（通常、Windowsにプリインストールされています）。
*   **VOICEVOX (オプション):** VOICEVOXを使用する場合、アプリケーションをダウンロードし、設定する必要があります。
*   **Azure Text-to-Speech サブスクリプション (オプション):** Azureを使用する場合、Speech Serviceリソースを含むアクティブなAzureサブスクリプションが必要です。

## 🛠️ セットアップ

1.  **リポジトリのクローンまたはダウンロード:**
    ```bash
    git clone https://github.com/tetsunosuke/pptx2movie.git
    ```
    または、ZIPファイルをダウンロードして展開します。

2.  **FFmpegの配置:**
    `ffmpeg.exe` をプロジェクトルートの `FFmpeg/` ディレクトリに配置し、`.env` ファイルでパスを指定します。

3.  **VOICEVOXの配置 (VOICEVOXを使用する場合):**
    VOICEVOXアプリケーションのファイルをプロジェクトルートの `VOICEVOX/` ディレクトリに配置するか、`.env` ファイルでパスを指定します。

4.  **.env ファイルの作成:**
    プロジェクトのルートディレクトリに `.env` という名前のファイルを作成し、以下の設定を記述します。

## ⚙️ 設定 (`.env`)

`.env` ファイルは、スクリプトの動作を制御するための環境変数を定義します。 `.env.example` を参考にしてください

```ini
# --- Azure Text-to-Speech の設定 ---
# VOICEVOX_SPEAKER_NAME が設定されていない場合に使用されます。
# Azure Speech Service のキーとリージョンを設定してください。
# SPEECH_KEY="YOUR_AZURE_SPEECH_KEY"
# SPEECH_REGION="YOUR_AZURE_SPEECH_REGION" # 例: eastus, japaneast

# --- VOICEVOX の設定 ---
# VOICEVOX を使用する場合に設定します。
# VOICEVOX_SPEAKER_NAME が設定されている場合、VOICEVOX が優先されます。
# VOICEVOX アプリケーションのルートディレクトリへのパス。
# VOICEVOX_FOLDER_PATH="C:/path/to/VOICEVOX"
# デフォルトの話者名とスタイル。VOICEVOX アプリケーションで確認できます。
# VOICEVOX_SPEAKER_NAME="ずんだもん"
# VOICEVOX_SPEAKER_STYLE="ノーマル"
# VOICEVOX エンジンのURL。通常は変更不要です。
# VOICEVOX_URL="http://127.0.0.1:50021"

# --- FFmpeg の設定 ---
# ffmpeg.exe が存在するフォルダへのパス。
# 設定しない場合、スクリプトはスクリプトと同じ階層またはFFmpeg/フォルダを検索します。
# FFMPEG_FOLDER_PATH="C:/path/to/FFmpeg"
```

**注意:**
*   `VOICEVOX_SPEAKER_NAME` が設定されている場合、VOICEVOXが優先的に使用されます。
*   `VOICEVOX_SPEAKER_NAME` が設定されていない場合、Azure TTSが使用されます。その際、`SPEECH_KEY` と `SPEECH_REGION` の設定が必須となります。

## 🚀 使い方

PowerShellを開き、スクリプトのディレクトリに移動します。

1.  **ファイル選択ダイアログから実行:**
    ```powershell
    .\pptx2movie.ps1
    ```
    実行すると、PowerPointファイルを選択するためのダイアログが表示されます。

2.  **パスを指定して実行:**
    ```powershell
    .\pptx2movie.ps1 -PptxPath "C:\Users\YourUser\Documents\MyPresentation.pptx"
    ```
    `"C:\Users\YourUser\Documents\MyPresentation.pptx"` を実際のPowerPointファイルへのパスに置き換えてください。

スクリプトの実行後、元のPowerPointファイルと同じディレクトリにMP4動画ファイルが生成されます。

## 🗣️ 音声のカスタマイズ

各スライドのノートに `(voice:...)` の形式で記述することで、スライドごとの音声（話者）を上書きできます。

*   **Azureの場合:**
    `ja-JP-NanamiNeural` のような音声名を使用します。
    例: `(voice:ja-JP-KeitaNeural) これは男性の声です。`

*   **VOICEVOXの場合:**
    話者IDを直接指定します。話者IDはVOICEVOXアプリケーションのAPIドキュメントや、`speakers` エンドポイントから取得できます。
    例: `(voice:8) これはVOICEVOXの特定の話者の声です。`

## トラブルシューティング

*   **`ffmpeg.exe` が見つかりません:**
    *   `ffmpeg.exe` が `FFmpeg/` ディレクトリに正しく配置されているか確認してください。
    *   `.env` ファイルの `FFMPEG_FOLDER_PATH` が正しいパスを指しているか確認してください。

*   **`SPEECH_KEY` が.envファイルに設定されていません。:**
    *   Azure TTSを使用している場合、`.env` ファイルに `SPEECH_KEY` と `SPEECH_REGION` が正しく設定されているか確認してください。

*   **VOICEVOXエンジンが指定時間内に起動しませんでした。:**
    *   `.env` ファイルの `VOICEVOX_FOLDER_PATH` がVOICEVOXアプリケーションの正しいルートディレクトリを指しているか確認してください。
    *   VOICEVOXアプリケーションがファイアウォールによってブロックされていないか確認してください。
    *   VOICEVOXアプリケーションを手動で起動してみて、正常に動作するか確認してください。

*   **指定された話者名またはスタイルが見つかりません。:**
    *   `.env` ファイルの `VOICEVOX_SPEAKER_NAME` と `VOICEVOX_SPEAKER_STYLE` がVOICEVOXアプリケーションで利用可能な話者名とスタイルに一致しているか確認してください。

