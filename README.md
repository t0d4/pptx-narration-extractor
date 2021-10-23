# 音声つきパワポから音声を抽出し、速度を調整して一つのファイルに結合する

## これは何？

世にあふれる音声付きpptxファイルから音声のみを取り出し、速度を調整した上で一つのmp3ファイルに結合します。
この際、各スライドに対応する開始タイミングを一覧形式で記載したテキストファイルが同時に出力されます。

音声つきpptxをPDFとmp3ファイルに分けることで容量の削減も期待されます。

## ダウンロード

[こちら](https://github.com/t0d4/pptx-narration-extractor/releases)からどうぞ

## 要求環境

### ソフトウェアの用意
- ffmpeg または libav

Linux

1. `sudo apt install ffmpeg libavcodec-extra`
2. 下の"Pythonパッケージのインストール"を行う

Mac

1. `brew install ffmpeg`
2. 下の"Pythonパッケージのインストール"を行う

Windows (**WSL環境が必要です**)
1. WSLをインストール済みであることを確認する。
2. ダウンロードしたフォルダの中にあるwin_setupフォルダ内のwsl_setup.ps1を右クリックし、メニュー中の"PowerShellで実行"をクリックする。
3. WSLが立ち上がるので、パスワードを入力する。
4. 終了。以下の"Pythonパッケージ"の欄を行う必要はありません。

### Pythonパッケージのインストール

Requirements
- pydub
- tqdm

パッケージのバージョンは任意です。以下のコマンドによって全てインストールできます。

`pip install -r requirements.txt`

## 使い方

Linux or Mac

1. `python extractor.py [OPTIONS] PPTX_FILE`
2. audioフォルダにmp3ファイルが生成される

オプション:

   --speed : 音声を何倍速にしたいかを指定します。(例: 1.2) そのままの速度にしたい場合は使用しないでください。 **注意: 1.0未満の数値を入力した場合、プログラムは正しく動作しません。**

Windows

1. pptxファイルを`win_extractor.bat`にドラッグアンドドロップする
2. Windowsによって警告が表示されるので、開発者を信頼する場合は"詳細情報"を押して"実行"を押す
3. 音声の速度を変更したい場合は数値を入力し、もとの速度で良い場合はそのままでEnterを押す
4. audioフォルダにmp3ファイルが生成される

## 注意

実装上、1.6倍速より速いファイルを生成する場合は処理にかかる時間が長くなります。