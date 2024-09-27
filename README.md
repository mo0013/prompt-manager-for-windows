## プロンプト管理アプリ

### 概要
---
本ツールは生成AIに入力するプロンプトを管理するためのアプリケーションです。
Windowsの通知トレイに常駐し、プロンプトが簡単に管理できるようになります。
プロンプトの保存を探す時間、プロンプトをコピーする時間を減らして、生成AIをより快適に使えるようになります。
![alt text](./img/sampleimage.png)
### 特徴
---
* **すぐに使える:** このアプリケーションは、Windowsに標準搭載されているPowerShellで開発されているため、追加のソフトウェアのインストールなしですぐに使用を開始できます。
* **軽量:** PowerShellスクリプトで記述されているため、システムリソースへの負荷が少なく、軽量です。
* **高い互換性:** Windows標準の機能を利用しているため、多くのWindowsバージョンで問題なく動作します。
* **スタートアップに登録:** スタートアップに登録することで、Windowsの起動時に自動的にタスクトレイ上に常駐します。

### 機能
---
* **プロンプト管理:**
    * Markdownファイル(.md)でプロンプトを保存します。
    * プロンプトをカテゴリで分類・管理できます。
    * プロンプトの新規作成、編集、削除が可能です。
* **プロンプトのプレビューとコピー:**
    * プロンプトを選択すると、プレビュー画面で詳細を確認できます。
    * プレビュー画面からプロンプトをクリップボードにコピーできます。

### 環境
---
このアプリケーションは、以下の環境で開発されています。
* **プログラミング言語:** PowerShell 5.1
* **UIフレームワーク:** .NET Framework
* **データ形式:** Markdown

### 実行方法
---
1. toolsフォルダ内の`setup_startup.bat` を管理者権限で実行します。このスクリプトは以下の処理を行います：
   - PowerShellの実行ポリシーを確認し、必要に応じて「RemoteSigned」に変更します。これは、ローカルで作成したスクリプトを実行可能にするために必要です。
   - アプリケーションをWindowsのスタートアップに登録します。

   注意: PowerShellの実行ポリシーの変更は、セキュリティ上の理由から必要です。この変更により、署名されていないローカルスクリプトの実行が可能になりますが、リモートからダウンロードしたスクリプトは引き続き署名が必要となります。
2. Windowsを再起動します。
3. 通知トレイにアプリケーションのアイコンが表示されます。

    ![タスクトレイのアイコン](./img/taskTray.png)

4. アイコンをクリックすると、プロンプト管理画面が表示されます。

    ![プロンプト管理画面](./img/promptManager.png)

### アンインストール方法
---
1. toolsフォルダ内の`delete_startup.bat` を管理者権限で実行します。このスクリプトは以下の処理を行います：
   - PowerShellの実行ポリシーを元の設定に戻します。
   - Windowsのスタートアップフォルダからアプリケーションのショートカットを削除します。

   注意:
    - このスクリプトを実行することで、セットアップ時に変更された設定が元に戻されます。
    - `powershell_policy_before_change.txt` はセットアップ前のPowerShellの実行ポリシーを保存しているファイルです。削除はしないでください。
