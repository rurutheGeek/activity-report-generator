# 活動報告書自動生成ツール (Activity Report Generator)

このツールは活動報告書を自動的に生成・更新しテストを行うPythonスクリプトです。月ごとの活動内容をランダムに生成し、Wordテンプレートに挿入します。

## 特徴

- 活動日と場所を入力すると自動的に報告書を生成
- 参加人数をランダムに設定
- 重み付けされたアクティビティリストから確率的に選択
- 年月情報の自動更新（テンプレート部分を維持）
- 八王子/新宿キャンパスの選択に対応
- ファイル名に場所と年月を含む
- 生成したファイルを自動的に開く機能
- PyInstaller対応: 実行ファイルとして配布

## 必要条件

### 実行ファイル版を使用する場合

- Windowsコンピュータ
- Microsoft Word（報告書を開くため）

### Pythonスクリプトとして実行する場合
- Python 3.6以上
- `python-docx` ライブラリ

## インストール方法

### 方法1: 実行ファイルを使用する場合

1. リリースページから最新の `ActivityReportGenerator.exe` をダウンロード
2. ダウンロードした `ActivityReportGenerator.exe` ファイルをダブルクリックして実行

### 方法2: Pythonスクリプトとして実行する場合

1. リポジトリをクローンまたはダウンロード：
   ```
   git clone https://github.com/ruruthegeek/activity-report-generator.git
   cd activity-report-generator
   ```

2. 必要なライブラリをインストール：
   ```
   pip install -r requirement.txt
   または
   pip install python-docx
   ```

3. テンプレートファイル「活動報告.docx」をスクリプトと同じディレクトリに配置

## 使用方法

1. ツールを実行：
   - 実行ファイル版: `ActivityReportGenerator.exe` をダブルクリック
   - Pythonスクリプト: `python activity_report_generator.py`

2. プロンプトに従って情報を入力：
   - 活動年と月を入力
   - 八王子または新宿を選択（h/s）
   - 活動日と場所を入力（最大6回）
   - 入力を終了するには日付欄を空白のままEnterを押す

3. ツールは自動的に：
   - 活動報告書を生成
   - 年月を現在の日付に更新
   - 新しいファイルを保存
   - 生成されたファイルを開く


## トラブルシューティング

### ファイルが更新されない場合

1. テンプレートファイルが「活動報告.docx」という名前でテンプレートファイルが実行ファイルと同じ階層に設置されていることを確認
2. テンプレートファイルが開かれていないことを確認

### ファイルが自動的に開かない場合

スクリプトは生成されたファイルの絶対パスを表示します。このパスを使用して手動でファイルを開いてください。

