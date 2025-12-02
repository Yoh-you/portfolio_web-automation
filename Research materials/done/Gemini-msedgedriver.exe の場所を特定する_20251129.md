# msedgedriver.exe の場所を特定する

**Exported:** 11/29/2025 14:20:24  

## Prompt:
msedgedriver.exe の保存場所をターミナルで特定させたい

## Response:
ターミナル（コマンドプロンプトやPowerShellなど）で **`msedgedriver.exe`** の保存場所を特定する方法はいくつかありますが、最も一般的なのは **`where`** コマンドまたは **`Get-Command`** コマンドを使用することです。

* * *

🔍 `where` コマンドによる検索（コマンドプロンプト/PowerShell）
------------------------------------------

`msedgedriver.exe` にパスが通っている（環境変数 `Path` にそのディレクトリが含まれている）場合に利用できます。

```
where msedgedriver.exe
```

*   **結果の例:**
    ```
    C:\Program Files\Selenium\msedgedriver.exe
    ```
*   **注意:** パスが通っていない場合は、ファイルが見つからなかったことを示すメッセージが表示されます。

* * *

🔍 `Get-Command` コマンドレットによる検索（PowerShell）
-----------------------------------------

PowerShellを使用している場合、`Get-Command` コマンドレットでもパスが通っている実行ファイルの場所を特定できます。

```
Get-Command msedgedriver.exe | Select-Object -ExpandProperty Path
```

*   **結果の例:**
    ```
    C:\Program Files\Selenium\msedgedriver.exe
    ```

* * *

🔍 ファイルシステム全体からの検索
------------------

上記の方法で見つからなかった場合、ファイルシステム全体を検索する必要があります。これは時間がかかる場合があります。

### 1\. PowerShell での検索（推奨）

`Get-ChildItem` コマンドレットを使ってファイルシステム全体を再帰的に検索します。

```
Get-ChildItem -Path C:\ -Filter msedgedriver.exe -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
```

*   **注意:** 検索対象のドライブ (`C:\`) は適宜変更してください。`-ErrorAction SilentlyContinue` は、アクセス権がないフォルダをスキップするために使用します。

### 2\. コマンドプロンプトでの検索

コマンドプロンプトでは `dir` コマンドと `findstr` を組み合わせて検索できますが、PowerShellの方が強力です。

```
dir C:\msedgedriver.exe /s /b
```

*   **注意:** こちらも検索対象のドライブ (`C:\`) は適宜変更してください。

* * *

💡 一般的な保存場所
-----------

`msedgedriver.exe` は通常、手動でダウンロードして配置されるため、特定の場所に決まっているわけではありませんが、一般的な配置場所としては以下のようなディレクトリが考えられます。

*   Seleniumなどのテストフレームワークを使用しているプロジェクトフォルダ内
*   `C:\Windows\System32` または `C:\Windows` (システム全体で利用可能にするため)
*   **WebDriver** を管理するための専用のフォルダ (例: `C:\WebDriver` や `C:\tools\webdriver`)

もし、`msedgedriver.exe` を自分でダウンロード・配置した覚えがあれば、その配置したディレクトリを直接確認するのが最も速い方法です。

もし、このファイルが見つからない場合、まずは **Microsoft Edge Insider Channels** のダウンロードページからダウンロードする必要があります。

`msedgedriver.exe` のダウンロード方法について知りたいですか？

## Prompt:
> ![Preview Image](blob:https://gemini.google.com/3ebfd479-4c5f-4e16-a9f5-cbd80967b35c)

この結果はないということだよね？

以下のコードをみて。Pythonでedgeを開くときに使ってるんだ。

self.driver = webdriver.Edge(options=options)

msedgedriver.exeはないのに、実際にPythonが動いてるのはなぜかな？

## Response:
はい、アップロードされた画像の結果は、ターミナルで実行した時点では、`where` コマンドが環境変数 **`Path`** に設定されているディレクトリの中から **`msedgedriver.exe`** を見つけられなかった、ということを示しています。

* * *

💡 `msedgedriver.exe` が見つからないのに Python が動く理由
--------------------------------------------

`msedgedriver.exe` の場所が特定できないにもかかわらず、Pythonのコードが動いて Microsoft Edge を開ける主な理由は、**Selenium の新しいバージョン**が、**Edge WebDriver を自動的に管理・ダウンロードする機能**を持っているからです。

### 1\. **WebDriver Manager（自動管理機能）**

最近の **Selenium (Python)** のバージョンでは、`webdriver.Edge()` や `webdriver.Chrome()` などのドライバを初期化する際、**必要な WebDriver の実行ファイル（この場合は `msedgedriver.exe`）がローカル環境で見つからない場合**に、**自動的にダウンロードしてキャッシュし、それを使用する**ようになっています。

*   **以前のバージョン**では、開発者が手動で `msedgedriver.exe` をダウンロードし、そのファイルのパスをコード内で指定するか、環境変数 `Path` に通す必要がありました。
*   **現在のバージョン**では、その手間が不要になり、`msedgedriver.Edge()` と書くだけで、**裏側でドライバのダウンロードと設定が実行**されています。

### 2\. **コードの動作**

あなたのコード

```
self.driver = webdriver.Edge(options=options)
```

を実行すると、Selenium のライブラリ内部では以下のプロセスが実行されていると考えられます。

1.  ローカルの `Path` をチェックし、`msedgedriver.exe` があるか探す。
2.  （`where` の結果からわかるように）**見つからない**。
3.  Selenium の内部的な **WebDriver Manager** が、現在インストールされている **Microsoft Edge のブラウザバージョンに合った** `msedgedriver.exe` を自動的にインターネットからダウンロードする。
4.  ダウンロードしたファイルを一時的なキャッシュディレクトリに保存する。
5.  そのキャッシュされた `msedgedriver.exe` を使って Edge ブラウザの制御を開始する。

### 3\. **「見つからない」ことの意味**

`where msedgedriver.exe` の結果が「見つかりませんでした」というのは、**システム全体で利用可能な**（`Path` に設定された）**場所には存在しない**、という意味です。Pythonが使っている `msedgedriver.exe` は、**Pythonライブラリが一時的に使っている別の場所（キャッシュ）に存在している**ため、コードの実行には問題がないのです。

* * *

✅ 確認方法
------

実際に自動でダウンロードされた `msedgedriver.exe` がどこにあるか確認したい場合は、Seleniumが使用する**キャッシュフォルダ**を探す必要があります。

使用しているライブラリ（例えば `selenium-manager` やサードパーティの WebDriver Manager）によって場所は異なりますが、一般的にはユーザーのプロファイルフォルダ（`~/.cache/selenium` や `~/.wdm` など）内に保存されています。



---
Powered by [Gem Chat Exporter](https://www.gem-chat-exporter.com)