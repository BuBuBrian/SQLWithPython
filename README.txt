DemoSQLQueryTool Python／EXE 使用說明
=====================================

一、內容
--------
app.py
    原生 Python tkinter 桌面程式。
    保留原 HTML 版的主要功能：
    - 兩組 WHERE 條件
    - AND / OR
    - 每組各自支援 NOT
    - LIKE、BETWEEN、IS NULL、IS NOT NULL
    - ORDER BY ASC / DESC
    - 文字數字轉換排序
    - SQL 預覽
    - Excel 匯出

demo.sql
    預設資料來源。

run_source.bat
    直接以 Python 執行 UI，適合先測試。

build_exe.bat
    一鍵安裝 PyInstaller 並建立 Windows EXE。

requirements.txt
    Python 套件清單。

original_html_reference.html
    最後一版含 AND / OR / NOT 的 HTML，方便比較。


二、先直接執行 Python 版本
--------------------------
1. 安裝 Python 3.12 64-bit。
2. 安裝時勾選「Add Python to PATH」。
3. 雙擊 run_source.bat。
4. 程式會開啟原生桌面 UI。


三、建立 Windows EXE
--------------------
1. 雙擊 build_exe.bat。
2. 第一次執行會下載 openpyxl 與 PyInstaller。
3. 完成後會產生：

   dist\DemoSQLQueryTool.exe
   dist\demo.sql

4. 雙擊 DemoSQLQueryTool.exe 即可啟動 UI。
5. 建置完成的 EXE 可放到沒有安裝 Python 的 Windows 電腦執行。


四、demo.sql 放置方式
---------------------
程式會優先讀取 EXE 或 app.py 同資料夾中的 demo.sql。

建置時 demo.sql 也會內嵌進 EXE，因此：
- EXE 旁邊有 demo.sql：優先讀取外部檔案，可自行替換。
- EXE 旁邊沒有 demo.sql：使用 EXE 內嵌的預設資料。


五、Excel 匯出
--------------
按「匯出目前結果 Excel」後：

Excel 最上方：
SELECT
FROM
WHERE
AND / OR / NOT
ORDER BY

空一列後：
欄位名稱與目前查詢結果。


六、重要說明
------------
這個程式不會連接 MySQL Server。
它解析 MySQL Dump 中的 CREATE TABLE 與 INSERT INTO，
再以 Python 模擬 SQL 查詢。

目前環境不是 Windows，因此提供的是已驗證語法的 Python 專案
與 Windows 一鍵建置檔，而不是在非 Windows 環境交叉編譯的 EXE。
