# rpa-accounting

這是 ZAZA 眨眨眼組織，開發用來將分散在不同系統的帳務資料整合到統一的 Google 表單中。
目前支援包含： NetiCRM 系統的捐款資料、藍新系統的手續費以及提領到帳的資訊，透過自動化瀏覽器或是 Gmail 信件的方式，將相關資訊自動爬出並整理到指定的 Google 表單中，並將相關的憑證上傳指定的資料夾。

## 設定細節

在使用之前，請先修改 `sample_config.yml`  內的每個欄位，以符合自己的所需的設定。

## 使用方法

可以直接用 `python rpa_accounting.py -a` 去蒐集昨天所有的資訊，並將它轉移到 Google 表單中。

其他選項包含：

* `--start_date YYYY-mm-dd` : 設定要搜尋的起始日期，預設為昨天。
* `--end_date YYYY-mm-dd` : 設定要搜尋的結束日期，預設為昨天。
* `--config PATH_TO_CONFIG_FILE` : 設定要參考的 config 檔案，預設為 `./config.yml`
* `--crawl_newebpay_invoice` : 只整理藍新金流的手續費資訊
* `--crawl_transfer_result` : 只整理藍新金流提領到帳戶的資訊
* `--crawl_neti_result` : 只處理 NetiCRM 的捐款資料