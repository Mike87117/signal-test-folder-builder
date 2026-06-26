# Signal Test Folder Builder

一個用於建立測試資料夾結構的桌面小工具，適合訊號量測、相容性測試、DUT 測試資料整理等場景使用。使用者可以依照 DUT、Version、主板板廠、小卡、小卡板廠、Cable、Connector、Slot、GEN 等條件，自動產生完整的測試資料夾階層，並可匯出測試清單 XLSX，方便後續記錄測試結果。

## 功能特色

- 使用 PySide6 建立桌面 GUI，不需要使用命令列操作。
- 支援 DUT Name 與 Version 作為最上層資料夾。
- 支援主板板廠、小卡、小卡板廠、Cable、Connector、Slot、GEN 多層級資料夾建立。
- 支援多組設定，可針對不同小卡、Slot 或測試組合分別建立規則。
- 支援 Slot / GEN 編號規則，例如 `1-5`、`1,3-5,8`。
- 提供資料夾結構預覽，可切換精簡模式與詳細模式。
- 可匯入 / 匯出 JSON 設定，方便重複使用同一組測試配置。
- 可匯出測試清單 XLSX，內含：
  - Test Matrix 工作表
  - Result 欄位下拉選單：PASS / FAIL / NA / BLOCK
  - PASS / FAIL 條件式格式
  - Summary 統計頁
- 建立資料夾時顯示進度視窗，並支援取消。

## 適合使用情境

此工具適合用於需要大量建立固定測試資料夾結構的工作，例如：

- PCIe / DP / eDP / USB / Thunderbolt 等訊號量測資料整理
- DUT 測試資料夾初始化
- 多 Slot / 多 GEN / 多 Cable / 多 Connector 的測試矩陣整理
- 測試前先產生標準化資料夾與測試清單

## 資料夾結構範例

假設輸入：

- DUT Name：`DUT_A`
- Version：`EVT1`
- 主板板廠：`ASUS`
- 小卡名稱：`LAN_CARD`
- 小卡板廠：`Vendor_A`
- Cable：`Cable_A`
- Connector：`Conn_A`
- Slot 名稱：`PCIE`
- Slot 編號：`1-2`
- GEN 編號：`4-5`

會產生類似以下結構：

```text
DUT_A/
└── EVT1/
    └── ASUS/
        └── LAN_CARD/
            └── Vendor_A/
                └── Cable_A/
                    └── Conn_A/
                        ├── PCIE1/
                        │   ├── GEN4/
                        │   └── GEN5/
                        └── PCIE2/
                            ├── GEN4/
                            └── GEN5/
```

## 編號輸入規則

Slot 編號與 GEN 編號支援以下格式：

| 輸入 | 代表意義 |
|---|---|
| `1-5` | 連續建立 1、2、3、4、5 |
| `1,5` | 只建立 1 和 5 |
| `1,3-5,8` | 建立 1、3、4、5、8 |

注意事項：

- Slot 名稱有填寫時，Slot 編號可空白；空白時會直接建立固定名稱資料夾。
- GEN 編號可空白；空白時不建立 GEN 資料夾。
- 編號只允許數字、逗號、分號、頓號、換行與 dash。
- 範圍不可反向，例如 `5-1` 會被判定為錯誤。

## 安裝方式

建議使用 Python 3.10 以上版本。

```bash
git clone https://github.com/your-name/signal-test-folder-builder.git
cd signal-test-folder-builder
pip install -r requirements.txt
```

如果沒有建立 `requirements.txt`，可以直接安裝必要套件：

```bash
pip install PySide6 openpyxl
```

## 執行方式

```bash
python main.py
```

## 建議專案結構

```text
signal-test-folder-builder/
├── README.md
├── requirements.txt
├── main.py
└── .gitignore
```

## 使用流程

1. 選擇輸出路徑。
2. 輸入 DUT Name 與 Version。
3. 輸入主板板廠清單。
4. 在「設定列表」新增一組或多組設定。
5. 針對每組設定填入小卡名稱、小卡板廠、Cable、Connector、Slot 與 GEN。
6. 透過右側預覽確認資料夾結構。
7. 需要保留配置時，可匯出 JSON。
8. 需要測試清單時，可匯出 XLSX。
9. 點擊「開始建立資料夾」產生資料夾。

## 匯出 XLSX 說明

匯出的 XLSX 會包含兩個工作表：

### Test Matrix

用於列出每一筆測試組合，欄位包含：

- Mainboard Vendor
- Card
- Card Vendor
- Cable
- Connector
- Slot
- GEN
- Result
- Tester
- Date
- Note

### Summary

用於統計：

- Total Cases
- PASS Count
- FAIL Count
- PASS Rate

## 開發備註

目前程式主要使用以下套件：

- `PySide6`：建立 GUI 介面
- `openpyxl`：產生 XLSX 測試清單
- `pathlib`：處理資料夾路徑
- `json`：匯入 / 匯出設定檔

## 後續可改進方向

- 將 UI 與資料處理邏輯拆成不同檔案，方便維護。
- 新增 `requirements.txt` 與版本鎖定。
- 新增範例 JSON 設定檔。
- 新增英文版介面或 README，方便公開分享。
- 加入 PyInstaller 打包流程，讓沒有 Python 環境的使用者也能執行。
- 加入單元測試，特別是 Slot / GEN 編號解析邏輯。
