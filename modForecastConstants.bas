Attribute VB_Name = "modForecastConstants"
Option Explicit

' ================================
' Worksheet Names
' ================================
Public Const WS_FORECAST_UI As String = "Forecast_UI"
Public Const WS_FORECAST_DB As String = "Forecast_DB"
Public Const WS_FORECAST_RULES_DB As String = "Forecast_Rules_DB"
Public Const WS_PRODUCTS_DB As String = "Products_DB"
Public Const WS_SETTINGS As String = "Settings"
Public Const WS_INVENTORY_LOG As String = "Inventory_Log"
Public Const WS_SALES_DB As String = "Sales_DB"

' ================================
' Table Names
' ================================
Public Const TBL_FORECAST As String = "tblForecast"
Public Const TBL_FORECAST_RULES As String = "tblForecastRules"

' ================================
' Forecast_UI - Scope / SKU Area
' ================================
Public Const CELL_APPLY_TO As String = "B5"
Public Const CELL_SKU_SELECTOR As String = "B7"
Public Const RNG_SELECTED_SKUS As String = "B8:B12"

' ================================
' Forecast_UI - Rule Input Area
' ================================
Public Const CELL_BASE_METHOD As String = "B14"
Public Const CELL_HIST_DAYS As String = "B15"
Public Const CELL_TREND_ON As String = "B16"
Public Const CELL_TREND_SENSITIVITY As String = "B17"
Public Const CELL_NOHISTORY_METHOD As String = "B18"
Public Const CELL_NOHISTORY_DEFAULT_QTY As String = "B19"
Public Const CELL_MANUAL_BASE_QTY As String = "B20"
Public Const CELL_PROMO_MULTIPLIER As String = "B21"
Public Const CELL_PROMO_ADDON As String = "B22"
Public Const CELL_SEASONALITY_ON As String = "B23"
Public Const CELL_ROUND_RULE As String = "B24"

' ================================
' Forecast_UI - Rule Info Display
' ================================
Public Const CELL_INFO_CURRENT_STOCK As String = "D14"
Public Const CELL_INFO_REORDER_LEVEL As String = "D15"
Public Const CELL_INFO_LEADTIME As String = "D16"
Public Const CELL_INFO_LAST7AVG As String = "D17"
Public Const CELL_INFO_LAST30AVG As String = "D18"
Public Const CELL_INFO_RULE_ID As String = "D19"
Public Const CELL_INFO_LAST_UPDATED As String = "D20"

' ================================
' Forecast_UI - Generate Area
' ================================
Public Const CELL_GEN_START_DATE As String = "G5"
Public Const CELL_GEN_HORIZON_DAYS As String = "G6"
Public Const CELL_GEN_REPLACE_EXISTING As String = "G7"
Public Const CELL_GEN_MISSING_ONLY As String = "G8"
Public Const CELL_GEN_RUN_PREVIEW As String = "G9"

' ================================
' Forecast_UI - View Control Area
' 你现在页面已经下移一行
' ================================
Public Const CELL_VIEW_START_DATE As String = "B28"
Public Const CELL_VIEW_DAYS As String = "D28"
Public Const CELL_VIEW_MODE As String = "F28"

' ================================
' Forecast_UI - Daily View Layout
' 你现在表头在第29行，数据从第30行开始
' ================================
Public Const ROW_VIEW_HEADER As Long = 29
Public Const ROW_VIEW_DATA_START As Long = 30

Public Const COL_VIEW_SKU As Long = 1          ' A
Public Const COL_VIEW_PRODUCT As Long = 2      ' B
Public Const COL_VIEW_STOCK As Long = 3        ' C
Public Const COL_VIEW_SOURCE As Long = 4       ' D
Public Const COL_VIEW_RULESUMMARY As Long = 5  ' E
Public Const COL_VIEW_DATE_START As Long = 7   ' G

' ================================
' General Limits
' ================================
Public Const MAX_SELECTED_SKUS As Long = 5

' ================================
' Apply To Options
' ================================
Public Const APPLY_TO_SELECTED As String = "SELECTED_SKUS"
Public Const APPLY_TO_ALL As String = "ALL_PRODUCTS"

' ================================
' Base Method Options
' ================================
Public Const BASE_METHOD_RECENT_AVG As String = "RECENT_AVG"
Public Const BASE_METHOD_WEIGHTED_AVG As String = "WEIGHTED_AVG"
Public Const BASE_METHOD_MANUAL_BASE As String = "MANUAL_BASE"

' ================================
' Trend Options
' ================================
Public Const YES_TEXT As String = "YES"
Public Const NO_TEXT As String = "NO"

Public Const TREND_CONSERVATIVE As String = "CONSERVATIVE"
Public Const TREND_NORMAL As String = "NORMAL"
Public Const TREND_AGGRESSIVE As String = "AGGRESSIVE"

' ================================
' No History Options
' ================================
Public Const NOHISTORY_ZERO As String = "ZERO"
Public Const NOHISTORY_MANUAL_DEFAULT As String = "MANUAL_DEFAULT"
Public Const NOHISTORY_REORDER_BASED As String = "REORDER_BASED"

' ================================
' Round Rule Options
' ================================
Public Const ROUND_RULE_ROUND As String = "ROUND"
Public Const ROUND_RULE_ROUNDUP As String = "ROUNDUP"
Public Const ROUND_RULE_ROUNDDOWN As String = "ROUNDDOWN"

' ================================
' View Mode Options
' ================================
Public Const VIEW_MODE_FINAL As String = "FINAL"
Public Const VIEW_MODE_SYSTEM As String = "SYSTEM"
Public Const VIEW_MODE_OVERRIDE As String = "OVERRIDE"

' ================================
' Default Values
' ================================
Public Const DEFAULT_HIST_DAYS As Long = 14
Public Const DEFAULT_VIEW_DAYS As Long = 30
Public Const DEFAULT_HORIZON_DAYS As Long = 90
Public Const DEFAULT_PROMO_MULTIPLIER As Double = 1
Public Const DEFAULT_PROMO_ADDON As Double = 0
Public Const DEFAULT_NOHISTORY_QTY As Double = 1
Public Const DEFAULT_SEASONALITY_ON As String = YES_TEXT

' Auto seasonality settings
Public Const DEFAULT_AUTO_SEASONALITY_WINDOW_DAYS As Long = 56
Public Const MIN_AUTO_SEASONALITY_TOTAL_DAYS As Long = 10
Public Const MIN_AUTO_SEASONALITY_WEEKDAY_COUNT As Long = 2
Public Const MIN_AUTO_SEASONALITY_FACTOR As Double = 0.8
Public Const MAX_AUTO_SEASONALITY_FACTOR As Double = 1.2
