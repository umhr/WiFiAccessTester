# 特定のターゲットリストを指定したいときには -target で指定可能
# powershell -ExecutionPolicy Bypass .\accesstester.ps1 -target "./newlist.csv"

# 新しくターゲットリストを読み込まずに、読み込み保存済みのターゲットリストを使う際は -init $false
# powershell -ExecutionPolicy Bypass .\accesstester.ps1 -init $false

powershell -ExecutionPolicy Bypass .\accesstester.ps1

pause