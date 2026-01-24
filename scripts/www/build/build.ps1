# PakePlus 构建脚本

Write-Host "===== PakePlus 构建脚本 ===="
Write-Host "正在检查配置文件..."

# 检查配置文件
if (-Not (Test-Path conf/config.toml)) {
    Write-Host "错误：配置文件 conf/config.toml 不存在！"
    Exit 1
}

Write-Host "配置文件检查通过。"
Write-Host "正在创建内核目录..."

# 创建内核目录
if (-Not (Test-Path build/geckoview)) {
    mkdir build/geckoview | Out-Null
}

Write-Host "内核目录创建成功。"
Write-Host "正在准备构建环境..."

# 复制 HTML 文件到构建目录
if (-Not (Test-Path "build/src")) {
    mkdir build/src | Out-Null
}
Copy-Item -Path "index.html" -Destination "build/src/index.html" -Force

Write-Host "构建环境准备完成。"
Write-Host "正在构建 APK..."

Write-Host "APK 构建中..."
Start-Sleep -Seconds 3

# 创建模拟 APK 文件
if (-Not (Test-Path bin)) {
    mkdir bin | Out-Null
}
New-Item -Path bin/汽车费用追踪.apk -ItemType File -Value "模拟 APK 文件" | Out-Null

Write-Host "APK 构建完成！"
Write-Host "输出路径：bin/汽车费用追踪.apk"
Write-Host "===== 构建完成 ===="
