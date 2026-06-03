# 安卓 APK 构建指南

## 方式一：在 Windows 上使用 WSL2（推荐）

### 1. 安装 WSL2
```powershell
wsl --install -d Ubuntu-24.04
```

### 2. 在 WSL 中安装依赖
```bash
# 进入项目目录
cd /mnt/c/Users/ccc12/Documents/Codex/2026-06-03/quiz-app

# 安装 Python 依赖
sudo apt update
sudo apt install python3 python3-pip python3-venv openjdk-17-jdk -y

# 创建虚拟环境
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 3. 安装 Android SDK
```bash
# 安装 Android 命令行工具
sudo apt install android-sdk -y
# 或手动下载 commandlinetools

# 设置环境变量
export ANDROID_HOME=$HOME/Android/Sdk
export PATH=$PATH:$ANDROID_HOME/cmdline-tools/latest/bin
export PATH=$PATH:$ANDROID_HOME/platform-tools

# 安装必要的 SDK 组件
sdkmanager "platform-tools" "platforms;android-34" "build-tools;34.0.0"
```

### 4. 构建 APK
```bash
flet build apk \
  --app-name "刷题" \
  --app-identifier "com.quiz.app" \
  --android-adaptive-icon assets/icon.png \
  --include-packages "$(cat requirements.txt | tr '\n' ' ')" \
  main.py
```

APK 文件将生成在 `build/apk/` 目录下。

---

## 方式二：使用 GitHub Actions（无需本地环境）

创建 `.github/workflows/build-apk.yml`：

```yaml
name: Build Android APK

on:
  push:
    branches: [main]
  workflow_dispatch:

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Setup Java
        uses: actions/setup-java@v4
        with:
          distribution: 'temurin'
          java-version: '17'

      - name: Setup Android SDK
        uses: android-actions/setup-android@v3

      - name: Install dependencies
        run: pip install -r requirements.txt

      - name: Build APK
        run: |
          flet build apk \
            --app-name "刷题" \
            --app-identifier "com.quiz.app" \
            --include-packages "$(cat requirements.txt | tr '\n' ' ')" \
            main.py

      - name: Upload APK
        uses: actions/upload-artifact@v4
        with:
          name: quiz-app-apk
          path: build/apk/*.apk
```

---

## 方式三：使用云端 Linux 虚拟机

任何 Ubuntu 22.04+ 的云服务器，按方式一的步骤操作即可。

---

## 快速本地测试（桌面端）

在打包 APK 之前，建议先在桌面端测试：

```powershell
# 使用 flet CLI 运行
flet run main.py

# 或直接 Python 运行
python main.py
```

---

## 文件结构

```
quiz-app/
├── main.py              # 主应用代码
├── questions.xls         # 题库文件
├── requirements.txt      # Python 依赖
├── assets/               # 资源目录（放图标等）
│   └── icon.png
└── BUILD.md             # 本文件
```
