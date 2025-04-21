import os
import platform
import subprocess
import sys
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def install_pyinstaller():
    """检查并安装 PyInstaller"""
    try:
        import pyinstaller
    except ImportError:
        logging.info("PyInstaller 未安装，正在安装...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
        logging.info("PyInstaller 安装完成。")

def install_wine():
    """检查并安装 wine"""
    try:
        # 检查 wine 是否支持 64 位
        wine_version = subprocess.check_output(['wine', '--version'], text=True)
        if 'wine-32' in wine_version.lower():
            logging.error("当前安装的 wine 是 32 位版本，不支持 64 位打包。")
            sys.exit(1)
    except FileNotFoundError:
        logging.info("wine 未安装，正在安装...")
        try:
            # 首先确保 Homebrew 已安装
            subprocess.check_call(['brew', '--version'])
            # 安装 wine 的依赖
            subprocess.check_call(['brew', 'install', 'gcenx/wine/wine-crossover'])
            logging.info("wine 安装完成。")
        except subprocess.CalledProcessError as e:
            logging.error(f"安装 wine 失败: {e}")
            sys.exit(1)

def build_installer():
    """主打包函数，负责调用其他模块完成打包任务"""
    # 检查并安装 PyInstaller
    install_pyinstaller()

    # 获取当前操作系统
    current_os = platform.system().lower()

    # 定义PyInstaller命令
    pyinstaller_cmd = [
        'pyinstaller',
        '--onefile',  # 打包成单个可执行文件
        '--windowed',  # 不显示控制台窗口（适用于GUI应用）
        '--name=SetAllChartsWhiteBackground',  # 可执行文件名称
        'main.py'  # 主程序入口文件
    ]

    # 根据操作系统调整命令
    if current_os == 'windows':
        pyinstaller_cmd.append('--icon=app.ico')  # Windows图标文件
    elif current_os == 'darwin':  # macOS
        # 在 macOS 下打包 Windows exe 安装程序
        install_wine()
        # 使用 wine 来执行 pyinstaller 命令
        pyinstaller_cmd = ['wine', 'pyinstaller'] + pyinstaller_cmd[1:] + ['--icon=app.ico', '--clean']
    else:
        logging.error(f"不支持的操作系统: {current_os}")
        return

    # 执行打包命令
    try:
        subprocess.run(pyinstaller_cmd, check=True)
        logging.info("打包完成，安装程序已生成在 dist 目录下。")
    except subprocess.CalledProcessError as e:
        logging.error(f"打包失败: {e}")
        sys.exit(1)

    # 在 macOS 上额外打包 pkg
    if current_os == 'darwin':
        build_pkg()

if __name__ == "__main__":
    build_installer()