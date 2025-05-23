import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from openpyxl.chart._chart import ChartBase  # 更新导入路径
from pathlib import Path
from tkinterdnd2 import TkinterDnD, DND_FILES  # 引入 tkinterdnd2 库
import logging
import sys

# 初始化日志
logging.basicConfig(
    filename='app.log',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def handle_exception(exc_type, exc_value, exc_traceback):
    """全局异常处理"""
    logging.error("未捕获异常", exc_info=(exc_type, exc_value, exc_traceback))
    messagebox.showerror("致命错误", "程序发生未预期错误，请查看日志文件")

sys.excepthook = handle_exception

def set_background_white(sheet):
    """将工作表的图表区和绘图区背景填充为纯白色"""
    # 使用 openpyxl 提供的公共方法获取图表对象
    if hasattr(sheet, '_charts'):  # 增加对私有属性的安全检查
        for chart in sheet._charts:
            try:
                if hasattr(chart.plotArea, 'spPr') and hasattr(chart.plotArea.spPr, 'solidFill'):
                    chart.plotArea.spPr.solidFill.srgbClr.val = "FFFFFF"
            except AttributeError as e:
                logging.warning(f"图表属性访问失败: {e}")  # 更细粒度的异常处理
                continue

# 定义版本号
__version__ = "1.0.0"

def process_file(file_path):
    """处理Excel文件，将所有工作表的图表区和绘图区背景填充为纯白色"""
    try:
        wb = load_workbook(file_path)
        for sheet in wb.worksheets:
            set_background_white(sheet)
        output_path = Path(file_path).parent / f"X_{Path(file_path).name}"
        wb.save(output_path)
        return output_path
    except Exception as e:
        logging.error(f"处理文件时出错: {e}")
        messagebox.showerror("错误", f"处理文件时出错: {str(e)}")

def on_drop(event):
    """处理文件拖拽事件"""
    file_path = event.data.strip('{}').strip('"')  # 清理拖拽路径中的多余字符
    if file_path.lower().endswith('.xlsx'):
        try:
            output_path = process_file(file_path)
            messagebox.showinfo("处理完成", f"文件已处理并保存为:\n{output_path}")
        except FileNotFoundError:
            messagebox.showerror("错误", "文件未找到，请检查路径是否正确")
        except ValueError as ve:
            messagebox.showerror("错误", f"文件格式错误: {str(ve)}")
        except Exception as e:
            messagebox.showerror("错误", f"处理文件时出错: {str(e)}")
    else:
        messagebox.showwarning("警告", "请选择一个有效的Excel文件")

def main():
    """主程序入口"""
    try:
        root = TkinterDnD.Tk()  # 使用 TkinterDnD 的 Tk 类
        # 在 GUI 界面显示版本
        version = f"v{__version__}"
        if hasattr(sys, '_MEIPASS'):
            version += " (Bundle)"
        root.title(f"Excel图表背景填充工具 - {version}")
        root.geometry("400x200")

        label = tk.Label(root, text="将Excel文件拖拽到此窗口", font=("Arial", 16))
        label.pack(pady=20)

        root.drop_target_register(DND_FILES)  # 使用 tkinterdnd2 的 DND_FILES
        root.dnd_bind('<<Drop>>', on_drop)

        root.mainloop()
    except Exception as e:
        handle_exception(type(e), e, e.__traceback__)

if __name__ == "__main__":
    main()