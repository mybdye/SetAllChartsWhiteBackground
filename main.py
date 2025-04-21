import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from openpyxl.chart._chart import ChartBase  # 更新导入路径
from pathlib import Path
from tkinterdnd2 import TkinterDnD, DND_FILES  # 引入 tkinterdnd2 库

def set_background_white(sheet):
    """将工作表的图表区和绘图区背景填充为纯白色"""
    # 使用 openpyxl 提供的公共方法获取图表对象
    for chart in sheet._charts:  # 修改：直接使用私有属性 sheet._charts 获取图表对象
        try:
            chart.plotArea.spPr.solidFill.srgbClr.val = "FFFFFF"
        except AttributeError:
            # 如果图表没有背景填充属性，跳过
            continue
    # 图片处理暂不支持，避免直接访问私有属性

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
        raise

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
    root = TkinterDnD.Tk()  # 使用 TkinterDnD 的 Tk 类
    root.title("Excel图表背景填充工具")
    root.geometry("400x200")

    label = tk.Label(root, text="将Excel文件拖拽到此窗口", font=("Arial", 16))
    label.pack(pady=20)

    root.drop_target_register(DND_FILES)  # 使用 tkinterdnd2 的 DND_FILES
    root.dnd_bind('<<Drop>>', on_drop)

    root.mainloop()

if __name__ == "__main__":
    main()