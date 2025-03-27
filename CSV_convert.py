import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from tkinter import ttk
import os

class CSVConverterApp:
    def __init__(self, master):
        """
        应用程序主类初始化
        参数：
            master -- Tkinter根窗口对象
        功能：
            1. 初始化界面布局
            2. 创建导航栏和页面容器
            3. 配置转换/合并功能组件
        """
        self.master = master  # 设置主窗口
        master.title('CSV转Excel工具')#标题的名称
        master.geometry('1000x700')#灵活调整大小
       

        # 导航栏
        self.nav_frame = tk.Frame(master)#创建一个叫做nav_frame的Frame对象，它是主窗口master的一个子窗口。
        self.nav_frame.pack(side='top', fill='x', padx=10, pady=10)#将nav_frame放置在主窗口的顶部，填充整个窗口的宽度，左右各留出10个像素的间距，上下各留出10个像素的间距。

        # 导航按钮
        self.btn_converter = tk.Button(self.nav_frame, text='转换', command=self.show_converter_page,width=10,border=3,activebackground='LightGray',font=("Microsoft YaHei", 10, "bold"))#创建一个按钮，文本为"转换"，点击时调用show_converter_page函数，按钮宽度为10个像素，边框宽度为3个像素，鼠标悬停时背景色为亮灰色，字体为微软雅黑，字号为10，加粗。
        self.btn_combine = tk.Button(self.nav_frame, text='合并', command=self.show_combine_page,width=10,border=3,activebackground='LightGray',font=("Microsoft YaHei", 10, "bold"))#创建一个按钮，文本为"合并"，点击时调用show_combine_page函数，按钮宽度为10个像素，边框宽度为3个像素，鼠标悬停时背景色为亮灰色，字体为微软雅黑，字号为10，加粗。
        self.btn_converter.pack(side='left', padx=2,pady=5)#将转换按钮放置在导航栏的左侧，左右各留出2个像素的间距，上下各留出5个像素的间距。
        self.btn_combine.pack(side='left', padx=2,pady=5)#将合并按钮放置在导航栏的左侧，左右各留出2个像素的间距，上下各留出5个像素的间距。
        
        # 版本号标签
        self.lbl_version = tk.Label(self.nav_frame, text='V1.0.0', fg='#666666')#创建一个标签，文本为"V1.0.0"，前景色为灰色。
        self.lbl_version.pack(side='right', padx=10)#将版本号标签放置在导航栏的右侧，左右各留出10个像素的间距。

#--------------------------------------------------------------------------------------------------------------------------------------

        # 页面容器
        self.page_container = tk.Frame(master)#创建一个Frame对象，它是主窗口master的一个子窗口。
        self.page_container.pack(fill='both', expand=True)#将页面容器放置在主窗口的中心，填充整个窗口，并且在窗口大小改变时自适应。

        # 转换页面
        self.converter_page = tk.Frame(self.page_container)
        self.converter_page.pack(fill='both', expand=True)
        
        # 列选择组件
        self.selection_frame = tk.LabelFrame(self.converter_page, text='选择需要保留的列')#创建一个标签框架，文本为"选择需要保留的列"。
        self.check_vars = {#创建一个字典，用于存储复选框的变量。
            '时间': tk.BooleanVar(value=True),
            '源码': tk.BooleanVar(value=False),
            '物理量': tk.BooleanVar(value=True),
            '判决': tk.BooleanVar(value=False),
            '单位': tk.BooleanVar(value=False),
            '按文件名分组': tk.BooleanVar(value=True)
        }
        
        # 创建复选框
        self.checkboxes = []#创建一个空列表，用于存储复选框。
        for idx, (col, var) in enumerate(self.check_vars.items()):#遍历check_vars字典，获取列名和对应的布尔变量。
            cb = tk.Checkbutton(self.selection_frame, text=col, variable=var)#创建一个复选框，文本为列名，变量为布尔变量。
            cb.grid(row=0, column=idx, sticky='w', padx=5)#将复选框放置在标签框架的第一行，第idx列，左对齐，左右各留出5个像素的间距。
            self.checkboxes.append(cb)#将复选框添加到checkboxes列表中。
        self.selection_frame.pack(pady=5, fill='x', padx=10)


        self.button_frame = tk.Frame(self.converter_page)#新建一个叫button_frame的容器
        
        self.btn_select = tk.Button(self.button_frame, text='选择CSV文件', command=self.select_file)#创建一个按钮，文本为"选择CSV文件"，点击时调用select_file函数。
        self.btn_output_dir = tk.Button(self.button_frame, text='选择输出目录', command=self.select_output_dir)#创建一个按钮，文本为"选择输出目录"，点击时调用select_output_dir函数。
        self.btn_clear = tk.Button(self.button_frame, text='清除选择', command=self.clear_selection)#创建一个按钮
        self.lbl_output = tk.Label(self.button_frame, text='输出目录：未选择', anchor='w')

        # 水平排列按钮
        self.btn_select.pack(side='left', padx=5)
        self.btn_output_dir.pack(side='left', padx=5)
        self.btn_clear.pack(side='left', padx=5)
        self.lbl_output.pack(side='left', padx=5, fill='x', expand=True)
        
        self.button_frame.pack(pady=5, fill='x', padx=10)

        # 创建UI组件
        
        self.btn_convert = tk.Button(self.converter_page, text='转换', command=self.convert_file, state=tk.DISABLED, width=10, height=1,border=3, font=('Microsoft YaHei', 12, 'bold'))#创建一个按钮，文本为"转换"，点击时调用convert_file函数，按钮宽度为10个像素，高度为1个像素，边框宽度为3个像素，字体为微软雅黑，字号为12，加粗。
        self.btn_convert.pack(pady=5)
        self.label_status = tk.Label(self.converter_page, text='未选择CSV输入文件')#创建一个标签，文本为"未选择CSV输入文件"。
        self.label_status.pack(pady=5)
        self.progress = tk.Label(self.converter_page, text='')#创建一个标签，文本为空。
        self.progress.pack(pady=5)

        # 创建横向按钮容器
        
        # 底部显示区域
        self.display_frame = tk.Frame(self.converter_page)

        # 左侧输入文件列表
        self.input_frame = tk.LabelFrame(self.display_frame, text='输入文件列表',relief='flat',font=('SimSun',12,'italic','bold'),labelanchor='n')#创建一个标签框架，文本为"输入文件列表"，边框为无，字体为宋体，字号为12，斜体，加粗，标签锚点为北。
        self.input_frame.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')#将标签框架放置在显示区域的第一行，第一列，左右各留出5个像素的间距，上下各留出5个像素的间距，并且在区域大小改变时自适应。
        
        self.input_text = tk.Text(self.input_frame, wrap=tk.NONE, height=8)#创建一个文本框，不换行，高度为8个像素。
        self.input_scroll = tk.Scrollbar(self.input_frame, orient='vertical', command=self.input_text.yview)#创建一个滚动条，垂直方向，命令为input_text.yview。
        self.input_text.configure(yscrollcommand=self.input_scroll.set,state=tk.DISABLED)#将滚动条与文本框关联。
        self.input_scroll_x = tk.Scrollbar(self.input_frame, orient=tk.HORIZONTAL, command=self.input_text.xview)#创建一个滚动条，水平方向，命令为input_text.xview。
        self.input_text.configure(xscrollcommand=self.input_scroll_x.set,state=tk.DISABLED)#将滚动条与文本框关联。
        
        self.input_scroll_x.pack(side='bottom', fill='x')#将滚动条放置在底部，填充x轴。
        self.input_scroll.pack(side='right', fill='y')#将滚动条放置在右侧，填充y轴。
        self.input_text.pack(side='left', fill='both', expand=True)#将文本框放置在左侧，填充整个区域，并且在区域大小改变时自适应。


        # 右侧输出文件列表
        self.output_frame = tk.LabelFrame(self.display_frame, text='输出文件列表',relief='flat',font=('SimSun',12,'italic','bold'),labelanchor='n')
        self.output_frame.grid(row=0, column=1, padx=5, pady=5, sticky='nsew')
        
        self.output_text = tk.Text(self.output_frame, wrap=tk.NONE, height=8)
        self.output_scroll = tk.Scrollbar(self.output_frame, orient='vertical', command=self.output_text.yview)
        self.output_text.configure(yscrollcommand=self.output_scroll.set,state=tk.DISABLED)
        self.output_scroll_x = tk.Scrollbar(self.output_frame, orient=tk.HORIZONTAL, command=self.output_text.xview)
        self.output_text.configure(xscrollcommand=self.output_scroll_x.set,state=tk.DISABLED)
       
        self.output_scroll_x.pack(side='bottom', fill='x')#将滚动条放置在底部，填充x轴。
        self.output_scroll.pack(side='right', fill='y')#将滚动条放置在右侧，填充y轴。
        self.output_text.pack(side='left', fill='both', expand=True)#将文本框放置在左侧，填充整个区域，并且在区域大小改变时自适应。
        
        # 配置grid布局权重
        self.display_frame.columnconfigure(0, weight=1)#将第一列的权重设置为1。
        self.display_frame.columnconfigure(1, weight=1)#将第二列的权重设置为1。
        self.display_frame.rowconfigure(0, weight=1)#将第一行的权重设置为1。
        
        self.display_frame.pack(pady=10, fill='both', expand=True)

        #--------------------------------------------------------------------------------------------------------------------------------------
        #--------------------------------------------------------------------------------------------------------------------------------------
        #--------------------------------------------------------------------------------------------------------------------------------------
        # 初始化分析页面
        self.combine_page = tk.Frame(self.page_container)
        
        # 分析页面控件
        self.combine_btn_frame = tk.Frame(self.combine_page)
        
        self.btn_select_files = tk.Button(self.combine_btn_frame, text='选择XLSX文件', command=self.select_combine_files)
        self.btn_output_dir = tk.Button(self.combine_btn_frame, text='选择输出位置', command=self.select_combine_output_dir)
        self.btn_clear_combine = tk.Button(self.combine_btn_frame, text='清除选择', command=self.clear_combine)
        self.lbl_combine_output = tk.Label(self.combine_btn_frame, text='输出目录：未选择', anchor='w')


        self.btn_select_files.pack(side='left',padx=5)
        self.btn_output_dir.pack(side='left',padx=5)
        self.btn_clear_combine.pack(side='left',padx=5)
        self.lbl_combine_output.pack(side='left', padx=5, anchor='w', fill='x', expand=True)
        self.combine_btn_frame.pack(padx=10,pady=5, anchor='nw', side='top')

    
        
        self.btn_merge = tk.Button(self.combine_page,text='合并文件', command=self.merge_files, state=tk.DISABLED,width=10, height=1,border=3,font=('Microsoft YaHei',12,'bold'))
        self.btn_merge.pack(padx=5)
        
        self.combine_display_frame = tk.Frame(self.combine_page)

        # 文件列表显示
        self.combine_input_frame = tk.LabelFrame(self.combine_display_frame, text='输入文件列表',relief='flat',font=('SimSun',12,'italic','bold'),labelanchor='n')
        self.combine_input_frame.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')
        
        self.combine_input_text = tk.Text(self.combine_input_frame, height=5, wrap=tk.NONE)
        self.combine_input_scroll = tk.Scrollbar(self.combine_input_frame,orient='vertical',command=self.combine_input_text.yview)
        self.combine_input_text.configure(yscrollcommand=self.combine_input_scroll.set,state=tk.DISABLED)
        self.combine_input_scroll_x = tk.Scrollbar(self.combine_input_frame,orient=tk.HORIZONTAL,command=self.combine_input_text.xview)
        self.combine_input_text.configure(yscrollcommand=self.combine_input_scroll_x.set,state=tk.DISABLED)

        # 布局文件列表
        self.combine_input_scroll_x.pack(side='bottom',fill='x')
        self.combine_input_scroll.pack(side='right',fill='y')
        self.combine_input_text.pack(side='left', fill='both', expand=True)
        
        self.combine_output_frame = tk.LabelFrame(self.combine_display_frame, text='输入文件列表',relief='flat',font=('SimSun',12,'italic','bold'),labelanchor='n')
        self.combine_output_frame.grid(row=0, column=1, padx=5, pady=5, sticky='nsew')
        
        self.combine_output_text = tk.Text(self.combine_output_frame, height=5, wrap=tk.NONE)
        self.combine_output_scroll = tk.Scrollbar(self.combine_output_frame,orient='vertical',command=self.combine_output_text.yview)
        self.combine_output_text.configure(yscrollcommand=self.combine_output_scroll.set,state=tk.DISABLED)
        self.combine_output_scroll_x = tk.Scrollbar(self.combine_output_frame,orient=tk.HORIZONTAL,command=self.combine_output_text.xview)
        self.combine_output_text.configure(yscrollcommand=self.combine_output_scroll_x.set,state=tk.DISABLED)

        # 布局文件列表
        self.combine_output_scroll_x.pack(side='bottom',fill='x')
        self.combine_output_scroll.pack(side='right',fill='y')
        self.combine_output_text.pack(side='left', fill='both', expand=True)


# 配置grid布局权重
        self.combine_display_frame.columnconfigure(0, weight=1)#将第一列的权重设置为1。
        self.combine_display_frame.columnconfigure(1, weight=1)#将第二列的权重设置为1。
        self.combine_display_frame.rowconfigure(0, weight=1)#将第一行的权重设置为1。
        # 统计信息展示
       
        self.combine_display_frame.pack(pady=10, fill='both', expand=True)



        # 初始化数据存储
        self.combine_files = []
        self.merged_data = None

    def show_converter_page(self):#转换页面
        self.combine_page.pack_forget()
        self.converter_page.pack(fill='both', expand=True)

    def show_combine_page(self):#合并页面
        self.converter_page.pack_forget()
        self.combine_page.pack(fill='both', expand=True)

        self.file_path = ''

    def select_file(self):
        file_paths = filedialog.askopenfilenames(
            filetypes=[('CSV文件', '*.csv')],
            title='选择多个CSV文件',
            multiple=True
        )
        if file_paths:
            self.file_paths = list(file_paths)
            # 强制设置输出目录为首个文件所在目录
            self.output_dir = os.path.dirname(file_paths[0])
            self.lbl_output.config(text=f'输出目录：{self.output_dir}')
            self.master.update_idletasks()  # 强制刷新界面
            self.btn_convert.config(state=tk.NORMAL)
            
            self.progress.config(text='')
            
            # 更新输入文件列表
            self.input_text.config(state=tk.NORMAL)
            self.input_text.delete(1.0, tk.END)
            for path in file_paths:
                self.input_text.insert(tk.END, f'{path}\n')
            self.input_text.config(state=tk.DISABLED)

    def select_output_dir(self):
        selected_dir = filedialog.askdirectory(title='选择输出目录')
        if selected_dir:
            self.output_dir = selected_dir
            self.lbl_output.config(text=f'输出目录：{self.output_dir}')
        else:
            # 未选择时自动设置为第一个输入文件的目录
            if hasattr(self, 'file_paths') and self.file_paths:
                self.output_dir = os.path.dirname(self.file_paths[0])
                self.lbl_output.config(text=f'输出目录：{self.output_dir}')

    def convert_file(self):
        try:
            total_files = len(self.file_paths)
            success_count = 0
            error_messages = []

            for index, file_path in enumerate(self.file_paths, 1):
                try:
                    self.progress.config(text=f'正在处理文件 {index}/{total_files}: {os.path.basename(file_path)}')
                    self.master.update_idletasks()

                    # 读取CSV并获取列名
                    df = pd.read_csv(file_path, nrows=1)
                    csv_columns = df.columns.tolist()

                    # 根据勾选项过滤列
                    selected_keywords = [keyword for keyword, var in self.check_vars.items() if var.get()]
                    matched_cols = [col for col in csv_columns if any(keyword in col for keyword in selected_keywords)]

                    if not matched_cols:
                        raise ValueError('没有找到匹配的列，请重新选择筛选条件')

                    # 重新读取完整CSV数据
                    df = pd.read_csv(file_path, usecols=matched_cols)

                    # 生成保存路径
                    filename = os.path.basename(file_path).rsplit('.', 1)[0] + '.xlsx'
                    output_dir = self.output_dir if self.output_dir else os.path.dirname(file_path)
                    
                    if self.check_vars['按文件名分组'].get() and len(filename) >= 20:
                        group_code = filename[26:28]  # 第19-20位字符
                        group_dir = os.path.join(output_dir, group_code)
                        os.makedirs(group_dir, exist_ok=True)
                        save_path = os.path.join(group_dir, filename)
                    else:
                        save_path = os.path.join(output_dir, filename)

                    # 保存为Excel
                    df.to_excel(save_path, index=False)
                    success_count += 1
                    
                    # 更新输出文件列表
                    self.output_text.config(state=tk.NORMAL)
                    self.output_text.insert(tk.END, f'{save_path}\n')
                    self.output_text.config(state=tk.DISABLED)

                except Exception as e:
                    error_msg = f'文件 {os.path.basename(file_path)} 转换失败: {str(e)}'
                    error_messages.append(error_msg)

            # 最终状态报告
            result_msg = f'已完成 {success_count}/{total_files} 个文件'
            if error_messages:
                result_msg += '\n错误详情:\n' + '\n'.join(error_messages)
            
            
            self.progress.config(text=result_msg)
            messagebox.showinfo('转换完成', result_msg)
            
        except Exception as e:
            messagebox.showerror('错误', f'转换失败: {str(e)}')
            self.progress.config(text='转换失败，请检查文件格式')

    def select_combine_files(self):
        file_paths = filedialog.askopenfilenames(
            filetypes=[('Excel文件', '*.xlsx')],
            title='选择分析用Excel文件',
            multiple=True
        )
        if file_paths:
            self.combine_files = list(file_paths)
            self.btn_merge.config(state=tk.NORMAL)
            self.combine_input_text.configure(state=tk.NORMAL)
            self.combine_input_text.delete(1.0, tk.END)
            for path in file_paths:
                self.combine_input_text.insert(tk.END, f'{path}\n')
            self.combine_input_text.configure(state=tk.DISABLED)

    def select_combine_output_dir(self):
        selected_dir = filedialog.askdirectory(title='选择合并文件输出目录')
        if selected_dir:
            self.combine_output_dir = selected_dir
            self.lbl_combine_output.config(text=f'输出目录：{self.combine_output_dir}')

    def merge_files(self):
        try:
            if not self.combine_files:
                return

            # 生成输出文件名
            first_file = os.path.basename(self.combine_files[0])
            output_name = f"{first_file.rsplit('.', 1)[0]}_合并.xlsx"
            # 获取首个输入文件所在目录
            output_dir = os.path.dirname(self.combine_files[0])
            output_path = os.path.join(output_dir, output_name)

            dfs = [pd.read_excel(f) for f in self.combine_files]
            merged_data = pd.concat(dfs, ignore_index=True)
            
            # 更新状态显示
            self.lbl_combine_output.config(text=f'输出目录：{output_dir}')

            # 按第二列排序
            if len(merged_data.columns) >= 2:
                sort_column = merged_data.columns[1]
                merged_data = merged_data.sort_values(by=sort_column)
            
            # 添加统计信息
            min_values = merged_data.select_dtypes(include='number').min()
            max_values = merged_data.select_dtypes(include='number').max()
            
            # 创建统计行并追加
            stats_df = pd.DataFrame({
                '统计类型': ['最小值', '最大值']
            })

            
            for col in merged_data.columns:
                if pd.api.types.is_numeric_dtype(merged_data[col]):
                    stats_df[col] = [min_values[col], max_values[col]]
                else:
                    stats_df[col] = ['', '']
            
            merged_data = pd.concat([merged_data, stats_df], ignore_index=True)

            merged_data.to_excel(output_path, index=False)
            messagebox.showinfo('完成', f'文件已保存至: {output_path}')
            self.clear_combine()
            self.lbl_combine_output.config(text='输出目录：未选择')
           

        except Exception as e:
            messagebox.showerror('错误', f'文件保存失败: {str(e)}')

   

    def clear_combine(self):
        self.combine_files = []
        self.merged_data = None
        self.combine_input_text.configure(state=tk.NORMAL)
        self.combine_input_text.delete(1.0, tk.END)
        self.combine_input_text.configure(state=tk.DISABLED)
        self.btn_merge.config(state=tk.DISABLED)
        self.lbl_combine_output.config(text='输出目录：未选择')

    def clear_selection(self):
        self.file_paths = []
        self.output_dir = ''
        self.lbl_output.config(text='输出目录：未选择')
        self.label_status.config(text='已清除文件选择')
        self.btn_convert.config(state=tk.DISABLED)
        self.progress.config(text='')
        
        # 清空输入文本框
        self.input_text.config(state=tk.NORMAL)
        self.input_text.delete(1.0, tk.END)
        self.input_text.config(state=tk.DISABLED)
        
        # 清空输出文本框
        self.output_text.config(state=tk.NORMAL)
        self.output_text.delete(1.0, tk.END)
        self.output_text.config(state=tk.DISABLED)

    def create_file_list_frame(self, parent_frame, title):
        frame = tk.LabelFrame(parent_frame, text=title)
        text_widget = tk.Text(frame, wrap=tk.NONE, height=8)
        scroll = tk.Scrollbar(frame, orient='vertical', command=text_widget.yview)
        text_widget.configure(yscrollcommand=scroll.set)
        text_widget.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')
        return frame, text_widget

    def create_control_buttons(self, parent_frame):
        button_frame = tk.Frame(parent_frame)
        button_frame.pack(pady=5, fill='x', padx=10)
        return button_frame

    def setup_converter_ui(self):
        # 转换页面布局
        self.selection_frame.pack(pady=10, fill='x', padx=10)
        
        # 创建控制按钮
        self.button_frame = self.create_control_buttons(self.converter_page)
        
        # 文件列表框架
        self.input_frame, self.input_text = self.create_file_list_frame(self.display_frame, '输入文件列表')
        self.output_frame, self.output_text = self.create_file_list_frame(self.display_frame, '输出文件列表')
        
        # 配置grid布局
        self.display_frame.grid_columnconfigure(0, weight=1)
        self.display_frame.grid_columnconfigure(1, weight=1)
        self.display_frame.grid_rowconfigure(0, weight=1)

    def setup_combine_ui(self):
        # 分析页面布局
        self.combine_btn_frame = self.create_control_buttons(self.combine_page)
        
        # 文件列表框架
        self.combine_input_frame, self.combine_input_text = self.create_file_list_frame(self.combine_page, '输入文件列表')
        
        # 配置grid布局
        self.combine_page.grid_columnconfigure(0, weight=1)
        self.combine_page.grid_rowconfigure(1, weight=1)

    def process_conversion(self, file_paths):
        """
        执行批量文件转换核心流程
        处理步骤：
            1. 遍历所有选中的CSV文件
            2. 加载并过滤数据列
            3. 生成输出路径并保存为Excel
            4. 更新界面进度显示
            5. 异常时调用错误处理
        """
        # 文件处理逻辑与界面分离
        total_files = len(file_paths)
        for index, path in enumerate(file_paths, 1):
            try:
                df = self.load_and_filter_data(path)
                output_path = self.generate_output_path(path)
                self.save_excel_file(df, output_path)
                self.update_ui_progress(index, total_files, path)
            except Exception as e:
                self.handle_conversion_error(path, e)

    def load_and_filter_data(self, path):
        # 分离数据加载逻辑
        df = pd.read_csv(path, nrows=1)
        selected_cols = self.get_selected_columns(df.columns)
        return pd.read_csv(path, usecols=selected_cols)

    def handle_conversion_error(self, path, error):
        """
        统一转换错误处理机制
        功能：
            1. 记录错误日志
            2. 收集错误信息
            3. 弹出错误提示框
            4. 更新错误消息列表
        """
        # 统一错误处理
        error_msg = f'文件 {os.path.basename(path)} 转换失败: {str(error)}'
        self.error_messages.append(error_msg)
        messagebox.showerror('转换错误', error_msg)
        self.log_error(error)

if __name__ == '__main__':
    root = tk.Tk()
    app = CSVConverterApp(root)
    root.mainloop()
