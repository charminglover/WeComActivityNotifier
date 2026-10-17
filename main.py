"""Native Windows-style desktop application. Tk widgets are touched only by the UI thread."""
from __future__ import annotations

import json
import os
import queue
import threading
import traceback
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from engine import (DEFAULT_FIELDS, read_excel, filter_records, text_of, number, rate_number,
                    make_jobs, prepare_messages, read_mapping, WeComSender, check_webhook)

APP_NAME = '企业微信活动通知工具'
CONFIG_PATH = (Path(os.environ.get('APPDATA', str(Path.home()))) / 'WeComActivityNotifier' / 'preferences.json')


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f'{APP_NAME}  v2.0')
        self.geometry('1250x820')
        self.minsize(980, 640)
        self.style = ttk.Style(self)
        if 'vista' in self.style.theme_names():
            self.style.theme_use('vista')
        elif 'xpnative' in self.style.theme_names():
            self.style.theme_use('xpnative')
        self.option_add('*Font', ('Microsoft YaHei UI', 9))
        self.pool = ThreadPoolExecutor(max_workers=2, thread_name_prefix='activity-worker')
        self.events = queue.Queue()
        self.records = []
        self.fields = []
        self.field_flags = {}
        self.filtered = []
        self.checked = set()
        self.inflight = False
        self.stop_event = threading.Event()
        self._table_generation = 0
        self.display_limit = 2000
        # Delay Pillow's native module loading until the user imports a workbook.
        self._checkbox_images = {}
        self.prefs = self._load_prefs()
        self._init_vars()
        self._layout()
        self.after(100, self._pump)
        self.protocol('WM_DELETE_WINDOW', self._close)

    def _create_checkbox_images(self):
        """DPI-stable real checkboxes as Treeview images, not font-dependent glyphs."""
        from PIL import Image, ImageDraw, ImageTk

        result = {}
        for enabled in (False, True):
            im = Image.new('RGBA', (22, 22), (255, 255, 255, 0))
            d = ImageDraw.Draw(im)
            outline = '#2B579A' if enabled else '#8793A1'
            d.rounded_rectangle((3, 3, 19, 19), radius=3, fill='#2B579A' if enabled else 'white',
                                outline=outline, width=2)
            if enabled:
                d.line([(7, 11), (10, 14), (16, 7)], fill='white', width=3, joint='curve')
            result[enabled] = ImageTk.PhotoImage(im, master=self)
        return result

    def _load_prefs(self):
        try:
            return json.loads(CONFIG_PATH.read_text(encoding='utf8'))
        except (OSError, ValueError):
            return {}

    def _init_vars(self):
        p = self.prefs
        self.file_var = tk.StringVar()
        self.shop_var = tk.StringVar()
        self.act_var = tk.StringVar()
        self.status_var = tk.StringVar(value='全部')
        self.metric_var = tk.StringVar(value='任务执行率')
        self.comp_var = tk.StringVar(value='<')
        self.threshold_var = tk.StringVar()
        self.qty_metric = tk.StringVar(value='任务台数')
        self.qty_comp = tk.StringVar(value='>')
        self.qty_threshold = tk.StringVar()
        self.msg_title = tk.StringVar(value=p.get('msg_title','售后活动执行进度通知'))
        self.mode_var = tk.StringVar(value=p.get('mode','图片表格'))
        self.page_var = tk.StringVar(value=str(p.get('page_size',12)))
        self.column_var = tk.StringVar(value=str(p.get('col_size',6)))
        self.warn_var = tk.StringVar(value=str(p.get('warn_rate',80)))
        self.group_var = tk.BooleanVar(value=p.get('group_by_shop',False))
        self.kind_var = tk.StringVar(value=p.get('kind','群机器人Webhook'))
        self.webhook_var = tk.StringVar()    # Credentials are deliberately NOT saved.
        self.corpid_var = tk.StringVar()
        self.secret_var = tk.StringVar()
        self.agentid_var = tk.StringVar()
        self.touser_var = tk.StringVar()
        self.mapping_path = tk.StringVar()
        self.count_var = tk.StringVar(value='尚未导入文件')
        self.status_line = tk.StringVar(value='就绪。请先导入 Excel 文件。')

    def _layout(self):
        root = ttk.Frame(self, padding=10)
        root.pack(fill='both', expand=True)
        top = ttk.Frame(root)
        top.pack(fill='x', pady=(0,8))
        ttk.Label(top,text='Excel 文件：').pack(side='left')
        ttk.Entry(top,textvariable=self.file_var,state='readonly').pack(side='left',fill='x',expand=True,padx=(2,6))
        ttk.Button(top,text='选择文件…',command=self.choose_excel).pack(side='left')
        ttk.Label(root,textvariable=self.count_var).pack(anchor='w',pady=(0,6))
        self.tabs = ttk.Notebook(root)
        self.tabs.pack(fill='both',expand=True)
        self.page_data = ttk.Frame(self.tabs,padding=10)
        self.page_fields = ttk.Frame(self.tabs,padding=10)
        self.page_send = ttk.Frame(self.tabs,padding=10)
        self.tabs.add(self.page_data,text='1  数据筛选与记录勾选')
        self.tabs.add(self.page_fields,text='2  推送内容与格式')
        self.tabs.add(self.page_send,text='3  企业微信配置与发送')
        self._data_tab()
        self._fields_tab()
        self._send_tab()
        foot = ttk.Frame(root)
        foot.pack(fill='x',pady=(8,0))
        ttk.Label(foot,textvariable=self.status_line).pack(side='left',fill='x',expand=True)
        self.bar=ttk.Progressbar(foot,mode='indeterminate',length=180)
        self.bar.pack(side='right')

    def _data_tab(self):
        filters = ttk.LabelFrame(self.page_data,text='筛选条件（同一行的条件按 AND 同时生效）',padding=10)
        filters.pack(fill='x')
        for i in range(12):
            filters.columnconfigure(i,weight=1 if i in (1,4) else 0)
        ttk.Label(filters,text='服务店/区域').grid(row=0,column=0,sticky='w')
        ttk.Entry(filters,textvariable=self.shop_var,width=23).grid(row=0,column=1,sticky='ew',padx=(4,15))
        ttk.Label(filters,text='活动名称/编码').grid(row=0,column=2,sticky='w')
        ttk.Entry(filters,textvariable=self.act_var,width=24).grid(row=0,column=3,sticky='ew',padx=(4,15))
        ttk.Label(filters,text='活动状态').grid(row=0,column=4,sticky='w')
        self.status_select=ttk.Combobox(filters,textvariable=self.status_var,values=['全部'],state='readonly',width=13)
        self.status_select.grid(row=0,column=5,sticky='w',padx=(4,15))
        ttk.Button(filters,text='执行筛选',command=self.apply_filters).grid(row=0,column=6,padx=5)
        ttk.Button(filters,text='清除条件',command=self.clear_filters).grid(row=0,column=7,padx=5)
        ttk.Label(filters,text='完成率条件').grid(row=1,column=0,sticky='w',pady=(10,0))
        self.metric_select=ttk.Combobox(filters,textvariable=self.metric_var,width=20,state='readonly')
        self.metric_select.grid(row=1,column=1,sticky='w',pady=(10,0))
        ttk.Combobox(filters,textvariable=self.comp_var,values=['<','<=','=','>=','>'],width=4,state='readonly').grid(row=1,column=2,sticky='e',pady=(10,0))
        ttk.Entry(filters,textvariable=self.threshold_var,width=8).grid(row=1,column=3,sticky='w',pady=(10,0))
        ttk.Label(filters,text='(例如 80，空白不限制)').grid(row=1,column=4,columnspan=2,sticky='w',pady=(10,0))
        ttk.Label(filters,text='数量条件').grid(row=2,column=0,sticky='w',pady=(10,0))
        self.qty_select=ttk.Combobox(filters,textvariable=self.qty_metric,width=20,state='readonly')
        self.qty_select.grid(row=2,column=1,sticky='w',pady=(10,0))
        ttk.Combobox(filters,textvariable=self.qty_comp,values=['<','<=','=','>=','>'],width=4,state='readonly').grid(row=2,column=2,sticky='e',pady=(10,0))
        ttk.Entry(filters,textvariable=self.qty_threshold,width=8).grid(row=2,column=3,sticky='w',pady=(10,0))
        ttk.Label(filters,text='(例如 5，空白不限制)').grid(row=2,column=4,columnspan=2,sticky='w',pady=(10,0))
        tools=ttk.Frame(self.page_data)
        tools.pack(fill='x',pady=(10,6))
        ttk.Button(tools,text='勾选当前筛选结果',command=lambda:self.select_filtered('all')).pack(side='left',padx=(0,5))
        ttk.Button(tools,text='反选当前筛选结果',command=lambda:self.select_filtered('invert')).pack(side='left',padx=(0,5))
        ttk.Button(tools,text='勾选高亮行',command=self.select_highlight).pack(side='left',padx=(0,5))
        ttk.Button(tools,text='取消所有勾选',command=lambda:self.select_filtered('none')).pack(side='left',padx=(0,5))
        ttk.Label(tools,text='单击左侧复选框，或选中行后按空格；双击行也可切换').pack(side='right')
        area=ttk.Frame(self.page_data)
        area.pack(fill='both',expand=True)
        display=['任务服务店编码','任务服务店名称','活动编码','活动名称','活动状态','任务台数','任务车辆执行台数','任务执行率','按期执行率']
        self.tree=ttk.Treeview(area,columns=display,show='tree headings',selectmode='extended')
        self.tree.heading('#0', text='勾选')
        self.tree.column('#0', width=62, minwidth=62, stretch=False, anchor='center')
        self.tree.tag_configure('checked', background='#EBF2FB')
        widths=[135,225,100,300,100,85,115,100,100]
        for f,w in zip(display,widths):
            self.tree.heading(f,text=f)
            self.tree.column(f,width=w,minwidth=60,anchor='center' if f in ('任务台数','任务车辆执行台数','任务执行率','按期执行率') else 'w')
        self.tree.grid(row=0,column=0,sticky='nsew')
        sv=ttk.Scrollbar(area,orient='vertical',command=self.tree.yview)
        sh=ttk.Scrollbar(area,orient='horizontal',command=self.tree.xview)
        self.tree.configure(yscrollcommand=sv.set,xscrollcommand=sh.set)
        sv.grid(row=0,column=1,sticky='ns'); sh.grid(row=1,column=0,sticky='ew')
        area.columnconfigure(0,weight=1); area.rowconfigure(0,weight=1)
        self.tree.bind('<ButtonRelease-1>', self.click_tree_checkbox)
        self.tree.bind('<Double-1>', self.toggle_tree_row)
        self.tree.bind('<space>', self.toggle_selected_rows)
        self.tree.bind('<Return>', self.toggle_selected_rows)
        self.btn_load_more=ttk.Button(self.page_data,text='加载更多行',command=self.show_more_rows)
        self.btn_load_more.pack(anchor='center',pady=(4,0))
        self.btn_load_more.pack_forget()

    def _fields_tab(self):
        ttk.Label(self.page_fields,text='勾选要推送的 Excel 列。支持改变列顺序；这些设置直接决定最终消息内容。').pack(anchor='w',pady=(0,10))
        row=ttk.Frame(self.page_fields)
        row.pack(fill='both',expand=True)
        holder=ttk.Frame(row)
        holder.pack(side='left',fill='both',expand=True)
        self.fields_tree=ttk.Treeview(holder,columns=('column',),show='tree headings',selectmode='browse')
        self.fields_tree.heading('#0',text='推送')
        self.fields_tree.column('#0',width=72,stretch=False,anchor='center')
        self.fields_tree.heading('column',text='Excel 原始列名'); self.fields_tree.column('column',width=340,anchor='w')
        self.fields_tree.tag_configure('checked', background='#EBF2FB')
        self.fields_tree.pack(side='left',fill='both',expand=True)
        fscroll=ttk.Scrollbar(holder,orient='vertical',command=self.fields_tree.yview)
        fscroll.pack(side='right',fill='y')
        self.fields_tree.configure(yscrollcommand=fscroll.set)
        self.fields_tree.bind('<Double-1>',self.double_click_field)
        self.fields_tree.bind('<ButtonRelease-1>',self.click_field_checkbox)
        self.fields_tree.bind('<space>', lambda event: self.toggle_field())
        self.fields_tree.bind('<Return>', lambda event: self.toggle_field())
        cmd=ttk.Frame(row,padding=(12,0))
        cmd.pack(side='left',fill='y')
        ttk.Button(cmd,text='切换勾选',command=self.toggle_field).pack(fill='x',pady=3)
        ttk.Button(cmd,text='上移 ↑',command=lambda:self.move_field(-1)).pack(fill='x',pady=3)
        ttk.Button(cmd,text='下移 ↓',command=lambda:self.move_field(1)).pack(fill='x',pady=3)
        ttk.Separator(cmd).pack(fill='x',pady=10)
        ttk.Button(cmd,text='推荐字段',command=lambda:self.set_fields('recommended')).pack(fill='x',pady=3)
        ttk.Button(cmd,text='全部勾选',command=lambda:self.set_fields('all')).pack(fill='x',pady=3)
        ttk.Button(cmd,text='全部清空',command=lambda:self.set_fields('none')).pack(fill='x',pady=3)
        settings=ttk.LabelFrame(self.page_fields,text='通知内容设置',padding=10)
        settings.pack(fill='x',pady=(12,0))
        ttk.Label(settings,text='通知标题').grid(row=0,column=0,sticky='w')
        ttk.Entry(settings,textvariable=self.msg_title,width=38).grid(row=0,column=1,sticky='w',padx=8)
        ttk.Label(settings,text='发送格式').grid(row=0,column=2,sticky='w',padx=(15,0))
        ttk.Combobox(settings,textvariable=self.mode_var,values=['图片表格','Markdown表格'],width=16,state='readonly').grid(row=0,column=3,sticky='w',padx=8)
        ttk.Label(settings,text='自动适配：按企业微信字节数与图片大小拆分行、列，无需设置每页条数。').grid(row=1,column=0,columnspan=4,sticky='w',pady=(12,0))
        ttk.Label(settings,text='低完成率警示阈值 %').grid(row=2,column=0,sticky='w',pady=(12,0))
        ttk.Entry(settings,textvariable=self.warn_var,width=8).grid(row=2,column=1,sticky='w',padx=8,pady=(12,0))
        ttk.Checkbutton(settings,text='按服务店分别分组（可以配合接收目标映射表使用）',variable=self.group_var).grid(row=3,column=0,columnspan=4,sticky='w',pady=(12,0))
        ttk.Button(settings,text='校验并预览消息',command=self.preview).grid(row=4,column=0,columnspan=2,sticky='w',pady=(12,4))
        ttk.Label(settings,text='说明：所有消息发送前会完整校验；Markdown 表格仅限群机器人。').grid(row=4,column=2,columnspan=4,sticky='w')

    def _send_tab(self):
        section=ttk.LabelFrame(self.page_send,text='企业微信发送方式',padding=12)
        section.pack(fill='x')
        ttk.Label(section,text='发送方式').grid(row=0,column=0,sticky='w')
        sender_combo=ttk.Combobox(section,values=['群机器人Webhook','自建应用消息'],textvariable=self.kind_var,
                     state='readonly',width=22)
        sender_combo.grid(row=0,column=1,sticky='w',padx=7)
        sender_combo.bind('<<ComboboxSelected>>', lambda event: self._toggle_sender())
        self.bot_form=ttk.Frame(section)
        self.bot_form.grid(row=1,column=0,columnspan=4,sticky='ew',pady=(12,0))
        ttk.Label(self.bot_form,text='Webhook URL').grid(row=0,column=0,sticky='w')
        ttk.Entry(self.bot_form,textvariable=self.webhook_var,width=82,show='').grid(row=0,column=1,padx=8,sticky='ew')
        self.app_form=ttk.Frame(section)
        self.app_form.grid(row=2,column=0,columnspan=4,sticky='ew',pady=(12,0))
        for i,(label,var) in enumerate((('CorpID',self.corpid_var),('Secret',self.secret_var),('AgentID',self.agentid_var),('接收人UserID（用 | 分隔）',self.touser_var))):
            ttk.Label(self.app_form,text=label).grid(row=i,column=0,sticky='w',pady=2)
            ttk.Entry(self.app_form,textvariable=var,width=76,show='●' if label=='Secret' else '').grid(row=i,column=1,sticky='ew',padx=8,pady=2)
        self._toggle_sender()
        m=ttk.LabelFrame(self.page_send,text='服务店接收目标映射（可选）',padding=12)
        m.pack(fill='x',pady=(12,0))
        ttk.Label(m,text='按店发送时可导入 CSV：任务服务店编码,Webhook（群机器人）或 任务服务店编码,UserID（应用）').pack(anchor='w')
        f=ttk.Frame(m)
        f.pack(fill='x',pady=(8,0))
        ttk.Entry(f,textvariable=self.mapping_path,state='readonly').pack(side='left',fill='x',expand=True)
        ttk.Button(f,text='导入 CSV…',command=self.choose_map).pack(side='left',padx=5)
        ttk.Button(f,text='移除映射',command=lambda:self.mapping_path.set('')).pack(side='left')
        ttk.Label(m,text='安全规则：只要导入映射表，就强制按服务店分组；缺少映射的服务店将阻止整批发送。').pack(anchor='w',pady=(8,0))
        actions=ttk.Frame(self.page_send)
        actions.pack(fill='x',pady=(12,8))
        self.btn_send=ttk.Button(actions,text='确认并发送',command=self.send_now)
        self.btn_send.pack(side='left')
        self.btn_stop=ttk.Button(actions,text='停止后续发送',command=self.stop_sending,state='disabled')
        self.btn_stop.pack(side='left',padx=8)
        ttk.Button(actions,text='预览消息',command=self.preview).pack(side='left',padx=8)
        ttk.Label(actions,text='只会在点击“确认并发送”且再次确认后联网发送。').pack(side='right')
        log_frame=ttk.LabelFrame(self.page_send,text='发送日志（不保存密钥或 Webhook）',padding=7)
        log_frame.pack(fill='both',expand=True)
        self.log_text=tk.Text(log_frame,height=15,wrap='word',state='disabled',font=('Consolas',10))
        self.log_text.pack(side='left',fill='both',expand=True)
        log_scr=ttk.Scrollbar(log_frame,command=self.log_text.yview)
        log_scr.pack(side='right',fill='y')
        self.log_text.configure(yscrollcommand=log_scr.set)

    def _toggle_sender(self):
        if self.kind_var.get() == '群机器人Webhook':
            self.bot_form.grid(); self.app_form.grid_remove()
        else:
            self.app_form.grid(); self.bot_form.grid_remove()

    def choose_excel(self):
        path=filedialog.askopenfilename(title='选择 Excel 文件',filetypes=[('Excel / CSV','*.xlsx *.xlsm *.xls *.csv'),('所有文件','*.*')])
        if not path: return
        self.file_var.set(path)
        self._submit(lambda: read_excel(path), 'loaded', '正在解析 Excel，请稍候…')

    def _submit(self, func, eventname, status):
        if self.inflight:
            messagebox.showwarning('当前有任务','请先完成或停止当前任务')
            return
        self.inflight=True
        self.status_line.set(status)
        self.bar.start(10)
        def task():
            try:
                self.events.put((eventname, func()))
            except Exception as e:
                self.events.put(('error', str(e)))
        self.pool.submit(task)

    def _pump(self):
        try:
            while True:
                typ, payload = self.events.get_nowait()
                if typ == 'loaded':
                    self._idle()
                    fields, records = payload
                    self.records=records
                    self.checked.clear()
                    self.fields=fields
                    self._fill_fields()
                    self.status_select.configure(values=['全部']+sorted({text_of(r.get('活动状态')) for r in records if text_of(r.get('活动状态'))}))
                    rfields=[c for c in fields if any(x in c for x in ('率','完成率'))]
                    cfields=[c for c in fields if ('台数' in c or '数量' in c)]
                    self.metric_select.configure(values=rfields)
                    self.qty_select.configure(values=cfields)
                    if self.metric_var.get() not in rfields and rfields: self.metric_var.set(rfields[0])
                    if self.qty_metric.get() not in cfields and cfields: self.qty_metric.set(cfields[0])
                    self.apply_filters()
                    self._log(f'导入成功：{len(records)} 行，{len(fields)} 列。')
                elif typ == 'filtered':
                    self._idle()
                    self.filtered=payload
                    self._render_records()
                elif typ == 'preview':
                    self._idle()
                    self._show_preview(*payload)
                elif typ == 'planned':
                    self._idle()
                    self._confirm_and_send(payload)
                elif typ == 'progress':
                    self._log(payload)
                    self.status_line.set(payload)
                elif typ == 'sent':
                    self._idle()
                    success,total,failed=payload
                    self.btn_stop.config(state='disabled')
                    self._log(f'发送结束：成功 {success}/{total} 条，失败 {failed} 条。')
                    messagebox.showinfo('发送结束',f'成功 {success}/{total} 条，失败 {failed} 条。\n详细内容请查看日志。')
                elif typ == 'error':
                    self._idle()
                    self._log('错误：'+payload)
                    messagebox.showerror('执行失败',payload)
        except queue.Empty:
            pass
        self.after(100,self._pump)

    def _idle(self):
        self.inflight=False
        self.bar.stop()
        self.status_line.set('就绪')
        self.btn_send.config(state='normal')
        self.btn_stop.config(state='disabled')

    def _fill_fields(self):
        if not self._checkbox_images:
            self._checkbox_images = self._create_checkbox_images()
        last=set(self.prefs.get('fields', DEFAULT_FIELDS))
        preferred=[x for x in self.prefs.get('order',[]) if x in self.fields]
        self.fields=preferred+[x for x in self.fields if x not in preferred]
        self.field_flags={x:x in last for x in self.fields}
        self._refresh_fields()

    def _refresh_fields(self, focus=None):
        self.fields_tree.delete(*self.fields_tree.get_children())
        for i,f in enumerate(self.fields):
            enabled=bool(self.field_flags.get(f))
            self.fields_tree.insert('', 'end', iid=str(i), text='', image=self._checkbox_images[enabled],
                                    tags=('checked',) if enabled else (), values=(f,))
        if focus is not None and 0 <= focus < len(self.fields):
            self.fields_tree.selection_set(str(focus)); self.fields_tree.see(str(focus))

    def click_field_checkbox(self,event):
        if self.fields_tree.identify_column(event.x) != '#0': return
        item=self.fields_tree.identify_row(event.y)
        if item:
            self.fields_tree.selection_set(item)
            self.toggle_field()

    def double_click_field(self,event):
        if self.fields_tree.identify_column(event.x) != '#0':
            item=self.fields_tree.identify_row(event.y)
            if item:
                self.fields_tree.selection_set(item)
                self.toggle_field()

    def toggle_field(self,event=None):
        selection=self.fields_tree.selection()
        if not selection: return
        index=int(selection[0]); f=self.fields[index]
        self.field_flags[f]=not self.field_flags.get(f,False)
        enabled=self.field_flags[f]
        self.fields_tree.item(str(index),image=self._checkbox_images[enabled],
                              tags=('checked',) if enabled else ())

    def move_field(self, delta):
        selection=self.fields_tree.selection()
        if not selection: return
        index=int(selection[0]); nextindex=index+delta
        if not 0<=nextindex<len(self.fields): return
        self.fields[index],self.fields[nextindex]=self.fields[nextindex],self.fields[index]
        self._refresh_fields(nextindex)

    def set_fields(self, mode):
        for f in self.fields:
            self.field_flags[f]=mode=='all' or (mode=='recommended' and f in DEFAULT_FIELDS)
        self._refresh_fields()

    def clear_filters(self):
        self.shop_var.set(''); self.act_var.set('');self.status_var.set('全部')
        self.threshold_var.set('');self.qty_threshold.set('')
        self.apply_filters()

    def apply_filters(self):
        if not self.records: return
        options=dict(shop=self.shop_var.get(),activity=self.act_var.get(),
                     status=self.status_var.get(),metric=self.metric_var.get(),
                     operator=self.comp_var.get(),threshold=self.threshold_var.get(),
                     count_metric=self.qty_metric.get(),count_operator=self.qty_comp.get(),
                     count_threshold=self.qty_threshold.get())
        if len(self.records) > 8000:
            self._submit(lambda: filter_records(self.records, **options), 'filtered', '正在后台筛选…')
            return
        try:
            self.filtered=filter_records(self.records,**options)
        except ValueError as e:
            messagebox.showwarning('筛选条件',str(e));return
        self._render_records()

    def _render_records(self):
        self._table_generation+=1
        tag=self._table_generation
        self.display_limit=2000
        self.tree.delete(*self.tree.get_children())
        self._render_chunk(0,tag)

    def _render_chunk(self,start,tag):
        if tag!=self._table_generation: return
        target=min(len(self.filtered),self.display_limit)
        end=min(start+125,target)
        keys=['任务服务店编码','任务服务店名称','活动编码','活动名称','活动状态','任务台数','任务车辆执行台数','任务执行率','按期执行率']
        for index in range(start,end):
            r=self.filtered[index]
            enabled=r['_id'] in self.checked
            vals=[text_of(r.get(f),f) for f in keys]
            self.tree.insert('', 'end', iid=str(index), text='', image=self._checkbox_images[enabled],
                             tags=('checked',) if enabled else (), values=vals)
        self._update_count()
        if end<target:
            self.after(1,lambda:self._render_chunk(end,tag))
        else:
            if target<len(self.filtered):
                if not self.btn_load_more.winfo_manager():self.btn_load_more.pack(anchor='center',pady=(4,0))
            else:
                if self.btn_load_more.winfo_manager():self.btn_load_more.pack_forget()

    def show_more_rows(self):
        start=min(self.display_limit,len(self.filtered))
        self.display_limit+=2000
        self._render_chunk(start,self._table_generation)

    def _update_count(self):
        total=len(self.filtered)
        visible=min(total,self.display_limit)
        more=f'  |  当前展示前 {visible} 条（选择操作仍作用于全部筛选结果）' if visible<total else ''
        self.count_var.set(f'数据总计 {len(self.records)} 条  |  当前筛选 {total} 条  |  已勾选 {len(self.checked)} 条'+more)

    def _toggle_row(self,item):
        r=self.filtered[int(item)]
        key=r['_id']
        if key in self.checked: self.checked.remove(key)
        else: self.checked.add(key)
        enabled=key in self.checked
        self.tree.item(item,image=self._checkbox_images[enabled],tags=('checked',) if enabled else ())

    def click_tree_checkbox(self,event):
        if self.tree.identify_column(event.x) == '#0':
            item=self.tree.identify_row(event.y)
            if item:self._toggle_row(item);self._update_count()

    def toggle_tree_row(self,event):
        if self.tree.identify_column(event.x) == '#0': return
        item=self.tree.identify_row(event.y)
        if item: self._toggle_row(item);self._update_count()

    def toggle_selected_rows(self,event=None):
        items=self.tree.selection()
        if not items:return 'break'
        for item in items:self._toggle_row(item)
        self._update_count()
        return 'break'

    def select_highlight(self):
        for item in self.tree.selection():
            self.checked.add(self.filtered[int(item)]['_id'])
            self.tree.item(item,image=self._checkbox_images[True],tags=('checked',))
        self._update_count()

    def select_filtered(self,mode):
        if not self.records:return
        keys={r['_id'] for r in self.filtered}
        if mode=='all':self.checked.update(keys)
        elif mode=='none':self.checked.clear()
        elif mode=='invert':self.checked.symmetric_difference_update(keys)
        for item in self.tree.get_children():
            enabled=self.filtered[int(item)]['_id'] in self.checked
            self.tree.item(item,image=self._checkbox_images[enabled],tags=('checked',) if enabled else ())
        self._update_count()

    def _snapshot(self):
        selected=[r for r in self.records if r['_id'] in self.checked]
        if not selected:
            raise ValueError('请在第 1 页勾选至少一条活动记录')
        fields=[f for f in self.fields if self.field_flags.get(f)]
        if not fields:
            raise ValueError('请在第 2 页勾选至少一个推送字段')
        warning=number(self.warn_var.get())
        if warning is None or not 0<=warning<=100:
            raise ValueError('警示阈值必须是 0-100 之间的数字')
        mode=self.mode_var.get()
        group=self.group_var.get() or bool(self.mapping_path.get())
        return dict(selected=selected,fields=fields,warning=warning,mode=mode,group=group,
                    title=self.msg_title.get().strip() or '售后活动执行进度通知')

    @staticmethod
    def _prepare(snapshot):
        jobs=make_jobs(snapshot['selected'],snapshot['fields'],group_by_shop=snapshot['group'],
                       page_size=12, columns_per_page=6,title=snapshot['title'])
        return prepare_messages(jobs,snapshot['mode'],snapshot['warning'])

    def preview(self):
        if self.inflight:return
        try:
            snapshot=self._snapshot()
        except ValueError as e:
            messagebox.showwarning('无法预览',str(e));return
        def process():
            messages=self._prepare(snapshot)
            return messages[0].payload,snapshot['mode'],len(messages)
        self._submit(process,'preview','后台校验并生成全部消息…')

    def _show_preview(self, preview, mode, total):
        w=tk.Toplevel(self)
        w.title(f'消息预览 - 自动拆分为 {total} 条消息')
        w.geometry('1050x740')
        w.transient(self)
        ttk.Label(w,text=f'已逐条校验，预计 {total} 条消息。这里显示第一条；只有确认发送后才会推送。',padding=10).pack(anchor='w')
        if mode=='图片表格':
            import io
            from PIL import Image, ImageTk

            with Image.open(io.BytesIO(preview)) as im: display=im.copy()
            frame=ttk.Frame(w)
            frame.pack(fill='both',expand=True)
            canvas=tk.Canvas(frame,background='white')
            vs=ttk.Scrollbar(frame,orient='vertical',command=canvas.yview)
            hs=ttk.Scrollbar(frame,orient='horizontal',command=canvas.xview)
            canvas.configure(xscrollcommand=hs.set,yscrollcommand=vs.set)
            canvas.grid(row=0,column=0,sticky='nsew');vs.grid(row=0,column=1,sticky='ns');hs.grid(row=1,column=0,sticky='ew')
            frame.columnconfigure(0,weight=1);frame.rowconfigure(0,weight=1)
            img=ImageTk.PhotoImage(display)
            canvas.create_image((0,0),image=img,anchor='nw')
            canvas.image=img
            canvas.configure(scrollregion=(0,0,display.width,display.height))
            def save_preview():
                dest=filedialog.asksaveasfilename(parent=w,defaultextension='.jpg',filetypes=[('JPEG 图片','*.jpg')])
                if dest:
                    Path(dest).write_bytes(preview)
                    messagebox.showinfo('已保存',dest,parent=w)
            ttk.Button(w,text='保存与实际推送相同的 JPEG…',command=save_preview).pack(pady=8)
        else:
            area=tk.Text(w,wrap='none',font=('Consolas',10))
            area.insert('1.0',preview);area.configure(state='disabled')
            area.pack(fill='both',expand=True,padx=10,pady=10)

    def choose_map(self):
        path=filedialog.askopenfilename(title='导入店铺编码映射',filetypes=[('CSV','*.csv')])
        if path:self.mapping_path.set(path)

    def _validate_routes(self,jobs,mapping,kind,webhook,corp,secret,agent,user):
        if mapping:
            missing=sorted({(job.store_code,job.store_name) for job in jobs if not job.store_code or job.store_code not in mapping})
            if missing:
                details='、'.join(f'{code or "（无编码）"} {name}' for code,name in missing[:8])
                raise ValueError('以下服务店没有接收映射，已阻止整批发送：'+details)
            for key,value in mapping.items():
                if kind=='群机器人Webhook':check_webhook(value)
        elif kind=='群机器人Webhook':
            check_webhook(webhook)
        elif not user.strip():
            raise ValueError('应用消息需要填写接收人 UserID，或者使用按店的映射 CSV')
        return True

    def send_now(self):
        if self.inflight:
            messagebox.showwarning('正在处理','请先完成当前任务');return
        try:
            snapshot=self._snapshot()
            kind=self.kind_var.get()
            if kind=='自建应用消息' and snapshot['mode']=='Markdown表格':
                raise ValueError('自建应用消息不支持 Markdown V2 表格。请使用图片表格。')
            options=dict(kind=kind,webhook=self.webhook_var.get(),corp=self.corpid_var.get(),
                         secret=self.secret_var.get(),agent=self.agentid_var.get(),user=self.touser_var.get(),
                         mapping_path=self.mapping_path.get())
            # Important: mapping and full image encoding both happen in a worker.
            def prepare():
                prepared=self._prepare(snapshot)
                mapping=read_mapping(options['mapping_path'],kind)
                self._validate_routes([m.job for m in prepared],mapping,kind,options['webhook'],
                                      options['corp'],options['secret'],options['agent'],options['user'])
                if kind=='自建应用消息' and not all((options['corp'].strip(), options['secret'].strip(),options['agent'].strip())):
                    raise ValueError('自建应用需要填写 CorpID、Secret、AgentID')
                return prepared,mapping,options,snapshot['warning']
            self._submit(prepare,'planned','正在后台校验全部图片/Markdown 与接收目标…')
        except (ValueError,KeyError) as e:
            messagebox.showwarning('发送配置',str(e))

    def _confirm_and_send(self,plan):
        prepared,mapping,options,warning=plan
        kind=options['kind'];mode=prepared[0].mode
        recipients=len({mapping[m.job.store_code] for m in prepared}) if mapping else 1
        record_keys={r['_id'] for m in prepared for r in m.job.rows}
        shop_codes={text_of(r.get('任务服务店编码')) for m in prepared for r in m.job.rows}
        confirm=(f'发送格式：{mode}（已完整校验）\n目标模式：{kind}\n'
                 f'本次勾选活动记录：{len(record_keys)} 条\n'
                 f'预计发送消息：{len(prepared)} 条\n接收目标：{recipients} 个\n\n'
                 '请确认接收对象正确。发送后不可撤回。')
        if not mapping and len(shop_codes)>1:
            confirm+='\n\n注意：未设置接收映射，多家服务店的数据将发送到同一目标。'
        if not messagebox.askyesno('正式发送前确认',confirm,icon='warning'):
            return
        self.stop_event.clear()
        self.inflight=True
        self.bar.start(10)
        self.btn_send.config(state='disabled');self.btn_stop.config(state='normal')
        self._log(f'开始发送 {len(prepared)} 条已校验消息。')
        def process():
            ok=0;failed=0
            try:
                sender=WeComSender(kind,webhook=options['webhook'],corpid=options['corp'],
                                   secret=options['secret'],agentid=options['agent'],touser=options['user'])
                for i,item in enumerate(prepared,1):
                    if self.stop_event.is_set():
                        self.events.put(('progress','操作已停止：后续未发送的消息已取消'))
                        break
                    job=item.job
                    recipient=mapping.get(job.store_code,'') if mapping else ''
                    try:
                        sender.send(item,mode,recipient,warning)
                        ok+=1
                        self.events.put(('progress',f'[{i}/{len(prepared)}] 发送成功 · {job.store_name or "汇总"} · 第 {job.part}/{job.parts} 页'))
                    except Exception as e:
                        failed+=1
                        ss=str(e)
                        for sec in (options['webhook'],options['secret'],recipient if kind=='群机器人Webhook' else ''):
                            if sec: ss=ss.replace(sec,'[已隐藏]')
                        self.events.put(('progress',f'[{i}/{len(prepared)}] 失败 · {job.store_name or "汇总"} · {ss}'))
                    if i<len(prepared) and kind=='群机器人Webhook':
                        self.stop_event.wait(3.2)
            except Exception as e:
                ss=str(e)
                for sec in (options['secret'],options['webhook']):
                    if sec:ss=ss.replace(sec,'[已隐藏]')
                self.events.put(('progress','发送初始化失败：'+ss))
            finally:
                self.events.put(('sent',(ok,len(prepared),failed)))
        self.pool.submit(process)

    def stop_sending(self):
        self.stop_event.set()
        self.status_line.set('正在停止…当前网络请求结束后不再发送后续消息')

    def _log(self,line):
        self.log_text.configure(state='normal')
        self.log_text.insert('end',f'{datetime.now().strftime("%H:%M:%S")}  {line}\n')
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    def _close(self):
        if self.inflight and not messagebox.askyesno('正在处理','当前任务仍在执行。确认退出吗？'):
            return
        self.stop_event.set()
        try:
            CONFIG_PATH.parent.mkdir(parents=True,exist_ok=True)
            data={'fields':[f for f in self.fields if self.field_flags.get(f)],
                  'order':self.fields, 'msg_title':self.msg_title.get(),
                  'mode':self.mode_var.get(), 'warn_rate':self.warn_var.get(),
                  'group_by_shop':self.group_var.get(), 'kind':self.kind_var.get()}
            CONFIG_PATH.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')
        except OSError:
            pass
        self.pool.shutdown(wait=False,cancel_futures=True)
        self.destroy()


if __name__=='__main__':
    App().mainloop()
