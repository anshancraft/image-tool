import os
import json
import shutil
from datetime import datetime
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from PIL import Image as PILImage, ImageDraw, ImageFont
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from io import BytesIO

SETTING_FILE = "config.json"

NAME_TEMPLATES = [
    ("1. 基础通用", "{prefix}{index}-{date}"),
    ("2. 款号标准", "{code}{index}-{date}"),
    ("3. 款号+颜色尺码", "{code}-{color}-{size}{index}-{date}"),
    ("4. 品类+季节", "{cate}-{season}{index}-{date}"),
    ("5. 用途分类", "{use}-{code}{index}-{date}")
]

WATERMARK_POS = ["左上角", "右上角", "左下角", "右下角", "居中"]
IMG_SUFFIX = [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"]
CONVERT_FORMAT_LIST = ["JPG", "PNG", "WebP", "BMP"]

class FileOrganizer:
    def __init__(self, root):
        self.root = root
        self.root.title("服装电商图片批量整理工具")
        self.root.geometry("780x620")
        self.root.resizable(True, True)

        # 滚动布局核心
        self.canvas = Canvas(root)
        self.scroll_frame = Frame(self.canvas)
        self.scroll_bar = Scrollbar(root, orient=VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scroll_bar.set)

        self.scroll_bar.pack(side=RIGHT, fill=Y)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)
        self.canvas.create_window((0,0), window=self.scroll_frame, anchor=NW)
        self.scroll_frame.bind("<Configure>", self.on_frame_config)
        self.canvas.bind_all("<MouseWheel>", self.on_mouse_wheel)

        # 基础变量
        self.folder_path = StringVar()
        self.prefix = StringVar(value="素材")
        self.code = StringVar(value="D款")
        self.color = StringVar(value="黑色")
        self.size = StringVar(value="M")
        self.cate = StringVar(value="连衣裙")
        self.season = StringVar(value="2026夏")
        self.use_type = StringVar(value="主图")

        self.auto_sort = IntVar(value=1)
        self.auto_rename = IntVar(value=1)
        self.add_date_switch = IntVar(value=1)
        self.include_sub_folder = IntVar(value=0)

        self.water_text = StringVar(value="原创女装 禁止盗图")
        self.water_pos = StringVar(value="右下角")
        self.water_size = IntVar(value=30)
        self.water_alpha = IntVar(value=50)
        self.logo_path = StringVar(value="")
        self.use_logo_water = IntVar(value=0)
        self.water_font = StringVar(value="黑体")
        self.water_angle = IntVar(value=0)

        self.compress_quality = IntVar(value=60)
        self.with_thumbnail = IntVar(value=0)
        self.convert_target_fmt = StringVar(value="JPG")

        # 裁剪尺寸
        self.crop_w = StringVar(value="750")
        self.crop_h = StringVar(value="1000")

        self.template_sel = StringVar()
        self.template_list = [t[0] for t in NAME_TEMPLATES]

        self.undo_list = []
        self.excel_data = []
        self.preview_list = []

        self.load_config()
        self.build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_frame_config(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox(ALL))

    def on_mouse_wheel(self, event):
        self.canvas.yview_scroll(-1 * (event.delta // 120), UNITS)

    def load_config(self):
        if not os.path.exists(SETTING_FILE):
            return
        try:
            with open(SETTING_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self.prefix.set(cfg.get("prefix", "素材"))
            self.code.set(cfg.get("code", "D款"))
            self.color.set(cfg.get("color", "黑色"))
            self.size.set(cfg.get("size", "M"))
            self.cate.set(cfg.get("cate", "连衣裙"))
            self.season.set(cfg.get("season", "2026夏"))
            self.use_type.set(cfg.get("use_type", "主图"))
            self.auto_sort.set(cfg.get("auto_sort", 1))
            self.auto_rename.set(cfg.get("auto_rename", 1))
            self.add_date_switch.set(cfg.get("add_date_switch", 1))
            self.include_sub_folder.set(cfg.get("include_sub_folder", 0))
            self.water_text.set(cfg.get("water_text", "原创女装 禁止盗图"))
            self.water_pos.set(cfg.get("water_pos", "右下角"))
            self.water_size.set(cfg.get("water_size", 30))
            self.water_alpha.set(cfg.get("water_alpha", 50))
            self.water_font.set(cfg.get("water_font", "黑体"))
            self.water_angle.set(cfg.get("water_angle", 0))
            self.compress_quality.set(cfg.get("compress_quality", 60))
            self.with_thumbnail.set(cfg.get("with_thumbnail", 0))
            self.convert_target_fmt.set(cfg.get("convert_target_fmt", "JPG"))
            self.crop_w.set(cfg.get("crop_w", "750"))
            self.crop_h.set(cfg.get("crop_h", "1000"))
        except:
            pass

    def save_config(self):
        cfg = {
            "prefix": self.prefix.get(),
            "code": self.code.get(),
            "color": self.color.get(),
            "size": self.size.get(),
            "cate": self.cate.get(),
            "season": self.season.get(),
            "use_type": self.use_type.get(),
            "auto_sort": self.auto_sort.get(),
            "auto_rename": self.auto_rename.get(),
            "add_date_switch": self.add_date_switch.get(),
            "include_sub_folder": self.include_sub_folder.get(),
            "water_text": self.water_text.get(),
            "water_pos": self.water_pos.get(),
            "water_size": self.water_size.get(),
            "water_alpha": self.water_alpha.get(),
            "water_font": self.water_font.get(),
            "water_angle": self.water_angle.get(),
            "compress_quality": self.compress_quality.get(),
            "with_thumbnail": self.with_thumbnail.get(),
            "convert_target_fmt": self.convert_target_fmt.get(),
            "crop_w": self.crop_w.get(),
            "crop_h": self.crop_h.get()
        }
        with open(SETTING_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

    def on_close(self):
        self.save_config()
        self.root.destroy()

    def log(self, msg):
        t = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(END, f"[{t}] {msg}\n")
        self.log_text.see(END)

    def build_ui(self):
        container = self.scroll_frame
        Label(container, text="📂 服装电商图片批量整理工具", font=("微软雅黑", 18, "bold")).pack(pady=10)

        f1 = Frame(container)
        f1.pack(pady=5, fill=X, padx=20)
        Label(f1, text="目标文件夹：", font=("微软雅黑", 11)).pack(side=LEFT)
        Entry(f1, textvariable=self.folder_path, width=48).pack(side=LEFT, padx=5)
        Button(f1, text="选择文件夹", command=self.select_folder, bg="#4299e1", fg="white").pack(side=LEFT)

        f_sub = Frame(container)
        f_sub.pack(pady=2, fill=X, padx=20)
        Checkbutton(f_sub, text="递归包含所有子文件夹", variable=self.include_sub_folder, font=("微软雅黑", 10, "bold"), fg="#e53e3e").pack(side=LEFT)

        f_temp = Frame(container)
        f_temp.pack(pady=6, fill=X, padx=20)
        Label(f_temp, text="📝 命名模板：", font=("微软雅黑", 11)).pack(side=LEFT)
        self.temp_cb = ttk.Combobox(f_temp, textvariable=self.template_sel, values=self.template_list, width=32, state="readonly")
        self.temp_cb.current(0)
        self.temp_cb.pack(side=LEFT, padx=10)
        Checkbutton(f_temp, text="名称带日期", variable=self.add_date_switch).pack(side=LEFT, padx=15)

        f2 = Frame(container)
        f2.pack(pady=6, fill=X, padx=20)
        Label(f2, text="前缀：").grid(row=0, column=0, sticky=W)
        Entry(f2, textvariable=self.prefix, width=10).grid(row=0, column=1, padx=3)
        Label(f2, text="款号：").grid(row=0, column=2, sticky=W)
        Entry(f2, textvariable=self.code, width=10).grid(row=0, column=3, padx=3)
        Label(f2, text="颜色：").grid(row=0, column=4, sticky=W)
        Entry(f2, textvariable=self.color, width=8).grid(row=0, column=5, padx=3)

        f3 = Frame(container)
        f3.pack(pady=6, fill=X, padx=20)
        Label(f3, text="尺码：").grid(row=0, column=0, sticky=W)
        Entry(f3, textvariable=self.size, width=8).grid(row=0, column=1, padx=3)
        Label(f3, text="品类：").grid(row=0, column=2, sticky=W)
        Entry(f3, textvariable=self.cate, width=10).grid(row=0, column=3, padx=3)
        Label(f3, text="季节：").grid(row=0, column=4, sticky=W)
        Entry(f3, textvariable=self.season, width=10).grid(row=0, column=5, padx=3)
        Label(f3, text="用途：").grid(row=0, column=6, sticky=W)
        Entry(f3, textvariable=self.use_type, width=10).grid(row=0, column=7, padx=3)

        # 批量水印
        f_water = Frame(container, bd=1, relief=SOLID)
        f_water.pack(pady=8, fill=X, padx=20)
        Label(f_water, text="💧 批量水印", font=("微软雅黑", 12, "bold")).grid(row=0, column=0, columnspan=8, pady=5)
        Label(f_water, text="水印文字：").grid(row=1, column=0, padx=5, pady=3)
        Entry(f_water, textvariable=self.water_text, width=18).grid(row=1, column=1)
        Label(f_water, text="位置：").grid(row=1, column=2)
        cb_pos = ttk.Combobox(f_water, textvariable=self.water_pos, values=WATERMARK_POS, width=8, state="readonly")
        cb_pos.grid(row=1, column=3)
        Label(f_water, text="字号：").grid(row=1, column=4)
        Entry(f_water, textvariable=self.water_size, width=5).grid(row=1, column=5)
        Label(f_water, text="透明度：").grid(row=2, column=0, pady=3)
        Entry(f_water, textvariable=self.water_alpha, width=5).grid(row=2, column=1)
        Label(f_water, text="字体：").grid(row=2, column=2)
        self.font_list = ["黑体", "微软雅黑", "宋体", "楷体"]
        cb_font = ttk.Combobox(f_water, textvariable=self.water_font, values=self.font_list, width=8, state="readonly")
        cb_font.grid(row=2, column=3)
        Label(f_water, text="倾斜角度：").grid(row=2, column=4)
        Entry(f_water, textvariable=self.water_angle, width=5).grid(row=2, column=5)
        Checkbutton(f_water, text="使用Logo水印", variable=self.use_logo_water).grid(row=1, column=6)
        Button(f_water, text="选择Logo", command=self.select_logo).grid(row=1, column=7, padx=2)

        # 批量压缩
        f_compress = Frame(container, bd=1, relief=SOLID)
        f_compress.pack(pady=8, fill=X, padx=20)
        Label(f_compress, text="🗜️ 批量压缩", font=("微软雅黑", 12, "bold")).grid(row=0, column=0, columnspan=4, pady=5)
        Label(f_compress, text="压缩质量(0-100)：").grid(row=1, column=0, padx=5, pady=3)
        Entry(f_compress, textvariable=self.compress_quality, width=5).grid(row=1, column=1)
        Button(f_compress, text="开始压缩", command=self.compress_images, bg="#20c997", fg="white").grid(row=1, column=2, padx=10)

        # 格式转换
        f_convert = Frame(container, bd=1, relief=SOLID)
        f_convert.pack(pady=8, fill=X, padx=20)
        Label(f_convert, text="🔄 格式转换", font=("微软雅黑", 12, "bold")).grid(row=0, column=0, columnspan=4, pady=5)
        Label(f_convert, text="转换为：").grid(row=1, column=0, padx=5, pady=3)
        cb_conv = ttk.Combobox(f_convert, textvariable=self.convert_target_fmt, values=CONVERT_FORMAT_LIST, width=8, state="readonly")
        cb_conv.grid(row=1, column=1)
        Button(f_convert, text="开始转换", command=self.convert_format, bg="#9370DB", fg="white").grid(row=1, column=2, padx=10)

        # 批量裁剪
        f_crop = Frame(container, bd=1, relief=SOLID)
        f_crop.pack(pady=8, fill=X, padx=20)
        Label(f_crop, text="✂️ 批量居中裁剪(不变形)", font=("微软雅黑", 12, "bold")).grid(row=0, column=0, columnspan=5, pady=5)
        Label(f_crop, text="宽度：").grid(row=1, column=0, padx=5)
        Entry(f_crop, textvariable=self.crop_w, width=8).grid(row=1, column=1)
        Label(f_crop, text="高度：").grid(row=1, column=2, padx=5)
        Entry(f_crop, textvariable=self.crop_h, width=8).grid(row=1, column=3)
        Button(f_crop, text="开始裁剪", command=self.crop_image, bg="#ff7f50", fg="white").grid(row=1, column=4, padx=10)

        f4 = Frame(container)
        f4.pack(pady=6, fill=X, padx=20)
        Checkbutton(f4, text="自动分类归档", variable=self.auto_sort).pack(side=LEFT, padx=8)
        Checkbutton(f4, text="自动批量改名", variable=self.auto_rename).pack(side=LEFT, padx=8)

        f_btn = Frame(container)
        f_btn.pack(pady=12)
        Button(f_btn, text="👁️ 预览效果", command=self.preview_rename, bg="#673ab7", fg="white", width=10, height=2).pack(side=LEFT, padx=3)
        Button(f_btn, text="💧 批量水印", command=self.batch_watermark, bg="#e53e3e", fg="white", width=10, height=2).pack(side=LEFT, padx=3)
        Button(f_btn, text="✅ 开始整理", command=self.start_organize, bg="#38a169", fg="white", width=10, height=2).pack(side=LEFT, padx=3)
        Button(f_btn, text="↩️ 一键撤销", command=self.undo_organize, bg="#dd6b20", width=10, height=2).pack(side=LEFT, padx=3)
        Button(f_btn, text="📊 导出台账", command=self.generate_excel, bg="#9c27b0", fg="white", width=10, height=2).pack(side=LEFT, padx=3)

        f_thumb = Frame(container)
        f_thumb.pack(pady=2)
        Checkbutton(f_thumb, text="Excel带图片缩略图", variable=self.with_thumbnail, font=("微软雅黑", 10)).pack()

        Label(container, text="处理日志 / 预览结果：", font=("微软雅黑", 10)).pack(padx=20, anchor=W)
        self.log_text = Text(container, height=12, width=98, font=("微软雅黑", 9))
        self.log_text.pack(padx=20, pady=5, fill=BOTH, expand=True)

    def get_all_image_files(self, root_dir):
        img_list = []
        if self.include_sub_folder.get():
            for root, _, files in os.walk(root_dir):
                for f in files:
                    if f.lower().endswith(tuple(IMG_SUFFIX)) and not f.startswith("."):
                        img_list.append((root, f))
        else:
            for f in os.listdir(root_dir):
                fp = os.path.join(root_dir, f)
                if os.path.isfile(fp) and f.lower().endswith(tuple(IMG_SUFFIX)) and not f.startswith("."):
                    img_list.append((root_dir, f))
        return img_list

    def select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path.set(path)
            self.log(f"已选择文件夹：{path}")

    def select_logo(self):
        path = filedialog.askopenfilename(filetypes=[("图片", "*.png;*.jpg;*.jpeg")])
        if path:
            self.logo_path.set(path)
            self.log(f"已选择Logo：{path}")

    def get_template_str(self):
        idx = self.temp_cb.current()
        return NAME_TEMPLATES[idx][1]

    def format_name(self, index_num, ext):
        tpl = self.get_template_str()
        if not self.add_date_switch.get():
            tpl = tpl.replace("-{date}", "")
            date_str = ""
        else:
            date_str = datetime.now().strftime("%Y%m%d")
        name = tpl.replace("{prefix}", self.prefix.get())\
                  .replace("{code}", self.code.get())\
                  .replace("{color}", self.color.get())\
                  .replace("{size}", self.size.get())\
                  .replace("{cate}", self.cate.get())\
                  .replace("{season}", self.season.get())\
                  .replace("{use}", self.use_type.get())\
                  .replace("{index}", f"{index_num:02d}")\
                  .replace("{date}", date_str)
        return name + ext

    def preview_rename(self):
        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择文件夹！")
            return
        self.log_text.delete(1.0, END)
        self.preview_list.clear()
        all_imgs = self.get_all_image_files(folder)
        if not all_imgs:
            self.log("文件夹内无有效图片文件")
            return
        self.log("========== 预览改名效果 ==========")
        index = 1
        for _, f in all_imgs:
            _, ext = os.path.splitext(f)
            new_name = self.format_name(index, ext)
            self.preview_list.append((f, new_name))
            self.log(f"原：{f}  →  新：{new_name}")
            index += 1
        self.log("========== 预览结束 ==========")

    def backup_original(self, folder):
        backup_dir = os.path.join(folder, "_原图备份")
        os.makedirs(backup_dir, exist_ok=True)
        for f in os.listdir(folder):
            fp = os.path.join(folder, f)
            if os.path.isfile(fp) and not f.startswith("."):
                shutil.copy2(fp, os.path.join(backup_dir, f))
        self.log("已自动备份原图到：_原图备份")

    def get_watermark_images(self, root_dir):
        return self.get_all_image_files(root_dir)

    # 批量格式转换
    def convert_format(self):
        root_folder = self.folder_path.get()
        if not root_folder or not os.path.isdir(root_folder):
            messagebox.showwarning("提示", "请先选择文件夹！")
            return
        target_fmt = self.convert_target_fmt.get().lower()
        out_dir = os.path.join(root_folder, "格式转换后")
        os.makedirs(out_dir, exist_ok=True)
        all_imgs = self.get_all_image_files(root_folder)
        if not all_imgs:
            self.log("❌ 未找到任何图片")
            return
        self.log(f"===== 开始转换为 {self.convert_target_fmt.get()} =====")
        for dir_path, fname in all_imgs:
            img_path = os.path.join(dir_path, fname)
            try:
                name_no_ext = os.path.splitext(fname)[0]
                save_name = f"{name_no_ext}.{target_fmt}"
                save_path = os.path.join(out_dir, save_name)
                im = PILImage.open(img_path)
                if target_fmt == "jpg":
                    im = im.convert("RGB")
                im.save(save_path)
                self.log(f"✅ 成功：{fname} → {save_name}")
            except Exception as e:
                self.log(f"❌ 失败：{fname} -> {str(e)}")
        self.log("===== 格式转换完成 =====")
        messagebox.showinfo("完成", f"转换完成！共处理{len(all_imgs)}张")

    # 批量居中裁剪
    def crop_image(self):
        root_folder = self.folder_path.get()
        if not root_folder or not os.path.isdir(root_folder):
            messagebox.showwarning("提示", "请先选择文件夹！")
            return
        try:
            tar_w = int(self.crop_w.get())
            tar_h = int(self.crop_h.get())
        except:
            messagebox.showerror("错误", "宽高必须是数字！")
            return
        if tar_w < 100 or tar_h < 100:
            messagebox.showwarning("提示", "宽高不能小于100")
            return
        out_dir = os.path.join(root_folder, "裁剪后")
        os.makedirs(out_dir, exist_ok=True)
        all_imgs = self.get_all_image_files(root_folder)
        if not all_imgs:
            self.log("❌ 未找到任何图片")
            return
        self.log(f"===== 开始批量裁剪 {tar_w}×{tar_h} =====")
        for dir_path, fname in all_imgs:
            img_path = os.path.join(dir_path, fname)
            try:
                im = PILImage.open(img_path).convert("RGB")
                w, h = im.size
                ratio_src = w / h
                ratio_tar = tar_w / tar_h
                if ratio_src > ratio_tar:
                    new_w = int(h * ratio_tar)
                    new_h = h
                    left = (w - new_w) // 2
                    top = 0
                else:
                    new_w = w
                    new_h = int(w / ratio_tar)
                    left = 0
                    top = (h - new_h) // 2
                im_crop = im.crop((left, top, left + new_w, top + new_h))
                im_resize = im_crop.resize((tar_w, tar_h), PILImage.Resampling.LANCZOS)
                save_path = os.path.join(out_dir, fname)
                im_resize.save(save_path)
                self.log(f"✅ 裁剪成功：{fname}")
            except Exception as e:
                self.log(f"❌ 裁剪失败：{fname} -> {str(e)}")
        self.log("===== 批量裁剪完成 =====")
        messagebox.showinfo("完成", f"裁剪完成！共处理{len(all_imgs)}张")

    def batch_watermark(self):
        if not messagebox.askyesno("确认", "开始批量加水印？"):
            return
        root_folder = self.folder_path.get()
        if not root_folder or not os.path.isdir(root_folder):
            messagebox.showwarning("提示", "请先选择文件夹！")
            return
        out_dir = os.path.join(root_folder, "加水印后")
        os.makedirs(out_dir, exist_ok=True)
        all_imgs = self.get_watermark_images(root_folder)
        if not all_imgs:
            self.log("❌ 未找到任何图片")
            return
        font_map = {"黑体":"simhei.ttf","微软雅黑":"msyh.ttf","宋体":"simsun.ttc","楷体":"simkai.ttf"}
        sel_font = self.water_font.get()
        font_path = font_map.get(sel_font, "simhei.ttf")
        angle = self.water_angle.get()
        self.log(f"===== 开始批量加水印 =====")
        for dir_path, fname in all_imgs:
            img_path = os.path.join(dir_path, fname)
            try:
                im = PILImage.open(img_path).convert("RGBA")
                w, h = im.size
                draw = ImageDraw.Draw(im)
                alpha = self.water_alpha.get()
                if not self.use_logo_water.get():
                    text = self.water_text.get()
                    fontsize = self.water_size.get()
                    try:
                        font = ImageFont.truetype(font_path, fontsize)
                    except:
                        font = ImageFont.load_default()
                    bbox = draw.textbbox((0,0), text, font=font)
                    tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
                    txt_img = PILImage.new("RGBA", (tw+40, th+40), (0,0,0,0))
                    txt_draw = ImageDraw.Draw(txt_img)
                    txt_draw.text((20,20), text, fill=(0,0,0,int(alpha*2.55)), font=font)
                    if angle != 0:
                        txt_img = txt_img.rotate(angle, expand=True)
                    gap = 30
                    pos = self.water_pos.get()
                    tx, ty = {"左上角":(gap,gap),"右上角":(w-txt_img.width-gap,gap),"左下角":(gap,h-txt_img.height-gap),"右下角":(w-txt_img.width-gap,h-txt_img.height-gap),"居中":((w-txt_img.width)//2,(h-txt_img.height)//2)}[pos]
                    im.paste(txt_img, (tx, ty), mask=txt_img)
                else:
                    if not os.path.exists(self.logo_path.get()):
                        self.log("⚠️ 未选Logo，跳过")
                        continue
                    logo = PILImage.open(self.logo_path.get()).convert("RGBA")
                    new_w = int(w * 0.25)
                    new_h = int(logo.height * new_w / logo.width)
                    logo = logo.resize((new_w, new_h), PILImage.Resampling.LANCZOS)
                    gap = 30
                    pos = self.water_pos.get()
                    tx, ty = {"左上角":(gap,gap),"右上角":(w-logo.width-gap,gap),"左下角":(gap,h-logo.height-gap),"右下角":(w-logo.width-gap,h-logo.height-gap),"居中":((w-logo.width)//2,(h-logo.height)//2)}[pos]
                    im.paste(logo, (tx, ty), mask=logo)
                if fname.lower().endswith(("jpg", "jpeg")):
                    im = im.convert("RGB")
                save_path = os.path.join(out_dir, fname)
                im.save(save_path)
                self.log(f"✅ 水印成功：{fname}")
            except Exception as e:
                self.log(f"❌ 水印失败：{fname} -> {str(e)}")
        self.log("===== 水印全部完成 =====")
        messagebox.showinfo("完成", f"水印完成！共处理{len(all_imgs)}张")

    def start_organize(self):
        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择文件夹！")
            return
        self.backup_original(folder)
        self.undo_list.clear()
        self.excel_data.clear()
        self.log("===== 开始整理 =====")
        all_imgs = self.get_all_image_files(folder)
        if not all_imgs:
            self.log("没有可处理图片")
            return
        index = 1
        for dir_path, file in all_imgs:
            old_path = os.path.join(dir_path, file)
            name, ext = os.path.splitext(file)
            size_mb = os.path.getsize(old_path) / 1024 / 1024
            if self.auto_sort.get():
                cat_path = os.path.join(folder, "整理后图片")
                os.makedirs(cat_path, exist_ok=True)
            else:
                cat_path = dir_path
            if self.auto_rename.get():
                new_name = self.format_name(index, ext)
                cnt = 1
                base_new = os.path.splitext(new_name)[0]
                while os.path.exists(os.path.join(cat_path, new_name)):
                    new_name = f"{base_new}-{cnt}{ext}"
                    cnt += 1
            else:
                new_name = file
            new_path = os.path.join(cat_path, new_name)
            shutil.move(old_path, new_path)
            self.undo_list.append((new_path, old_path))
            self.log(f"处理：{file} → {new_name}")
            self.excel_data.append({
                "原文件名": file,"新文件名": new_name,"文件分类":"图片",
                "存放路径": cat_path,"大小(MB)": round(size_mb,2),
                "修改时间": datetime.fromtimestamp(os.path.getmtime(new_path)).strftime("%Y-%m-%d %H:%M:%S"),
                "格式": ext,"完整路径": new_path
            })
            index += 1
        self.log(f"===== 完成共处理 {index-1} 张 =====")
        messagebox.showinfo("完成", "整理完成，已自动备份原图")

    def undo_organize(self):
        if not self.undo_list:
            messagebox.showinfo("提示", "无操作可撤销")
            return
        count = 0
        self.log("===== 开始撤销 =====")
        for new_path, old_path in reversed(self.undo_list):
            if os.path.exists(new_path):
                shutil.move(new_path, old_path)
                count += 1
        self.undo_list.clear()
        self.log(f"===== 撤销完成，恢复{count}个文件 =====")
        messagebox.showinfo("撤销完成", f"已恢复 {count} 个文件")

    def generate_excel(self):
        if not self.excel_data:
            messagebox.showinfo("提示", "请先执行整理再导出！")
            return
        folder = self.folder_path.get()
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")],initialdir=folder,initialfile="图片台账.xlsx")
        if not save_path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "图片台账"
        headers = ["原文件名","新文件名","分类","路径","大小(MB)","修改时间","格式"]
        has_thumb = self.with_thumbnail.get()
        if has_thumb:
            headers.append("缩略图")
        ws.append(headers)

        # 固定列宽、行高
        ws.column_dimensions["A"].width = 25
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["D"].width = 30
        if has_thumb:
            ws.column_dimensions["H"].width = 18

        row_num = 2
        for row in self.excel_data:
            ws.append([
                row["原文件名"],row["新文件名"],row["文件分类"],
                row["存放路径"],row["大小(MB)"],row["修改时间"],row["格式"]
            ])
            if has_thumb:
                img_path = row["完整路径"]
                if os.path.exists(img_path) and any(img_path.lower().endswith(s) for s in IMG_SUFFIX):
                    try:
                        # 生成固定大小缩略图
                        thumb = PILImage.open(img_path)
                        thumb.thumbnail((120, 120))
                        buffer = BytesIO()
                        thumb.save(buffer, format="PNG")
                        buffer.seek(0)
                        excel_img = ExcelImage(buffer)
                        # 适配旧版 openpyxl，直接指定单元格
                        cell_addr = f"H{row_num}"
                        ws.add_image(excel_img, cell_addr)
                        excel_img.width = 110
                        excel_img.height = 110
                        # 行高适配
                        ws.row_dimensions[row_num].height = 100
                    except Exception as e:
                        self.log(f"缩略图失败：{e}")
            row_num += 1

        wb.save(save_path)
        self.log("Excel台账已保存，缩略图已锚定单元格（需Excel手动勾选“随单元格改变位置和大小”）")
        messagebox.showinfo("完成", "台账导出成功！缩略图需在Excel中手动勾选一次随单元格缩放")

    def compress_images(self):
        if not messagebox.askyesno("确认", "开始压缩图片？"):
            return
        root_folder = self.folder_path.get()
        if not root_folder or not os.path.isdir(root_folder):
            messagebox.showwarning("提示", "请先选择文件夹！")
            return
        out_dir = os.path.join(root_folder, "压缩后")
        os.makedirs(out_dir, exist_ok=True)
        all_imgs = self.get_watermark_images(root_folder)
        if not all_imgs:
            self.log("❌ 未找到图片")
            return
        quality = self.compress_quality.get()
        for dir_path, fname in all_imgs:
            try:
                im = PILImage.open(os.path.join(dir_path, fname)).convert("RGB")
                im.save(os.path.join(out_dir, fname), quality=quality, optimize=True)
                self.log(f"✅ 压缩成功：{fname}")
            except:
                self.log(f"❌ 压缩失败：{fname}")
        messagebox.showinfo("完成", "压缩完成")

if __name__ == "__main__":
    root = Tk()
    app = FileOrganizer(root)
    root.mainloop()
