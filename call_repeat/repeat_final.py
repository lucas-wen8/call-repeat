# -*- coding: utf-8 -*-
import os, traceback
import pandas as pd
from datetime import timedelta
import tkinter as tk
from tkinter import filedialog, messagebox

EXCLUDE_GROUP = "线上运营组"
TIME_WINDOW_HOURS = 24  # 邻接配对时间窗口（小时）


def process_excel(input_path: str, output_dir: str):
    """
    读取 Excel（第2行为表头，dtype=str 保护大数字），清洗列名；
    过滤：呼入/接听技能组/坐席分机/坐席姓名 任一为空或 NaN，或 呼入技能组=“线上运营组” 的行；
    同一主叫号码内按开始时间做【邻接配对】（相邻两通且 <=24h）；
    导出仅包含源表整行原样数据（不新增任何列），分别写入“没跨组/跨组”文件。
    返回：(没跨组对数, 跨组对数)
    """
    # 1) 读第2行为表头；全列按字符串读取
    try:
        df_str = pd.read_excel(input_path, engine="openpyxl", header=1, dtype=str)
    except Exception:
        df_str = pd.read_excel(input_path, header=1, dtype=str)

    # 清洗表头空白/换行
    def _clean_col(c):
        s = str(c).replace("\u3000", " ").replace("\xa0", " ")
        s = s.replace("\r", "").replace("\n", "").strip()
        return s
    df_str.columns = [_clean_col(c) for c in df_str.columns]

    # 2) 列校验（必须存在）
    REQUIRED_COLS = ["开始时间", "主叫号码", "接听技能组", "呼入技能组", "坐席分机", "坐席姓名"]
    missing = [c for c in REQUIRED_COLS if c not in df_str.columns]
    if missing:
        raise ValueError(f"Excel 缺少必需列：{missing}\n必须包含：{REQUIRED_COLS}")

    # 3) 规范化 & 过滤（在字符串表上进行，保证导出原样）
    # 主叫号码去掉可能的前缀（TEL: 等）
    df_str["主叫号码"] = df_str["主叫号码"].astype(str).str.replace(r"^[A-Za-z：: ]+", "", regex=True).str.strip()

    # 去空白字符 -> 统一为空串判定
    for col in ["呼入技能组", "接听技能组", "坐席分机", "坐席姓名"]:
        df_str[col] = df_str[col].astype(str).str.replace(r"\s+", "", regex=True)
        df_str[col] = df_str[col].where(df_str[col].str.lower() != "nan", "")  # 防“nan”字符串

    # 过滤：任一为空/NaN 或 呼入技能组 = 线上运营组
    df_str = df_str[
        (df_str["呼入技能组"] != "") &
        (df_str["接听技能组"] != "") &
        (df_str["坐席分机"] != "") &
        (df_str["坐席姓名"] != "") &
        (df_str["呼入技能组"] != EXCLUDE_GROUP)
    ].copy()

    # 4) 创建处理副本做时间解析与排序（不影响 df_str 的原样文本）
    df_proc = df_str.copy()
    df_proc["_rowid"] = range(len(df_proc))  # 绑定原样行位置，后续反查 df_str
    df_proc["开始时间"] = pd.to_datetime(df_proc["开始时间"], errors="coerce")
    if df_proc["开始时间"].isna().any():
        raise ValueError("存在无法解析的‘开始时间’，请检查时间格式。")
    df_proc.sort_values(["主叫号码", "开始时间"], inplace=True, kind="stable")
    df_proc.reset_index(drop=True, inplace=True)

    # 5) 邻接配对：同号码相邻两通且 <= 24 小时
    pairs = []  # 元组：(标签, rowid_a, rowid_b)
    for number, g in df_proc.groupby("主叫号码", sort=False):
        g = g.sort_values("开始时间", kind="stable").reset_index(drop=True)
        for i in range(len(g) - 1):
            a_time = g.loc[i, "开始时间"]
            b_time = g.loc[i + 1, "开始时间"]
            if (b_time - a_time) <= pd.Timedelta(hours=TIME_WINDOW_HOURS):
                rid_a = int(g.loc[i, "_rowid"])
                rid_b = int(g.loc[i + 1, "_rowid"])
                cross = (str(g.loc[i, "接听技能组"]) != str(g.loc[i + 1, "呼入技能组"]))
                pairs.append(("跨组" if cross else "没跨组", rid_a, rid_b))

    # 6) 仅以“原字符串表 df_str”导出整行（不新增任何列）
    out_nc = os.path.join(output_dir, "重复来电_没跨组.xlsx")
    out_c  = os.path.join(output_dir, "重复来电_跨组.xlsx")
    orig_cols = list(df_str.columns)

    if not pairs:
        df_str.head(0).to_excel(out_nc, index=False)
        df_str.head(0).to_excel(out_c, index=False)
        return 0, 0

    def pairs_to_original_rows(pairs_subset):
        rows = []
        for _, rid_a, rid_b in pairs_subset:
            rows.append(df_str.iloc[rid_a].to_dict())  # A 行（原样字符串）
            rows.append(df_str.iloc[rid_b].to_dict())  # B 行（原样字符串）
        return pd.DataFrame(rows, columns=orig_cols)

    pairs_nc = [p for p in pairs if p[0] == "没跨组"]
    pairs_c  = [p for p in pairs if p[0] == "跨组"]

    df_nc_only = pairs_to_original_rows(pairs_nc)
    df_c_only  = pairs_to_original_rows(pairs_c)

    df_nc_only.to_excel(out_nc, index=False)
    df_c_only.to_excel(out_c, index=False)

    return len(pairs_nc), len(pairs_c)


# ================== GUI ==================
class App:
    def __init__(self, root):
        self.root = root
        root.title("重复来电分类工具（24h 邻接配对｜跨组/没跨组）")
        root.geometry("1280x360")
        root.configure(bg="#FFFFFF")
        try:
            # macOS：启动时前置窗口，避免文件对话框被遮挡
            root.lift()
            root.attributes("-topmost", True)
            root.after(600, lambda: root.attributes("-topmost", False))
            root.tk.call('tk', 'scaling', 1.25)
        except Exception:
            pass

        self.in_var  = tk.StringVar(value="（未选择）")
        self.out_var = tk.StringVar(value=os.path.expanduser("~"))

        # 输入框（横向居中 + 加宽）
        tk.Label(root, text="Excel 文件：", bg="#FFFFFF", fg="#000000", font=("Arial", 14)).place(x=24, y=36)
        tk.Entry(root, textvariable=self.in_var, bg="#F7F7F7", fg="#000000",
                 relief="solid", bd=1, font=("Arial", 12), justify="left"
                 ).place(relx=0.5, y=36, anchor="center", width=1000, height=30)

        tk.Label(root, text="输出目录：", bg="#FFFFFF", fg="#000000", font=("Arial", 14)).place(x=24, y=96)
        tk.Entry(root, textvariable=self.out_var, bg="#F7F7F7", fg="#000000",
                 relief="solid", bd=1, font=("Arial", 12), justify="left"
                 ).place(relx=0.5, y=96, anchor="center", width=1000, height=30)

        # —— 横向居中 + 放大的白色按钮（显示完整路径；按下/松开反馈）——
        self.btn_file = tk.Button(
            root, text="选择 Excel…", command=self.choose_file,
            bg="#FFFFFF", fg="#333333",
            activebackground="#EDEDED", activeforeground="#111111",
            relief="raised", bd=1, font=("Arial", 12),
            anchor="w", justify="left"
        )
        self.btn_file.place(relx=0.5, y=70, anchor="center", width=1000, height=36)

        self.btn_dir = tk.Button(
            root, text="选择目录…", command=self.choose_dir,
            bg="#FFFFFF", fg="#333333",
            activebackground="#EDEDED", activeforeground="#111111",
            relief="raised", bd=1, font=("Arial", 12),
            anchor="w", justify="left"
        )
        self.btn_dir.place(relx=0.5, y=130, anchor="center", width=1000, height=36)

        # 操作区
        self.btn_run = tk.Button(root, text="开始处理", command=self.run,
                                 bg="#34A853", fg="#FFFFFF", activebackground="#278A41",
                                 activeforeground="#FFFFFF", relief="raised", bd=2, font=("Arial", 13))
        self.btn_run.place(x=150, y=200, width=120, height=36)
        tk.Button(root, text="关闭", command=root.quit,
                  bg="#777777", fg="#FFFFFF", activebackground="#5F5F5F",
                  activeforeground="#FFFFFF", relief="raised", bd=2, font=("Arial", 13)
                  ).place(x=290, y=200, width=120, height=36)

        self.status = tk.StringVar(value="就绪")
        tk.Label(root, textvariable=self.status, bg="#FFFFFF", fg="#333333", font=("Arial", 11)).place(x=24, y=270)

    # —— 按钮反馈（白底风格） ——
    def _btn_press(self, btn, pressed_bg="#D3D3D3", pressed_fg="#111111"):
        try:
            btn.config(relief="sunken", bg=pressed_bg, fg=pressed_fg)
            btn.update_idletasks()
        except Exception:
            pass

    def _btn_release(self, btn, text=None, released_bg="#FFFFFF", released_fg="#333333"):
        try:
            if text is not None:
                btn.config(text=text)
            btn.config(relief="raised", bg=released_bg, fg=released_fg)
            btn.update_idletasks()
        except Exception:
            pass

    # —— 文件/目录选择 —— 
    def choose_file(self):
        self._btn_press(self.btn_file)
        try:
            p = filedialog.askopenfilename(
                title="选择 Excel 文件",
                filetypes=[("Excel", "*.xlsx *.xls *.xlsm *.xlsb"), ("所有文件", "*.*")]
            )
            if p:
                self.in_var.set(p)
                self._btn_release(self.btn_file, text=p, released_bg="#FFFFFF", released_fg="#333333")
            else:
                self._btn_release(self.btn_file, text="选择 Excel…", released_bg="#FFFFFF", released_fg="#333333")
        except Exception:
            self._btn_release(self.btn_file, text="选择 Excel…", released_bg="#FFFFFF", released_fg="#333333")
            raise

    def choose_dir(self):
        self._btn_press(self.btn_dir)
        try:
            d = filedialog.askdirectory(
                title="选择输出文件夹",
                initialdir=self.out_var.get() or os.path.expanduser("~")
            )
            if d:
                self.out_var.set(d)
                self._btn_release(self.btn_dir, text=d, released_bg="#FFFFFF", released_fg="#333333")
            else:
                self._btn_release(self.btn_dir, text="选择目录…", released_bg="#FFFFFF", released_fg="#333333")
        except Exception:
            self._btn_release(self.btn_dir, text="选择目录…", released_bg="#FFFFFF", released_fg="#333333")
            raise

    # —— 执行 —— 
    def run(self):
        in_path = self.in_var.get().strip()
        out_dir = self.out_var.get().strip()
        if not (in_path and os.path.isfile(in_path)):
            messagebox.showwarning("提示", "请先选择要处理的 Excel 文件。")
            return
        if not (out_dir and os.path.isdir(out_dir)):
            messagebox.showwarning("提示", "请先选择输出文件夹。")
            return
        try:
            self._btn_press(self.btn_run, pressed_bg="#2E6EF7", pressed_fg="#FFFFFF")
            self.btn_run.config(text="处理中…")
            self.status.set("处理中…"); self.root.update_idletasks()

            n_nc, n_c = process_excel(in_path, out_dir)
            total = n_nc + n_c
            if total:
                ratio_nc = n_nc / total
                ratio_c = n_c / total
            else:
                ratio_nc = ratio_c = 0.0

            messagebox.showinfo(
                "完成",
                (f"处理完成！\n\n"
                 f"没跨组：{n_nc}（{ratio_nc:.1%}）\n"
                 f"跨组：{n_c}（{ratio_c:.1%}）\n\n"
                 f"输出目录：\n{out_dir}\n已生成：\n"
                 f"- 重复来电_没跨组.xlsx\n- 重复来电_跨组.xlsx")
            )
            self.status.set("完成")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("出错了", str(e))
            self.status.set("出错")
        finally:
            self._btn_release(self.btn_run, text="开始处理", released_bg="#34A853", released_fg="#FFFFFF")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
