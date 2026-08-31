import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def process_exam(selected_subject, input_file_path):
    try:
        # 決定科目標題名稱 (英聽不加「科」)
        if selected_subject == "英聽":
            subj_cover = "英聽試題本"
            subj_text = "英聽"
        else:
            subj_cover = f"{selected_subject}科試題本"
            subj_text = f"{selected_subject}科"

        # 讀取 Excel 檔案
        df = pd.read_excel(input_file_path)
        df = df.replace(np.nan, '', regex=True)

        # ==========================================
        # 1. 建立「題本」導言區 (使用佔位符 [SUBJECT_COVER] 等)
        # ==========================================
        exam_latex = r"""\documentclass[12pt, a4paper]{article}
\usepackage{geometry}
\geometry{top=1.8cm, bottom=1.8cm, left=2.5cm, right=2.5cm}
\usepackage{graphicx}
\usepackage{xeCJK}
\setCJKmainfont{Noto Serif CJK TC}

\usepackage{fancyhdr}
\usepackage{lastpage}
\usepackage{ifthen}
\usepackage{pifont}
\usepackage{enumitem}
\usepackage{tikz}
\usepackage{tcolorbox}

\newcommand{\cW}[1]{\tikz[baseline=-0.1ex] \draw (0,0) circle (1.2ex) node {\small #1};}
\newcommand{\cB}[1]{\tikz[baseline=-0.1ex] \draw[fill=black] (0,0) circle (1.2ex) node[text=white] {\small #1};}
\newcommand{\cH}[1]{\tikz[baseline=-0.1ex] { \draw (0,0) circle (1.2ex) node {\small #1}; \fill[black] (-0.8ex,-1ex) rectangle (0ex,1ex); }}
\newcommand{\cG}[1]{\tikz[baseline=-0.1ex] { \fill[gray!40] (0,0) circle (1.2ex); \draw (0,0) circle (1.2ex) node {\small #1}; }}
\newcommand{\cO}[1]{\tikz[baseline=-0.1ex] { \draw (0,0) circle (1.2ex) node {\small #1}; \fill[black] (0.3ex,-0.3ex) circle (1.3ex); }}

\pagestyle{fancy}
\fancyhf{} 
\renewcommand{\headrulewidth}{0pt} 
\fancyfoot[C]{\thepage}
\fancyfoot[R]{%
    \ifthenelse{\value{page}=\pageref{LastPage}}%
    {\fbox{\textbf{試題結束}}}%
    {\fbox{\textbf{請翻到下一頁}}}%
}
\setlength{\parindent}{0pt}

\begin{document}
\pagestyle{empty} 

\begin{center}
    {\fontsize{24pt}{30pt}\selectfont \textbf{115年吾思國中會考模擬考}\ding{172}}\\[0.4cm] 
    {\fontsize{28pt}{36pt}\selectfont \textbf{[SUBJECT_COVER]}}\\[0.6cm] 

    \begin{tcolorbox}[colframe=black, colback=white, sharp corners, boxrule=1pt, width=1.0\textwidth, center, top=4mm, bottom=4mm] 
        \begin{center}
            {\fontsize{20pt}{24pt}\selectfont 請不要翻到次頁！}\\[0.3cm]
            {\fontsize{16pt}{22pt}\selectfont 讀完本頁的說明，聽從監試委員的指示才開始作答！}\\[0.5cm]
            {\fontsize{12pt}{16pt}\selectfont ※ 請先確認你的答案卡、准考證與座位號碼是否一致無誤。}
        \end{center}
    \end{tcolorbox}
\end{center}

\vspace{0.2cm} 

\begin{tcolorbox}[colframe=black, colback=white, sharp corners, boxrule=1pt, width=1.0\textwidth, center, top=4mm, bottom=4mm]
    \noindent\textbf{\Large 請閱讀以下測驗作答說明：}\\[0.2cm]
    \textbf{\large 測驗說明：}\\
    這是國中教育會考[SUBJECT_TEXT]試題本，試題本採雙面印刷，共18頁，有54題選擇題，每題都只有一個正確或最佳的答案。測驗時間從10:40到11:50，共70分鐘。作答開始與結束請聽從監試委員的指示。\\[0.2cm]
    
    \textbf{\large 注意事項：}
    \begin{enumerate}[label=\arabic*., leftmargin=1.5em, itemsep=2pt, parsep=0pt]
        \item 所有試題均為四選一的選擇題，答錯不倒扣。
        \item 試題中所附圖形，如有附上比例尺，以比例尺為依據作答；若無比例尺，則該圖僅作為參考，不代表實際大小。
        \item 可利用試題本中空白部分計算，切勿在答案卡上計算。
        \item 故意損壞試題本，或於答案卡上挖補、汙損、折疊、作標記、顯示自己身分，均屬違反試場規則行為，依簡章違規處理要點論處。
    \end{enumerate}

    \vspace{0.2cm} 

    \textbf{\large 作答方式：}\\[0.1cm]
    \hspace*{1.5em}請依照題意從四個選項中選出\underline{一個}正確或最佳的答案，並用 \textbf{2B} 鉛筆在答案卡上相應的位置畫記，請務必將選項塗黑、塗滿。如果需要修改答案，請使用橡皮擦擦拭乾淨，重新塗黑答案。例如答案為 \textbf{B}，則將 \cW{B} 選項塗黑、塗滿，即： \cW{A} \cB{B} \cW{C} \cW{D}\\[0.2cm]
    \hspace*{1.5em}以下為錯誤的畫記方式，可能導致電腦無法正確判讀。如：

    \vspace{0.1cm}

    \begin{itemize}[label={}, leftmargin=3.5em, itemsep=2pt]
        \item \cW{A} \cH{B} \cW{C} \cW{D} \quad —未將選項塗滿
        \item \cW{A} \cG{B} \cW{C} \cW{D} \quad —未將選項塗黑
        \item \cW{A} \cB{B} \cG{C} \cW{D} \quad —未擦拭乾淨
        \item \cW{A} \cO{B} \cW{C} \cW{D} \quad —塗出選項外
        \item \cW{A} \cB{B} \cB{C} \cW{D} \quad —同時塗兩個選項
    \end{itemize}
\end{tcolorbox}

\newpage
\null 
\newpage
\pagestyle{fancy} 
\setcounter{page}{1} 
"""
        
        # 替換題本中的科目變數
        exam_latex = exam_latex.replace("[SUBJECT_COVER]", subj_cover).replace("[SUBJECT_TEXT]", subj_text)

        # ==========================================
        # 2. 建立「解答本」導言區
        # ==========================================
        ans_latex = r"""\documentclass[11pt, a4paper]{article}
\usepackage{geometry}
\geometry{top=2cm, bottom=2cm, left=1.5cm, right=1.5cm}
\usepackage{xeCJK}
\setCJKmainfont{Noto Serif CJK TC}
\usepackage{multicol}
\usepackage{tcolorbox}
\usepackage{tabularx}

\setlength{\parindent}{0pt}
\setlength{\columnsep}{1cm}
\setlength{\columnseprule}{0.5pt}

\begin{document}
\begin{center}
    {\fontsize{18pt}{24pt}\selectfont \textbf{115年吾思國中會考模擬考}}\\[0.2cm] 
    {\fontsize{22pt}{28pt}\selectfont \textbf{[SUBJECT_TEXT] 解答與解析}}
\end{center}
\vspace{0.3cm}
"""
        # 替換解答本中的科目變數
        ans_latex = ans_latex.replace("[SUBJECT_TEXT]", subj_text)

        # ==========================================
        # 3. 處理資料與產出題目
        # ==========================================
        answers_data = [] 

        for index, row in df.iterrows():
            row = row.rename(index=lambda x: str(x).strip())
            q_type = str(row.get('資料類型', '')).strip()
            if not q_type: continue

            q_num = str(row.get('題號', '')).replace('.0', '')
            q_text = str(row.get('題目敘述', '')).strip()
            stem_img = str(row.get('題幹圖片', '')).strip()
            stem_pos = str(row.get('題幹圖位置', '')).strip()
            opt_type = str(row.get('選項類型', '')).strip()

            opt_a = str(row.get('選項 (A)', row.get('選項A', ''))).strip()
            opt_b = str(row.get('選項 (B)', row.get('選項B', ''))).strip()
            opt_c = str(row.get('選項 (C)', row.get('選項C', ''))).strip()
            opt_d = str(row.get('選項 (D)', row.get('選項D', ''))).strip()

            ans = str(row.get('答案', '')).strip().upper()
            exp = str(row.get('詳解', '')).strip()

            if q_num and ans:
                answers_data.append({"num": q_num, "ans": ans, "exp": exp})

            if q_type == '題組文章':
                exam_latex += "\\textbf{閱讀下列資訊並回答問題}\\\\\n"
                if stem_img:
                    if stem_pos == '靠右':
                        exam_latex += "\\begin{minipage}[t]{0.65\\textwidth}\n" + f"{q_text}\n" + "\\end{minipage}%\n"
                        exam_latex += "\\begin{minipage}[t]{0.3\\textwidth}\n" + f"\\vspace{{0pt}}\\includegraphics[width=\\linewidth]{{{stem_img}}}\n" + "\\end{minipage}\n\n"
                    else: 
                        exam_latex += f"{q_text}\\\\\n\\begin{{center}}\n\\includegraphics[width=0.5\\linewidth]{{{stem_img}}}\n\\end{{center}}\n"
                else:
                    exam_latex += f"{q_text}\n\n"
                exam_latex += "\\vspace{0.3cm}\n"
                continue

            exam_latex += "\\begin{minipage}{\\linewidth}\n"
            exam_latex += f"\\textbf{{{q_num}.}} "

            if stem_img:
                if stem_pos == '靠右':
                    exam_latex += "\\begin{minipage}[t]{0.65\\textwidth}\n" + f"{q_text}\n" + "\\end{minipage}%\n"
                    exam_latex += "\\begin{minipage}[t]{0.3\\textwidth}\n" + f"\\vspace{{0pt}}\\includegraphics[width=\\linewidth]{{{stem_img}}}\n" + "\\end{minipage}\n\n"
                else:
                    exam_latex += f"{q_text}\\\\\n\\begin{{center}}\n\\includegraphics[width=0.5\\linewidth]{{{stem_img}}}\n\\end{{center}}\n"
            else:
                exam_latex += f"{q_text}\n\n"

            if opt_type == '圖片':
                exam_latex += f"\\makebox[0.24\\textwidth][l]{{(A) \\includegraphics[width=2.5cm]{{{opt_a}}}}} "
                exam_latex += f"\\makebox[0.24\\textwidth][l]{{(B) \\includegraphics[width=2.5cm]{{{opt_b}}}}} "
                exam_latex += f"\\makebox[0.24\\textwidth][l]{{(C) \\includegraphics[width=2.5cm]{{{opt_c}}}}} "
                exam_latex += f"\\makebox[0.24\\textwidth][l]{{(D) \\includegraphics[width=2.5cm]{{{opt_d}}}}}\\\\\n"
            else:
                max_len = max(len(opt_a), len(opt_b), len(opt_c), len(opt_d))
                if max_len > 14:
                    exam_latex += f"(A) {opt_a}\\\\\n(B) {opt_b}\\\\\n(C) {opt_c}\\\\\n(D) {opt_d}\\\\\n"
                elif max_len > 6:
                    exam_latex += f"\\makebox[0.48\\textwidth][l]{{(A) {opt_a}}} \\makebox[0.48\\textwidth][l]{{(B) {opt_b}}}\\\\\n"
                    exam_latex += f"\\makebox[0.48\\textwidth][l]{{(C) {opt_c}}} \\makebox[0.48\\textwidth][l]{{(D) {opt_d}}}\\\\\n"
                else:
                    exam_latex += f"\\makebox[0.24\\textwidth][l]{{(A) {opt_a}}} \\makebox[0.24\\textwidth][l]{{(B) {opt_b}}} "
                    exam_latex += f"\\makebox[0.24\\textwidth][l]{{(C) {opt_c}}} \\makebox[0.24\\textwidth][l]{{(D) {opt_d}}}\\\\\n"

            exam_latex += "\\end{minipage}\n\\vspace{0.8cm}\n\n"

        exam_latex += r"\end{document}"

        # ==========================================
        # 4. 自動生成解答本內容
        # ==========================================
        ans_latex += "\\begin{tcolorbox}[colframe=black, colback=gray!10, sharp corners, boxrule=1pt, title=\\textbf{選擇題解答一覽表}]\n"
        ans_latex += "\\renewcommand{\\arraystretch}{1.5}\n"
        ans_latex += "\\begin{tabularx}{\\textwidth}{|c|X|X|X|X|X|X|X|X|X|X|}\n\\hline\n"
        
        for i in range(0, len(answers_data), 10):
            chunk = answers_data[i:i+10]
            start_num = str(chunk[0]['num']).zfill(2)
            end_num = str(chunk[-1]['num']).zfill(2)
            row_label = f"\\textbf{{{start_num}-{end_num}}}"
            
            row_answers = [f"({item['ans']})" for item in chunk]
            while len(row_answers) < 10:
                row_answers.append("")
                
            ans_latex += f"    {row_label} & " + " & ".join(row_answers) + " \\\\ \\hline\n"
            
        ans_latex += "\\end{tabularx}\n\\end{tcolorbox}\n\\vspace{0.5cm}\n"

        ans_latex += "\\textbf{\\Large 試題詳解}\n\\vspace{0.3cm}\n\\begin{multicols*}{2}\n"
        for item in answers_data:
            exp_text = item['exp'] if item['exp'] else "略"
            ans_latex += f"\\textbf{{{item['num']}. ({item['ans']})}}\\\\\n解析：{exp_text}\n\n\\vspace{{0.3cm}}\n"
        ans_latex += "\\end{multicols*}\n\\end{document}"

        # ==========================================
        # 5. 自動儲存在 Excel 所在目錄
        # ==========================================
        base_dir = os.path.dirname(input_file_path)
        exam_filename = os.path.join(base_dir, f"{subj_text}_Exam_Paper.tex")
        ans_filename = os.path.join(base_dir, f"{subj_text}_Answer_Key.tex")
        
        with open(exam_filename, 'w', encoding='utf-8') as f:
            f.write(exam_latex)
        with open(ans_filename, 'w', encoding='utf-8') as f:
            f.write(ans_latex)

        messagebox.showinfo("轉換成功！", f"🎉 題本與解答本已產生完成！\n\n檔案儲存於：\n{base_dir}")

    except Exception as e:
        messagebox.showerror("發生錯誤", f"轉換過程中出現錯誤：\n{str(e)}")

# ==========================================
# 建立 GUI 視窗介面
# ==========================================
def main():
    root = tk.Tk()
    root.title("吾思會考題本生成器")
    root.geometry("320x350")
    
    # 標題
    tk.Label(root, text="請選擇要生成的科目", font=("Arial", 14, "bold")).pack(pady=15)
    
    # 建立單選按鈕變數
    subject_var = tk.StringVar(value="社會")
    subjects = ["國文", "英文", "數學", "社會", "自然", "英聽"]
    
    # 繪製單選按鈕
    frame = tk.Frame(root)
    frame.pack()
    for sub in subjects:
        tk.Radiobutton(frame, text=sub, variable=subject_var, value=sub, font=("Arial", 12)).pack(anchor="w", pady=2)
        
    def start_conversion():
        selected_subject = subject_var.get()
        input_file_path = filedialog.askopenfilename(
            title="請選擇題庫 Excel 檔案",
            filetypes=[("Excel 活頁簿", "*.xlsx"), ("所有檔案", "*.*")]
        )
        if input_file_path:
            process_exam(selected_subject, input_file_path)

    # 執行按鈕
    tk.Button(root, text="選擇 Excel 檔案並轉檔", font=("Arial", 12), bg="#4CAF50", fg="white", command=start_conversion).pack(pady=20)
    
    root.mainloop()

if __name__ == "__main__":
    main()