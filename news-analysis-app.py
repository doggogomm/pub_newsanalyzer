import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime
import os
import re
from collections import Counter

# Optional imports with error handling
try:
    from wordcloud import WordCloud
    wordcloud_available = True
except ImportError:
    wordcloud_available = False

try:
    from konlpy.tag import Okt
    okt_available = True
except ImportError:
    okt_available = False

class NewsAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("뉴스 분석 애플리케이션")
        self.root.geometry("1200x800")
        
        # Initialize variables
        self.df = None
        self.date_column = None
        self.title_column = None
        self.content_column = None
        self.source_column = None
        self.stopwords = set()
        
        # Create main frame
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create top frame for file loading and settings
        self.top_frame = ttk.Frame(self.main_frame)
        self.top_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Create file loading section
        self.file_frame = ttk.LabelFrame(self.top_frame, text="파일 로드")
        self.file_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        
        self.file_path_var = tk.StringVar()
        self.file_path_entry = ttk.Entry(self.file_frame, textvariable=self.file_path_var, width=50)
        self.file_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        
        self.browse_button = ttk.Button(self.file_frame, text="찾아보기", command=self.browse_file)
        self.browse_button.pack(side=tk.LEFT, padx=5, pady=5)
        
        self.load_button = ttk.Button(self.file_frame, text="로드", command=self.load_file)
        self.load_button.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Create middle frame for column selection
        self.middle_frame = ttk.Frame(self.main_frame)
        self.middle_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Create column selection section
        self.column_frame = ttk.LabelFrame(self.middle_frame, text="컬럼 선택")
        self.column_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Date column
        ttk.Label(self.column_frame, text="날짜/시간 컬럼:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.date_column_var = tk.StringVar()
        self.date_column_combo = ttk.Combobox(self.column_frame, textvariable=self.date_column_var, state="readonly", width=20)
        self.date_column_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Title column
        ttk.Label(self.column_frame, text="제목 컬럼:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.title_column_var = tk.StringVar()
        self.title_column_combo = ttk.Combobox(self.column_frame, textvariable=self.title_column_var, state="readonly", width=20)
        self.title_column_combo.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
        # Content column
        ttk.Label(self.column_frame, text="내용 컬럼:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.content_column_var = tk.StringVar()
        self.content_column_combo = ttk.Combobox(self.column_frame, textvariable=self.content_column_var, state="readonly", width=20)
        self.content_column_combo.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Source column
        ttk.Label(self.column_frame, text="매체명 컬럼:").grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        self.source_column_var = tk.StringVar()
        self.source_column_combo = ttk.Combobox(self.column_frame, textvariable=self.source_column_var, state="readonly", width=20)
        self.source_column_combo.grid(row=1, column=3, padx=5, pady=5, sticky=tk.W)
        
        # Date format
        ttk.Label(self.column_frame, text="날짜 형식:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.date_format_var = tk.StringVar(value="%Y-%m-%d %H:%M:%S")
        self.date_format_entry = ttk.Entry(self.column_frame, textvariable=self.date_format_var, width=20)
        self.date_format_entry.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(self.column_frame, text="(예: %Y-%m-%d %H:%M:%S)").grid(row=2, column=2, padx=5, pady=5, sticky=tk.W)
        
        # Period settings
        self.period_frame = ttk.LabelFrame(self.middle_frame, text="기간 설정")
        self.period_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Period type
        ttk.Label(self.period_frame, text="기간 유형:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.period_type_var = tk.StringVar(value="월별")
        self.period_type_combo = ttk.Combobox(self.period_frame, textvariable=self.period_type_var, values=["일별", "월별", "년별"], state="readonly", width=15)
        self.period_type_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Date range
        ttk.Label(self.period_frame, text="시작일:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.start_date_var = tk.StringVar()
        self.start_date_entry = ttk.Entry(self.period_frame, textvariable=self.start_date_var, width=15)
        self.start_date_entry.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(self.period_frame, text="종료일:").grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        self.end_date_var = tk.StringVar()
        self.end_date_entry = ttk.Entry(self.period_frame, textvariable=self.end_date_var, width=15)
        self.end_date_entry.grid(row=0, column=5, padx=5, pady=5, sticky=tk.W)
        ttk.Label(self.period_frame, text="(형식: YYYY-MM-DD)").grid(row=0, column=6, padx=5, pady=5, sticky=tk.W)
        
        # Stopwords section
        self.stopwords_frame = ttk.LabelFrame(self.middle_frame, text="불용어 설정")
        self.stopwords_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.stopwords_text = scrolledtext.ScrolledText(self.stopwords_frame, width=60, height=5)
        self.stopwords_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.stopwords_button_frame = ttk.Frame(self.stopwords_frame)
        self.stopwords_button_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)
        
        self.add_stopwords_button = ttk.Button(self.stopwords_button_frame, text="추가", command=self.add_stopwords)
        self.add_stopwords_button.pack(pady=5)
        
        self.clear_stopwords_button = ttk.Button(self.stopwords_button_frame, text="초기화", command=self.clear_stopwords)
        self.clear_stopwords_button.pack(pady=5)
        
        # Default stopwords
        self.default_stopwords = ["뉴진스", "멤버", "그룹", "데뷔", "앨범", "활동", "팬", "음악", "무대", 
                              "스타일", "패션", "인기", "사랑", "응원", "기대", "관심", "컴백", "노래"]
        self.stopwords_text.insert(tk.END, ", ".join(self.default_stopwords))
        self.add_stopwords()
        
        # Analysis button
        self.analysis_button_frame = ttk.Frame(self.middle_frame)
        self.analysis_button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.analysis_button = ttk.Button(self.analysis_button_frame, text="분석 실행", command=self.run_analysis)
        self.analysis_button.pack(padx=5, pady=5)
        
        # Create notebook for results
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create tabs for different visualizations
        self.time_tab = ttk.Frame(self.notebook)
        self.source_tab = ttk.Frame(self.notebook)
        self.wordcloud_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.time_tab, text="기간별 분석")
        self.notebook.add(self.source_tab, text="매체별 분석")
        if wordcloud_available:
            self.notebook.add(self.wordcloud_tab, text="워드클라우드")
        
        # Create frames to hold graphs
        self.time_graph_frame = ttk.Frame(self.time_tab)
        self.time_graph_frame.pack(fill=tk.BOTH, expand=True)
        
        self.source_graph_frame = ttk.Frame(self.source_tab)
        self.source_graph_frame.pack(fill=tk.BOTH, expand=True)
        
        if wordcloud_available:
            self.wordcloud_frame = ttk.Frame(self.wordcloud_tab)
            self.wordcloud_frame.pack(fill=tk.BOTH, expand=True)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_var.set("준비 완료")
        
        # Initialize graph figures and canvases
        self.time_fig = plt.Figure(figsize=(10, 6), dpi=100)
        self.time_canvas = FigureCanvasTkAgg(self.time_fig, self.time_graph_frame)
        self.time_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        self.source_fig = plt.Figure(figsize=(10, 6), dpi=100)
        self.source_canvas = FigureCanvasTkAgg(self.source_fig, self.source_graph_frame)
        self.source_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        if wordcloud_available:
            self.wordcloud_fig = plt.Figure(figsize=(10, 6), dpi=100)
            self.wordcloud_canvas = FigureCanvasTkAgg(self.wordcloud_fig, self.wordcloud_frame)
            self.wordcloud_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV 파일", "*.csv"), ("Excel 파일", "*.xlsx;*.xls")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            
    def load_file(self):
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showerror("오류", "파일을 선택해주세요.")
            return
            
        try:
            # Load file based on extension
            if file_path.lower().endswith('.csv'):
                self.df = pd.read_csv(file_path, encoding='utf-8')
            elif file_path.lower().endswith(('.xlsx', '.xls')):
                self.df = pd.read_excel(file_path)
            else:
                messagebox.showerror("오류", "지원하지 않는 파일 형식입니다.")
                return
                
            # Update column selection dropdowns
            columns = self.df.columns.tolist()
            
            self.date_column_combo['values'] = columns
            self.title_column_combo['values'] = columns
            self.content_column_combo['values'] = columns
            self.source_column_combo['values'] = columns
            
            # Try to guess the correct columns based on common names
            for col in columns:
                lower_col = col.lower()
                if any(date_term in lower_col for date_term in ['날짜', 'date', 'time', '시각']):
                    self.date_column_var.set(col)
                elif any(title_term in lower_col for title_term in ['제목', 'title']):
                    self.title_column_var.set(col)
                elif any(content_term in lower_col for content_term in ['내용', 'content', 'text']):
                    self.content_column_var.set(col)
                elif any(source_term in lower_col for source_term in ['매체', 'source', 'media']):
                    self.source_column_var.set(col)
            
            # Set default date range
            if self.date_column_var.get():
                try:
                    date_col = self.date_column_var.get()
                    date_format = self.date_format_var.get()
                    
                    # Convert date column to datetime
                    self.df[date_col] = pd.to_datetime(self.df[date_col], format=date_format, errors='coerce')
                    
                    # Set date range
                    min_date = self.df[date_col].min().strftime('%Y-%m-%d')
                    max_date = self.df[date_col].max().strftime('%Y-%m-%d')
                    
                    self.start_date_var.set(min_date)
                    self.end_date_var.set(max_date)
                except:
                    pass
            
            # Show sample in status bar
            self.status_var.set(f"파일 로드 완료: {file_path} (행: {len(self.df)}, 열: {len(self.df.columns)})")
            messagebox.showinfo("성공", f"파일 로드 완료: {os.path.basename(file_path)}")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일 로드 중 오류 발생: {str(e)}")
            
    def add_stopwords(self):
        text = self.stopwords_text.get("1.0", tk.END).strip()
        words = [word.strip() for word in re.split(r'[,\s]+', text) if word.strip()]
        self.stopwords = set(words)
        self.status_var.set(f"불용어 {len(self.stopwords)}개 설정됨")
        
    def clear_stopwords(self):
        self.stopwords_text.delete("1.0", tk.END)
        self.stopwords = set()
        self.status_var.set("불용어 초기화 완료")
        
    def run_analysis(self):
        if self.df is None:
            messagebox.showerror("오류", "먼저 파일을 로드해주세요.")
            return
            
        date_column = self.date_column_var.get()
        title_column = self.title_column_var.get()
        content_column = self.content_column_var.get()
        source_column = self.source_column_var.get()
        
        if not date_column or not title_column:
            messagebox.showerror("오류", "날짜와 제목 컬럼은 필수 선택사항입니다.")
            return
            
        try:
            # Process date column
            if isinstance(self.df[date_column].iloc[0], str):
                date_format = self.date_format_var.get()
                self.df[date_column] = pd.to_datetime(self.df[date_column], format=date_format, errors='coerce')
                
            # Filter by date range if provided
            start_date = self.start_date_var.get()
            end_date = self.end_date_var.get()
            
            df_filtered = self.df.copy()
            
            if start_date:
                start_date = pd.to_datetime(start_date)
                df_filtered = df_filtered[df_filtered[date_column] >= start_date]
                
            if end_date:
                end_date = pd.to_datetime(end_date)
                df_filtered = df_filtered[df_filtered[date_column] <= end_date]
                
            period_type = self.period_type_var.get()
            
            # Create period column based on selection
            if period_type == "일별":
                df_filtered['period'] = df_filtered[date_column].dt.date
            elif period_type == "월별":
                df_filtered['period'] = df_filtered[date_column].dt.to_period('M').astype(str)
            elif period_type == "년별":
                df_filtered['period'] = df_filtered[date_column].dt.year
                
            # Generate time-based analysis
            self.generate_time_chart(df_filtered, title_column)
            
            # Generate source-based analysis
            if source_column:
                self.generate_source_chart(df_filtered, source_column, title_column)
                
            # Generate word cloud
            if wordcloud_available and content_column and content_column in df_filtered.columns:
                self.generate_wordcloud(df_filtered, content_column)
                
            self.status_var.set(f"분석 완료 - {len(df_filtered)}개 항목 분석됨")
            
        except Exception as e:
            messagebox.showerror("오류", f"분석 중 오류 발생: {str(e)}")
            
    def generate_time_chart(self, df, title_column):
        self.time_fig.clear()
        ax = self.time_fig.add_subplot(111)
        
        # Group by period and count
        period_counts = df.groupby('period')[title_column].count()
        
        # Plotting
        period_counts.plot(kind='line', marker='o', ax=ax)
        
        ax.set_title('기간별 기사 수')
        ax.set_xlabel('기간')
        ax.set_ylabel('기사 수')
        ax.grid(True)
        
        # Rotate x-axis labels if there are many periods
        if len(period_counts) > 10:
            plt.setp(ax.get_xticklabels(), rotation=45, ha='right')
            
        self.time_fig.tight_layout()
        self.time_canvas.draw()
        
    def generate_source_chart(self, df, source_column, title_column):
        self.source_fig.clear()
        ax = self.source_fig.add_subplot(111)
        
        # Group by source and count
        source_counts = df.groupby(source_column)[title_column].count().sort_values(ascending=False).head(15)
        
        # Plotting
        source_counts.plot(kind='bar', ax=ax, color='skyblue')
        
        ax.set_title('매체별 기사 수 (상위 15개)')
        ax.set_xlabel('매체명')
        ax.set_ylabel('기사 수')
        ax.grid(True, axis='y')
        
        # Rotate x-axis labels
        plt.setp(ax.get_xticklabels(), rotation=45, ha='right')
            
        self.source_fig.tight_layout()
        self.source_canvas.draw()
        
    def generate_wordcloud(self, df, content_column):
        if not wordcloud_available:
            return
            
        self.wordcloud_fig.clear()
        ax = self.wordcloud_fig.add_subplot(111)
        
        # Combine all text
        all_text = ' '.join(df[content_column].dropna().astype(str))
        
        # Process text
        if okt_available:
            # Use Konlpy/Okt for Korean text processing
            okt = Okt()
            words = []
            try:
                pos_tagged = okt.pos(all_text)
                words = [word for word, pos in pos_tagged if pos == 'Noun' and word not in self.stopwords and len(word) > 1]
            except Exception as e:
                messagebox.showwarning("경고", f"한국어 형태소 분석 중 오류 발생: {str(e)}\n기본 텍스트 분석으로 대체합니다.")
                # Fall back to basic tokenization
                words = [word for word in re.findall(r'\w+', all_text) if word not in self.stopwords and len(word) > 1]
        else:
            # Basic tokenization for non-Korean text
            words = [word for word in re.findall(r'\w+', all_text) if word not in self.stopwords and len(word) > 1]
            
        word_counts = Counter(words)
        
        # Generate word cloud
        if word_counts:
            try:
                # Try to use a Korean font if available
                font_path = None
                for font in ['malgun', 'NanumGothic', 'AppleGothic', None]:
                    try:
                        if font is None:
                            wordcloud = WordCloud(width=800, height=400, background_color='white', 
                                             max_words=200, contour_width=1, contour_color='steelblue')
                        else:
                            wordcloud = WordCloud(font_path=font, width=800, height=400, background_color='white', 
                                             max_words=200, contour_width=1, contour_color='steelblue')
                        wordcloud.generate_from_frequencies(word_counts)
                        break
                    except:
                        continue
                
                ax.imshow(wordcloud, interpolation='bilinear')
                ax.set_title('워드클라우드')
                ax.axis('off')
                
                self.wordcloud_fig.tight_layout()
                self.wordcloud_canvas.draw()
            except Exception as e:
                messagebox.showwarning("경고", f"워드클라우드 생성 중 오류 발생: {str(e)}")
        else:
            ax.text(0.5, 0.5, "텍스트 데이터가 충분하지 않습니다.", ha='center', va='center', fontsize=12)
            ax.axis('off')
            self.wordcloud_canvas.draw()

if __name__ == "__main__":
    root = tk.Tk()
    app = NewsAnalysisApp(root)
    root.mainloop()
