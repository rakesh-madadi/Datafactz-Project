import pandas as pd
import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import seaborn as sns

class InteractiveSalesDashboard:
    def __init__(self, master):
        self.quantity_sort_order = 'desc'
        self.revenue_sort_order = 'desc'
        self.current_sort_button = None  # Track the current sort button
        self.master = master
        self.master.title("Sales Data Analysis Of Electronics")
        self.master.geometry("1200x800")  # Increased window size
        
        self.style = ttk.Style("darkly")
        
        self.create_widgets()
        self.load_and_process_data()
        self.current_chart = None

    def create_widgets(self):
        # Main frame
        self.main_frame = ttk.Frame(self.master, padding="15")  # Increased padding
        self.main_frame.pack(fill=BOTH, expand=YES)

        # Left frame for summary and controls
        self.left_frame = ttk.Frame(self.main_frame, width=400)  # Increased width
        self.left_frame.pack(side=LEFT, fill=Y, padx=(0, 15))  # Increased padding

        # Summary treeview
        self.tree_frame = ttk.Frame(self.left_frame)
        self.tree_frame.pack(fill=BOTH, expand=YES, pady=(0, 15))  # Increased padding

        self.tree = ttk.Treeview(self.tree_frame, columns=("Product Name", "Total Quantity", "Total Revenue"), show='headings', height=20)  # Increased height
        self.tree.heading("Product Name", text="Product Name")
        self.tree.heading("Total Quantity", text="Total Quantity")
        self.tree.heading("Total Revenue", text="Total Revenue")
        self.tree.column("Product Name", width=200)  # Increased width
        self.tree.column("Total Quantity", width=150)  # Increased width
        self.tree.column("Total Revenue", width=200)  # Increased width

        # Increase font size
        self.tree.tag_configure('large', font=('Helvetica', 14))  # Larger font size
        self.tree.configure(style="Custom.Treeview")
        self.style.configure("Custom.Treeview", font=('Helvetica', 14))

        self.tree.pack(side=LEFT, fill=BOTH, expand=YES)

        self.tree_scroll = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree_scroll.pack(side=RIGHT, fill=Y)
        self.tree.configure(yscrollcommand=self.tree_scroll.set)

        # Control buttons
        self.btn_frame = ttk.Frame(self.left_frame)
        self.btn_frame.pack(fill=X)

        ttk.Button(self.btn_frame, text="Pie Chart", command=self.show_pie_chart).pack(side=TOP, fill=X, pady=5)  # Increased padding
        ttk.Button(self.btn_frame, text="Quantity Bar", command=self.show_quantity_bar_chart).pack(side=TOP, fill=X, pady=5)  # Increased padding
        ttk.Button(self.btn_frame, text="Revenue Bar", command=self.show_revenue_bar_chart).pack(side=TOP, fill=X, pady=5)  # Increased padding
        ttk.Button(self.btn_frame, text="Time Series", command=self.show_quantity_line_graph).pack(side=TOP, fill=X, pady=5)  # Increased padding

        # Right frame for charts
        self.right_frame = ttk.Frame(self.main_frame)
        self.right_frame.pack(side=LEFT, fill=BOTH, expand=YES)

        # Chart area
        self.chart_frame = ttk.Frame(self.right_frame)
        self.chart_frame.pack(fill=BOTH, expand=YES)

    def load_and_process_data(self):
        try:
            file_path = "C:\\Users\\Rakesh Reddy\\Desktop\\sales_data.xlsx"
            self.df = pd.read_excel(file_path)
            self.df = self.clean_data(self.df)
            if self.df is not None:
                self.summary_df = self.format_data(self.df)
                self.display_data()
        except Exception as e:
            messagebox.showerror("Error", f"Error loading or processing data: {str(e)}")

    def clean_data(self, df):
        try:
            df = df.dropna()
            df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
            df['Sales Price'] = pd.to_numeric(df['Sales Price'], errors='coerce')
            df['Date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y', errors='coerce')
            df = df.dropna()
            return df
        except Exception as e:
            messagebox.showerror("Error", f"Error during data cleaning: {str(e)}")
            return None

    def format_data(self, df):
        try:
            df['Total Revenue'] = df['Quantity'] * df['Sales Price'] / 1e6  # Convert to millions
            summary_df = df.groupby('Product Name').agg(
                Total_Quantity_Sold=('Quantity', 'sum'),
                Total_Revenue_M=('Total Revenue', 'sum')
            ).reset_index()
            return summary_df
        except Exception as e:
            messagebox.showerror("Error", f"Error during data formatting: {str(e)}")
            return None

    def display_data(self):
        if self.summary_df is not None and not self.summary_df.empty:
            self.tree.delete(*self.tree.get_children())
            for _, row in self.summary_df.iterrows():
                self.tree.insert("", "end", values=(row['Product Name'], f"{row['Total_Quantity_Sold']:,.0f}", f"${row['Total_Revenue_M']:,.2f}M"), tags=('large',))
        else:
            messagebox.showinfo("No Data", "No summary data to display.")

    def show_pie_chart(self):
        if self.summary_df is not None:
            fig, ax = plt.subplots(figsize=(8, 6), dpi=100)  # Increased figure size
            colors = sns.color_palette('pastel')
            wedges, texts, autotexts = ax.pie(self.summary_df['Total_Revenue_M'], labels=self.summary_df['Product Name'], 
                                              autopct='%1.1f%%', startangle=90, colors=colors)
            ax.set_title('Sales Distribution by Product')
            plt.tight_layout()
            
            def hover(event):
                for wedge in wedges:
                    if wedge.contains_point([event.x, event.y]):
                        wedge.set_edgecolor('white')
                        wedge.set_linewidth(3)
                    else:
                        wedge.set_edgecolor('black')
                        wedge.set_linewidth(1)
                fig.canvas.draw_idle()
            
            fig.canvas.mpl_connect('motion_notify_event', hover)
            
            self.display_plot(fig)

    def show_quantity_bar_chart(self):
        if self.summary_df is not None:
            self.create_sort_buttons(chart_type='quantity')
            self.plot_quantity_bar_chart(order='desc')  # Default to descending order

    def show_revenue_bar_chart(self):
        if self.summary_df is not None:
            self.create_sort_buttons(chart_type='revenue')
            self.plot_revenue_bar_chart(order='desc')  # Default to descending order

    def create_sort_buttons(self, chart_type):
        # Destroy existing sort button if any
        if self.current_sort_button is not None:
            self.current_sort_button.destroy()

        # Add a common sort button that toggles between ascending and descending
        sort_frame = ttk.Frame(self.left_frame)
        sort_frame.pack(fill=X, pady=10)  # Increased padding

        sort_button = ttk.Button(sort_frame, text="Sort Descending", command=lambda: self.toggle_sort_order(chart_type, sort_button))
        sort_button.pack(side=LEFT, padx=10)  # Increased padding

        self.current_sort_button = sort_frame  # Update the current sort button reference

    def toggle_sort_order(self, chart_type, button):
        if chart_type == 'quantity':
            if self.quantity_sort_order == 'desc':
                self.quantity_sort_order = 'asc'
                button.config(text="Sort Ascending")
            else:
                self.quantity_sort_order = 'desc'
                button.config(text="Sort Descending")
            self.plot_quantity_bar_chart(order=self.quantity_sort_order)
        
        elif chart_type == 'revenue':
            if self.revenue_sort_order == 'desc':
                self.revenue_sort_order = 'asc'
                button.config(text="Sort Ascending")
            else:
                self.revenue_sort_order = 'desc'
                button.config(text="Sort Descending")
            self.plot_revenue_bar_chart(order=self.revenue_sort_order)

    def plot_quantity_bar_chart(self, order='desc'):
        sorted_df = self.summary_df.sort_values(by='Total_Quantity_Sold', ascending=(order == 'asc'))

        fig, ax = plt.subplots(figsize=(8, 6), dpi=100)  # Increased figure size
        bars = ax.bar(sorted_df['Product Name'], sorted_df['Total_Quantity_Sold'], color=sns.color_palette('pastel'))

        # Add labels to each bar
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2.0, height,
                    f'{height:,.0f}', ha='center', va='bottom')

        ax.set_title('Total Quantity Sold by Product')
        ax.set_ylabel('Quantity Sold')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()

        self.display_plot(fig)

    def plot_revenue_bar_chart(self, order='desc'):
        sorted_df = self.summary_df.sort_values(by='Total_Revenue_M', ascending=(order == 'asc'))

        fig, ax = plt.subplots(figsize=(8, 6), dpi=100)  # Increased figure size
        bars = ax.bar(sorted_df['Product Name'], sorted_df['Total_Revenue_M'], color=sns.color_palette('pastel'))

        # Add labels to each bar
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2.0, height,
                    f'${height:,.2f}M', ha='center', va='bottom')

        ax.set_title('Total Revenue by Product')
        ax.set_ylabel('Revenue in $M')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()

        self.display_plot(fig)

    def show_quantity_line_graph(self):
        if self.df is not None:
            monthly_df = self.df.groupby(self.df['Date'].dt.to_period('M')).agg({'Quantity': 'sum'}).reset_index()
            monthly_df['Date'] = monthly_df['Date'].dt.to_timestamp()
            
            fig, ax = plt.subplots(figsize=(8, 6), dpi=100)  # Increased figure size
            line = ax.plot(monthly_df['Date'], monthly_df['Quantity'], marker='o', color='b')[0]
            ax.set_title('Total Quantity Sold Over Time')
            ax.set_xlabel('Date')
            ax.set_ylabel('Quantity Sold')
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            
            # Add value labels to each marker
            for x, y in zip(monthly_df['Date'], monthly_df['Quantity']):
                ax.text(x, y, f'{y:,.0f}', ha='right', va='top', fontsize=15, color='black')

            self.display_plot(fig)

    def display_plot(self, fig):
        for widget in self.chart_frame.winfo_children():
            widget.destroy()

        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=YES)

        toolbar = NavigationToolbar2Tk(canvas, self.chart_frame)
        toolbar.update()
        canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=YES)

        self.current_chart = fig

    def update_tree_selection(self, product):
        for item in self.tree.get_children():
            if self.tree.item(item)['values'][0] == product:
                self.tree.selection_set(item)
                self.tree.focus(item)
                self.tree.see(item)
                break

if __name__ == "__main__":
    root = ttk.Window("Interactive Sales Dashboard", "darkly")
    app = InteractiveSalesDashboard(root)
    root.mainloop()
