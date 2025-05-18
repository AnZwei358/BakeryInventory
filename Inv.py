import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
from matplotlib.gridspec import GridSpec
import seaborn as sns


class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestão de Estoque de Padaria")
        self.root.geometry("1400x800")

        self.portuguese_months = [
            'Janeiro', 'Fevereiro', 'Março', 'Abril',
            'Maio', 'Junho', 'Julho', 'Agosto',
            'Setembro', 'Outubro', 'Novembro', 'Dezembro'
        ]

        self.monthly_data = {month: [] for month in self.portuguese_months}
        self.current_month = self.translate_month(datetime.now().month)

        self.create_widgets()
        self.change_analysis()

    def translate_month(self, month_number):
        return self.portuguese_months[month_number - 1]

    def create_widgets(self):
        # Control frame
        self.control_frame = tk.Frame(self.root)
        self.control_frame.pack(pady=10)

        # Dropdown selector
        self.analysis_var = tk.StringVar(value=self.current_month)
        self.selector = ttk.Combobox(
            self.control_frame,
            textvariable=self.analysis_var,
            values=self.portuguese_months + ["Médias", "Totais"],
            state="readonly",
            width=15
        )
        self.selector.pack(side=tk.LEFT, padx=10)
        self.selector.bind("<<ComboboxSelected>>", self.change_analysis)

        # Treeview setup
        self.tree_frame = tk.Frame(self.root)
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.tree_scroll = ttk.Scrollbar(self.tree_frame)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Columns configuration
        self.columns_config = {
            "Monthly": [
                ("ID", 50), ("Item", 150), ("Unidades Produzidas", 120),
                ("Unidades Vendidas", 120), ("Unidades em Estoque", 120),
                ("Preço (R$)", 100), ("Custo (R$)", 100),
                ("Custo Mensal (R$)", 120), ("Vendas Mensais (R$)", 120),
                ("Lucro (R$)", 100)
            ],
            "Médias": [
                ("Item", 150), ("Média Unid. Produzidas", 140),
                ("Média Unid. Vendidas", 140), ("Preço Médio (R$)", 120),
                ("Custo Médio (R$)", 120), ("Lucro Médio (R$)", 120)
            ],
            "Totais": [
                ("Item", 150), ("Total Unid. Produzidas", 140),
                ("Total Unid. Vendidas", 140), ("Vendas Totais (R$)", 140),
                ("Custo Total (R$)", 140), ("Lucro Total (R$)", 140)
            ]
        }

        self.inventory_tree = ttk.Treeview(
            self.tree_frame,
            columns=[col[0] for col in self.columns_config["Monthly"]],
            show="headings",
            yscrollcommand=self.tree_scroll.set
        )
        self.tree_scroll.config(command=self.inventory_tree.yview)

        for col in self.columns_config["Monthly"]:
            self.inventory_tree.heading(col[0], text=col[0])
            self.inventory_tree.column(col[0], width=col[1], anchor=tk.CENTER)
        self.inventory_tree.pack(fill=tk.BOTH, expand=True)

        # Entry fields
        self.entry_frame = tk.Frame(self.root)
        self.entry_frame.pack(pady=15)

        self.labels = ["Item", "Unidades Produzidas", "Unidades Vendidas", "Preço (R$)", "Custo (R$)"]
        self.entries = {}
        for i, label in enumerate(self.labels):
            tk.Label(self.entry_frame, text=label).grid(row=0, column=i, padx=5)
            entry = tk.Entry(self.entry_frame, width=20)
            entry.grid(row=1, column=i, padx=5)
            self.entries[label] = entry

        # Buttons
        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(pady=10)
        buttons = [
            ("📁 Importar", self.import_excel),
            ("➕ Adicionar", self.add_item),
            ("🔄 Atualizar", self.update_item),
            ("🗑️ Remover", self.remove_item),
            ("💾 Salvar", self.save_all_months),
            ("📊 Gráficos", self.open_graphs_window)
        ]
        for i, (text, cmd) in enumerate(buttons):
            btn = tk.Button(self.button_frame, text=text, command=cmd)
            btn.grid(row=0, column=i, padx=5)

    def change_analysis(self, event=None):
        selection = self.analysis_var.get()
        if selection in self.portuguese_months:
            self.configure_columns("Monthly")
            self.load_month_data(selection)
        elif selection == "Médias":
            self.configure_columns("Médias")
            self.load_averages()
        elif selection == "Totais":
            self.configure_columns("Totais")
            self.load_totals()

    def configure_columns(self, view_type):
        """Configure treeview columns (Portuguese version)"""
        # Clear existing columns
        for col in self.inventory_tree["columns"]:
            self.inventory_tree.heading(col, text="")
            self.inventory_tree.column(col, width=0)

        # Set new columns
        new_columns = self.columns_config[view_type]
        self.inventory_tree["columns"] = [col[0] for col in new_columns]

        for col in new_columns:
            self.inventory_tree.heading(col[0], text=col[0])
            self.inventory_tree.column(col[0], width=col[1], anchor=tk.CENTER)

    def load_month_data(self, month):
        self.clear_tree()
        for idx, item in enumerate(self.monthly_data[month], start=1):
            item_name, units_produced, units_sold, price, cost = item
            units_stock = units_produced - units_sold
            monthly_cost = units_produced * cost
            monthly_sales = units_sold * price
            profit = monthly_sales - monthly_cost

            self.inventory_tree.insert("", "end", values=(
                idx, item_name, units_produced, units_sold,
                units_stock,
                f"{price:.2f}", f"{cost:.2f}",
                f"{monthly_cost:.2f}", f"{monthly_sales:.2f}",
                f"{profit:.2f}"
            ))

    def load_averages(self):
        self.clear_tree()
        product_data = {}
        grand_total = {
            "produced": 0, "sold": 0,
            "price": 0.0, "cost": 0.0,
            "profit": 0.0
        }
        total_entries = 0

        for month in self.portuguese_months:
            for item in self.monthly_data[month]:
                name = item[0]
                if name not in product_data:
                    product_data[name] = {
                        "produced": [], "sold": [],
                        "price": [], "cost": [],
                        "profit": []
                    }

                units_produced, units_sold, price, cost = item[1], item[2], item[3], item[4]
                profit = (units_sold * price) - (units_produced * cost)

                product_data[name]["produced"].append(units_produced)
                product_data[name]["sold"].append(units_sold)
                product_data[name]["price"].append(price)
                product_data[name]["cost"].append(cost)
                product_data[name]["profit"].append(profit)

                grand_total["produced"] += units_produced
                grand_total["sold"] += units_sold
                grand_total["price"] += price
                grand_total["cost"] += cost
                grand_total["profit"] += profit
                total_entries += 1

        for product, data in product_data.items():
            count = len(data["produced"])
            self.inventory_tree.insert("", "end", values=(
                product,
                f"{sum(data['produced']) / count:.1f}",
                f"{sum(data['sold']) / count:.1f}",
                f"{sum(data['price']) / count:.2f}",
                f"{sum(data['cost']) / count:.2f}",
                f"{sum(data['profit']) / count:.2f}"
            ))

        if total_entries > 0:
            self.inventory_tree.insert("", "end", tags=('grand_total',), values=(
                "Todos os Produtos",
                f"{grand_total['produced'] / total_entries:.1f}",
                f"{grand_total['sold'] / total_entries:.1f}",
                f"{grand_total['price'] / total_entries:.2f}",
                f"{grand_total['cost'] / total_entries:.2f}",
                f"{grand_total['profit'] / total_entries:.2f}"
            ))
            self.inventory_tree.tag_configure('grand_total', background='#e0e0e0')

    def load_totals(self):
        self.clear_tree()
        totals = {}
        grand_total = {"produced": 0, "sold": 0, "sales": 0.0, "cost": 0.0, "profit": 0.0}

        for month in self.portuguese_months:
            for item in self.monthly_data[month]:
                name = item[0]
                if name not in totals:
                    totals[name] = {
                        "produced": 0, "sold": 0,
                        "sales": 0.0, "cost": 0.0,
                        "profit": 0.0
                    }

                units_produced, units_sold, price, cost = item[1], item[2], item[3], item[4]
                monthly_cost = units_produced * cost
                monthly_sales = units_sold * price
                profit = monthly_sales - monthly_cost

                totals[name]["produced"] += units_produced
                totals[name]["sold"] += units_sold
                totals[name]["sales"] += monthly_sales
                totals[name]["cost"] += monthly_cost
                totals[name]["profit"] += profit

                grand_total["produced"] += units_produced
                grand_total["sold"] += units_sold
                grand_total["sales"] += monthly_sales
                grand_total["cost"] += monthly_cost
                grand_total["profit"] += profit

        for product, data in totals.items():
            self.inventory_tree.insert("", "end", values=(
                product,
                data["produced"],
                data["sold"],
                f"{data['sales']:.2f}",
                f"{data['cost']:.2f}",
                f"{data['profit']:.2f}"
            ))

        self.inventory_tree.insert("", "end", tags=('grand_total',), values=(
            "Todos os Produtos",
            grand_total["produced"],
            grand_total["sold"],
            f"{grand_total['sales']:.2f}",
            f"{grand_total['cost']:.2f}",
            f"{grand_total['profit']:.2f}"
        ))
        self.inventory_tree.tag_configure('grand_total', background='#e0e0e0')

    def import_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not file_path:
            return

        try:
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                if sheet in self.monthly_data.keys():
                    df = pd.read_excel(xls, sheet_name=sheet)
                    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                    self.monthly_data[sheet] = df[
                        ["Item", "Unidades Produzidas", "Unidades Vendidas",
                         "Preço (R$)", "Custo (R$)"]].values.tolist()

            self.change_analysis()
            messagebox.showinfo("Sucesso", "Dados importados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha na importação:\n{str(e)}")

    def add_item(self):
        try:
            item = self.entries["Item"].get().strip()
            units_produced = int(self.entries["Unidades Produzidas"].get())
            units_sold = int(self.entries["Unidades Vendidas"].get())
            price = float(self.entries["Preço (R$)"].get())
            cost = float(self.entries["Custo (R$)"].get())

            if units_sold > units_produced:
                raise ValueError("Unidades vendidas não podem exceder unidades produzidas")
            if not item:
                raise ValueError("Nome do item não pode estar vazio")
            if any(val < 0 for val in [units_produced, units_sold, price, cost]):
                raise ValueError("Valores devem ser números positivos")

            month = self.analysis_var.get() if self.analysis_var.get() in self.portuguese_months else self.current_month
            self.monthly_data[month].append([item, units_produced, units_sold, price, cost])
            self.load_month_data(month)
            self.clear_entries()
        except ValueError as e:
            messagebox.showerror("Erro de Entrada", str(e))

    def update_item(self):
        selected = self.inventory_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um item para atualizar")
            return

        try:
            item = self.entries["Item"].get().strip()
            units_produced = int(self.entries["Unidades Produzidas"].get())
            units_sold = int(self.entries["Unidades Vendidas"].get())
            price = float(self.entries["Preço (R$)"].get())
            cost = float(self.entries["Custo (R$)"].get())

            if not item:
                raise ValueError("Nome do item não pode estar vazio")
            if units_sold > units_produced:
                raise ValueError("Unidades vendidas não podem exceder unidades produzidas")
            if any(val < 0 for val in [units_produced, units_sold, price, cost]):
                raise ValueError("Valores devem ser números positivos")

            month = self.analysis_var.get() if self.analysis_var.get() in self.portuguese_months else self.current_month
            index = int(self.inventory_tree.item(selected, 'values')[0]) - 1
            self.monthly_data[month][index] = [item, units_produced, units_sold, price, cost]
            self.load_month_data(month)
            self.clear_entries()
        except ValueError as e:
            messagebox.showerror("Erro de Entrada", str(e))

    def remove_item(self):
        selected = self.inventory_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Selecione um item para remover")
            return

        month = self.analysis_var.get() if self.analysis_var.get() in self.portuguese_months else self.current_month
        index = int(self.inventory_tree.item(selected, 'values')[0]) - 1
        del self.monthly_data[month][index]
        self.load_month_data(month)

    def save_all_months(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not file_path:
            return

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                for month, data in self.monthly_data.items():
                    if data:
                        df = pd.DataFrame(data, columns=["Item", "Unidades Produzidas",
                                                         "Unidades Vendidas", "Preço (R$)",
                                                         "Custo (R$)"])
                        df.insert(0, "ID", range(1, len(df) + 1))
                        df["Unidades em Estoque"] = df["Unidades Produzidas"] - df["Unidades Vendidas"]
                        df["Custo Mensal (R$)"] = df["Unidades Produzidas"] * df["Custo (R$)"]
                        df["Vendas Mensais (R$)"] = df["Unidades Vendidas"] * df["Preço (R$)"]
                        df["Lucro (R$)"] = df["Vendas Mensais (R$)"] - df["Custo Mensal (R$)"]
                        df["Month"] = month
                        df.to_excel(writer, index=False, sheet_name=month)
            messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar:\n{str(e)}")

    def open_graphs_window(self):
        graph_window = tk.Toplevel(self.root)
        graph_window.title("Painel Analítico")
        graph_window.geometry("1200x800")

        control_frame = tk.Frame(graph_window)
        control_frame.pack(pady=10)

        ttk.Label(control_frame, text="Tipo de Gráfico:").pack(side=tk.LEFT, padx=5)
        self.graph_type = ttk.Combobox(control_frame, values=[
            "Produção vs Vendas",
            "Top 5 Produtos",
            "Distribuição de Vendas",
            "Tendência Anual",
            "Margens de Lucro",
            "Níveis de Estoque",
            "Custos vs Receitas",
            "Sazonalidade"
        ], width=25)
        self.graph_type.pack(side=tk.LEFT, padx=5)
        self.graph_type.set("Produção vs Vendas")

        generate_btn = ttk.Button(control_frame, text="Gerar Gráfico",
                                  command=lambda: self.generate_graph(graph_window))
        generate_btn.pack(side=tk.LEFT, padx=10)

        self.graph_canvas = tk.Canvas(graph_window)
        self.graph_canvas.pack(fill=tk.BOTH, expand=True)

    def generate_graph(self, window):
        for widget in window.winfo_children():
            if isinstance(widget, FigureCanvasTkAgg):
                widget.destroy()

        fig = plt.figure(figsize=(12, 6), dpi=100)
        selected_type = self.graph_type.get()

        try:
            if selected_type == "Tendência Anual":
                self.plot_yearly_trend(fig, "sales")
            elif selected_type == "Margens de Lucro":
                self.plot_profit_margins(fig)
            elif selected_type == "Níveis de Estoque":
                self.plot_stock_levels(fig)
            elif selected_type == "Custos vs Receitas":
                self.plot_cost_revenue(fig)
            elif selected_type == "Sazonalidade":
                self.plot_seasonality(fig)
            else:
                month = self.analysis_var.get() if self.analysis_var.get() in self.portuguese_months else self.current_month
                if selected_type == "Produção vs Vendas":
                    self.plot_production_vs_sales(fig, month)
                elif selected_type == "Top 5 Produtos":
                    self.plot_top_items(fig, month)
                elif selected_type == "Distribuição de Vendas":
                    self.plot_sales_distribution(fig, month)

            plt.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=window)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao gerar gráfico:\n{str(e)}")

    def plot_profit_margins(self, fig):
        ax = fig.add_subplot(111)
        margins = []
        products = []

        for month in self.portuguese_months:
            for item in self.monthly_data[month]:
                product = item[0]
                price = item[3]
                cost = item[4]
                if price > 0:
                    margin = ((price - cost) / price) * 100
                    margins.append(margin)
                    products.append(product)

        avg_margins = {}
        for product, margin in zip(products, margins):
            if product not in avg_margins:
                avg_margins[product] = []
            avg_margins[product].append(margin)

        avg_margins = {k: np.mean(v) for k, v in avg_margins.items()}
        products = list(avg_margins.keys())
        values = list(avg_margins.values())

        colors = ['#4CAF50' if v >= 30 else '#FFC107' if v >= 20 else '#F44336' for v in values]

        bars = ax.barh(products, values, color=colors)
        ax.set_title("Margem de Lucro por Produto (%)", fontsize=14)
        ax.set_xlabel("Margem Média (%)", fontsize=12)
        ax.bar_label(bars, fmt='%.1f%%', padding=3)
        ax.grid(axis='x', linestyle='--', alpha=0.7)

    def plot_stock_levels(self, fig):
        ax = fig.add_subplot(111)
        months = []
        stock_levels = []

        for month in self.portuguese_months:
            total_stock = 0
            for item in self.monthly_data[month]:
                total_stock += item[1] - item[2]
            months.append(month)
            stock_levels.append(total_stock)

        ax.fill_between(months, stock_levels, color='#2196F3', alpha=0.3)
        ax.plot(months, stock_levels, marker='o', color='#0D47A1', linewidth=2)
        ax.set_title("Níveis de Estoque Mensais", fontsize=14)
        ax.set_ylabel("Unidades em Estoque", fontsize=12)
        ax.grid(True, linestyle='--', alpha=0.6)
        plt.xticks(rotation=45)

    def plot_cost_revenue(self, fig):
        gs = GridSpec(2, 1, height_ratios=[3, 1])
        ax1 = fig.add_subplot(gs[0])
        ax2 = fig.add_subplot(gs[1])

        months = self.portuguese_months
        revenues = []
        costs = []
        profits = []

        for month in months:
            monthly_revenue = 0
            monthly_cost = 0
            for item in self.monthly_data[month]:
                monthly_revenue += item[2] * item[3]
                monthly_cost += item[1] * item[4]
            revenues.append(monthly_revenue)
            costs.append(monthly_cost)
            profits.append(monthly_revenue - monthly_cost)

        ax1.plot(months, revenues, marker='o', color='#4CAF50', label='Receitas')
        ax1.plot(months, costs, marker='o', color='#F44336', label='Custos')
        ax1.set_title("Custos vs Receitas Mensais", fontsize=14)
        ax1.legend()
        ax1.grid(True, linestyle='--', alpha=0.6)

        ax2.bar(months, profits, color='#FFC107')
        ax2.set_title("Lucro Líquido Mensal", fontsize=12)
        ax2.axhline(0, color='black', linewidth=0.8)
        plt.xticks(rotation=45)

    def plot_seasonality(self, fig):
        ax = fig.add_subplot(111)
        data_matrix = []
        products = list(set(item[0] for month in self.portuguese_months for item in self.monthly_data[month]))

        for product in products:
            product_sales = []
            for month in self.portuguese_months:
                monthly_sales = 0
                for item in self.monthly_data[month]:
                    if item[0] == product:
                        monthly_sales = item[2] * item[3]
                product_sales.append(monthly_sales)
            data_matrix.append(product_sales)

        sns.heatmap(data_matrix,
                    xticklabels=self.portuguese_months,
                    yticklabels=products,
                    cmap='YlGnBu',
                    ax=ax)
        ax.set_title("Sazonalidade de Vendas por Produto", fontsize=14)
        ax.set_xlabel("Mês")
        ax.set_ylabel("Produto")
        plt.xticks(rotation=45)

    def plot_yearly_trend(self, fig, trend_type):
        ax = fig.add_subplot(111)
        months = self.portuguese_months
        values = []

        for month in months:
            total = 0
            for item in self.monthly_data[month]:
                if trend_type == "sales":
                    total += item[2] * item[3]
                else:
                    total += item[1] * item[4]
            values.append(total)

        color = '#4CAF50' if trend_type == "sales" else '#FF5722'
        ax.plot(months, values, marker='o', color=color, linewidth=1.5)
        ax.set_title(f"Tendência Anual de {'Vendas' if trend_type == 'sales' else 'Custos'}", fontsize=14)
        ax.set_ylabel(f"Total {'Vendas' if trend_type == 'sales' else 'Custos'} (R$)", fontsize=12)
        ax.tick_params(axis='both', labelsize=10)
        ax.grid(True, linestyle='--', alpha=0.6)
        plt.xticks(rotation=45)

    def plot_production_vs_sales(self, fig, month):
        ax = fig.add_subplot(111)
        data = self.monthly_data[month]
        items = [item[0] for item in data]
        produced = [item[1] for item in data]
        sold = [item[2] for item in data]

        width = 0.35
        x = range(len(items))

        ax.bar(x, produced, width, label='Produzidas', color='#2196F3')
        ax.bar([p + width for p in x], sold, width, label='Vendidas', color='#FF9800')

        ax.set_title(f"Produção vs Vendas - {month}", fontsize=14)
        ax.set_ylabel("Quantidade", fontsize=12)
        ax.set_xticks([p + width / 2 for p in x])
        ax.set_xticklabels(items, rotation=45, ha='right', fontsize=10)
        ax.legend(fontsize=10)
        ax.tick_params(axis='y', labelsize=10)
        ax.grid(True, linestyle='--', alpha=0.6)

    def plot_top_items(self, fig, month):
        ax = fig.add_subplot(111)
        data = self.monthly_data[month]
        sales_data = [(item[0], item[2] * item[3]) for item in data]
        sorted_data = sorted(sales_data, key=lambda x: x[1], reverse=True)[:5]

        items = [x[0] for x in sorted_data]
        sales = [x[1] for x in sorted_data]

        ax.barh(items[::-1], sales[::-1], color='#9C27B0')
        ax.set_title(f"Top 5 Produtos - {month}", fontsize=14)
        ax.set_xlabel("Vendas (R$)", fontsize=12)
        ax.tick_params(axis='both', labelsize=10)
        ax.grid(True, linestyle='--', alpha=0.6)

    def plot_sales_distribution(self, fig, month):
        ax = fig.add_subplot(111)
        data = self.monthly_data[month]
        sales = [item[2] * item[3] for item in data]
        items = [item[0] for item in data]

        colors = plt.cm.tab20.colors
        ax.pie(sales, labels=items, autopct='%1.1f%%',
               startangle=90, colors=colors, textprops={'fontsize': 8})
        ax.set_title(f"Distribuição de Vendas - {month}", fontsize=14)

    def clear_entries(self):
        for entry in self.entries.values():
            entry.delete(0, tk.END)

    def clear_tree(self):
        for item in self.inventory_tree.get_children():
            self.inventory_tree.delete(item)


if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()