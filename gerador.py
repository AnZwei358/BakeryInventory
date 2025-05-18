import pandas as pd
import random

# Itens e meses em Português
items = [
    "Pão Francês", "Bolo de Chocolate", "Brigadeiro",
    "Pão de Queijo", "Coxinha", "Quindim",
    "Pastel de Carne", "Empada de Frango", "Sonho",
    "Torta de Nozes", "Pão Doce", "Biscoito de Polvilho",
    "Churros", "Bolo de Cenoura", "Pão de Mel"
]

meses = [
    'Janeiro', 'Fevereiro', 'Março', 'Abril',
    'Maio', 'Junho', 'Julho', 'Agosto',
    'Setembro', 'Outubro', 'Novembro', 'Dezembro'
]

with pd.ExcelWriter('estoque_aleatorio.xlsx', engine='openpyxl') as writer:
    for mes in meses:
        dados_mes = []
        for item in items:
            produzidas = random.randint(100, 2000)
            vendidas = random.randint(50, produzidas - 1)
            preco = round(random.uniform(1.0, 50.0), 2)
            custo = round(random.uniform(0.5, preco * 0.8), 2)

            dados_mes.append({
                "Item": item,
                "Unidades Produzidas": produzidas,
                "Unidades Vendidas": vendidas,
                "Preço (R$)": preco,
                "Custo (R$)": custo
            })

        df = pd.DataFrame(dados_mes)
        df.to_excel(writer, sheet_name=mes, index=False)

print("Planilha gerada com sucesso!")