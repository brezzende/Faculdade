import xlsxwriter

# Criando um novo arquivo Excel para cálculos combinatórios
file_path = "Calculos_Combinatorios_Corrigido.xlsx"
workbook = xlsxwriter.Workbook(file_path)
worksheet = workbook.add_worksheet("Cálculos")

# Definição de formatações
bold = workbook.add_format({"bold": True})

# Definição das áreas de cálculo corrigidas
sections = [
    ("Arranjo Simples", "A2", "A(n, k) = n! / (n-k)!", "B2", "B3", "B4", "FACT(B2)/FACT(B2-B3)"),
    ("Arranjo com Repetição", "A6", "A'(n, k) = n^k", "B6", "B7", "B8", "B6^B7"),
    ("Permutação Simples", "A10", "P(n) = n!", "B10", "", "B11", "FACT(B10)"),
    ("Combinação Simples", "A14", "C(n, k) = n! / (k!(n-k)!)", "B14", "B15", "B16", "FACT(B14)/(FACT(B15)*FACT(B14-B15))")
]

# Preenchendo a planilha com os cálculos corrigidos
for section in sections:
    title, label_cell, formula_text, input_n, input_k, result_cell, formula = section
    
    worksheet.write(label_cell, title, bold)  # Escreve o título da seção
    worksheet.write(label_cell[:-1] + str(int(label_cell[-1]) + 1), formula_text)  # Escreve a fórmula matemática
    
    # Definição de rótulos e entrada de valores em colunas separadas
    worksheet.write(input_n.replace("B", "A"), "n:", bold)
    worksheet.write(input_n, "")  # Célula para entrada do usuário
    
    if input_k:  # Se precisar de n e k
        worksheet.write(input_k.replace("B", "A"), "k:", bold)
        worksheet.write(input_k, "")  # Célula para entrada do usuário
    
    worksheet.write(result_cell.replace("B", "A"), "Resultado:", bold)  # Rótulo de resultado
    worksheet.write_formula(result_cell, formula)  # Fórmula corrigida apenas com referências numéricas

# Ajustando a largura das colunas
worksheet.set_column("A:A", 20)  # Largura da coluna de rótulos
worksheet.set_column("B:B", 15)  # Largura da coluna de valores

# Salvando e fechando a planilha
workbook.close()

print(f"Planilha criada: {file_path}")
