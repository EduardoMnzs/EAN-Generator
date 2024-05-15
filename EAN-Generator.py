import random
from openpyxl import Workbook

def gerar_numero_ean():
    prefixo = "789"  
    digitos_centrais = ''.join(random.choices('0123456789', k=9))
    soma_pares = sum(int(d) for d in digitos_centrais[::2])
    soma_impares = sum(int(d) for d in digitos_centrais[1::2])
    soma_total = soma_pares * 3 + soma_impares
    digito_verificador = (10 - (soma_total % 10)) % 10
    codigo_ean = prefixo + digitos_centrais + str(digito_verificador)
    return codigo_ean

# Criando uma instância do Workbook
wb = Workbook()

# Selecionando a planilha ativa (por padrão, é a primeira planilha)
ws = wb.active

# Adicionando os cabeçalhos
ws.append(["Código EAN"])

ean = int(input("Quantos códigos deseja gerar: "))

# Gerando 100 códigos e adicionando à planilha
for _ in range(ean):
    codigo = gerar_numero_ean()
    ws.append([codigo])

# Salvando o arquivo Excel
wb.save(f"codigos_ean{ean}.xlsx")

print(f"Códigos EAN gerados e salvos com sucesso em 'codigos_ean{ean}.xlsx'.")
