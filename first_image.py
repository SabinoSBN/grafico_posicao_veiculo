import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.drawing.image import Image

# Caminho do arquivo Excel
file_path = 'test_1.xlsx'  # Atualize com o caminho correto

# Carregar o arquivo Excel
df = pd.read_excel(file_path)

# Transformar o formato da hora para string para facilitar a plotagem
df['hors'] = pd.to_datetime(df['hors']).dt.strftime('%H:%M:%S')

# Mapeamento do estado para valores numéricos
df['stato machine num'] = df['stato machine'].map({'ON': 1, 'OFF': 0})

# Criar um gráfico de linhas para o estado da máquina e a área
fig, ax1 = plt.subplots(figsize=(12, 8))

# Plotar o estado da máquina
color = 'tab:orange'
ax1.set_xlabel('Hora')
ax1.set_ylabel('Stato Machine', color=color)
ax1.plot(df['hors'], df['stato machine num'], label='Stato Machine', color=color, marker='o')
ax1.tick_params(axis='y', labelcolor=color)
ax1.set_ylim(-0.1, 1.1)
ax1.set_yticks([0, 1])
ax1.set_yticklabels(['OFF', 'ON'])

# Plotar a área em um segundo eixo y
ax2 = ax1.twinx()
color = 'tab:blue'
ax2.set_ylabel('Área', color=color)
ax2.plot(df['hors'], df['area'], label='Área', color=color, marker='x')
ax2.tick_params(axis='y', labelcolor=color)
ax2.set_yticks(range(df['area'].min(), df['area'].max() + 1))  # Definir valores inteiros para área

# Melhorar a nitidez dos horários
fig.autofmt_xdate()

# Adicionar título e legenda
fig.suptitle('Estado da Máquina e Área ao Longo do Tempo')
fig.tight_layout()
fig.legend(loc='upper left', bbox_to_anchor=(0.1, 0.9))

plt.xticks(rotation=45)

# Salvar o gráfico como imagem
image_path = 'stato_machine_and_area_chart.png'  # Atualize com o caminho correto
fig.savefig(image_path, dpi=300)  # Aumentar a resolução da imagem

# Carregar o arquivo Excel e inserir o gráfico
wb = openpyxl.load_workbook(file_path)
ws = wb.active

# Inserir o gráfico de estado da máquina e área na célula A1
img = Image(image_path)
ws.add_image(img, 'A1')

# Salvar o arquivo Excel atualizado
updated_file_path = 'browser_1_updated_with_chart.xlsx'
wb.save(updated_file_path)

print(f"Gráfico salvo em: {image_path}")
print(f"Arquivo Excel atualizado salvo em: {updated_file_path}")
