import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd

def timedelta_to_hhmmss(timedelta_obj):
    # Converte um objeto Timedelta para o formato HH:MM:SS
    total_seconds = int(timedelta_obj.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return '{:02}:{:02}:{:02}'.format(hours, minutes, seconds)

def calcular_horas_extras(row):
    horas_trabalhadas = row['Horas Trabalhadas']
    horas_extras = max(pd.Timedelta(hours=0), pd.to_datetime(horas_trabalhadas) - pd.to_datetime('08:50:00'))
    return horas_extras

def formatar_horas_extras(horas_extras):
    if pd.isnull(horas_extras):
        return '00:00:00'
    else:
        return str(horas_extras).split()[-1]

def process_file():
    # Abrir a janela de seleção de arquivo
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    
    if file_path:
        # Atualizar mensagem para informar o usuário sobre o progresso
        result_label.config(text="Processando arquivo...")

        # Ler o arquivo de texto
        df = pd.read_csv(file_path, delimiter='\t')

        # Criar uma nova coluna com o nome do colaborador
        df['Nome Colaborador'] = df['Nome']

        # Obter uma lista de todos os colaboradores únicos
        colaboradores_unicos = df['Nome'].unique()

        total_colaboradores = len(colaboradores_unicos)
        progresso = 0

        for colaborador in colaboradores_unicos:
            # Atualizar a barra de progresso
            progresso += 1
            progress_bar["value"] = (progresso / total_colaboradores) * 100
            root.update_idletasks()

            # Filtrar os registros do colaborador atual
            colaborador_df = df[df['Nome'] == colaborador]

            # Criar um dicionário para armazenar as batidas de horário de cada dia
            batidas_dict = {}

            # Iterar pelos registros do colaborador atual
            for index, row in colaborador_df.iterrows():
                data = row['Tempo'].split()[0]  # Extrair apenas a data
                hora = row['Tempo'].split()[1]  # Extrair apenas a hora

                if data in batidas_dict:
                    batidas_dict[data].append(hora)
                else:
                    batidas_dict[data] = [hora]

            # Obter o número máximo de batidas de horário em um único dia
            max_batidas = max(len(batidas) for batidas in batidas_dict.values())

            # Criar as colunas correspondentes às batidas de horário
            colunas = [f'hora{i+1}' for i in range(max_batidas)]

            # Converter o dicionário em um DataFrame pandas
            df_final = pd.DataFrame.from_dict(batidas_dict, orient='index', columns=colunas)

            # Resetar o índice para que a data se torne uma coluna
            df_final.reset_index(inplace=True)
            # Renomear a coluna de data
            df_final.rename(columns={'index': 'data'}, inplace=True)

            # Adicionar a coluna com o nome do colaborador em todas as linhas
            df_final['Nome Colaborador'] = colaborador

            # Reorganizar as colunas
            df_final = df_final[['Nome Colaborador'] + list(df_final.columns[:-1])]

            # Calcular o total de horas trabalhadas no formato HH:MM:SS
            horas_trabalhadas = []
            for i in range(max_batidas // 2):
                horas_trabalhadas.append(pd.to_datetime(df_final[f'hora{i*2+2}']) - pd.to_datetime(df_final[f'hora{i*2+1}']))
            horas_trabalhadas = sum(horas_trabalhadas, pd.Timedelta(0))
            df_final['Horas Trabalhadas'] = horas_trabalhadas.dt.total_seconds()
            df_final['Horas Trabalhadas'] = pd.to_datetime(df_final['Horas Trabalhadas'], unit='s').dt.strftime('%H:%M:%S')

            # Calcular horas extras
            df_final['Horas Extras'] = df_final.apply(calcular_horas_extras, axis=1)

            # Formatar horas extras sem a contagem de dias
            df_final['Horas Extras'] = df_final['Horas Extras'].apply(formatar_horas_extras)

            # Salvar os registros em um arquivo Excel
            file_name = f'batidas_{colaborador}.xlsx'
            df_final.to_excel(file_name, index=False)

            # Calcular a soma das horas trabalhadas e das horas extras
            df_final['Horas Trabalhadas Timedelta'] = pd.to_timedelta(df_final['Horas Trabalhadas'])
            df_final['Horas Extras Timedelta'] = pd.to_timedelta(df_final['Horas Extras'])

            total_horas_trabalhadas = df_final['Horas Trabalhadas Timedelta'].sum()
            total_horas_extras = df_final['Horas Extras Timedelta'].sum()

            # Converter Timedelta para o formato HH:MM:SS
            str_total_horas_trabalhadas = timedelta_to_hhmmss(total_horas_trabalhadas)
            str_total_horas_extras = timedelta_to_hhmmss(total_horas_extras)

            # Criar um novo DataFrame para as somas
            df_somas = pd.DataFrame({
                'Nome Colaborador': ['Total'],
                'Horas Trabalhadas': [str_total_horas_trabalhadas],
                'Horas Extras': [str_total_horas_extras]
            })

            # Remover as colunas de Timedelta usadas para cálculo das somas
            df_final.drop(columns=['Horas Trabalhadas Timedelta', 'Horas Extras Timedelta'], inplace=True)

            # Concatenar o DataFrame original com o DataFrame das somas
            df_final_com_somas = pd.concat([df_final, df_somas], ignore_index=True)

            # Salvar os registros em um arquivo Excel
            file_name = f'batidas_{colaborador}.xlsx'
            df_final_com_somas.to_excel(file_name, index=False)

        # Atualizar mensagem após o processamento ser concluído
        result_label.config(text="Arquivos Excel gerados com sucesso.")

# Criar a janela principal
root = tk.Tk()
root.title("Selecione o Arquivo de Entrada")

# Criar um botão para selecionar o arquivo
select_button = tk.Button(root, text="Selecionar Arquivo", command=process_file)
select_button.pack(pady=20)

# Adicionar uma barra de progresso
progress_bar = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
progress_bar.pack(pady=10)

# Label para exibir o resultado do processamento
result_label = tk.Label(root, text="")
result_label.pack()

# Iniciar o loop principal do Tkinter
root.mainloop()