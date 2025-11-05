import os
import sys
import pandas as pd
import tkinter as tk
import customtkinter as ctk
from datetime import datetime
from PIL import Image, ImageTk
from tkinter import ttk, filedialog, messagebox, PhotoImage

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class ProgramaTesouraria:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Acerto em Atraso")
        self.window.geometry("650x450")
        self.font_subtitle = ctk.CTkFont(family="Roboto", size=15, weight="bold")

        self.style = ttk.Style()
        self.style.configure('Header.TLabel', font=('Arial', 12, 'bold'))

        # Configuração do Ícone para aplicação
        try:
            icon_path = resource_path('Imagens/relatorio.png')
            if os.path.exists(icon_path):
                icon_image = PhotoImage(file=icon_path)
                self.window.iconphoto(True, icon_image)
                self.icon_image = icon_image
            else:
                print(f"Arquivo de ícone não encontrado em: {icon_path}")
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")

        logo_path = resource_path('Imagens/logo_arcom.jpeg')
        if os.path.exists(logo_path):
            logo_img = Image.open(logo_path)
            basewidth = 40
            wpercent = (basewidth/float(logo_img.size[0]))
            hsize = int((float(logo_img.size[1])*float(wpercent)))
            logo_img = logo_img.resize((basewidth, hsize), Image.Resampling.LANCZOS)
            logo_photo = ImageTk.PhotoImage(logo_img)
            
            logo_label = ttk.Label(self.window, image=logo_photo, background="#ffffff")
            logo_label.image = logo_photo
            logo_label.pack(pady=(10, 5))
            
            title_label = ttk.Label(self.window, text="Análise de Acertos", 
                                    style='Header.TLabel')
            title_label.pack(pady=(0, 10))
        
        # Frame Principal
        main_frame = ttk.Frame(self.window)
        main_frame.pack(padx=0.1, pady=0.1, fill='both', expand=True)
        
        # Configuração de entrada e saída de dados
        config_frame = ttk.LabelFrame(main_frame, text="Defina o caminho de entrada e saída:")
        config_frame.pack(fill='x', padx=1, pady=1)

        # Seleção do arquivo de entrada
        ttk.Label(config_frame, text="Importar Arquivo:").grid(row=0, column=0, padx=5, pady=5)
        self.input_path_var = tk.StringVar()
        self.input_path_var.set(os.path.join(self.get_base_path(), 'DADOS.xls'))
        ttk.Entry(config_frame, textvariable=self.input_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(config_frame, text="Procurar", command=self.choose_input_file).grid(row=0, column=2, padx=5, pady=5)

        # Seleção de pasta de saída
        ttk.Label(config_frame, text="Exportar Arquivos:").grid(row=1, column=0, padx=5, pady=5)
        self.output_path_var = tk.StringVar()
        self.output_path_var.set(os.path.join(self.get_base_path(), 'Arquivos-Analise'))
        ttk.Entry(config_frame, textvariable=self.output_path_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(config_frame, text="Escolher", command=self.choose_output_folder).grid(row=1, column=2, padx=5, pady=5)
        
        # Frame para filtros
        filters_frame = ttk.LabelFrame(main_frame, text="Filtros:")
        filters_frame.pack(fill='x', padx=2, pady=2)
        
        # Criar frame para lista e scrollbar
        list_container = ttk.Frame(filters_frame)
        list_container.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        
        # Lista de filtros predefinidos com scrollbar
        self.filters_listbox = tk.Listbox(list_container, width=80)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.filters_listbox.yview)
        self.filters_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.filters_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Botões para gerenciar filtros
        buttons_frame = ttk.Frame(filters_frame)
        buttons_frame.pack(side='left', padx=5, pady=5)
        ttk.Button(buttons_frame, text="Editar Filtro", command=self.edit_filter).pack(pady=2)

        # Botões de ação
        action_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        action_frame.pack(fill='x', padx=10, pady=10)
        
        button_container = ctk.CTkFrame(action_frame, fg_color="transparent")
        button_container.pack() 

        ctk.CTkButton(button_container, text="Processar Arquivo", command=self.load_file, 
                      fg_color="#388E3C", hover_color="#4CAF50", corner_radius=5, font=self.font_subtitle).pack(side='left', padx=5, ipady=5)
        
        self.execute_button = ctk.CTkButton(button_container, text="Executar Análise", command=self.execute_analysis, state='disabled',
                                            fg_color="#1976D2", hover_color="#2196F3", corner_radius=5, font=self.font_subtitle)
        self.execute_button.pack(side='left', padx=5, ipady=5)

        # Variável para armazenar o DataFrame
        self.df = None

        # Inicializar filtros predefinidos
        self.initialize_filters()

    def get_base_path(self):
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))
        
    def choose_input_file(self):
        file_types = [
            ('Arquivos Excel', '*.xls;*.xlsx;*.xlsm;*.xlsb')
        ]
        file_path = filedialog.askopenfilename(
            initialdir=os.path.dirname(self.input_path_var.get()),
            title="Selecione o arquivo de dados",
            filetypes=file_types
        )
        if file_path:
            self.input_path_var.set(file_path)

    def choose_output_folder(self):
        folder = filedialog.askdirectory(initialdir=self.output_path_var.get())
        if folder:
            self.output_path_var.set(folder)
        
    def initialize_filters(self):
        drop_columns = ['nome', 'acerto_cda', 'veiculo_traslado', 'nome_acertador', 'func_filial', 'data_entrada_matriz', 'tempo_acertador', 'data_liberacao', 'veiculo', 'impressao']
        self.filters = [
            {
                'name': 'TRANSPORTADORA',
                'tipo_acerto': 'M',
                'atraso': 30,
                'situacao': 1,
                'tipo_viagem': 4,
                'gerente': None,
                'local_tempo': None,
                'drop_columns': drop_columns,
            },
            {
                'name': 'DEVOLUCAO',
                'tipo_acerto': 'M',
                'atraso': 30,
                'situacao': 3,
                'tipo_viagem': 3,
                'gerente': None,
                'local_tempo': ['Devolucao', 'Dev.Udi'],
                'drop_columns': drop_columns,
            },
            {
                'name': 'ACERTO_DE_CAIXA',
                'tipo_acerto': 'C',
                'atraso': 1,
                'situacao': 3,
                'tipo_viagem': 3,
                'gerente': None,
                'local_tempo': None,
                'drop_columns': drop_columns,
            }
        ]
        self.update_filters_list()
        
    def update_filters_list(self):
        self.filters_listbox.delete(0, tk.END)
        for f in self.filters:
            self.filters_listbox.insert(tk.END, f"{f['name']} (Tipo: {f['tipo_acerto']}, Atraso: {f['atraso']})")
            
            
    def edit_filter(self):
        if not self.filters_listbox.curselection():
            messagebox.showwarning("Aviso", "Por favor, selecione um filtro para editar.")
            return
            
        idx = self.filters_listbox.curselection()[0]
        dialog = FilterDialog(self.window, "Editar Filtro", self.filters[idx])
        if dialog.result:
            self.filters[idx] = dialog.result
            self.update_filters_list()
    
    def load_file(self):
        input_path = self.input_path_var.get()
        if not os.path.exists(input_path):
            messagebox.showerror("Erro", "O arquivo de dados não foi encontrado. Por favor, selecione um arquivo válido.")
            return
            
        try:
            # Ler arquivo Excel
            self.df = pd.read_excel(input_path)
            
            # Para obter lista única de gerentes
            if 'gerente_transporte' not in self.df.columns:
                messagebox.showerror("Erro", "Coluna 'gerente_transporte' não encontrada no arquivo.")
                return
            
            # Tratar valores nulos e converter para string
            self.df['gerente_transporte'] = self.df['gerente_transporte'].fillna('Não Informado')
            self.df['gerente_transporte'] = self.df['gerente_transporte'].astype(str)
            
            # Filtrar e ordenar gerentes únicos, removendo strings vazias
            gerentes = sorted([g for g in self.df['gerente_transporte'].unique().tolist() 
                              if g and g.strip() and g != 'Não Informado'])
            
            self.update_gerentes_filters(gerentes)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar arquivo:\n{str(e)}")
            self.df = None

    def update_gerentes_filters(self, gerentes):
        """Atualiza a lista de filtros com os gerentes selecionados"""       
        self.filters = [f for f in self.filters if not f.get('gerente')]    
        # Adiciona Filtros baseado na lista de gerentes do arquivo.xls
        for gerente in gerentes:
            self.filters.append({
                'name': gerente,
                'tipo_acerto': 'M',
                'atraso': 15,
                'situacao': 1,
                'tipo_viagem': 4,
                'gerente': gerente,
                'local_tempo': None,
                'drop_columns': ['nome', 'acerto_cda', 'veiculo_traslado', 'nome_acertador', 'func_filial', 'data_entrada_matriz', 'tempo_acertador', 'data_liberacao', 'veiculo', 'impressao']
            })

        self.update_filters_list()
        # Habilitar botão de execução
        self.execute_button.configure(state='normal')

    def execute_analysis(self):
        if self.df is None:
            messagebox.showerror("Erro", "Por favor, carregue um arquivo de dados primeiro.")
            return
        
        try:
            
            # Criar pasta de saída
            data_atual = datetime.today().strftime('%Y-%m-%d')
            output_folder = os.path.join(self.output_path_var.get(), data_atual)
            
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
                
            # Processar cada filtro
            for filtro in self.filters:
                
                # Converter e validar campos numéricos
                df_process = self.df.copy()
                df_process['atraso'] = pd.to_numeric(df_process['atraso'], errors='coerce')
                df_process['situacao'] = pd.to_numeric(df_process['situacao'], errors='coerce')
                df_process['tipo_viagem'] = pd.to_numeric(df_process['tipo_viagem'], errors='coerce')
                
                # Aplicar filtros com máscaras booleanas para maior segurança
                mask = (
                    (df_process['tipo_acerto'].astype(str) == filtro["tipo_acerto"]) &
                    (df_process['atraso'] >= filtro["atraso"]) &
                    (df_process['situacao'] == filtro["situacao"]) &
                    (df_process['tipo_viagem'] == filtro["tipo_viagem"])
                )
                
                df_result = df_process[mask]
                
                if filtro['gerente']:
                    df_result = df_result[df_result['gerente_transporte'] == filtro["gerente"]]
                if filtro['local_tempo']:
                    df_result = df_result[df_result['local_tempo'].isin(filtro["local_tempo"])]
                    
                # Ordenar por atraso
                df_result = df_result.sort_values(by='atraso', ascending=False)
                
                # Remover colunas
                drop_columns = [col for col in filtro['drop_columns'] if col in df_result.columns]
                df_result.drop(drop_columns, axis=1, inplace=True)
                
                # Formatar colunas - apenas capitalizar os nomes das colunas
                df_result.columns = [col.title() for col in df_result.columns]
                        
                # Salvar arquivo
                output_file = os.path.join(output_folder, f'{filtro["name"]}.xlsx')
                df_result.to_excel(output_file, index=False, engine='openpyxl')

            messagebox.showinfo("Sucesso", "Processamento concluído! Verifique a pasta de saída.")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento:\n{str(e)}", 'error')


class FilterDialog:
    def __init__(self, parent, title, initial_values=None):
        self.result = None
        
        # Criar janela de diálogo
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.geometry("250x250")
        
        # Campos do Editar Filtro        
        ttk.Label(dialog, text="Nome do Filtro:").pack(pady=2)
        name_var = tk.StringVar(value=initial_values['name'] if initial_values else "")
        ttk.Entry(dialog, textvariable=name_var).pack(pady=2)
        
        ttk.Label(dialog, text="Tipo de Acerto (M/C):").pack(pady=2)
        tipo_acerto_var = tk.StringVar(value=initial_values['tipo_acerto'] if initial_values else "M")
        ttk.Entry(dialog, textvariable=tipo_acerto_var).pack(pady=2)
        
        ttk.Label(dialog, text="Atraso Mínimo (dias):").pack(pady=2)
        atraso_var = tk.StringVar(value=str(initial_values['atraso']) if initial_values else "0")
        ttk.Entry(dialog, textvariable=atraso_var).pack(pady=2)
        
        ttk.Label(dialog, text="Situação (0-3):").pack(pady=2)
        situacao_var = tk.StringVar(value=str(initial_values['situacao']) if initial_values else "0")
        ttk.Entry(dialog, textvariable=situacao_var).pack(pady=2)
        
        ttk.Label(dialog, text="Tipo de Viagem (0-4):").pack(pady=2)
        tipo_viagem_var = tk.StringVar(value=str(initial_values['tipo_viagem']) if initial_values else "0")
        ttk.Entry(dialog, textvariable=tipo_viagem_var).pack(pady=2)

        def save():
            self.result = {
                'name': name_var.get(),
                'tipo_acerto': tipo_acerto_var.get(),
                'atraso': int(atraso_var.get()),
                'situacao': int(situacao_var.get()),
                'tipo_viagem': int(tipo_viagem_var.get()),
                'local_tempo': None,
                'gerente': initial_values.get('gerente'),
                'drop_columns': ['nome', 'acerto_cda', 'veiculo_traslado', 'nome_acertador', 'func_filial', 
                               'data_entrada_matriz', 'tempo_acertador', 'data_liberacao', 'veiculo', 'impressao']
            }
            dialog.destroy()

        ttk.Button(dialog, text="Salvar", command=save).pack(pady=10)
        
        dialog.transient(parent)
        dialog.grab_set()
        parent.wait_window(dialog)


if __name__ == "__main__":
    app = ProgramaTesouraria()
    app.window.mainloop()