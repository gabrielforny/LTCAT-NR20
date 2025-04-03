import traceback
import os
import pythoncom
import time
import win32com.client as win32
import win32com.client

USERNAME = os.getenv("USERNAME")


def get_first_word_doc(folder_path):
    # List all files in the directory
    files = os.listdir(folder_path)
    
    # Filter for Word documents (.rtf)
    word_files = [f for f in files if f.endswith(('.rtf', '.docx'))]
    
    if not word_files:
        return None
    
    # Return the first Word document found
    return os.path.join(folder_path, word_files[0])
    
def limpar_tabela(tabela):
    """Limpa o conteúdo de todas as células da tabela, mantendo o cabeçalho"""
    try:
        # Obtém o número total de linhas
        num_rows = tabela.Rows.Count
        
        # Começa da segunda linha (após o cabeçalho)
        for i in range(2, num_rows + 1):
            row = tabela.Rows.Item(i)
            for j in range(1, row.Cells.Count + 1):
                # Limpa o conteúdo da célula
                cell = row.Cells.Item(j)
                cell.Range.Text = ""
                # Centraliza o conteúdo horizontal e verticalmente
                cell.Range.ParagraphFormat.Alignment = 1  # 1 = Centralizado
                cell.VerticalAlignment = 1  # 1 = Centralizado
        
        print("Tabela limpa com sucesso")
        return True
    except Exception as e:
        print(f"Erro ao limpar tabela: {str(e)}")
        return False

def formatar_celula(cell):
    """Aplica formatação padrão à célula"""
    try:
        # Centraliza horizontalmente
        cell.Range.ParagraphFormat.Alignment = 0  # 1 = Centralizado
        # Centraliza verticalmente
        cell.VerticalAlignment = 1  # 1 = Centralizado
        # Define a fonte como Verdana
        cell.Range.Font.Name = "Verdana"
        # Define o tamanho da fonte
        cell.Range.Font.Size = 8
        # Adiciona negrito
        cell.Range.Font.Bold = False
    except Exception as e:
        print(f"Erro ao formatar célula: {str(e)}")

def preencher_dados_tabelas_funcao(first_word_doc_in, first_word_doc_out):
    try:
        # Inicializar COM
        pythoncom.CoInitialize()
        
        try:
            # Inicializar Word
            word = win32.Dispatch("Word.Application")
            word.Visible = False
        except:
            word = win32com.client.gencache.EnsureDispatch('Word.Application')
            word.Visible = False
        
        print(f"Abrindo documento de entrada: {first_word_doc_in}")
        print(f"Abrindo documento de saída: {first_word_doc_out}")
        
        # Abrir documentos
        doc_input = word.Documents.Open(first_word_doc_in)
        doc_output = word.Documents.Open(first_word_doc_out)
        
        # Dicionário para armazenar os cargos e suas informações
        dados_cargos = {}
        cargo_atual = None
        setor_atual = None
        encontrou_cargo = False
        
        print(f"Número total de tabelas no documento de entrada: {doc_input.Tables.Count}")
        
        # Procurar nas tabelas do documento de entrada
        i = 1
        while i <= doc_input.Tables.Count:
            try:
                table = doc_input.Tables.Item(i)
                print(f"\nAnalisando tabela {i} de {doc_input.Tables.Count}")
                
                # Primeiro, procurar por "Cargo:"
                cargo_encontrado = False
                for row_idx in range(1, table.Rows.Count + 1):
                    try:
                        row = table.Rows.Item(row_idx)
                        for cell_idx in range(1, row.Cells.Count + 1):
                            try:
                                cell = row.Cells.Item(cell_idx)
                                cell_text = cell.Range.Text.strip().replace('\r', '').replace('\n', ' ').replace('\x07', '')
                                
                                # Procurar por células com "Setor:"
                                if "Setor:" in cell_text:
                                    setor_atual = cell_text.replace("Setor:", "").strip().upper()
                                    print(f"Setor encontrado: {setor_atual}")
                                
                                # Procurar por células com "Cargo:"
                                if "Cargo:" in cell_text:
                                    cargo_atual = cell_text.replace("Cargo:", "").strip()
                                    print(f"Cargo encontrado: {cargo_atual}")
                                    cargo_encontrado = True
                                    nome_formatado = f"{setor_atual} / {cargo_atual}" if setor_atual else cargo_atual
                                    dados_cargos[nome_formatado] = {}
                                    break
                            except Exception as cell_error:
                                continue
                        if cargo_encontrado:
                            break
                    except Exception:
                        continue
                
                # Se encontrou cargo, procurar a próxima tabela com "Agente"
                if cargo_encontrado:
                    i += 1  # Avançar para próxima tabela
                    while i <= doc_input.Tables.Count:
                        try:
                            next_table = doc_input.Tables.Item(i)
                            found_agente = False
                            insalubridade_found = False
                            periculosidade_found = False
                            aposentadoria_found = False
                            
                            # Verificar se esta tabela contém "Agente"
                            for row_idx in range(1, next_table.Rows.Count + 1):
                                try:
                                    row = next_table.Rows.Item(row_idx)
                                    for cell_idx in range(1, row.Cells.Count + 1):
                                        try:
                                            cell = row.Cells.Item(cell_idx)
                                            cell_text = cell.Range.Text.strip().replace('\r', '').replace('\n', ' ').replace('\x07', '')
                                            
                                            if "Agente" in cell_text:
                                                found_agente = True
                                            
                                            # Se encontrou a tabela com Agente, procurar os valores
                                            if found_agente:
                                                if "Insalubridade" in cell_text and not insalubridade_found:
                                                    # Pegar o valor da próxima célula se existir
                                                    if cell_idx + 1 <= row.Cells.Count:
                                                        valor = row.Cells.Item(cell_idx + 1).Range.Text.strip().replace('\r', '').replace('\n', ' ').replace('\x07', '')
                                                        dados_cargos[nome_formatado]['insalubridade'] = valor
                                                        insalubridade_found = True
                                                        print(f"Insalubridade encontrada para {nome_formatado}: {valor}")
                                                
                                                if "Periculosidade" in cell_text and not periculosidade_found:
                                                    # Pegar o valor da próxima célula se existir
                                                    if cell_idx + 1 <= row.Cells.Count:
                                                        valor = row.Cells.Item(cell_idx + 1).Range.Text.strip().replace('\r', '').replace('\n', ' ').replace('\x07', '')
                                                        dados_cargos[nome_formatado]['periculosidade'] = valor
                                                        periculosidade_found = True
                                                        print(f"Periculosidade encontrada para {nome_formatado}: {valor}")
                                                
                                                if "Aposentadoria Especial" in cell_text and not aposentadoria_found:
                                                    # Pegar o valor da próxima célula se existir
                                                    if cell_idx + 1 <= row.Cells.Count:
                                                        valor = row.Cells.Item(cell_idx + 1).Range.Text.strip().replace('\r', '').replace('\n', ' ').replace('\x07', '')
                                                        dados_cargos[nome_formatado]['aposentadoria_especial'] = valor
                                                        aposentadoria_found = True
                                                        print(f"Aposentadoria Especial encontrada para {nome_formatado}: {valor}")
                                        except Exception:
                                            continue
                                except Exception:
                                    continue
                            
                            if found_agente:
                                break  # Encontrou e processou a tabela com Agente, pode parar
                            i += 1  # Se não encontrou "Agente", continua para próxima tabela
                        except Exception:
                            i += 1
                            continue
                else:
                    i += 1  # Se não encontrou cargo, continua para próxima tabela
            except Exception as table_error:
                print(f"Erro ao processar tabela {i}: {str(table_error)}")
                i += 1
                continue
        
        print(f"\nDados coletados de todos os cargos: {dados_cargos}")
        
        # Procurar a tabela de destino no documento de saída
        print(f"\nProcurando tabela de destino no documento de saída...")
        print(f"Número total de tabelas no documento de saída: {doc_output.Tables.Count}")
        
        tabela_destino = None
        for i in range(1, doc_output.Tables.Count + 1):
            try:
                table = doc_output.Tables.Item(i)
                print(f"\nVerificando tabela {i} de {doc_output.Tables.Count}")
                
                # Verificar se é a tabela correta
                primeira_celula = table.Cell(1, 1).Range.Text.strip()
                print(f"Conteúdo da primeira célula: {primeira_celula}")
                
                if "FUNÇÃO/CARGO" in primeira_celula:
                    tabela_destino = table
                    print("Tabela de destino encontrada!")
                    break
                    
            except Exception as table_error:
                print(f"Erro ao verificar tabela {i}: {str(table_error)}")
                continue
        
        if tabela_destino:
            print("\nLimpando a tabela de destino...")
            if not limpar_tabela(tabela_destino):
                print("Erro ao limpar a tabela de destino!")
                return False
            
            print("\nAjustando o número de linhas da tabela...")
            num_cargos = len(dados_cargos)
            num_linhas_atual = tabela_destino.Rows.Count
            
            # Ajustar o número de linhas (mantendo a primeira linha do cabeçalho)
            if num_linhas_atual < num_cargos + 1:  # +1 para o cabeçalho
                # Adicionar linhas se necessário
                for _ in range(num_cargos + 1 - num_linhas_atual):
                    tabela_destino.Rows.Add()
            elif num_linhas_atual > num_cargos + 1:
                # Remover linhas extras
                for _ in range(num_linhas_atual - (num_cargos + 1)):
                    tabela_destino.Rows.Item(tabela_destino.Rows.Count).Delete()
            
            print(f"Tabela ajustada para {num_cargos + 1} linhas (incluindo cabeçalho)")
            
            # Formatar o cabeçalho
            header_row = tabela_destino.Rows.Item(1)
            for j in range(1, header_row.Cells.Count + 1):
                cell = header_row.Cells.Item(j)
                formatar_celula(cell)
            
            # Preencher os dados
            idx = 2
            for cargo_formatado, dados in dados_cargos.items():
                try:
                    row = tabela_destino.Rows.Item(idx)
                    
                    # Preencher cada linha com os valores
                    try:
                        # Preencher o nome do cargo (coluna 1)
                        cell = row.Cells.Item(1)
                        cell.Range.Text = cargo_formatado
                        formatar_celula(cell)
                        
                        # Preencher Insalubridade (coluna 2)
                        cell = row.Cells.Item(2)
                        cell.Range.Text = dados['insalubridade']
                        formatar_celula(cell)
                        
                        # Preencher Periculosidade (coluna 3)
                        cell = row.Cells.Item(3)
                        cell.Range.Text = dados['periculosidade']
                        formatar_celula(cell)
                        
                        # Preencher Aposentadoria Especial (coluna 4)
                        cell = row.Cells.Item(4)
                        cell.Range.Text = dados['aposentadoria_especial']
                        formatar_celula(cell)
                        
                    except Exception as cell_error:
                        print(f"Erro ao preencher células da linha {idx}: {str(cell_error)}")
                        continue
                        
                except Exception as row_error:
                    print(f"Erro ao processar linha {idx}: {str(row_error)}")
                    continue
                    
                idx += 1
            
            print("\nSalvando alterações...")
            doc_output.Save()
            time.sleep(2)  # Aguardar um pouco para garantir que o salvamento seja concluído
            
        else:
            print("ERRO: Tabela de destino não encontrada!")
        
        print("\nFechando documentos...")
        doc_input.Close()
        doc_output.Close()
        word.Quit()
        
        pythoncom.CoUninitialize()
        return True
        
    except Exception as e:
        print(f"Erro ao preencher dados das tabelas: {str(e)}")
        traceback.print_exc()
        
        # Garantir que os documentos sejam fechados em caso de erro
        try:
            doc_input.Close()
        except:
            pass
        try:
            doc_output.Close()
        except:
            pass
        try:
            word.Quit()
        except:
            pass
        
        try:
            pythoncom.CoUninitialize()
        except:
            pass
            
        return False
