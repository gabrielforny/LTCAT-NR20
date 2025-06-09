import win32com.client as win32
import traceback
import os
import pythoncom
import time
import re

USERNAME = os.getenv("USERNAME")


# def get_first_word_doc(folder_path):
#     # List all files in the directory
#     files = os.listdir(folder_path)
    
#     # Filter for Word documents (.rtf)
#     word_files = [f for f in files if f.endswith(('.rtf', '.docx'))]
    
#     if not word_files:
#         return None
    
#     # Return the first Word document found
#     return os.path.join(folder_path, word_files[0])


# input = fr"C:\Users\Gabriel\tecnico\PGR - GRO\FORMATAÇÃO\LTCAT NR 20"
# output = fr"C:\Users\Gabriel\tecnico\PGR - GRO\FORMATAÇÃO\TEMPLATE\LTACT NR 20"

# Get the first Word document
# first_word_doc_in = get_first_word_doc(input)
# first_word_doc_out = get_first_word_doc(output)

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
        pythoncom.CoInitialize()
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        print(f"Abrindo documento de entrada: {first_word_doc_in}")
        print(f"Abrindo documento de saída: {first_word_doc_out}")
        doc_input = word.Documents.Open(first_word_doc_in)
        doc_output = word.Documents.Open(first_word_doc_out)
        dados_cargos = {}
        cargo_atual = None
        setor_atual = None
        nome_formatado = None
        print(f"Número total de tabelas no documento de entrada: {doc_input.Tables.Count}")
        i = 1
        while i <= doc_input.Tables.Count:
            try:
                table = doc_input.Tables.Item(i)
                print(f"\nAnalisando tabela {i} de {doc_input.Tables.Count}")
                cargo_encontrado = False
                for row_idx in range(1, table.Rows.Count + 1):
                    try:
                        row = table.Rows.Item(row_idx)
                        for cell_idx in range(1, row.Cells.Count + 1):
                            try:
                                cell = row.Cells.Item(cell_idx)
                                cell_text = cell.Range.Text.strip().replace('\r', '').replace('\n', ' ').replace('\x07', '')
                                if "Setor:" in cell_text:
                                    setor_atual = cell_text.replace("Setor:", "").strip().upper()
                                    print(f"Setor encontrado: {setor_atual}")
                                if "Cargo:" in cell_text:
                                    cargo_atual = cell_text.replace("Cargo:", "").strip()
                                    print(f"Cargo encontrado: {cargo_atual}")
                                    cargo_encontrado = True
                                    nome_formatado = f"{setor_atual} / {cargo_atual}" if setor_atual else cargo_atual
                                    dados_cargos[nome_formatado] = {
                                        'insalubridade': [],
                                        'periculosidade': [],
                                        'aposentadoria_especial': []
                                    }
                                    break
                            except Exception:
                                continue
                        if cargo_encontrado:
                            break
                    except Exception:
                        continue
                if cargo_encontrado:
                    i += 1
                    while i <= doc_input.Tables.Count:
                        next_table = doc_input.Tables.Item(i)
                        encontrou_novo_cargo = False
                        for row_idx in range(1, next_table.Rows.Count + 1):
                            try:
                                row = next_table.Rows.Item(row_idx)
                                textos_linha = []
                                for cell_idx in range(1, row.Cells.Count + 1):
                                    try:
                                        cell = row.Cells.Item(cell_idx)
                                        cell_text = cell.Range.Text.strip().replace('\r', '').replace('\n', ' ').replace('\x07', '')
                                        textos_linha.append(cell_text)
                                    except Exception:
                                        textos_linha.append("")
                                if any(("Cargo:" in t or "Setor:" in t) for t in textos_linha):
                                    encontrou_novo_cargo = True
                                    break
                                for campo in ['Insalubridade', 'Periculosidade', 'Aposentadoria Especial']:
                                    for idx, texto in enumerate(textos_linha):
                                        if campo.upper() in texto.upper():
                                            valor = None
                                            for v_idx in range(idx+1, len(textos_linha)):
                                                v_text = textos_linha[v_idx].strip()
                                                if v_text and not any(c.upper() in v_text.upper() for c in ['Insalubridade', 'Periculosidade', 'Aposentadoria Especial']):
                                                    valor = v_text
                                                    break
                                            if valor:
                                                campo_key = campo.lower().replace(' ', '_')
                                                dados_cargos[nome_formatado][campo_key].append(valor)
                                                print(f"{campo_key} encontrado para {nome_formatado}: {valor}")
                                            break
                            except Exception:
                                continue
                        if encontrou_novo_cargo:
                            break
                        i += 1
                else:
                    i += 1
            except Exception as table_error:
                print(f"Erro ao processar tabela {i}: {str(table_error)}")
                i += 1
                continue
        # Aplicar prioridade para cada campo
        for cargo, campos in dados_cargos.items():
            # INSALUBRIDADE: prioridade com maior porcentagem
            campo = 'insalubridade'
            valores = campos[campo]
            escolhido = ''
            # 1. PREJUDICADO
            for v in valores:
                if v.upper().startswith('PREJUDICADO'):
                    escolhido = v
                    break
            # 2. Maior porcentagem SIM
            if not escolhido:
                maiores = []
                for v in valores:
                    m = re.search(r'SIM\s*[-–]?\s*(\d+)[%％]', v.upper())
                    if m:
                        maiores.append((int(m.group(1)), v))
                if maiores:
                    maiores.sort(reverse=True)
                    escolhido = maiores[0][1]
            # 3. SIM sem porcentagem
            if not escolhido:
                for v in valores:
                    if v.upper().startswith('SIM'):
                        escolhido = v
                        break
            # 4. NÃO
            if not escolhido:
                for v in valores:
                    if v.upper().startswith('NÃO'):
                        escolhido = v
                        break
            campos[campo] = escolhido

            # PERICULOSIDADE: prioridade com maior porcentagem
            campo = 'periculosidade'
            valores = campos[campo]
            escolhido = ''
            # 1. PREJUDICADO
            for v in valores:
                if v.upper().startswith('PREJUDICADO'):
                    escolhido = v
                    break
            # 2. Maior porcentagem SIM
            if not escolhido:
                maiores = []
                for v in valores:
                    m = re.search(r'SIM\s*[-–]?\s*(\d+)[%％]', v.upper())
                    if m:
                        maiores.append((int(m.group(1)), v))
                if maiores:
                    maiores.sort(reverse=True)
                    escolhido = maiores[0][1]
            # 3. SIM sem porcentagem
            if not escolhido:
                for v in valores:
                    if v.upper().startswith('SIM'):
                        escolhido = v
                        break
            # 4. NÃO
            if not escolhido:
                for v in valores:
                    if v.upper().startswith('NÃO'):
                        escolhido = v
                        break
            campos[campo] = escolhido

            # APOSENTADORIA ESPECIAL: prioridade com maior ano
            campo = 'aposentadoria_especial'
            valores = campos[campo]
            escolhido = ''
            # 1. PREJUDICADO
            for v in valores:
                if v.upper().startswith('PREJUDICADO'):
                    escolhido = v
                    break
            # 2. Maior ano SIM
            if not escolhido:
                maiores = []
                for v in valores:
                    m = re.search(r'SIM\s*[-–]?\s*(\d+)\s*ANOS?', v.upper())
                    if m:
                        maiores.append((int(m.group(1)), v))
                if maiores:
                    maiores.sort(reverse=True)
                    escolhido = maiores[0][1]
            # 3. SIM sem ano
            if not escolhido:
                for v in valores:
                    if v.upper().startswith('SIM'):
                        escolhido = v
                        break
            # 4. NÃO
            if not escolhido:
                for v in valores:
                    if v.upper().startswith('NÃO'):
                        escolhido = v
                        break
            campos[campo] = escolhido

        print(f"\nDados coletados de todos os cargos: {dados_cargos}")
        print(f"\nProcurando tabela de destino no documento de saída...")
        print(f"Número total de tabelas no documento de saída: {doc_output.Tables.Count}")
        tabela_destino = None
        for i in range(1, doc_output.Tables.Count + 1):
            try:
                table = doc_output.Tables.Item(i)
                print(f"\nVerificando tabela {i} de {doc_output.Tables.Count}")
                primeira_celula = table.Cell(1, 1).Range.Text.strip()
                print(f"Conteúdo da primeira célula: {primeira_celula}")
                if "CARGO/ATIVIDADE" in primeira_celula:
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
            if num_linhas_atual < num_cargos + 1:
                for _ in range(num_cargos + 1 - num_linhas_atual):
                    tabela_destino.Rows.Add()
            elif num_linhas_atual > num_cargos + 1:
                for _ in range(num_linhas_atual - (num_cargos + 1)):
                    tabela_destino.Rows.Item(tabela_destino.Rows.Count).Delete()
            print(f"Tabela ajustada para {num_cargos + 1} linhas (incluindo cabeçalho)")
            header_row = tabela_destino.Rows.Item(1)
            for j in range(1, header_row.Cells.Count + 1):
                cell = header_row.Cells.Item(j)
                formatar_celula(cell)
            idx = 2
            for cargo_formatado, dados in dados_cargos.items():
                try:
                    row = tabela_destino.Rows.Item(idx)
                    cell = row.Cells.Item(1)
                    cell.Range.Text = cargo_formatado
                    formatar_celula(cell)
                    cell = row.Cells.Item(2)
                    cell.Range.Text = dados['insalubridade'] if dados['insalubridade'] else ''
                    formatar_celula(cell)
                    cell = row.Cells.Item(3)
                    cell.Range.Text = dados['periculosidade'] if dados['periculosidade'] else ''
                    formatar_celula(cell)
                    cell = row.Cells.Item(4)
                    cell.Range.Text = dados['aposentadoria_especial'] if dados['aposentadoria_especial'] else ''
                    formatar_celula(cell)
                except Exception as cell_error:
                    print(f"Erro ao preencher células da linha {idx}: {str(cell_error)}")
                    continue
                idx += 1
            print("\nSalvando alterações...")
            doc_output.Save()
            time.sleep(2)
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

# if __name__ == "__main__":
#     sucesso = preencher_dados_tabelas_funcao(
#         r"C:\Users\Nitro\tecnico\PGR - GRO\FORMATAÇÃO\LTCAT\MAIO 2025 - LTCAT - 22067325000124 - SETT SINALIZACAO E EQUIPAMENTOS DE TRANSITO E COMERCIO LTDA (1).rtf",
#         r"C:\Users\Nitro\tecnico\PGR - GRO\FORMATAÇÃO\TEMPLATE\template_ltcat_padrao.docx"
#     )
    
#     if sucesso:
#         print("Processamento concluído com sucesso!")
#     else:
#         print("Ocorreu um erro durante o processamento.")