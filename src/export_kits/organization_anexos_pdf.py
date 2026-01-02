import glob
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import PyPDF2
import re
from tqdm import tqdm
import os
import pandas as pd
import shutil


def process_anexo_pdfs(caminho_base="C:/Projeto/data/Resumo", limite_mb=5, contador_inicial=289):
    # Caminho do arquivo .xlsx
    # file_path = "filtro2.xlsx"
    # data = pd.read_excel(file_path)
    # data["Contrato"] = data["Contrato"].astype(str)
    # contratos_invalidos = data["Contrato"].values.tolist()
    contratos_invalidos = []  # Zerado conforme o código original
    limite_bytes = limite_mb * 1024 * 1024
    lote = os.path.basename(os.path.dirname(caminho_base))
    anexo = os.path.basename(caminho_base)

    arquivos_info = []
    for f in os.listdir(caminho_base):
        if f.endswith(".pdf"):
            caminho_pdf = os.path.join(caminho_base, f)
            tamanho = os.path.getsize(caminho_pdf)
            # Verifica se o número do contrato (apenas dígitos do nome) não está na lista de contratos inválidos
            if re.sub(r'\D', '', f) not in contratos_invalidos:
                arquivos_info.append((re.sub(r'\D', '', f), caminho_pdf, tamanho))

    # Cria a pasta de saída para os lotes
    pasta_final = f"{caminho_base}_{lote}_partes"
    os.makedirs(pasta_final, exist_ok=True)

    # Função para dividir PDFs que excedem o tamanho máximo permitido
    def dividir_pdf(origem, tamanho_max, label):
        partes = []
        with open(origem, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            writer = None
            parte_indice = 1
            tamanho_acumulado = 0
            for page in reader.pages:
                temp_writer = PyPDF2.PdfWriter()
                temp_writer.add_page(page)
                # Salva temporariamente para verificar o tamanho da página
                with open("temp_part.pdf", "wb") as temp_out:
                    temp_writer.write(temp_out)
                parte_tamanho = os.path.getsize("temp_part.pdf")
                if not writer or (tamanho_acumulado + parte_tamanho > tamanho_max):
                    if writer:
                        nome_parte = os.path.join(pasta_final, f"{label}_parte_{parte_indice}.pdf")
                        with open(nome_parte, "wb") as out:
                            writer.write(out)
                        partes.append(nome_parte)
                        parte_indice += 1
                    writer = PyPDF2.PdfWriter()
                    tamanho_acumulado = 0
                writer.add_page(page)
                tamanho_acumulado += parte_tamanho
            if writer:
                nome_parte = os.path.join(pasta_final, f"{label}_parte_{parte_indice}.pdf")
                with open(nome_parte, "wb") as out:
                    writer.write(out)
                partes.append(nome_parte)
        os.remove("temp_part.pdf")
        return partes

    # Agrupa os PDFs (ou partes deles) em grupos cujo tamanho total não exceda 'limite_bytes'
    grupos = []
    grupo_atual = []
    tamanho_atual = 0
    for pasta, caminho_pdf, tamanho in arquivos_info:
        if tamanho > limite_bytes:
            nome_base = os.path.splitext(os.path.basename(caminho_pdf))[0]
            partes = dividir_pdf(caminho_pdf, limite_bytes, nome_base)
            for parte in partes:
                p_size = os.path.getsize(parte)
                parte_nome = os.path.basename(parte)
                if p_size + tamanho_atual <= limite_bytes:
                    grupo_atual.append((pasta, parte_nome))
                    tamanho_atual += p_size
                else:
                    if grupo_atual:
                        grupos.append(grupo_atual)
                    grupo_atual = [(pasta, parte_nome)]
                    tamanho_atual = p_size
        else:
            if tamanho + tamanho_atual <= limite_bytes:
                grupo_atual.append((pasta, caminho_pdf))
                tamanho_atual += tamanho
            else:
                if grupo_atual:
                    grupos.append(grupo_atual)
                grupo_atual = [(pasta, caminho_pdf)]
                tamanho_atual = tamanho
    if grupo_atual:
        grupos.append(grupo_atual)

    dados_excel = []
    final_pdfs = []  # Lista para armazenar os caminhos dos PDFs finais (lotes)
    for i, grupo in enumerate(tqdm(grupos, desc="Processando grupos"), start=1):
        # Registra informações para o Excel (sem criar o TXT)
        for (p, arquivo) in grupo:
            parte = f" - Parte {arquivo.split('_parte_')[-1].split('.pdf')[0]}" if "_parte_" in arquivo else ""
            dados_excel.append({"Contrato Número": p, "Lote": f"{anexo}_{lote}_parte_{i}.pdf", "Parte do contrato": parte})

        # Cria o PDF final para o grupo
        escritor_final = PyPDF2.PdfWriter()
        for (p, arquivo) in grupo:
            # Cria uma página de cabeçalho com o número do contrato
            buffer = io.BytesIO()
            c = canvas.Canvas(buffer, pagesize=letter)
            c.setFont("Helvetica-Bold", 14)
            width, height = letter
            texto = f"Contrato Número: {p}"
            c.drawCentredString(width / 2, height / 2, texto)
            c.showPage()
            c.save()
            buffer.seek(0)

            leitor_temp = PyPDF2.PdfReader(buffer)
            for pagina_temp in leitor_temp.pages:
                escritor_final.add_page(pagina_temp)

            # Determina o caminho completo do PDF a ser mesclado
            if os.path.exists(arquivo):
                caminho_completo = arquivo
            else:
                caminho_completo = os.path.join(pasta_final, arquivo)

            with open(caminho_completo, "rb") as pdf_in:
                leitor = PyPDF2.PdfReader(pdf_in)
                for pagina in leitor.pages:
                    escritor_final.add_page(pagina)

        nome_pdf_final = os.path.join(pasta_final, f"{anexo}_{lote}_parte_{i}.pdf")
        with open(nome_pdf_final, "wb") as saida:
            escritor_final.write(saida)
        final_pdfs.append(nome_pdf_final)

    # Agrupa os PDFs finais em subpastas: cada subpasta terá, no máximo, 14 documentos e tamanho total <= 140 MB.
    # Cria um mapeamento de cada PDF final para a sua subpasta.
    pdf_to_subfolder = {}
    max_docs = 5
    max_folder_size = 140 * 1024 * 1024  # 140 MB em bytes
    current_group_files = []
    current_group_size = 0
    folder_index = contador_inicial

    for pdf_file in final_pdfs:
        file_size = os.path.getsize(pdf_file)
        if (len(current_group_files) >= max_docs) or (current_group_size + file_size > max_folder_size):
            subfolder = os.path.join(pasta_final, f"{folder_index}_{lote}_grupo")
            os.makedirs(subfolder, exist_ok=True)
            for f_pdf in current_group_files:
                destino = os.path.join(subfolder, os.path.basename(f_pdf))
                os.rename(f_pdf, destino)
                pdf_to_subfolder[os.path.basename(f_pdf)] = f"{folder_index}_{lote}_grupo"
            folder_index += 1
            current_group_files = []
            current_group_size = 0
        current_group_files.append(pdf_file)
        current_group_size += file_size

    if current_group_files:
        subfolder = os.path.join(pasta_final, f"{folder_index}_{lote}_grupo")
        os.makedirs(subfolder, exist_ok=True)
        for f_pdf in current_group_files:
            destino = os.path.join(subfolder, os.path.basename(f_pdf))
            os.rename(f_pdf, destino)
            pdf_to_subfolder[os.path.basename(f_pdf)] = f"{folder_index}_{lote}_grupo"

    # Cria o arquivo Excel consolidado com os dados dos contratos e lotes,
    # incluindo a coluna "Pasta Interna" que indica a subpasta onde o PDF foi salvo.
    df = pd.DataFrame(dados_excel)
    df["Pasta Interna"] = df["Lote"].apply(lambda lote: pdf_to_subfolder.get(lote, ""))
    df.to_excel(os.path.join(caminho_base + f"_{lote}_partes", f"{anexo}_{lote}_partes.xlsx"), index=False)

    # Remove arquivos temporários
    temp_files = ["temp_part.pdf"]
    for temp_file in temp_files:
        if os.path.exists(temp_file):
            os.remove(temp_file)

    # Compacta as pastas criadas em arquivos zip
    for p in glob.glob(caminho_base + f"_{lote}_partes/*.pdf"):
        os.remove

    shutil.make_archive(caminho_base + f"_{lote}_partes", 'zip', caminho_base + f"_{lote}_partes")
    shutil.rmtree(caminho_base + f"_{lote}_partes")
    return folder_index+1


# contador_inicial = process_anexo_pdfs(caminho_base="./listade_lotes/lote_8/Anexo_N_C")
# contador_inicial = process_anexo_pdfs(caminho_base="./listade_lotes/lote_8/Resumo",contador_inicial=contador_inicial)
# contador_inicial = process_anexo_pdfs(caminho_base="./listade_lotes/lote_8/Anexo_N_A",contador_inicial=contador_inicial)
contador_inicial = 0
# contador_inicial = process_anexo_pdfs(caminho_base="C:/Projeto/data/Resumo",contador_inicial=contador_inicial)

# contador_inicial = process_anexo_pdfs(caminho_base="data/Anexo N_A",contador_inicial=contador_inicial)
contador_inicial = process_anexo_pdfs(caminho_base="C:/Projeto/data/Resumo",contador_inicial=contador_inicial)

# contador_inicial = process_anexo_pdfs(caminho_base="./entrega_parcial/Resumo",contador_inicial=contador_inicial)
# contador_inicial = process_anexo_pdfs(caminho_base="data/filtro/Anexo N_A",contador_inicial=contador_inicial)