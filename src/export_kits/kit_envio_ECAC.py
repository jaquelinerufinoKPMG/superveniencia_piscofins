import glob
import io
import os
import re
import shutil
from typing import List, Dict, Any, Tuple

import pandas as pd
import PyPDF2
from tqdm import tqdm
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

import pikepdf
import fitz  
#Junta os pdf, coloca uma folha de rosto e zipa com limite de 14mb 

# Helpers  Detectar "visual blank" 

def _page_is_visually_blank(page: fitz.Page, zoom: float = 1.5) -> bool:
    """
    Renderiza a página e verifica se ela é praticamente toda branca.
    Isso detecta o caso: PDF original ok, mas quando juntado fica branco.
    """
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)

    # Amostra simples: pega bytes e checa se quase tudo é branco.
    # pix.samples é bytes em RGB.
    samples = pix.samples
    if not samples:
        return True

    # Conta pixels não-brancos com tolerância.

    nonwhite = 0
    total = 0
    for i in range(0, len(samples), 3 * stride):
        r = samples[i]
        g = samples[i + 1]
        b = samples[i + 2]
        total += 1
        if r < 245 or g < 245 or b < 245:
            nonwhite += 1
            if nonwhite > 10:  # achou conteúdo suficiente
                return False

    # Se praticamente tudo é branco, considera em branco.
    return nonwhite == 0


def pdf_is_visually_blank(caminho_pdf: str, check_pages: int = 2) -> bool:
    """
    Abre com PyMuPDF e verifica se as primeiras páginas são visualmente brancas.
    """
    try:
        doc = fitz.open(caminho_pdf)
    except Exception:
        # Verificar se abriu
        return True

    try:
        if doc.page_count == 0:
            return True

        n = min(doc.page_count, check_pages)
        for i in range(n):
            page = doc.load_page(i)
            if not _page_is_visually_blank(page):
                return False
        return True
    finally:
        doc.close()


def flatten_pdf_to_images(input_pdf: str, output_pdf: str, dpi: int = 200) -> None:
    """
    "Achata" o PDF: renderiza cada página para imagem e recria um PDF.
    Isso elimina qualquer dependência de objetos/camadas que o merge possa perder.
    """
    src = fitz.open(input_pdf)
    try:
        out = fitz.open()
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)

        for i in range(src.page_count):
            page = src.load_page(i)
            pix = page.get_pixmap(matrix=mat, alpha=False)

            # Cria uma página no PDF de saída com o tamanho correto
            width_pt = pix.width / zoom
            height_pt = pix.height / zoom
            new_page = out.new_page(width=width_pt, height=height_pt)

            # Insere a imagem rasterizada ocupando a página toda
            img_rect = fitz.Rect(0, 0, width_pt, height_pt)
            new_page.insert_image(img_rect, pixmap=pix)

        out.save(output_pdf, deflate=True)
        out.close()
    finally:
        src.close()


def extrair_contrato(nome_arquivo: str) -> str:
    base = os.path.splitext(os.path.basename(nome_arquivo))[0]
    dig = re.sub(r"\D", "", base)
    return dig if dig else base


def criar_pdf_capa(contrato: str) -> io.BytesIO:
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    c.setFont("Helvetica-Bold", 14)
    width, height = letter
    texto = f"Contrato Número: {contrato}"
    c.drawCentredString(width / 2, height / 2, texto)
    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer


def process_anexo_pdfs(
    caminho_base=r"C:\projetos\superveniencia_piscofins\data\Output\Resumos\pdf",
    limite_mb=14,
    contador_inicial=0
):
    contratos_invalidos = []  
    limite_bytes = int(limite_mb * 1024 * 1024)

    lote = os.path.basename(os.path.dirname(caminho_base))
    anexo = os.path.basename(caminho_base)

    # Pasta de saída para os lotes
    pasta_final = f"{caminho_base}_{lote}_partes"
    os.makedirs(pasta_final, exist_ok=True)

    # Coleta PDFs
    arquivos_info: List[Tuple[str, str, int]] = []
    for f in sorted(os.listdir(caminho_base)):
        if f.lower().endswith(".pdf"):
            caminho_pdf = os.path.join(caminho_base, f)
            tamanho = os.path.getsize(caminho_pdf)
            contrato_num = extrair_contrato(f)
            if contrato_num not in contratos_invalidos:
                arquivos_info.append((contrato_num, caminho_pdf, tamanho))

    # Divide PDFs grandes
    def dividir_pdf(origem: str, tamanho_max: int, label: str) -> List[str]:
        partes = []
        temp_path = os.path.join(pasta_final, "temp_part.pdf")

        with open(origem, "rb") as f:
            reader = PyPDF2.PdfReader(f, strict=False)
            if getattr(reader, "is_encrypted", False):
                try:
                    reader.decrypt("")
                except Exception:
                    raise RuntimeError(f"PDF criptografado (decrypt falhou): {origem}")

            writer = None
            parte_indice = 1
            tamanho_acumulado = 0

            for page in reader.pages:
                temp_writer = PyPDF2.PdfWriter()
                temp_writer.add_page(page)

                with open(temp_path, "wb") as temp_out:
                    temp_writer.write(temp_out)

                parte_tamanho = os.path.getsize(temp_path)

                if (writer is None) or (tamanho_acumulado + parte_tamanho > tamanho_max):
                    if writer is not None:
                        nome_parte = os.path.join(pasta_final, f"{label}_parte_{parte_indice}.pdf")
                        with open(nome_parte, "wb") as out:
                            writer.write(out)
                        partes.append(nome_parte)
                        parte_indice += 1

                    writer = PyPDF2.PdfWriter()
                    tamanho_acumulado = 0

                writer.add_page(page)
                tamanho_acumulado += parte_tamanho

            if writer is not None:
                nome_parte = os.path.join(pasta_final, f"{label}_parte_{parte_indice}.pdf")
                with open(nome_parte, "wb") as out:
                    writer.write(out)
                partes.append(nome_parte)

        if os.path.exists(temp_path):
            os.remove(temp_path)

        return partes

    # Agrupa em lotes de até 14MB
    grupos: List[List[Tuple[str, str]]] = []
    grupo_atual: List[Tuple[str, str]] = []
    tamanho_atual = 0

    for contrato, caminho_pdf, tamanho in arquivos_info:
        if tamanho > limite_bytes:
            nome_base = os.path.splitext(os.path.basename(caminho_pdf))[0]
            partes = dividir_pdf(caminho_pdf, limite_bytes, nome_base)

            for parte in partes:
                p_size = os.path.getsize(parte)
                parte_nome = os.path.basename(parte)

                if p_size + tamanho_atual <= limite_bytes:
                    grupo_atual.append((contrato, parte_nome))
                    tamanho_atual += p_size
                else:
                    if grupo_atual:
                        grupos.append(grupo_atual)
                    grupo_atual = [(contrato, parte_nome)]
                    tamanho_atual = p_size
        else:
            if tamanho + tamanho_atual <= limite_bytes:
                grupo_atual.append((contrato, caminho_pdf))
                tamanho_atual += tamanho
            else:
                if grupo_atual:
                    grupos.append(grupo_atual)
                grupo_atual = [(contrato, caminho_pdf)]
                tamanho_atual = tamanho

    if grupo_atual:
        grupos.append(grupo_atual)

    dados_excel: List[Dict[str, Any]] = []
    final_pdfs: List[str] = []
    erros: List[Dict[str, Any]] = []

    # Merge por grupo (PIKEPDF + fallback)
  
    for i, grupo in enumerate(tqdm(grupos, desc="Processando grupos"), start=1):
        # Registra para Excel
        for (p, arquivo) in grupo:
            parte = f" - Parte {arquivo.split('_parte_')[-1].split('.pdf')[0]}" if "_parte_" in arquivo else ""
            dados_excel.append({
                "Contrato Número": p,
                "Lote": f"{anexo}_{lote}_parte_{i}.pdf",
                "Parte do contrato": parte
            })

        pdf_final = pikepdf.Pdf.new()

        for (p, arquivo) in grupo:
            # Capa
            capa_buffer = criar_pdf_capa(p)
            capa_pdf = pikepdf.Pdf.open(capa_buffer)
            pdf_final.pages.extend(capa_pdf.pages)

            # Resolve caminho real
            if os.path.exists(arquivo):
                caminho_completo = arquivo
            else:
                caminho_completo = os.path.join(pasta_final, arquivo)

            # 1) tenta juntar direto com pikepdf
            try:
                with pikepdf.open(caminho_completo) as pdf_contrato:
                    # anexa
                    start_pages = len(pdf_final.pages)
                    pdf_final.pages.extend(pdf_contrato.pages)

                # 2) valida visualmente: se o contrato entrou "em branco", faz fallback 
                if not pdf_is_visually_blank(caminho_completo, check_pages=1):
                    pass

                # Validação do resultado por contrato:
                # salva um temp do grupo até aqui e checa as últimas páginas adicionadas (mais caro, mas garante)
                tmp_valid = os.path.join(pasta_final, f"__tmp_validacao_grupo_{i}.pdf")
                pdf_final.save(tmp_valid)

                # A página do contrato começa após as páginas atuais menos as do contrato;
                doc_tmp = fitz.open(tmp_valid)
                try:
                    if doc_tmp.page_count >= 2:
                        # checa a última página do temp (geralmente parte do contrato)
                        last_page = doc_tmp.load_page(doc_tmp.page_count - 1)
                        if _page_is_visually_blank(last_page):
                            raise RuntimeError("Conteúdo ficou em branco após merge (fallback ativado)")
                finally:
                    doc_tmp.close()
                    try:
                        os.remove(tmp_valid)
                    except Exception:
                        pass

            except Exception:
                # Fallback A:juntar
                try:
                    nome_base = os.path.splitext(os.path.basename(caminho_completo))[0]
                    caminho_achatado = os.path.join(pasta_final, f"{nome_base}__FLATTEN.pdf")

                    if not os.path.exists(caminho_achatado):
                        flatten_pdf_to_images(caminho_completo, caminho_achatado, dpi=200)

                    with pikepdf.open(caminho_achatado) as pdf_flat:
                        pdf_final.pages.extend(pdf_flat.pages)

                    erros.append({
                        "Contrato Número": p,
                        "Arquivo": os.path.basename(caminho_completo),
                        "Caminho": caminho_completo,
                        "Ação": "Fallback: achatado (raster)",
                        "Lote": f"{anexo}_{lote}_parte_{i}.pdf"
                    })

                except Exception as e2:
                    erros.append({
                        "Contrato Número": p,
                        "Arquivo": os.path.basename(caminho_completo),
                        "Caminho": caminho_completo,
                        "Ação": "Falhou até no fallback",
                        "Erro": str(e2),
                        "Lote": f"{anexo}_{lote}_parte_{i}.pdf"
                    })

        # Salva o PDF do grupo
        nome_pdf_final = os.path.join(pasta_final, f"{anexo}_{lote}_parte_{i}.pdf")
        pdf_final.save(nome_pdf_final)
        final_pdfs.append(nome_pdf_final)

   
    # máx 14MB
  
    pdf_to_subfolder: Dict[str, str] = {}
    max_docs = 14
    max_folder_size = 140 * 1024 * 1024

    current_group_files: List[str] = []
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

    # 
    # Excel consolidado + aba Erros/Fallback
    # 
    df = pd.DataFrame(dados_excel)
    df["Pasta Interna"] = df["Lote"].apply(lambda lote_nome: pdf_to_subfolder.get(lote_nome, ""))

    excel_path = os.path.join(pasta_final, f"{anexo}_{lote}_partes.xlsx")
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Lotes")
        if erros:
            pd.DataFrame(erros).to_excel(writer, index=False, sheet_name="Erros")

    # Compacta a pasta em ZIP e remove a pasta
    shutil.make_archive(pasta_final, "zip", pasta_final)
    shutil.rmtree(pasta_final)

    return folder_index + 1

# Execução


contador_inicial = 0
contador_inicial = process_anexo_pdfs(
    caminho_base=r"C:\projetos\superveniencia_piscofins\data\Output\Resumos\pdf",
    contador_inicial=contador_inicial
)

