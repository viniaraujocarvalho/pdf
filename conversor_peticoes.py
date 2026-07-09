#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversor e Compactador de Petições
====================================

Programa com interface gráfica para advogados que precisam:

1. Converter petições feitas no Word (.docx / .doc) para PDF em LOTE
   (várias de uma vez).
2. Compactar PDFs até ficarem abaixo de 5 MB por arquivo, sem se
   preocupar com a qualidade das imagens/informações (foco em caber no
   limite de upload de tribunais como PJe/eproc).

Dependências (instale com: pip install -r requirements.txt):
    - PyMuPDF  (compactação dos PDFs)
    - docx2pdf (opcional, converte Word usando o Microsoft Word — Windows/Mac)
    Alternativa de conversão: LibreOffice instalado (comando "soffice"),
    detectado automaticamente. Funciona em Windows, Mac e Linux.

Executar:  python conversor_peticoes.py
"""

import os
import sys
import shutil
import subprocess
import threading
import queue
import tempfile
import traceback

# O Tkinter só é necessário para a janela gráfica. Se não estiver disponível
# (ambiente sem interface), o módulo ainda pode ser importado e as funções de
# conversão/compactação continuam utilizáveis; apenas a janela não abre.
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    _TK_BASE = tk.Tk
except Exception:  # pragma: no cover
    tk = ttk = filedialog = messagebox = None
    _TK_BASE = object

# ---------------------------------------------------------------------------
# Configurações
# ---------------------------------------------------------------------------
LIMITE_MB = 5.0                       # limite por arquivo (megabytes)
LIMITE_BYTES = int(LIMITE_MB * 1024 * 1024)

# Sequência de DPI usada ao rasterizar as páginas na compactação agressiva.
# Vai reduzindo até o PDF caber no limite. Como a qualidade não importa,
# podemos descer bastante.
DPIS = [200, 150, 120, 100, 80, 72, 60, 50, 42, 36, 30, 24]
# Qualidade JPEG usada ao recomprimir as páginas rasterizadas.
JPEG_QUALIDADES = [50, 35, 25, 15, 10]


# ---------------------------------------------------------------------------
# Conversão Word -> PDF
# ---------------------------------------------------------------------------
def _achar_libreoffice():
    """Retorna o caminho do executável do LibreOffice, se instalado."""
    candidatos = ["soffice", "libreoffice"]
    if sys.platform.startswith("win"):
        candidatos += [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
    elif sys.platform == "darwin":
        candidatos += ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    for c in candidatos:
        achado = shutil.which(c) if os.path.basename(c) == c else (c if os.path.exists(c) else None)
        if achado:
            return achado
    return None


def converter_com_libreoffice(caminho_docx, pasta_saida, soffice):
    """Converte um Word para PDF usando o LibreOffice em modo headless."""
    proc = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir",
         pasta_saida, caminho_docx],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=300,
    )
    nome = os.path.splitext(os.path.basename(caminho_docx))[0] + ".pdf"
    destino = os.path.join(pasta_saida, nome)
    if not os.path.exists(destino):
        raise RuntimeError(
            "LibreOffice não gerou o PDF.\n" +
            proc.stderr.decode(errors="ignore")[:500]
        )
    return destino


def converter_com_docx2pdf(caminho_docx, pasta_saida):
    """Converte um Word para PDF usando o Microsoft Word (via docx2pdf)."""
    from docx2pdf import convert
    nome = os.path.splitext(os.path.basename(caminho_docx))[0] + ".pdf"
    destino = os.path.join(pasta_saida, nome)
    convert(caminho_docx, destino)
    if not os.path.exists(destino):
        raise RuntimeError("A conversão via Microsoft Word falhou.")
    return destino


def obter_conversor(log):
    """Escolhe o mecanismo de conversão disponível.

    Retorna uma função conversora(caminho_docx, pasta_saida) -> caminho_pdf,
    ou None se nenhum mecanismo estiver disponível.
    """
    soffice = _achar_libreoffice()
    if soffice:
        log(f"Conversor de Word: LibreOffice ({soffice})")
        return lambda docx, out: converter_com_libreoffice(docx, out, soffice)

    try:
        import docx2pdf  # noqa: F401
        log("Conversor de Word: Microsoft Word (docx2pdf)")
        return lambda docx, out: converter_com_docx2pdf(docx, out)
    except Exception:
        pass

    return None


# ---------------------------------------------------------------------------
# Compactação de PDF
# ---------------------------------------------------------------------------
def _tamanho(caminho):
    return os.path.getsize(caminho)


def compactar_pdf(caminho_pdf, caminho_saida, limite_bytes=LIMITE_BYTES, log=print):
    """Compacta um PDF até ficar abaixo do limite (em bytes).

    Estratégia:
      1. Reescreve o PDF com limpeza e compressão de streams (deflate).
         Muitas vezes já resolve.
      2. Se ainda estiver acima do limite, rasteriza as páginas em imagens
         JPEG, reduzindo o DPI e a qualidade progressivamente até caber.
         Como a qualidade não importa, isso garante o resultado.

    Retorna o tamanho final em bytes.
    """
    import fitz  # PyMuPDF

    tam_original = _tamanho(caminho_pdf)

    # --- Passo 1: reescrita otimizada (sem perda de conteúdo) ---
    try:
        doc = fitz.open(caminho_pdf)
        doc.save(
            caminho_saida,
            garbage=4,       # remove objetos não usados / duplicados
            deflate=True,    # comprime os streams
            clean=True,
            deflate_images=True,
            deflate_fonts=True,
        )
        doc.close()
        if _tamanho(caminho_saida) <= limite_bytes:
            log(f"    otimização simples bastou "
                f"({_tamanho(caminho_saida)/1024/1024:.2f} MB)")
            return _tamanho(caminho_saida)
    except Exception as e:
        log(f"    aviso: otimização simples falhou ({e}); "
            f"seguindo para rasterização.")
        # Se a reescrita falhar, copia o original para prosseguir.
        shutil.copyfile(caminho_pdf, caminho_saida)

    # --- Passo 2: rasterização progressiva ---
    fonte = caminho_pdf
    melhor_resultado = None  # (tamanho, caminho_temp)

    for dpi in DPIS:
        for qualidade in JPEG_QUALIDADES:
            tmp = caminho_saida + f".tmp_{dpi}_{qualidade}.pdf"
            try:
                tam = _rasterizar(fonte, tmp, dpi, qualidade)
            except Exception as e:
                log(f"    erro rasterizando dpi={dpi} q={qualidade}: {e}")
                if os.path.exists(tmp):
                    os.remove(tmp)
                continue

            # Guarda o menor resultado obtido até agora (rede de segurança).
            if melhor_resultado is None or tam < melhor_resultado[0]:
                if melhor_resultado is not None and os.path.exists(melhor_resultado[1]):
                    os.remove(melhor_resultado[1])
                melhor_resultado = (tam, tmp)
            else:
                os.remove(tmp)

            log(f"    dpi={dpi:>3} qualidade={qualidade:>2} -> "
                f"{tam/1024/1024:.2f} MB")

            if tam <= limite_bytes:
                shutil.move(melhor_resultado[1], caminho_saida)
                return tam

    # Nenhuma combinação ficou abaixo do limite: usa o menor resultado.
    if melhor_resultado is not None:
        shutil.move(melhor_resultado[1], caminho_saida)
        log("    aviso: não foi possível ficar abaixo de "
            f"{limite_bytes/1024/1024:.0f} MB; usando o menor tamanho "
            f"({_tamanho(caminho_saida)/1024/1024:.2f} MB).")
        return _tamanho(caminho_saida)

    # Fallback total: mantém o que já estava salvo.
    log("    aviso: rasterização não produziu resultado; "
        "mantendo a versão otimizada.")
    return _tamanho(caminho_saida) if os.path.exists(caminho_saida) else tam_original


def _rasterizar(caminho_pdf, caminho_saida, dpi, qualidade):
    """Gera um novo PDF onde cada página é uma imagem JPEG rasterizada."""
    import fitz  # PyMuPDF

    origem = fitz.open(caminho_pdf)
    novo = fitz.open()
    try:
        zoom = dpi / 72.0
        matriz = fitz.Matrix(zoom, zoom)
        for pagina in origem:
            pix = pagina.get_pixmap(matrix=matriz, alpha=False)
            img_jpeg = pix.tobytes("jpeg", jpg_quality=qualidade)
            nova_pag = novo.new_page(width=pagina.rect.width,
                                     height=pagina.rect.height)
            nova_pag.insert_image(nova_pag.rect, stream=img_jpeg)
        novo.save(caminho_saida, garbage=4, deflate=True, clean=True)
    finally:
        novo.close()
        origem.close()
    return _tamanho(caminho_saida)


# ---------------------------------------------------------------------------
# Interface gráfica
# ---------------------------------------------------------------------------
class Aplicacao(_TK_BASE):
    def __init__(self):
        super().__init__()
        self.title("Conversor e Compactador de Petições")
        self.geometry("760x560")
        self.minsize(680, 480)

        self.fila_log = queue.Queue()
        self.trabalhando = False

        self._montar_interface()
        self.after(100, self._processar_fila_log)

    # ---- construção da interface ----
    def _montar_interface(self):
        pad = {"padx": 8, "pady": 4}

        topo = ttk.Frame(self)
        topo.pack(fill="x", **pad)

        ttk.Label(
            topo,
            text="Converta petições do Word para PDF em lote e compacte para até 5 MB por arquivo.",
            font=("Segoe UI", 10, "bold"),
            wraplength=720,
        ).pack(anchor="w")

        # Modo de operação
        modo_frame = ttk.LabelFrame(self, text="O que você quer fazer?")
        modo_frame.pack(fill="x", **pad)

        self.modo = tk.StringVar(value="converter_compactar")
        ttk.Radiobutton(
            modo_frame, text="Converter Word → PDF e já compactar (recomendado)",
            variable=self.modo, value="converter_compactar",
            command=self._atualizar_estado,
        ).pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(
            modo_frame, text="Somente converter Word → PDF",
            variable=self.modo, value="converter",
            command=self._atualizar_estado,
        ).pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(
            modo_frame, text="Somente compactar PDFs que já existem",
            variable=self.modo, value="compactar",
            command=self._atualizar_estado,
        ).pack(anchor="w", padx=8, pady=2)

        # Pasta de entrada
        ent_frame = ttk.Frame(self)
        ent_frame.pack(fill="x", **pad)
        ttk.Label(ent_frame, text="Pasta com os arquivos:").pack(anchor="w")
        linha1 = ttk.Frame(ent_frame)
        linha1.pack(fill="x")
        self.pasta_entrada = tk.StringVar()
        ttk.Entry(linha1, textvariable=self.pasta_entrada).pack(
            side="left", fill="x", expand=True)
        ttk.Button(linha1, text="Escolher...",
                   command=self._escolher_entrada).pack(side="left", padx=4)

        # Pasta de saída
        sai_frame = ttk.Frame(self)
        sai_frame.pack(fill="x", **pad)
        ttk.Label(sai_frame, text="Pasta de destino (onde salvar os resultados):").pack(anchor="w")
        linha2 = ttk.Frame(sai_frame)
        linha2.pack(fill="x")
        self.pasta_saida = tk.StringVar()
        ttk.Entry(linha2, textvariable=self.pasta_saida).pack(
            side="left", fill="x", expand=True)
        ttk.Button(linha2, text="Escolher...",
                   command=self._escolher_saida).pack(side="left", padx=4)

        # Botão de ação + progresso
        acao = ttk.Frame(self)
        acao.pack(fill="x", **pad)
        self.botao = ttk.Button(acao, text="Iniciar", command=self._iniciar)
        self.botao.pack(side="left")
        self.progresso = ttk.Progressbar(acao, mode="determinate")
        self.progresso.pack(side="left", fill="x", expand=True, padx=8)

        # Log
        log_frame = ttk.LabelFrame(self, text="Andamento")
        log_frame.pack(fill="both", expand=True, **pad)
        self.texto_log = tk.Text(log_frame, height=12, wrap="word",
                                 state="disabled", font=("Consolas", 9))
        scroll = ttk.Scrollbar(log_frame, command=self.texto_log.yview)
        self.texto_log.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.texto_log.pack(side="left", fill="both", expand=True)

        self._atualizar_estado()

    def _atualizar_estado(self):
        pass  # espaço reservado para regras futuras de interface

    # ---- callbacks ----
    def _escolher_entrada(self):
        pasta = filedialog.askdirectory(title="Escolha a pasta com os arquivos")
        if pasta:
            self.pasta_entrada.set(pasta)
            if not self.pasta_saida.get():
                self.pasta_saida.set(os.path.join(pasta, "saida"))

    def _escolher_saida(self):
        pasta = filedialog.askdirectory(title="Escolha a pasta de destino")
        if pasta:
            self.pasta_saida.set(pasta)

    def log(self, msg):
        self.fila_log.put(msg)

    def _processar_fila_log(self):
        try:
            while True:
                msg = self.fila_log.get_nowait()
                self.texto_log.configure(state="normal")
                self.texto_log.insert("end", msg + "\n")
                self.texto_log.see("end")
                self.texto_log.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._processar_fila_log)

    def _iniciar(self):
        if self.trabalhando:
            return
        entrada = self.pasta_entrada.get().strip()
        saida = self.pasta_saida.get().strip()

        if not entrada or not os.path.isdir(entrada):
            messagebox.showerror("Erro", "Escolha uma pasta de entrada válida.")
            return
        if not saida:
            messagebox.showerror("Erro", "Escolha uma pasta de destino.")
            return
        os.makedirs(saida, exist_ok=True)

        self.trabalhando = True
        self.botao.configure(state="disabled")
        self.progresso.configure(value=0)

        t = threading.Thread(
            target=self._executar, args=(entrada, saida, self.modo.get()),
            daemon=True,
        )
        t.start()

    # ---- processamento (roda em thread separada) ----
    def _executar(self, entrada, saida, modo):
        try:
            self._processar(entrada, saida, modo)
        except Exception:
            self.log("ERRO INESPERADO:\n" + traceback.format_exc())
        finally:
            self.trabalhando = False
            self.botao.configure(state="normal")

    def _processar(self, entrada, saida, modo):
        precisa_converter = modo in ("converter", "converter_compactar")
        precisa_compactar = modo in ("compactar", "converter_compactar")

        conversor = None
        if precisa_converter:
            conversor = obter_conversor(self.log)
            if conversor is None:
                self.log(
                    "ERRO: nenhum conversor de Word encontrado.\n"
                    "  Instale o LibreOffice (gratuito) OU, no Windows/Mac com "
                    "o Microsoft Word, rode: pip install docx2pdf"
                )
                messagebox.showerror(
                    "Sem conversor",
                    "Não encontrei LibreOffice nem Microsoft Word (docx2pdf).\n\n"
                    "Instale o LibreOffice (gratuito) para converter Word em PDF."
                )
                return

        # Coleta a lista de arquivos a processar
        if precisa_converter:
            exts = (".docx", ".doc")
        else:
            exts = (".pdf",)
        arquivos = [
            os.path.join(entrada, f) for f in sorted(os.listdir(entrada))
            if f.lower().endswith(exts)
        ]

        if not arquivos:
            self.log(f"Nenhum arquivo {' / '.join(exts)} encontrado na pasta.")
            messagebox.showinfo("Nada a fazer",
                                f"Nenhum arquivo {' / '.join(exts)} encontrado.")
            return

        self.log(f"Encontrados {len(arquivos)} arquivo(s).\n")
        self.progresso.configure(maximum=len(arquivos))

        temp_dir = tempfile.mkdtemp(prefix="peticoes_")
        sucesso, falha = 0, 0
        try:
            for i, caminho in enumerate(arquivos, start=1):
                nome = os.path.basename(caminho)
                self.log(f"[{i}/{len(arquivos)}] {nome}")
                try:
                    if precisa_converter:
                        pdf = conversor(caminho, temp_dir)
                        self.log("    convertido para PDF.")
                    else:
                        pdf = caminho

                    nome_pdf = os.path.splitext(nome)[0] + ".pdf"
                    destino = os.path.join(saida, nome_pdf)

                    if precisa_compactar:
                        tam = compactar_pdf(pdf, destino, LIMITE_BYTES, self.log)
                        self.log(f"    salvo: {nome_pdf} "
                                 f"({tam/1024/1024:.2f} MB)")
                    else:
                        shutil.copyfile(pdf, destino)
                        self.log(f"    salvo: {nome_pdf} "
                                 f"({_tamanho(destino)/1024/1024:.2f} MB)")
                    sucesso += 1
                except Exception as e:
                    falha += 1
                    self.log(f"    ERRO: {e}")
                finally:
                    self.progresso.configure(value=i)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

        self.log(f"\nConcluído. Sucesso: {sucesso} | Falhas: {falha}")
        self.log(f"Arquivos salvos em: {saida}")
        messagebox.showinfo(
            "Concluído",
            f"Processamento finalizado.\n\n"
            f"Sucesso: {sucesso}\nFalhas: {falha}\n\nPasta: {saida}"
        )


def main():
    if tk is None:
        print(
            "Este programa precisa do Tkinter (interface gráfica), que não "
            "está disponível nesta instalação do Python.\n"
            "No Windows/Mac o Tkinter já vem com o Python padrão de "
            "python.org. No Linux, instale com, por exemplo:\n"
            "    sudo apt install python3-tk"
        )
        sys.exit(1)
    app = Aplicacao()
    app.mainloop()


if __name__ == "__main__":
    main()
