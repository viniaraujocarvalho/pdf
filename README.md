# Conversor e Compactador de Petições

Programa com **janela gráfica** (não precisa usar o terminal no dia a dia)
para advogados que precisam:

1. **Converter petições do Word (`.docx` / `.doc`) para PDF em lote** —
   várias de uma vez, apontando a pasta onde estão os arquivos.
2. **Compactar PDFs até no máximo 5 MB por arquivo**, ignorando a
   qualidade das imagens e informações — ideal para respeitar o limite de
   upload de sistemas como **PJe** e **eproc**.

---

## 🚀 Jeito mais fácil: baixar o programa pronto (.exe, sem instalar Python)

Se você não quer mexer com Python, use a versão pronta para Windows:

1. Acesse a página de **Releases** do projeto:
   <https://github.com/viniaraujocarvalho/pdf/releases>
2. Abra a versão **"Conversor e Compactador de Petições (Windows)"**.
3. Em **Assets**, clique em **`ConversorPeticoes.exe`** para baixar.
4. Dê **dois cliques** no arquivo baixado para abrir o programa.

> O Windows pode mostrar um aviso azul ("O Windows protegeu o computador").
> Isso é normal para programas novos: clique em **"Mais informações"** e depois
> em **"Executar assim mesmo"**.

Mesmo nessa versão, para **converter Word em PDF** você ainda precisa ter o
**LibreOffice** (gratuito) ou o **Microsoft Word** instalado — veja o item (c)
abaixo. A **compactação de PDFs** já funciona sozinha.

---

## 1. O que você precisa instalar (uma vez só)  — modo "com Python"

### a) Python

Baixe em <https://www.python.org/downloads/> e, na instalação (Windows),
marque a opção **"Add Python to PATH"**.

### b) Bibliotecas do programa

Abra o terminal (Prompt de Comando no Windows) **na pasta do programa** e rode:

```
pip install -r requirements.txt
```

### c) Um conversor de Word (escolha UM)

O programa converte Word para PDF usando um destes — ele detecta sozinho:

- **LibreOffice** (gratuito, funciona em Windows, Mac e Linux) —
  recomendado. Baixe em <https://pt-br.libreoffice.org/>. Não precisa abrir,
  só estar instalado.
- **Microsoft Word** (Windows/Mac): já é usado automaticamente via a
  biblioteca `docx2pdf`, que o comando acima instala.

> Se você só vai **compactar PDFs** (sem converter Word), não precisa de
> nenhum conversor — apenas do `PyMuPDF`.

---

## 2. Como usar

Na pasta do programa, execute:

```
python conversor_peticoes.py
```

A janela abre. Então:

1. Escolha **o que fazer**:
   - Converter Word → PDF **e já compactar** (recomendado);
   - Somente converter;
   - Somente compactar PDFs que já existem.
2. Clique em **"Escolher..."** na **pasta de entrada** (onde estão os `.docx`
   ou os `.pdf`).
3. Clique em **"Escolher..."** na **pasta de destino** (onde os resultados
   serão salvos — por padrão, uma subpasta `saida`).
4. Clique em **"Iniciar"**.

O andamento aparece na parte de baixo da janela e, ao final, uma mensagem
mostra quantos arquivos foram concluídos.

---

## 3. Como funciona a compactação para 5 MB

Para cada PDF:

1. Primeiro tenta uma **otimização sem perdas** (remove dados desnecessários
   e recomprime). Muitas vezes já resolve.
2. Se ainda passar de 5 MB, **rasteriza as páginas em imagens JPEG**,
   reduzindo a resolução (DPI) e a qualidade progressivamente **até o
   arquivo ficar abaixo de 5 MB**. Como a qualidade não importa, isso
   garante o resultado.

Se, mesmo no menor tamanho possível, o arquivo não couber em 5 MB (caso raro,
PDFs com centenas de páginas), o programa salva a **menor versão obtida** e
avisa no andamento.

---

## 4. Perguntas comuns

- **Preciso deixar o Word aberto?** Não.
- **Funciona com `.doc` antigo?** Sim, desde que o conversor (LibreOffice ou
  Word) dê conta do arquivo.
- **Onde ficam os arquivos originais?** Intactos. O programa nunca altera os
  originais; sempre grava novos arquivos na pasta de destino.
