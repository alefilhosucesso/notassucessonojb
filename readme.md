# Lançamento automático de notas tomadas no JB Cepil

## 1. Instalação (só uma vez)

Abra o Prompt de Comando e rode:

```
pip install pdfplumber openpyxl pyautogui pygetwindow pillow opencv-python
```

Deixe o arquivo `lancar_notas_jb.py` na pasta:

```
\\SUCESSOSERVER\dados\SCAN\AARQUIVOS TRANSITÓRIOS\ALE\Automacoes\Lançamento de nota tomada no jb\
```

A planilha e o relatório ficam na subpasta `Relatório e planilha`, como você combinou.

## 2. Antes de rodar

- O relatório PDF precisa se chamar **Report.pdf** (ou mude `ARQ_RELATORIO` no script)
- A planilha precisa se chamar **SUFIXOS NOTAS SUCESSO.xlsx**
- O script cria sozinho as subpastas `logs` e `ancoras`

## 3. Os três modos

### `--conferir` — comece sempre por aqui
```
python lancar_notas_jb.py --conferir
```
Lê os arquivos, mostra o plano completo e a lista de pendências. **Não digita nada.**
Use todo mês antes de lançar, para ver o que vai e o que não vai.

### `--ancoras` — configuração opcional (recomendada)
```
python lancar_notas_jb.py --ancoras
```
Ajuda a capturar as imagens de referência de cada tela. Sem elas o script confere
apenas o título da janela; com elas a verificação fica bem mais rigorosa.

Você precisa de 4 recortes pequenos salvos na pasta `ancoras`:

| Arquivo | O que recortar |
|---|---|
| `pesquisa_empresa.png` | o rótulo "Empresa:" da tela de pesquisa |
| `campo_pacote.png` | o rótulo "Pacote:" da barra superior |
| `digitacao_servicos.png` | o rótulo "Modelo:" da tela de digitação |
| `duplicatas.png` | o rótulo "Vencimento Título:" da tela de duplicatas |

Recortes pequenos e de alto contraste funcionam melhor que telas inteiras.
Recorte limpo, sem círculo/seta desenhado em cima — o script compara pixel
a pixel, então qualquer marcação nunca vai bater com a tela real.

### `--lancar` — a execução de verdade
```
python lancar_notas_jb.py --lancar
```
Pede o vencimento do mês, mostra o plano, pede confirmação digitada e depois
o script procura a janela do JB Cepil sozinho e traz para frente — não é
mais preciso clicar nela a tempo.

**Importante: não clique na janela do console/PowerShell enquanto o script
roda.** Isso ativa o modo "Selecionar" do console, que rouba o foco do
teclado — daí a digitação para de ir para o JB. Se acontecer, o script
detecta e avisa; para destravar, clique de novo no JB ou aperte Esc no
console. Para evitar de vez: nas propriedades do PowerShell/Prompt de
Comando, desmarque "Modo de Seleção Rápida" (Quick Edit Mode).

## 4. Parada de emergência

Jogue o mouse no **canto superior esquerdo** da tela. O script aborta na hora.

## 5. Como ele para sozinho

O script interrompe o lote inteiro e salva um print da tela se:

- o total de notas lidas não bater com o total do rodapé do PDF
- a janela ativa não for a esperada em qualquer etapa
- uma imagem âncora não for encontrada na tela
- o nome de uma empresa não puder ser remontado do PDF

E **pula a nota** (registrando no log de pendências) se:

- a empresa não bater exatamente com nenhum nome da planilha
- o mesmo nome aparecer com dois códigos diferentes
- a empresa tiver observação na planilha (PIS E COFINS, CONTA etc.)
- o sufixo for 0, vazio ou inválido
- a situação da nota não for "Emitido"

## 6. Arquivos gerados a cada execução

Na subpasta `logs`:

- `lancamento_AAAA-MM-DD_HHMMSS.log` — tudo que aconteceu
- `pendencias_AAAA-MM-DD_HHMMSS.csv` — notas para lançar manualmente (abre no Excel)
- `erro_*.png` — print da tela, se algo deu errado

## 7. Dois ajustes para conferir no primeiro teste

Estão no topo do script, na seção CONFIGURAÇÃO:

**`CFOP_SUFIXO_COM_ENTER`** — o manual diz que o sufixo entra "junto do CFOP".
Deixei como `False` (digita 1933 e o sufixo em sequência, sem enter entre eles).
Se na prática o sufixo for um campo separado, mude para `True`.

**`PAUSA_TELA` e `PAUSA_CAMPO`** — se o JB estiver lento e o script começar a
digitar antes da tela abrir, aumente esses valores.

## 8. Primeira vez: teste com poucas notas

Faça um PDF de teste com 2 ou 3 notas apenas, rode o `--lancar` e confira no JB
se ficou tudo certo antes de soltar o lote inteiro.
