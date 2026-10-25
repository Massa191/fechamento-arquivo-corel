# Fechamento de Arquivos CorelDRAW para Gráfica

Macro para **CorelDRAW** (VBA) que prepara o documento aberto para envio à gráfica: converte textos em curvas, separa e converte sombras, transforma transparências e efeitos em objetos seguros para impressão, passa todas as cores para **CMYK** e gera um **PDF/X-3** com sangria e marcas de corte opcionais.

O arquivo original **nunca é alterado**: o fechamento é feito numa cópia `nome_FECHADO.cdr`.

> **Status:** em uso real, testado com **CorelDRAW Graphics Suite 2022** no Windows. Outras versões podem funcionar, mas ainda não foram validadas (veja [Limitações](#limitações-conhecidas)).

---

## O que a macro faz

| Etapa | Tratamento |
|---|---|
| **Pré-verificação** | Mostra, sem alterar nada, o que será feito: textos, sombras, transparências, cores não-CMYK, bitmaps, objetos ocultos, efeitos e o maior excesso de objetos além da página (ajuda a escolher a sangria). |
| **Proteção do original** | Salva uma cópia `nome_FECHADO.cdr` na mesma pasta **antes** de qualquer alteração. Se o CDR final não for salvo, a cópia de segurança vai para a pasta temporária. |
| **Objetos ocultos** | Objetos e camadas ocultas são removidos. |
| **Objetos bloqueados** | São desbloqueados antes do processamento. |
| **Texto** | Convertido em curvas (inclusive dentro de grupos e PowerClips). Textos que não convertem na primeira tentativa, como o texto de uma linha de dimensão, são tentados de novo no fim da camada. |
| **Sombras** | A sombra é separada do objeto e **somente a sombra** vira bitmap CMYK com fundo transparente. O objeto (curvas, PNG, grupo) continua como estava. |
| **Transparências** | Viram bitmap CMYK com fundo transparente. |
| **Preenchimentos especiais** | Padrão, textura, PostScript e malha viram bitmap CMYK. |
| **Objetos de efeito** | Extrusão, mistura, contorno, chanfradura, sombra de bloco, mídia artística, envelope, linhas de dimensão etc.: convertidos em curvas quando o Corel permite via VBA; caso contrário, viram bitmap CMYK em alta resolução (1200 dpi). Os "originais" vazios que sobram são removidos. |
| **Código de barras** | Convertido em bitmap em alta resolução. |
| **QR Code e tabelas** | Convertidos em curvas. |
| **Cores** | Tudo para CMYK (RGB 0,0,0 vira K100). Cores especiais (Pantone etc.) também viram CMYK. |
| **Bitmaps** | RGB/paletizados são convertidos para CMYK; aviso para resolução efetiva abaixo de 150 dpi. |
| **Contornos finos** | Ajustados para a espessura mínima (0,1 mm). |
| **PDF** | PDF/X-3, fontes embutidas e texto em curvas, com **sangria em mm** e **marcas de corte** opcionais. |
| **Relatório** | Resumo com pendências em destaque e log completo por página (opcionalmente salvo em `.txt`). |

A janela de opções e a janela de **progresso** (página X de Y, tempo decorrido e restante estimado, botão *Parar*) são um formulário do VBA.

---

## Instalação

### Opção 1 – Instalador (recomendado)

1. Baixe `Instalador_FechamentoGrafica_x.y.z.exe` na página de [Releases](../../releases).
2. **Feche o CorelDRAW** e execute o instalador (ele pede permissão de administrador, pois grava em `Program Files`).
3. Marque as instalações do CorelDRAW em que deseja instalar a macro.
4. Abra o CorelDRAW. A macro `FecharDocumentoAtivo` fica disponível em **Ferramentas > Macros > Gerenciador de macros**.

> O instalador não é assinado digitalmente; o Windows SmartScreen ou o antivírus podem mostrar um aviso na primeira execução.

**Botão com ícone:** o arquivo `MeuBotao.cdws` fica na pasta de instalação (`C:\Program Files\FechamentoGrafica`). Em **Ferramentas > Opções > Personalização > Área de trabalho > Importar**, escolha esse arquivo. *Atenção: importar uma área de trabalho pode alterar a sua configuração atual do Corel.* Alternativa: em **Ferramentas > Personalização > Comandos > Macros**, arraste `FecharDocumentoAtivo` para uma barra de ferramentas ou menu.

**Desinstalar:** pelo "Aplicativos instalados" do Windows (com o Corel fechado). A desinstalação remove o `.gms` das pastas onde foi instalado.

### Opção 2 – Manual (importando o código)

1. Abra o CorelDRAW e aperte **Alt+F11**.
2. No Project Explorer, clique com o botão direito no projeto desejado (por exemplo `GlobalMacros`) > **Importar arquivo...** e selecione `src/FechamentoGraficaPro.bas`.
3. Repita com `src/frmFechamento.frm` (o arquivo `frmFechamento.frx` precisa estar **na mesma pasta**).
4. Salve (**Ctrl+S**) e feche o editor.

---

## Como usar

1. Abra o documento no CorelDRAW e deixe-o **ativo**.
2. Execute `FecharDocumentoAtivo` (pelo botão ou pelo Gerenciador de macros).
3. Leia a **pré-verificação** e escolha:
   - salvar o **CDR fechado** (`nome_FECHADO.cdr`);
   - gerar o **PDF** (informe a **sangria em mm** e se quer **marcas de corte**);
   - salvar o **relatório** (`.txt`).
4. Acompanhe a barra de progresso. Em arquivos grandes, o processo pode levar vários minutos (um catálogo de 26 páginas levou cerca de 12). A tela pode parar por instantes em objetos grandes ou na exportação do PDF; isso é normal.
5. Ao final, confira o resumo (as pendências aparecem em **ATENÇÃO**) e, se quiser, abra o log completo.

Se o documento tem alterações não salvas, a macro pergunta se deseja salvá-las no original antes. Se nunca foi salvo, pede uma pasta e um nome.

> **Faça sempre os primeiros testes em cópias de arquivos.**

---

## Configurações

As opções ficam no início do módulo `FechamentoGraficaPro.bas` (constantes `Const`). As principais:

| Constante | Padrão | Função |
|---|---|---|
| `SANGRIA_PADRAO_MM` | `2` | Sangria sugerida na janela |
| `PDF_PRESET` | `"PDF/X-3"` | Preset de PDF carregado antes de aplicar sangria/marcas |
| `ESPESSURA_MIN_MM` | `0.1` | Espessura mínima de contorno |
| `RES_PADRAO` / `RES_PEQUENO` / `RES_GRANDE_FORMATO` | `300` / `600` / `150` | Resolução (dpi) dos bitmaps gerados, por tamanho do objeto |
| `RES_SOMBRA` | `300` | Resolução máxima do bitmap de sombra |
| `RES_FALLBACK_ESPECIAL` | `1200` | Resolução quando um objeto de efeito precisa virar bitmap |
| `REMOVER_ORIGINAIS_VAZIOS` | `True` | Remove curvas/retângulos vazios que sobram de efeitos |
| `RASTERIZAR_EFEITO_PERSISTENTE` | `False` | Rasteriza objetos em que um efeito continua depois da conversão |
| `FALLBACK_SOMBRA` | `0` | Se a sombra não puder ser separada: `0` mantém, `1` rasteriza (perde a sombra) |
| `LIMITE_PERDAS` | `3` | Limite de falhas de rasterização por página antes de parar de rasterizar naquela página |

---

## Estrutura do repositório

```
.
├── src/
│   ├── FechamentoGraficaPro.bas   # módulo principal (macro FecharDocumentoAtivo)
│   ├── frmFechamento.frm          # formulário (opções + progresso)
│   └── frmFechamento.frx          # recursos binários do formulário (manter junto do .frm)
├── dist/
│   ├── FechamentoGrafica.gms      # projeto de macros empacotado pelo instalador
│   ├── icone.ico                  # ícone do instalador e do botão
│   └── MeuBotao.cdws              # área de trabalho com o botão personalizado
├── dist_setup/                    # saída do instalador (não versionar o .exe; use Releases)
├── Arquivos_auxiliares/
│   └── MapearDocumento.bas        # ferramenta de desenvolvimento (mapeia objetos e efeitos em CSV)
└── FechamentoGrafica.iss          # script do Inno Setup 6
```

### Ferramentas de desenvolvimento

- **`DiagnosticoEfeitos`** (dentro do módulo principal): selecione objetos e execute para listar tipo, preenchimento, transparência, sombra e números de efeito. O resultado vai para `%TEMP%\diagnostico_corel.txt`.
- **`MapearDocumento`** (`Arquivos_auxiliares/`): percorre o documento inteiro e grava um CSV (`%TEMP%\mapa_objetos.csv`) com tipo, preenchimento, contorno, transparência, efeitos, modelo de cor, sobreposição e posição de cada objeto. Útil para descobrir como o Corel representa cada recurso e para comparar o arquivo antes e depois do fechamento.

---

## Desenvolvimento

### Gerar o instalador

1. Atualize `dist/FechamentoGrafica.gms` com a versão mais recente do projeto de macros. O arquivo fica em `C:\Program Files\Corel\<versão>\Draw\GMS\` quando o projeto é criado no Corel.
2. Abra `FechamentoGrafica.iss` no **Inno Setup 6** e use *Build > Compile*.
3. O instalador é gerado em `dist_setup/`.

O `.gms` só pode ser gerado pelo CorelDRAW (é um projeto VBA binário). Ele deve ser criado na **menor versão do Corel** que se deseja suportar.

### Como a macro trata sombras

Em testes, `Shape.Separate` não separou a sombra. O método que funcionou foi `Effects(i).Separate` sobre o efeito de sombra (tipo 5). O objeto separado aparece como um retângulo com transparência em degradê e é então rasterizado em CMYK.

---

## Limitações conhecidas

- **Versão do Corel:** testado apenas no CorelDRAW Graphics Suite 2022. Um `.gms` criado numa versão nova pode não abrir em versões mais antigas.
- **Requer o componente VBA** do CorelDRAW instalado.
- **Efeitos que persistem:** em testes, três efeitos (tipos 1, 9 e 0 no Corel utilizado) continuaram registrados em objetos depois da conversão em curvas, sem diferença visual aparente. Ainda não foram identificados.
- **Objetos de efeito viram bitmap** quando o `ConvertToCurves` não funciona via VBA (extrusão, mistura, contorno etc. nos testes). O resultado é bom, mas deixa de ser vetorial.
- **Cores especiais** são sempre convertidas para CMYK; **sobreposição (overprint)** é mantida, não alterada.
- Em objetos muito grandes ou na exportação do PDF, a janela de progresso não atualiza até o fim da operação do Corel.
- A macro **não substitui a conferência no Preflight** do Acrobat (Output Preview) antes de enviar à gráfica.

---

## Roteiro

- Identificar os efeitos persistentes (1, 9 e 0) e tratá-los.
- Tentar preservar o vetor em mais objetos de efeito (separar e desagrupar antes de rasterizar).
- Opção para remover sobreposição (overprint).
- Criação automática do botão com ícone via API do Corel.
- Validação em outras versões do CorelDRAW.

---

## Aviso

Esta ferramenta altera o documento aberto de forma **destrutiva** (curvas, bitmaps, cores). O arquivo original é preservado, mas **sempre confira o resultado** (CDR e PDF) antes de enviar à gráfica. Use por sua conta e risco.

## Licença

MIT.
