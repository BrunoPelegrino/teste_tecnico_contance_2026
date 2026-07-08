# Robô VR — Constance Calçados

Automação em Python para cálculo e geração do relatório mensal de Vale-Refeição (VR)
dos colaboradores, a partir de `colaboradores.xlsx`.

## Requisitos

- Python 3.12+
- Dependências em `requirements.txt`

## Instalação

```bash
pip install -r requirements.txt
```

## Como rodar

1. Coloque `colaboradores.xlsx` na mesma pasta de `robo_vr.py` (ou do `.exe`).
2. Execute:

```bash
python robo_vr.py
```

3. O robô imprime o progresso no terminal:

```
====================================
ROBÔ DE CÁLCULO DE VALE REFEIÇÃO
====================================
Lendo arquivo...
Validando colunas...
Aplicando regras...
Gerando relatório...
Formatando planilha...
Relatório salvo com sucesso.
====================================
Arquivo: relatorio_vr_072026.xlsx
Colaboradores processados: 15
Total a pagar pela empresa: R$ 6,896.00
Em alerta de absenteísmo: 4
Em período de experiência: 3
```

Se algum dado tiver um problema não fatal (ex.: setor não cadastrado, data de
admissão futura), o robô continua o processamento e mostra avisos como:

```
  [AVISO] Linha 7 (Fulano de Tal): setor 'Marketing' não reconhecido na lista de
  setores cadastrados (comercial, financeiro, logística, rh, ti) — aplicado
  valor de VR padrão. Verifique se não é erro de digitação.
```

4. O arquivo `relatorio_vr_MMAAAA.xlsx` é gerado na mesma pasta, com o mês/ano
   da **data de referência** (por padrão, a data atual do sistema).

### Argumentos opcionais

| Argumento | Descrição | Padrão |
|---|---|---|
| `--entrada` | Nome do arquivo de entrada | `colaboradores.xlsx` |
| `--saida` | Nome do arquivo de saída | `relatorio_vr_MMAAAA.xlsx` |
| `--data-ref` | Data de referência (AAAA-MM-DD) para elegibilidade e nome do arquivo | data atual |

Exemplo, reprocessando um mês fechado:

```bash
python robo_vr.py --data-ref 2026-06-30 --saida relatorio_vr_062026.xlsx
```

## Testes

```bash
python -m unittest test_robo_vr.py -v
```

## Como o `.exe` foi (e deve ser) gerado

O executável **precisa ser gerado em uma máquina Windows**, pois o PyInstaller
não faz cross-compilation: um `.exe` Windows só pode ser produzido rodando o
PyInstaller a partir de um Windows real (ou uma VM/CI com esse SO).

Este ambiente de desenvolvimento usado para construir e validar a solução é
Linux, então o binário gerado aqui localmente é um executável **ELF Linux**
(testado e funcional, mesma lógica), usado apenas para validar que o comando
de empacotamento funciona de ponta a ponta. Ele **não roda no Windows** e por
isso não foi incluído como `robo_vr.exe` no repositório — isso seria enganoso.

Para gerar o `.exe` real, rode em um Windows com Python 3.12+ instalado:

```bash
pip install -r requirements.txt
pyinstaller --onefile robo_vr.py
```

O PyInstaller cria duas pastas (`build/` e `dist/`) e um arquivo `robo_vr.spec`.
O executável final fica em `dist\robo_vr.exe`. Copie **apenas** o `robo_vr.exe`
para a pasta desejada junto com `colaboradores.xlsx` e execute normalmente
(duplo clique, ou via terminal `robo_vr.exe`).

Mais detalhes e a justificativa completa dessa decisão estão em `DECISOES.md`.

## Estrutura do relatório de saída

**Aba "Detalhamento"**: uma linha por colaborador, com cabeçalho colorido,
linhas zebradas, colunas monetárias formatadas em R$, e a coluna `Status`
destacada em vermelho claro quando `ALERTA` (absenteísmo > 10%).

**Aba "Resumo"**: total de colaboradores processados, total a pagar pela
empresa, total em alerta de absenteísmo, total em período de experiência, e
data/hora de geração.