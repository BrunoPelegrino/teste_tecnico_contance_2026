# Decisões técnicas — Robô VR

## 1. Data de referência para elegibilidade e nome do arquivo

O enunciado não define qual data deve ser usada para contar os "90 dias desde
a admissão" (regra 5.1), nem qual mês/ano compõe `relatorio_vr_MMAAAA.xlsx`.
Não existe coluna `mes_referencia` no arquivo de entrada.

**Decisão**: usar a data de execução do robô (`date.today()`) como padrão,
com override opcional via `--data-ref AAAA-MM-DD`. Isso cobre tanto o uso
mensal normal (RH roda no início do mês) quanto reprocessamentos de meses
anteriores, sem exigir mudança no arquivo de entrada.

**Risco assumido**: se o robô rodar em um dia diferente do esperado pelo RH
(ex.: dia 28 em vez do dia 1), colaboradores no limite dos 90 dias podem ser
classificados diferente do que seria "oficialmente" o corte do mês. Mitigado
pelo argumento `--data-ref`.

## 2. Matrícula tratada como texto

`matricula` é lida e armazenada como string (`str(valor).strip()`), nunca
convertida para número, para preservar zeros à esquerda (ex.: `001`).
Confirmado no arquivo real fornecido, onde os valores já vêm como texto.

## 3. Normalização do setor "Logística"

A comparação para aplicar R$ 35,00/dia usa `.strip().lower()` antes de
comparar com `"logística"` (com acento). **Limitação assumida**: se o RH
digitar "Logistica" sem acento em um mês futuro, o robô vai classificar como
"demais setores" (R$ 30,00/dia), pois não implementei normalização de
acentuação (unicodedata). Isso não ocorreu nos dados de teste fornecidos,
mas é um ponto de melhoria futura listado abaixo.

## 3.1 Setores desconhecidos

Mantive uma constante `SETORES_CONHECIDOS` (Logística, RH, TI, Financeiro,
Comercial — os setores observados no cadastro real fornecido). Se aparecer um
setor fora dessa lista, o robô **não trava**: aplica o valor de VR padrão
(R$ 30,00/dia, regra 5.3 já prevê "demais setores") e emite um aviso no
console apontando a linha, para que o RH confirme se não é erro de
digitação (ex.: "Comercail" em vez de "Comercial"). Tratei como aviso, e não
erro fatal, porque a regra de negócio já define um valor-padrão para
qualquer setor que não seja Logística — travar a execução seria mais
rígido do que o próprio enunciado exige.

## 4. Arredondamento monetário

Uso `round(valor, 2)` (arredondamento padrão do Python, "round half to
even" em casos de empate). Não implementei arredondamento bancário
específico nem truncamento, por não haver especificação no enunciado. Para
uso em folha de pagamento real, recomendo validar com o time financeiro se
esse comportamento é aceitável.

## 5. Linhas com dados inconsistentes não travam o processamento

Se `dias_trabalhados + faltas > dias_uteis_mes`, ou se `dias_uteis_mes = 0`,
ou se a data de admissão for posterior à data de referência, o robô **não
interrompe a execução**: registra um aviso no console (prefixo `[AVISO]`)
apontando a linha e o colaborador, e segue processando as demais linhas.

**Justificativa**: o time de RH precisa do relatório completo todo mês; travar
o robô por causa de uma linha suspeita obrigaria reprocessar tudo. Já erros
estruturais (coluna obrigatória ausente, arquivo não encontrado, arquivo
corrompido) **interrompem** a execução com mensagem clara, pois nesses casos
o relatório inteiro seria inconfiável.

## 6. Absenteísmo com `dias_uteis_mes = 0`

Protegido explicitamente contra `ZeroDivisionError`: retorna `0%` sem alerta
e registra aviso. Não ocorre no arquivo de teste fornecido (todos os
colaboradores têm `dias_uteis_mes = 22`), mas é tratado porque este robô
será reutilizado com dados reais em meses futuros.

## 7. Leitura de colunas por nome, não por posição

O robô lê o cabeçalho da linha 1 e monta um mapa `{nome_coluna: índice}`, em
vez de assumir que a coluna X está sempre na posição N. Isso torna o robô
resiliente a reordenação de colunas pelo RH.

## 7.1 Onde o robô procura o arquivo de entrada

O PDF pede que o `.exe` procure `colaboradores.xlsx` na mesma pasta do
executável (seção 6.3). Mas rodar `robo_vr.py` durante o desenvolvimento a
partir de outro diretório (com `--entrada caminho/para/arquivo.xlsx`) é um
uso legítimo e comum de uma ferramenta de linha de comando. Por isso, a busca
tenta, nesta ordem: (1) o caminho informado relativo ao diretório de trabalho
atual e (2) a mesma pasta do script/executável. Isso satisfaz o requisito do
`.exe` sem prejudicar o uso normal via CLI.

## 8. Executável `.exe`

PyInstaller não permite cross-compilation: um `.exe` Windows só pode ser
gerado rodando o PyInstaller **em um Windows real**. O ambiente usado para
desenvolver e validar esta solução é Linux, portanto o binário gerado aqui é
um executável ELF Linux — validado e funcional, mas **não é um `.exe`**.

Optei por **não** incluir um arquivo renomeado `robo_vr.exe` que na verdade
é um binário Linux, pois isso seria enganoso e não funcionaria na máquina do
avaliador. Em vez disso, documentei no `README.md` o comando exato
(`pyinstaller --onefile robo_vr.py`) para gerar o `.exe` verdadeiro em um
Windows, e testei esse mesmo fluxo de empacotamento localmente (com o
binário Linux equivalente, `dist/robo_vr`) para garantir que o comando e o
código funcionam ponta a ponta sem erros de empacotamento — inclusive a
lógica de localizar `colaboradores.xlsx` ao lado do executável, via
`sys.frozen`/`sys.executable`.

## 8.1 Experiência do terminal: pausas de progresso e espera antes de sair

Dois ajustes feitos após teste real no Windows, onde o `.exe` fechava a
janela imediatamente ao terminar (sucesso ou erro), impedindo o usuário de
RH de ler qualquer mensagem:

- **Pausa curta (1s) entre cada mensagem de progresso** (`Lendo arquivo...`,
  `Validando colunas...` etc.), via `time.sleep`, para dar sensação de
  processamento real em vez de tudo aparecer instantaneamente.
- **`input("Pressione ENTER para sair...")` ao final da execução**, tanto no
  caminho de sucesso quanto nos de erro, mantendo a janela do console aberta
  até o usuário confirmar que leu a mensagem.

Essa pausa final só é ativada quando `sys.frozen` é `True` (ou seja, quando
o código está rodando como `.exe` empacotado pelo PyInstaller). Ao rodar
`python robo_vr.py` durante o desenvolvimento, em CI, ou nos testes
automatizados, `sys.frozen` não existe e a pausa é pulada — assim o robô
continua não-interativo nesses contextos, e só vira interativo no cenário
real de uso (usuário de RH dando duplo clique no `.exe`).

## 9. Padrões de código adotados

- **Python 3.12+**, com `from __future__ import annotations` para permitir
  type hints modernos (`list[str]`, `tuple[float, str]`) sem depender de
  versões específicas do `typing`.
- **`pathlib.Path`** em vez de strings de caminho, para resolução de arquivo
  de entrada/saída de forma portável entre sistemas operacionais.
- **`dataclasses`** para os três modelos de dados (`Colaborador`, `ResultadoVR`,
  `Estatisticas`), evitando dicionários soltos e deixando os campos e tipos
  explícitos.
- **Funções pequenas e puras** para cada regra de negócio (`calcular_elegibilidade`,
  `calcular_dias_a_pagar`, `calcular_valor_vr`, `calcular_absenteismo`,
  `valor_dia_por_setor`), sem efeitos colaterais nem I/O, o que as torna
  testáveis isoladamente (ver `test_robo_vr.py`).
- **Constantes nomeadas** para todos os valores fixos do negócio (valores de
  VR, percentual de desconto, limite de experiência, limite de absenteísmo,
  cores de formatação), nunca "números mágicos" espalhados pelo código.
- **Exceções customizadas** (`RoboVRError` e subclasses) em vez de exceções
  genéricas, para que `main()` diferencie erros conhecidos (mensagem amigável)
  de falhas inesperadas (rede de segurança final, sem traceback exposto).

## Limitações conhecidas

- Não trata acentuação alternativa de nomes de setor (ver item 3).
- Não valida duplicidade de matrícula/nome (assume que o arquivo de entrada
  já vem sem duplicatas, pois essa validação não estava no escopo do
  enunciado).
- Não lê arquivos `.xls` antigos (apenas `.xlsx`), pois o enunciado especifica
  esse formato.
- Arredondamento monetário usa o padrão do Python, sem regra bancária
  específica (ver item 4).

## Melhorias futuras

- Normalizar acentuação na comparação de setores (`unicodedata.normalize`).
- Externalizar valores de VR por setor e o percentual de desconto em um
  arquivo de configuração (`config.json` ou variáveis de ambiente), em vez de
  constantes fixas no código, para facilitar reajustes sem alterar o script.
- Adicionar log em arquivo (`logging`), além do console, para auditoria em
  produção.
- Validar duplicidade de matrícula como erro estrutural.