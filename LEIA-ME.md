# Calculadora do Cidadão (BCB) — v2.8

Replicação local da Calculadora do Cidadão do Banco Central + cálculos
previdenciários (Demonstrativo, Cobrança Amigável, Atraso de Parcela,
Simulador IRPF).

---

## PRIMEIRA VEZ NA MAQUINA? (de casa)

### 1. Instalar Python (se ainda não tiver)

Verifica se já tem: abre o **Prompt de Comando** (cmd) e digita:

    python --version

Se aparecer `Python 3.x.x`, está OK, pula pro passo 2.
Se der erro, baixa em https://www.python.org/downloads/ e durante a
instalação **marque "Add Python to PATH"**.

### 2. Extrair os arquivos

Descompacta o ZIP em qualquer pasta, por exemplo:

    C:\Users\<seu_usuario>\Desktop\calculadora-bcb-app\

### 3. Rodar (modo rápido — sem gerar .exe)

Dá duplo clique em **executar.bat**.

A primeira vez pode pedir pra instalar módulos (openpyxl, reportlab).
Se isso acontecer, abre o cmd nessa pasta e roda:

    pip install openpyxl reportlab

Depois dá duplo clique de novo em executar.bat. O app abre.

### 4. Gerar o .exe (pra rodar sem precisar de Python)

Dá duplo clique em **regerar-exe.bat**.

Vai demorar uns 3-5 minutos. No final, o executável fica em
`dist\CalculadoraBCB.exe`. Esse arquivo é autocontido, pode mandar pra
qualquer Windows que ele roda sozinho.

---

## ARQUIVOS

| Arquivo | Pra que serve |
|---|---|
| calculadora_bcb.py | Código-fonte do app (Python). Esse é o coração. |
| executar.bat | Roda o app direto (modo desenvolvimento, sem .exe) |
| executar-debug.bat | Roda mostrando o console (útil pra ver erros) |
| regerar-exe.bat | Compila o .exe final em `dist\CalculadoraBCB.exe` |
| LEIA-ME.md | Este arquivo |

---

## FUNCIONALIDADES (v2.8)

### Abas:

1. **Índices de preços** — correção de valor entre datas pelo IPCA, IGP-M,
   INPC, IPCA-E, IPC-BR, IPC-SP
2. **Selic** — correção pela taxa Selic acumulada
3. **Lote (várias correções)** — processa várias linhas de uma vez
4. **Demonstrativo Previdenciário** — cálculo completo de contribuições em
   atraso (várias competências mensais, multa, juros, honorários)
5. **Cobrança Amigável** — atualização de débito único do termo, com multa,
   juros (a partir de notificação), honorários, abatimento de parcelas
   pagas e abatimento de 13º como crédito (IPC-FIPE)
6. **Atraso de Parcela** — atualização de uma única parcela em atraso
   (Lei 13.275/2002)

### Botão "Simulador IRPF (Receita Federal)" na aba Cobrança Amigável:
- Popup com simulador de alíquota efetiva igual ao da Receita Federal
- Tabelas embutidas: 2015 → 2026 (vigências oficiais)
- Regime mais benéfico automaticamente

### Exportações:
- CSV (todos)
- XLSX (todos, formatação institucional BCB)
- PDF (Demonstrativo — outros em breve)

---

## DADOS

A calculadora consulta a **API pública do Banco Central** (api.bcb.gov.br)
em tempo real. Precisa de internet. Os índices vêm das séries oficiais:

- IPCA = 433, IGP-M = 189, IGP-DI = 190, INPC = 188
- IPCA-E = 10764, IPC-BR = 191, IPC-SP (FIPE) = 193
- Selic diária = 11

---

## TROUBLESHOOTING

**App abre e o botão PDF some.**
O Python que rodou o `.py` não tem `reportlab` instalado:

    pip install reportlab

Ou gera o `.exe` (autocontido, tem tudo dentro).

**Erro "module not found: openpyxl".**

    pip install openpyxl

**SmartScreen do Windows reclama do .exe.**
Normal — o .exe não é assinado. Clica em "Mais informações" →
"Executar assim mesmo".

**Valores não atualizam quando muda a data.**
A partir da v2.3 o cache é limpo automaticamente. Se persistir, fecha e
abre o app.

---

Reprodução não-oficial. Para cálculos oficiais:
- BCB CALCIDADAO: https://www3.bcb.gov.br/CALCIDADAO/
- Simulador IRPF: https://www27.receita.fazenda.gov.br/simulador-irpf/
