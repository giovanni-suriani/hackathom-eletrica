
# Fórmulas de Dimensionamento de Cabos de Média Tensão

## 1. Cálculo da Bitola pelo Método de Condução de Corrente

### Expressão Matemática:
\`
I_{corrigida} = \frac{I_{nominal}}{n}
\`
- `I_{corrigida}` = Corrente corrigida para o número de condutores
- `I_{nominal}` = Corrente nominal do circuito (A)
- `n` = Número de condutores por fase

### Seleção de Bitola:
A bitola é selecionada a partir das tabelas de capacidade de corrente, de forma que:
\`
Bitola \geq mínima \ bitola \ que \ suporte \ I_{corrigida}
\`

## 2. Cálculo da Bitola pelo Método de Curto-Circuito

### Expressão Matemática:
\`
Bitola_{curto} = \frac{1000 \times I_{curto} \times \sqrt{t}}{K}
\`
- `I_{curto}` = Corrente de curto-circuito (A)
- `t` = Tempo de operação do dispositivo de proteção (s)
- `K` = Constante do material e tipo de isolação (ex.: 142 para EPR de cobre)

### Seleção de Bitola:
A bitola é escolhida de modo que:
\`
Bitola \geq Bitola_{curto}
\`

## 3. Cálculo da Bitola pelo Método de Queda de Tensão

### Expressão Matemática para Sistemas Trifásicos:
\`
\Delta V = \frac{\sqrt{3} \times I_{corrigida} \times L \times \left( R_{ca} \times FP + X_l \times \sin(\theta) \right)}{V_{fase} \times 1000}
\`
- `\Delta V` = Queda de tensão (V)
- `I_{corrigida}` = Corrente corrigida por fatores de instalação (A)
- `L` = Comprimento do cabo (m)
- `R_{ca}` = Resistência do condutor (Ω/km)
- `X_l` = Reatância indutiva do condutor (Ω/km)
- `FP` = Fator de potência
- `\sin(\theta)` = Seno do ângulo de fase, calculado como: 
  \`
  \sin(\theta) = \sqrt{1 - FP^2}
  \`
- `V_{fase}` = Tensão de fase do sistema (V)

### Limite de Queda de Tensão:
A bitola é determinada de modo que:
\`
\Delta V \leq \Delta V_{limite}
\`
- `\Delta V_{limite}` = Limite normativo de queda de tensão para o circuito

## 4. Seleção Final da Bitola

### Expressão Matemática:
\`
Bitola_{Final} = \max(Bitola_{capacidade}, Bitola_{curto}, Bitola_{queda})
\`
- `Bitola_{capacidade}` = Bitola escolhida pelo critério de capacidade de corrente
- `Bitola_{curto}` = Bitola escolhida pelo critério de curto-circuito
- `Bitola_{queda}` = Bitola escolhida pelo critério de queda de tensão

## 5. Cálculo de Seções de Neutro e Terra

### Regras de Dimensionamento:
1. **Para sistema Trifásico (3F):**
   - Neutro: "-"
   - Terra: "-"

2. **Para sistema Trifásico com Neutro (3F + N):**
   \`
   Neutro = Tabela(S_{fase})
   \`
   - Onde `S_{fase}` é a seção da fase escolhida.

3. **Para sistema Trifásico com Terra (3F + T):**
   \`
   Terra = Tabela(S_{fase})
   \`
   - Onde `S_{fase}` é a seção da fase escolhida.

## Observações
- Todas as bitolas são escolhidas de acordo com as tabelas normativas de capacidade de corrente, curto-circuito e queda de tensão.
- Os limites normativos para queda de tensão e fatores de correção de capacidade de corrente e curto-circuito devem ser respeitados para garantir a segurança do sistema.
