
# Cálculo das Fórmulas de Dimensionamento de Cabos

## 1. Cálculo da Corrente

### Expressões Matemáticas

1. **Sistema Trifásico com Neutro (3F + N), Trifásico com Terra (3F + T) ou Trifásico (3F):**
   ```
   I = (P * 1000) / (sqrt(3) * V * FP)
   ```
   - `I` = Corrente
   - `P` = Potência em kW
   - `V` = Tensão em volts
   - `FP` = Fator de Potência

2. **Sistema de Corrente Contínua (CC):**
   ```
   I = (P * 1000) / V
   ```

3. **Sistema Bifásico (FF):**
   ```
   I = (P * 1000) / (V * FP)
   ```

## 2. Capacidade de Corrente

A capacidade de corrente é verificada com base na seção mínima do cabo:
```
Seção >= I_nominal / Capacidade de Corrente
```

## 3. Curto-Circuito

A bitola mínima do cabo é calculada considerando a corrente de curto-circuito:
```
Bitola = (1000 * I_curto * sqrt(t)) / K
```
- `I_curto` = Corrente de Curto-Circuito (A)
- `t` = Tempo de operação do dispositivo de proteção (s)
- `K` = Constante do material e isolação (ex.: 143 para cobre com isolação de PVC)

## 4. Queda de Tensão

### Expressões Matemáticas

1. **Sistema Monofásico (CC):**
   ```
   Queda de Tensão = (2 * Rca * (I / n) * (L / 1000)) / V
   ```
   - `Rca` = Resistência do condutor (Ω/km)
   - `I` = Corrente (A)
   - `n` = Número de condutores
   - `L` = Comprimento do cabo (m)
   - `V` = Tensão (V)

2. **Sistema Trifásico (3F):**
   ```
   Queda de Tensão = (sqrt(3) * ((Rca * FP) + (Xl * sen(theta))) * (I / n) * (L / 1000)) / V
   ```
   - `Xl` = Reatância indutiva do condutor (Ω/km)
   - `sen(theta)` = sqrt(1 - FP^2)

## 5. Correção por Agrupamento e Temperatura

A capacidade de corrente é corrigida por fatores de agrupamento e temperatura:
```
I_corrigido = I_nominal * F_agrup * F_temp
```
- `F_agrup` = Fator de correção por agrupamento
- `F_temp` = Fator de correção por temperatura

## 6. Escolha de Seção Final (M_E)

A seção final é selecionada com base na maior das três bitolas:
```
Seção Final = max(Seção_capacidade, Seção_curto, Seção_queda)
```

## 7. Neutro e Terra (NeT)

As regras para definição da seção do condutor neutro e terra são:
- Se `Seção <= 16 mm²`, o neutro é igual à seção da fase.
- Se `16 < Seção <= 35 mm²`, o neutro é limitado a 16 mm².
- Se `Seção > 35 mm²`, o neutro é igual à metade da seção da fase.