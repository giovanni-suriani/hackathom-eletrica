<script type="text/x-mathjax-config">
MathJax.Hub.Config({
    tex2jax: {
        inlineMath: [['$$','$$'], ['\\(','\\)']],
        skipTags: ['script', 'noscript', 'style', 'textarea', 'pre'] // removed 'code' entry
    }
});
MathJax.Hub.Queue(function() {
    var all = MathJax.Hub.getAllJax(), i;
    for(i = 0; i < all.length; i += 1) {
        all[i].SourceElement().parentNode.className += ' has-jax';
    }
});
</script> 
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.4/MathJax.js?config=TeX-AMS_HTML-full"></script>

# Cálculo das Fórmulas de Dimensionamento de Cabos de Baixa Tensão

## 1. Cálculo da Corrente

### Expressões Matemáticas

1. **Sistema Trifásico com Neutro (3F + N), Trifásico com Terra (3F + T) ou Trifásico (3F):**
   $$
   I = \frac{P \times 1000}{\sqrt{3} \times V \times FP}
   $$
   - Onde:
     $$
     I = \text{Corrente (A)}
     $$
     $$
     P = \text{Potência (kW)}
     $$
     $$
     V = \text{Tensão (V)}
     $$
     $$
     FP = \text{Fator de Potência}
     $$

2. **Sistema de Corrente Contínua (CC):**
   $$
   I = \frac{P \times 1000}{V}
   $$

3. **Sistema Bifásico (FF):**
   $$
   I = \frac{P \times 1000}{V \times FP}
   $$

## 2. Capacidade de Corrente

A capacidade de corrente é verificada com base na seção mínima do cabo:
$$
\text{Seção} \geq \frac{I_{nominal}}{\text{Capacidade de Corrente}}
$$

## 3. Curto-Circuito

A bitola mínima do cabo é calculada considerando a corrente de curto-circuito:
$$
\text{Bitola}_{curto} = \frac{1000 \times I_{curto} \times \sqrt{t}}{K}
$$
- Onde:
  $$
  I_{curto} = \text{Corrente de Curto-Circuito (A)}
  $$
  $$
  t = \text{Tempo de operação do dispositivo de proteção (s)}
  $$
  $$
  K = \text{Constante do material e isolação (ex.: 143 para cobre com isolação de PVC)}
  $$

## 4. Queda de Tensão

### Expressões Matemáticas

1. **Sistema Monofásico (CC):**
   $$
   \Delta V = \frac{2 \times R_{ca} \times \left(\frac{I}{n}\right) \times \left(\frac{L}{1000}\right)}{V}
   $$
   - Onde:
     $$
     \Delta V = \text{Queda de Tensão (V)}
     $$
     $$
     R_{ca} = \text{Resistência do condutor (Ω/km)}
     $$
     $$
     I = \text{Corrente (A)}
     $$
     $$
     n = \text{Número de condutores}
     $$
     $$
     L = \text{Comprimento do cabo (m)}
     $$
     $$
     V = \text{Tensão (V)}
     $$

2. **Sistema Trifásico (3F):**
   $$
   \Delta V = \frac{\sqrt{3} \times \left((R_{ca} \times FP) + (X_l \times \sin(\theta))\right) \times \left(\frac{I}{n}\right) \times \left(\frac{L}{1000}\right)}{V}
   $$
   - Onde:
     $$
     X_l = \text{Reatância indutiva do condutor (Ω/km)}
     $$
     $$
     \sin(\theta) = \sqrt{1 - FP^2} \text{ (seno do ângulo de fase)}
     $$

## 5. Correção por Agrupamento e Temperatura

A capacidade de corrente é corrigida por fatores de agrupamento e temperatura:
$$
I_{corrigido} = I_{nominal} \times F_{agrup} \times F_{temp}
$$
- Onde:
  $$
  I_{corrigido} = \text{Corrente corrigida (A)}
  $$
  $$
  F_{agrup} = \text{Fator de correção por agrupamento}
  $$
  $$
  F_{temp} = \text{Fator de correção por temperatura}
  $$

## 6. Escolha de Seção Final (M_E)

A seção final é selecionada com base na maior das três bitolas:
$$\text{Seção Final} = \max(\text{Seção_capacidade}, \text{Seção_curto}, \text{Seção_queda})$$

## 7. Neutro e Terra (NeT)

As regras para definição da seção do condutor neutro e terra são:

1. Se 
   $$
   \text{Seção} \leq 16 \; \text{mm}^2
   $$ 
   o neutro é igual à seção da fase.
   
2. Se 
   $$
   16 < \text{Seção} \leq 35 \; \text{mm}^2
   $$
   o neutro é limitado a 16 mm².

3. Para 
   $$
   \text{Seção} > 35 \; \text{mm}^2
   $$
   o neutro é igual à metade da seção da fase.
