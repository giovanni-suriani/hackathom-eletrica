class TipoIsolacaoCondutor:
    PVC = "PVC"
    LSHF_A = "LSHF/A"
    EPR_HEPR = "EPR/HEPR"
    XLPE = "XLPE"

class TipoMedicaoTemperatura:
    AMBIENTE = "AMBIENTE"
    SOLO = "SOLO"

class TipoIsolacaoCondutor:
    PVC = "PVC"
    LSHF_A = "LSHF/A"
    EPR_HEPR = "EPR/HEPR"
    XLPE = "XLPE"

class TemperaturaAmbiente:
    TEMPERATURAS = [10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80]
    
    
def make_fator_temperatura():
    # Correcao Temp
    conjunto_correcao_temperatura = [
    # Temperatura Ambiente
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 10, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 1.22 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 10, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 1.15 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 15, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 1.17 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 15, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 1.12 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 20, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 1.12 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 20, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 1.08 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 25, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 1.06 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 25, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 1.04 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 30, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 1.00 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 30, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.96 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 35, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.94 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 35, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.96 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 40, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.87 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 40, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.91 
    },
    # Temperatura no Solo
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 10, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 1.10 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 10, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 1.07 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 15, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 1.05 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 15, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 1.04 
    },
    {
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO,
        "temperatura": 20,
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A],
        "correcao_temperatura": 1.00
    },
    {
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO,
        "temperatura": 20,
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE],
        "correcao_temperatura": 0.98
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 25, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.95 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 25, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.96 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 30, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.89 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 30, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.93 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 35, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.84 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 35, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.89 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 40, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.77 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 40, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.85 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 45, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.71 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 45, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.80 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 50, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.63 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 50, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.76 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 55, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.55 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 55, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.71 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 60, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.45 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 60, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.65 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 65, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.60 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 70, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.53 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 75, 
        "isolacao_condutor": [TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.46 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 75, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.46 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 80, 
        "isolacao_condutor": [[TipoIsolacaoCondutor.PVC, TipoIsolacaoCondutor.LSHF_A], TipoIsolacaoCondutor.LSHF_A], 
        "correcao_temperatura": 0.38 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 80, 
        "isolacao_condutor": [TipoIsolacaoCondutor.EPR_HEPR, TipoIsolacaoCondutor.XLPE], 
        "correcao_temperatura": 0.38 
    }
]
    
    return conjunto_correcao_temperatura

def get_fator_temperatura(tipo_medicacao_temperatura, temperatura, isolacao_condutor):
    conjunto_correcao_temperatura = make_fator_temperatura()
    
    # Busca horrorosa, O(n), passar os dados como chave e o valor ser a correcao de temperatura vai melhorar imensamente
    for correcao_temperatura in conjunto_correcao_temperatura:
        if correcao_temperatura["tipo_medicao_temperatura"] == tipo_medicacao_temperatura and correcao_temperatura["temperatura"] == temperatura and isolacao_condutor in correcao_temperatura["isolacao_condutor"]:
            return correcao_temperatura["correcao_temperatura"]
    
print(get_fator_temperatura(TipoMedicaoTemperatura.AMBIENTE, 40, TipoIsolacaoCondutor.EPR_HEPR))