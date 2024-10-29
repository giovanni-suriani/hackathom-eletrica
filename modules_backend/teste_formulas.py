import math

class MetodosNorma:
    METODOS_NORMA_MR_NBR410 = ["A1", "A2", "B1", "B2", "C", "E", "F"]

class TipoCabo:
    CONDUTORES_OU_CABOS_CIRCULAR = "Condutores isolados ou cabos unipolares em eletroduto de secao circular"
    CABO_MULTIPOLAR_CIRCULAR = "Cabo multipolar em eletroduto de secao circular"
    CONDUTORES_OU_CABOS_NAO_CIRCULAR = "Condutores isolados ou cabos unipolares em eletroduto de secao nao circular"
    CABO_MULTIPOLAR_NAO_CIRCULAR = "Cabo multipolar em eletroduto de secao nao circular"
    CABO_UNIPOLAR = "Cabos unipolares"
    CABO_UNIPOLAR_TRIFOLIO = "Cabos unipolares em trifólio"
    CABO_MULTIPOLAR = "Cabo multipolar"

class ManeiraInstalar:
    EMBUTIDO_PAREDE = "Embutido em parede"
    SOBRE_PAREDE = "Sobre parede"
    LEITO = "Leito"
    EMBUTIDO_ALVENARIAS = "Embutido em alvenarias"
    BANDEJA_PERFURADA = "Bandeja perfurada"
    BANDEJA_NAO_PERFURADA = "Bandeja nao-perfurada"
    FIXADO_DIRETAMENTE_TETO = "Fixado diretamente no teto"
    BANDEJA_NAO_PERFURADA_PERFILADO = "Bandeja nao-perfurada, perfilado ou prateleiras"

NormasMRNBR = [
    {
        "tipo_cabo": TipoCabo.CONDUTORES_OU_CABOS_CIRCULAR,
        "maneira_instalar": ManeiraInstalar.EMBUTIDO_PAREDE,
        "metodo": "A1"
    },
    {
        "tipo_cabo": TipoCabo.CABO_MULTIPOLAR,
        "maneira_instalar": ManeiraInstalar.EMBUTIDO_PAREDE,
        "metodo": "A2"
    },
    {
        "tipo_cabo": TipoCabo.CONDUTORES_OU_CABOS_NAO_CIRCULAR,
        "maneira_instalar": ManeiraInstalar.SOBRE_PAREDE,
        "metodo": "B1"
    },
    {
        "tipo_cabo": TipoCabo.CABO_MULTIPOLAR,
        "maneira_instalar": ManeiraInstalar.SOBRE_PAREDE,
        "metodo": "B2"
    },
    {
        "tipo_cabo": TipoCabo.CABO_UNIPOLAR,
        "maneira_instalar": ManeiraInstalar.LEITO,
        "metodo": "F"
    },
    {
        "tipo_cabo": TipoCabo.CABO_UNIPOLAR_TRIFOLIO,
        "maneira_instalar": ManeiraInstalar.LEITO,
        "metodo": "F"
    },
    {
        "tipo_cabo": TipoCabo.CABO_MULTIPOLAR,
        "maneira_instalar": ManeiraInstalar.LEITO,
        "metodo": "E"
    },
    {
        "tipo_cabo": TipoCabo.CABO_MULTIPOLAR,
        "maneira_instalar": ManeiraInstalar.EMBUTIDO_ALVENARIAS,
        "metodo": "B2"
    },
    {
        "tipo_cabo": TipoCabo.CABO_MULTIPOLAR,
        "maneira_instalar": ManeiraInstalar.BANDEJA_NAO_PERFURADA,
        "metodo": "E"
    },
    {
        "tipo_cabo": TipoCabo.CABO_UNIPOLAR,
        "maneira_instalar": ManeiraInstalar.BANDEJA_NAO_PERFURADA,
        "metodo": "E"
    },
    {
        "tipo_cabo": TipoCabo.CABO_UNIPOLAR_TRIFOLIO,
        "maneira_instalar": ManeiraInstalar.BANDEJA_NAO_PERFURADA,
        "metodo": "E"
    },
    {
        "tipo_cabo": TipoCabo.CABO_MULTIPOLAR,
        "maneira_instalar": ManeiraInstalar.BANDEJA_PERFURADA,
        "metodo": "F"
    },
    {
        "tipo_cabo": TipoCabo.CABO_UNIPOLAR,
        "maneira_instalar": ManeiraInstalar.BANDEJA_PERFURADA,
        "metodo": "F"
    },
    {
        "tipo_cabo": TipoCabo.CABO_UNIPOLAR_TRIFOLIO,
        "maneira_instalar": ManeiraInstalar.BANDEJA_PERFURADA,
        "metodo": "F"
    },
    {
        "tipo_cabo": TipoCabo.CONDUTORES_OU_CABOS_CIRCULAR,
        "maneira_instalar": ManeiraInstalar.EMBUTIDO_ALVENARIAS,
        "metodo": "B1"
    },
    {
        "tipo_cabo": TipoCabo.CABO_MULTIPOLAR,
        "maneira_instalar": ManeiraInstalar.FIXADO_DIRETAMENTE_TETO,
        "metodo": "C"
    },
    {
        "tipo_cabo": TipoCabo.CABO_UNIPOLAR,
        "maneira_instalar": ManeiraInstalar.BANDEJA_NAO_PERFURADA,
        "metodo": "C"
    }
]

class NormaMRNBR:
    # nome ruim, pode ser que haja novas normas, mas eh isso
    def __init__(self, tipo_cabo:TipoCabo , maneira_instalar, metodo):
        self.tipo_cabo = tipo_cabo
        self.maneira_instalar = maneira_instalar
        if metodo not in MetodosNorma.METODOS_NORMA_MR_NBR410:
            raise ValueError("Método de norma nao reconhecido")
        self.metodo = metodo
        
    def __eq__(self, value):
        if self.tipo_cabo == value.tipo_cabo and self.maneira_instalar == value.maneira_instalar and self.metodo == value.metodo:
            return True

class ConjuntoNormasMRNBR:
    def __init__(self):
        self.normas = []
        
    def lista_cabos_condutores(self):
        lista_cabos_condutores = []
        for norma in self.normas:
            lista_cabos_condutores.append(norma.tipo_cabo)
        return lista_cabos_condutores

    def lista_maneiras_instalar(self):
        lista_maneiras_instalar = []
        for norma in self.normas:
            lista_maneiras_instalar.append(norma.maneira_instalar)
        return lista_maneiras_instalar
    
    def lista_metodos(self):
        lista_metodos = []
        for norma in self.normas:
            lista_metodos.append(norma.metodo)
        return lista_metodos

    def metodo_norma(self, tipo_cabo, maneira_instalar):
        # Retorna o metodo de uma norma (tipo_cabo, maneira_instalar)
        for norma in self.normas:
            if norma.tipo_cabo == tipo_cabo and norma.maneira_instalar == maneira_instalar:
                return norma.metodo
        raise ValueError("Tipo de cabo ou maneira de instalar nao reconhecido")

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

class FatorCorrecaoTemperatura:
    # Classe usada para fazer o conjunto de correcao de temperatura
    def __init__(self, tipo_medicao_temperatura:TipoMedicaoTemperatura, temperatura, isolacao_condutor:TipoIsolacaoCondutor, correcao_temperatura=None):
        self.tipo_medicao_temperatura = tipo_medicao_temperatura
        if temperatura not in TemperaturaAmbiente.TEMPERATURAS:
            raise ValueError("Temperatura nao reconhecida")
        self.temperatura = temperatura
        self.isolacao_condutor = isolacao_condutor
        self.correcao_temperatura = None
        
    def __eq__(self, value):
        if self.tipo_medicao_temperatura == value.tipo_medicao_temperatura and self.temperatura == value.temperatura and self.isolacao_condutor == value.isolacao_condutor:
            return True

conjunto_correcao_temperatura = [
    # Temperatura Ambiente
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 10, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 1.22 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 10, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 1.15 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 15, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 1.17 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 15, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 1.12 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 20, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 1.12 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 20, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 1.08 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 25, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 1.06 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 25, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 1.04 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 30, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 1.00 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 30, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.96 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 35, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.94 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 35, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.96 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 40, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.87 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.AMBIENTE, 
        "temperatura": 40, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.91 
    },
    # Temperatura no Solo
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 10, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 1.10 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 10, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 1.07 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 15, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 1.05 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 15, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 1.04 
    },
    {
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO,
        "temperatura": 20,
        "isolacao_condutor": TipoIsolacaoCondutor.PVC,
        "correcao_temperatura": 1.00
    },
    {
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO,
        "temperatura": 20,
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE,
        "correcao_temperatura": 0.98
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 25, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.95 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 25, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.96 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 30, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.89 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 30, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.93 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 35, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.84 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 35, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.89 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 40, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.77 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 40, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.85 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 45, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.71 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 45, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.80 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 50, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.63 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 50, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.76 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 55, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.55 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 55, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.71 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 60, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.45 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 60, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.65 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 65, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.60 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 70, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.53 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 75, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.46 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 75, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.46 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 80, 
        "isolacao_condutor": TipoIsolacaoCondutor.PVC, 
        "correcao_temperatura": 0.38 
    },
    { 
        "tipo_medicao_temperatura": TipoMedicaoTemperatura.SOLO, 
        "temperatura": 80, 
        "isolacao_condutor": TipoIsolacaoCondutor.EPR_XLPE, 
        "correcao_temperatura": 0.38 
    }
]

class ConjuntoCorrecaoTemperatura: 
    def __init__(self):
        self.fatores = []       

    def fator_correcao_temperatura(self, fator_correcao_temperatura: FatorCorrecaoTemperatura):
        for fator in self.fatores:
            if fator_correcao_temperatura == fator:
                return fator.correcao_temperatura
        raise ValueError("Fator de correcao de temperatura nao encontrado")

class QuantidadeCamadas:
    CAMADAS = [1,2,3]

""" class FatorAgrupamento:
    def __init__(self, quantidade_circuitos_ou_multipolares, quantidade_camadas_infraestrutura, metodo_norma,maneira agrupamento=None):
        self.quantidade_circuitos_ou_multipolares = quantidade_circuitos_ou_multipolares
        self.quantidade_camadas_infraestrutura = quantidade_camadas_infraestrutura
        self.metodo_norma = metodo_norma
        self.agrupamento = agrupamento
        
    def __eq__(self, value):
        if self.quantidade_circuitos_ou_multipolares == value.quantidade_circuitos_ou_multipolares and self.quantidade_camadas_infraestrutura == value.quantidade_camadas_infraestrutura and self.metodo_norma == value.metodo_norma:
            return True
 """

""" class ConjuntoFatoresAgrupamento:
    def __init__(self):
        self.fatores = []
        
    def fator_agrupamento(self, fator_agrupamento: FatorAgrupamento):
        for fator in self.fatores:
            if fator_agrupamento == fator:
                return fator.agrupamento
        raise ValueError("Fator de agrupamento nao encontrado")
 """


referencia_agrupamento = {
    1: {
        "formas_agrupamento": ["Em feixe ao ar livre", "Em feixe sobre superficie", "Embutidos", "Em conduto fechado"],
        "metodos_norma": ["A1", "A2", "B1", "B2", "C", "E", "F"]
    },
    2: {
        "formas_agrupamento": ["Camada unica sobre parede", "Camada unica sobre piso", "Camada unica em bandeja nao perfurada", 
                               "Camada unica sobre prateleira"],
        "metodos_norma": ["C"]  
    },
    3: {
        "formas_agrupamento": ["Camada unica no teto"],
        "metodos_norma": ["C"]  
    },
    4: {
        "formas_agrupamento": ["Camada unica em bandeja perfurada"],
        "metodos_norma": ["E", "F"]  
    },
    5: {
        "formas_agrupamento": ["Camada unica sobre leito", "Camada unica sobre suporte"],
        "metodos_norma": []  
    }
}

class MetodoCalcularCabos:
    CAPACIDADE_CORRENTE = "Capacidade de corrente"
    CORRENTE_CURTO_CIRCUITO = "Corrente de curto-circuito"
    QUEDA_TENSAO = "Queda de tensao"

class NumeroCondutoresCarregados:
    NUMERO_CONDUTORES = [2, 3]

class TamanhosBitolaCabo:
    TAMANHOS_MAXIMO_BITOLA = [16 , 25, 35, 50, 70, 95, 120, 150, 185, 240, 300, 400, 500, 630, 800, 1000]
    TAMANHOS_MINIMO_BITOLA = [0.5, 0.75, 1, 1.5, 2.5, 4, 6, 10, 16, 25, 35]

class FatorCabos:
    def __init__(self, metodo_calcular_cabos:MetodoCalcularCabos, metodo_norma, 
                 isolacao_condutor:TipoIsolacaoCondutor, numero_condutores_carregados, 
                 corrente, quantidade_cabos=None):
        self.metodo_calcular_cabos = metodo_calcular_cabos
        self.metodo_norma = metodo_norma
        self.numero_condutores_carregados = numero_condutores_carregados
        self.corrente = corrente
        self.isolacao_condutor = isolacao_condutor
        self.quantidade_cabos = quantidade_cabos
        
    def __eq__(self, value):
        if self.metodo_calcular_cabos == value.metodo_calcular_cabos and self.metodo_norma == value.metodo_norma and self.isolacao_condutor == value.isolacao_condutor and self.numero_condutores_carregados == value.numero_condutores_carregados and self.corrente == value.corrente:
            return True
        
conjunto_fatores_cabos = [
    # Métodos A a F - Referência 1
    {
        "numero_circuitos_ou_multipolares": 1,
        "referencia": 1,
        "fator_agrupamento": 1.0
    },
    {
        "numero_circuitos_ou_multipolares": 2,
        "referencia": 1,
        "fator_agrupamento": 0.8
    },
    {
        "numero_circuitos_ou_multipolares": 3,
        "referencia": 1,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 4,
        "referencia": 1,
        "fator_agrupamento": 0.65
    },
    {
        "numero_circuitos_ou_multipolares": 5,
        "referencia": 1,
        "fator_agrupamento": 0.6
    },
    {
        "numero_circuitos_ou_multipolares": 6,
        "referencia": 1,
        "fator_agrupamento": 0.57
    },
    {
        "numero_circuitos_ou_multipolares": 7,
        "referencia": 1,
        "fator_agrupamento": 0.54
    },
    {
        "numero_circuitos_ou_multipolares": 8,
        "referencia": 1,
        "fator_agrupamento": 0.52
    },
    {
        "numero_circuitos_ou_multipolares": 9,
        "referencia": 1,
        "fator_agrupamento": 0.5
    },
    {
        "numero_circuitos_ou_multipolares": 10,
        "referencia": 1,
        "fator_agrupamento": 0.5
    },
    {
        "numero_circuitos_ou_multipolares": 11,
        "referencia": 1,
        "fator_agrupamento": 0.5
    },
    {
        "numero_circuitos_ou_multipolares": 12,
        "referencia": 1,
        "fator_agrupamento": 0.5
    },
    {
        "numero_circuitos_ou_multipolares": 13,
        "referencia": 1,
        "fator_agrupamento": 0.45
    },
    {
        "numero_circuitos_ou_multipolares": 14,
        "referencia": 1,
        "fator_agrupamento": 0.45
    },
    {
        "numero_circuitos_ou_multipolares": 15,
        "referencia": 1,
        "fator_agrupamento": 0.45
    },
    {
        "numero_circuitos_ou_multipolares": 16,
        "referencia": 1,
        "fator_agrupamento": 0.41
    },
    {
        "numero_circuitos_ou_multipolares": 17,
        "referencia": 1,
        "fator_agrupamento": 0.41
    },
    {
        "numero_circuitos_ou_multipolares": 18,
        "referencia": 1,
        "fator_agrupamento": 0.41
    },
    {
        "numero_circuitos_ou_multipolares": 19,
        "referencia": 1,
        "fator_agrupamento": 0.41
    },
    {
        "numero_circuitos_ou_multipolares": 20,
        "referencia": 1,
        "fator_agrupamento": 0.38
    },
    # Referencia igual a 2
    {
        "numero_circuitos_ou_multipolares": 1,
        "referencia": 2,
        "fator_agrupamento": 0.85
    },
    {
        "numero_circuitos_ou_multipolares": 2,
        "referencia": 2,
        "fator_agrupamento": 0.79
    },
    {
        "numero_circuitos_ou_multipolares": 3,
        "referencia": 2,
        "fator_agrupamento": 0.75
    },
    {
        "numero_circuitos_ou_multipolares": 4,
        "referencia": 2,
        "fator_agrupamento": 0.73
    },
    {
        "numero_circuitos_ou_multipolares": 5,
        "referencia": 2,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 6,
        "referencia": 2,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 7,
        "referencia": 2,
        "fator_agrupamento": 0.71
    },
    {
        "numero_circuitos_ou_multipolares": 8,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 9,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 10,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 11,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 12,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 13,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 14,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 15,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 16,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 17,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 18,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 19,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    {
        "numero_circuitos_ou_multipolares": 20,
        "referencia": 2,
        "fator_agrupamento": 0.7
    },
    # Referencia igual a 3
    {
        "numero_circuitos_ou_multipolares": 1,
        "referencia": 3,
        "fator_agrupamento": 0.95
    },
    {
        "numero_circuitos_ou_multipolares": 2,
        "referencia": 3,
        "fator_agrupamento": 0.81
    },
    {
        "numero_circuitos_ou_multipolares": 3,
        "referencia": 3,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 4,
        "referencia": 3,
        "fator_agrupamento": 0.68
    },
    {
        "numero_circuitos_ou_multipolares": 5,
        "referencia": 3,
        "fator_agrupamento": 0.66
    },
    {
        "numero_circuitos_ou_multipolares": 6,
        "referencia": 3,
        "fator_agrupamento": 0.64
    },
    {
        "numero_circuitos_ou_multipolares": 7,
        "referencia": 3,
        "fator_agrupamento": 0.63
    },
    {
        "numero_circuitos_ou_multipolares": 8,
        "referencia": 3,
        "fator_agrupamento": 0.62
    },
    {
        "numero_circuitos_ou_multipolares": 9,
        "referencia": 3,
        "fator_agrupamento": 0.61
    },
    {
        "numero_circuitos_ou_multipolares": 10,
        "referencia": 3,
        "fator_agrupamento": 0.61
    },
    {
        "numero_circuitos_ou_multipolares": 11,
        "referencia": 3,
        "fator_agrupamento": 0.61
    },
    {
        "numero_circuitos_ou_multipolares": 12,
        "referencia": 3,
        "fator_agrupamento": 0.61
    },
    {
        "numero_circuitos_ou_multipolares": 13,
        "referencia": 3,
        "fator_agrupamento": 0.61
    },
    {
        "numero_circuitos_ou_multipolares": 14,
        "referencia": 3,
        "fator_agrupamento": 0.61
    },
    {
        "numero_circuitos_ou_multipolares": 15,
        "referencia": 3,
        "fator_agrupamento": 0.61
    },
    {
        "numero_circuitos_ou_multipolares": 16,
        "referencia": 3,
        "fator_agrupamento": 0.61
    },
    {
        "numero_circuitos_ou_multipolares": 17,
        "referencia": 3,
        "fator_agrupamento": 0.61
    },
    {
        "numero_circuitos_ou_multipolares": 18,
        "referencia": 3,
        "fator_agrupamento": 0.61
    },
    {
        "numero_circuitos_ou_multipolares": 19,
        "referencia": 3,
        "fator_agrupamento": 0.61
    },
    {
        "numero_circuitos_ou_multipolares": 20,
        "referencia": 3,
        "fator_agrupamento": 0.61
    },
    # Referencia igual a 4
    {
        "numero_circuitos_ou_multipolares": 1,
        "referencia": 4,
        "fator_agrupamento": 0.88
    },
    {
        "numero_circuitos_ou_multipolares": 2,
        "referencia": 4,
        "fator_agrupamento": 0.82
    },
    {
        "numero_circuitos_ou_multipolares": 3,
        "referencia": 4,
        "fator_agrupamento": 0.77
    },
    {
        "numero_circuitos_ou_multipolares": 4,
        "referencia": 4,
        "fator_agrupamento": 0.75
    },
    {
        "numero_circuitos_ou_multipolares": 5,
        "referencia": 4,
        "fator_agrupamento": 0.73
    },
    {
        "numero_circuitos_ou_multipolares": 6,
        "referencia": 4,
        "fator_agrupamento": 0.73
    },
    {
        "numero_circuitos_ou_multipolares": 7,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 8,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 9,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 10,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 11,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 12,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 13,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 14,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 15,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 16,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 17,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 18,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 19,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    {
        "numero_circuitos_ou_multipolares": 20,
        "referencia": 4,
        "fator_agrupamento": 0.72
    },
    
    # Referencia igual a 5
    {
        "numero_circuitos_ou_multipolares": 1,
        "referencia": 5,
        "fator_agrupamento": 0.87
    },
    
    {
        "numero_circuitos_ou_multipolares": 2,
        "referencia": 5,
        "fator_agrupamento": 0.82
    },
    {
        "numero_circuitos_ou_multipolares": 3,
        "referencia": 5,
        "fator_agrupamento": 0.8
    },
    {
        "numero_circuitos_ou_multipolares": 4,
        "referencia": 5,
        "fator_agrupamento": 0.8
    },
    {
        "numero_circuitos_ou_multipolares": 5,
        "referencia": 5,
        "fator_agrupamento": 0.79
    },
    {
        "numero_circuitos_ou_multipolares": 6,
        "referencia": 5,
        "fator_agrupamento": 0.79
    },
    {
        "numero_circuitos_ou_multipolares": 7,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 8,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 9,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 10,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 11,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 12,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 13,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 14,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 15,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 16,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 17,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 18,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 19,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
    {
        "numero_circuitos_ou_multipolares": 20,
        "referencia": 5,
        "fator_agrupamento": 0.78
    },
]

# Para a parte de capacidade de conducao de corrente

capacidade_conducao_corrente_cobre_A2 = [
    (10, 9),
    (12, 11),
    (14, 13),
    (18.5, 16.5),
    (25, 22),
    (33, 30),
    (42, 38),
    (57, 51),
    (76, 68),
    (99, 89),
    (121, 109),
    (145, 130),
    (183, 164),
    (220, 197),
    (253, 227),
    (290, 259),
    (329, 295),
    (386, 346),
    (442, 396),
    (527, 472),
    (604, 541),
    (696, 623),
    (805, 721),
    (923, 826)
]
     
capacidade_conducao_corrente_cobre_B1 = [
    (12, 10),
    (15, 13),
    (18, 16),
    (23, 20),
    (31, 28),
    (42, 37),
    (54, 48),
    (75, 66),
    (100, 88),
    (133, 117),
    (164, 144),
    (198, 175),
    (253, 222),
    (306, 269),
    (354, 312),
    (407, 358),
    (464, 408),
    (546, 481),
    (628, 553),
    (751, 661),
    (864, 760),
    (998, 879),
    (1158, 1020),
    (1332, 1173)
]    

capacidade_conducao_corrente_cobre_B2 = [
    (11, 10),
    (15, 13),
    (17, 15),
    (22, 19.5),
    (30, 26),
    (40, 35),
    (51, 44),
    (69, 60),
    (91, 80),
    (119, 105),
    (146, 128),
    (175, 154),
    (221, 194),
    (265, 233),
    (305, 268),
    (349, 307),
    (395, 348),
    (462, 407),
    (529, 465),
    (628, 552),
    (718, 631),
    (825, 725),
    (952, 837),
    (1088, 957)
]

capacidade_conducao_corrente_cobre_C = [
    (12, 11),
    (16, 14),
    (19, 17),
    (24, 22),
    (33, 30),
    (45, 40),
    (58, 52),
    (80, 71),
    (107, 96),
    (138, 119),
    (171, 147),
    (209, 179),
    (269, 229),
    (328, 278),
    (382, 322),
    (441, 371),
    (506, 424),
    (599, 500),
    (693, 576),
    (835, 692),
    (966, 797),
    (1122, 923),
    (1311, 1074),
    (1515, 1237)
]

capacidade_conducao_corrente_cobre_D = [
    (14, 12),
    (18, 15),
    (21, 17),
    (26, 22),
    (34, 29),
    (44, 37),
    (56, 46),
    (73, 61),
    (95, 79),
    (121, 101),
    (146, 122),
    (173, 144),
    (213, 178),
    (252, 211),
    (287, 240),
    (324, 271),
    (363, 304),
    (419, 351),
    (474, 396),
    (555, 464),
    (627, 525),
    (711, 596),
    (811, 679),
    (916, 767)
]

capacidade_conducao_correntes_cobre_dict = {
    "A2": capacidade_conducao_corrente_cobre_A2,
    "B1": capacidade_conducao_corrente_cobre_B1,
    "B2": capacidade_conducao_corrente_cobre_B2,
    "C": capacidade_conducao_corrente_cobre_C,
    "D": capacidade_conducao_corrente_cobre_D
}
     
secoes_nominais = [
    0.5, 0.75, 1, 1.5, 2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 
    120, 150, 185, 240, 300, 400, 500, 630, 800, 1000
]
     
capacidade_conducao_corrente_cobre = [
    # Método A1
    {
        "metodo_norma": "A1",
        "secao_nominal": 0.5,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 10
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 0.5,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 9
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 0.75,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 12
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 0.75,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 11
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 1,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 15
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 1,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 13
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 1.5,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 19
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 1.5,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 17
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 2.5,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 26
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 2.5,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 23
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 4,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 35
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 4,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 31
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 6,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 45
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 6,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 40
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 10,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 61
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 10,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 54
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 16,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 81
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 16,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 73
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 25,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 106
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 25,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 95
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 35,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 131
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 35,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 117
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 50,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 158
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 50,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 141
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 70,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 198
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 70,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 175
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 95,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 241
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 95,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 216
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 120,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 278
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 120,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 249
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 150,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 318
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 150,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 285
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 185,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 362
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 185,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 324
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 240,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 424
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 240,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 385
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 300,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 486
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 300,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 435
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 400,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 579
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 400,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 519
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 400,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 579
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 400,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 519
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 500,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 664
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 500,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 595
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 630,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 789
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 630,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 685
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 800,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 885
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 800,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 792
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 1_000,
        "numero_condutores_carregados": 2,
        "capacidade_conducao_corrente": 1_014
    },
    {
        "metodo_norma": "A1",
        "secao_nominal": 1_000,
        "numero_condutores_carregados": 3,
        "capacidade_conducao_corrente": 908
    },
    # Continue para as demais seções e métodos A2,
]
   
def main():
    for metodo_norma, capacidades in capacidade_conducao_correntes_cobre_dict.items():
        for secao in secoes_nominais:
            for cap_2_cond, cap_3_cond in capacidades:
                var1 = {
                    "metodo_norma": metodo_norma,
                    "secao_nominal": secao,
                    "numero_condutores_carregados": 2,
                    "capacidade_conducao_corrente": cap_2_cond
                }
                var2 = {
                    "metodo_norma": metodo_norma,
                    "secao_nominal": secao,
                    "numero_condutores_carregados": 3,
                    "capacidade_conducao_corrente": cap_3_cond
                }
                capacidade_conducao_corrente_cobre.append(var1)
                capacidade_conducao_corrente_cobre.append(var2)
    
   
main() 
     
""" conjunto_conducao_de_corrente = [
   {"EPR/HEPR_XLPE":{ "Cobre": 
                     
     }
    
    }
]
      """
def metodo_capacidade_corrente(corremte, tam_max_bitola, tam_min_bitola):
    pass
     
        
class ConjuntoFatoresCabos:
    def __init__(self):
        self.fatores = []
        
    def fator_cabos(self, fator_cabos: FatorCabos):
        for fator in self.fatores:
            if fator_cabos == fator:
                return fator.quantidade_cabos
        raise ValueError("Fator de cabos nao encontrado")

def calcular_corrente_nominal(tensao, potencia_nominal, fp, sistema):
    """ 
    Calcula a corrente com base no sistema e atualiza a coluna 'H' na planilha.
    """  
    if sistema is "Trifásico com Neutro" or sistema is "Trifásico com Terra" or sistema is "Trifásico":
        corrente = (potencia_nominal * 1000) / (math.sqrt(3) * tensao * fp)
    elif sistema is "Corrente Contínua":
        corrente = (potencia_nominal * 1000) / tensao
    elif sistema is "Fase-Fase":
        corrente = (potencia_nominal * 1000) / (tensao * fp)
    else:
        raise ValueError("Sistema nao reconhecido")

    return corrente

""" If sistema = "3F + N" Or sistema = "3F + T" Or sistema = "3F" Then
            corrente = (potencia * 1000) / (Sqr(3) * tensao * fp)
            wsCalculosBT.Cells(i, "H").Value = corrente
        ElseIf sistema = "CC" Then
            corrente = (potencia * 1000) / tensao
            wsCalculosBT.Cells(i, "H").Value = corrente
        ElseIf sistema = "FF" Then
            corrente = (potencia * 1000) / (tensao * fp)
            wsCalculosBT.Cells(i, "H").Value = corrente
        End If
        """

def metodo_norma(normas_mrnbr:ConjuntoNormasMRNBR, tipo_cabo, maneira_instalar):
    #norma brasileira que estabelece regras para instalacões elétricas de baixa tensao
    # mr_nbr= { cabo unipolar: {maneira_instalar: {leito: F} } }    tipo_cabo: maneira_instalar: metodo 
    metodo = normas_mrnbr.metodo_norma(tipo_cabo, maneira_instalar)
    return metodo

def fator_correcao_temp(conjunto_correcao_temperatura: ConjuntoCorrecaoTemperatura, metodo_norma, 
                        temperatura, isolacao_condutor:TipoIsolacaoCondutor):
    
    lista_metodo_normas_ambiente = ["B1", "B2", "C", "E", "F"]
    lista_metodo_normas_solo = ["A1", "A2"]
    
    if metodo_norma in lista_metodo_normas_ambiente:
        tipo_medicao_temperatura = TipoMedicaoTemperatura.AMBIENTE
    elif metodo_norma in lista_metodo_normas_solo:
        tipo_medicao_temperatura = TipoMedicaoTemperatura.SOLO
    else:
        raise ValueError("Método de norma nao reconhecido")
    
    fator = FatorCorrecaoTemperatura(tipo_medicao_temperatura, temperatura, isolacao_condutor)
    fator_correcao_temp = conjunto_correcao_temperatura.fator_correcao_temperatura(fator)
    return fator_correcao_temp
    
def fator_agrupamento(conjunto_fatores_agrupamento: ConjuntoFatoresAgrupamento, quantidade_circuitos_ou_multipolares, 
                      quantidade_camadas_infraestrutura, metodo_norma):
    
    fator = FatorAgrupamento(quantidade_circuitos_ou_multipolares, quantidade_camadas_infraestrutura, metodo_norma)
    fator_agrupamento = conjunto_fatores_agrupamento.fator_agrupamento(fator)
    return fator_agrupamento

def calcula_corrente(corrente_nominal, fator_agrupamento, fator_temperatura):
    return corrente_nominal / ( fator_agrupamento * fator_temperatura)

def quantidade_cabos(metodo_calcular_cabos):
    if metodo_calcular_cabos not in MetodoCalcularCabos.METODOS_CALCULAR_CABOS:
        raise ValueError("Método de cálculo de cabos nao reconhecido")
    pass
