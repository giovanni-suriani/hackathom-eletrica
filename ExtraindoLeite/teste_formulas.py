import math



class MetodosNorma:
    METODOS_NORMA_MR_NBR410 = ["A1", "A2", "B1", "B2", "C", "E", "F"]

class TipoCabo:
    CONDUTORES_OU_CABOS_CIRCULAR = "Condutores isolados ou cabos unipolares em eletroduto de seção circular"
    CABO_MULTIPOLAR_CIRCULAR = "Cabo multipolar em eletroduto de seção circular"
    CONDUTORES_OU_CABOS_NAO_CIRCULAR = "Condutores isolados ou cabos unipolares em eletroduto de seção não circular"
    CABO_MULTIPOLAR_NAO_CIRCULAR = "Cabo multipolar em eletroduto de seção não circular"
    CABO_UNIPOLAR = "Cabos unipolares"
    CABO_UNIPOLAR_TRIFOLIO = "Cabos unipolares em trifólio"
    CABO_MULTIPOLAR = "Cabo multipolar"

class ManeiraInstalar:
    pass

class MetodoNorma:
    pass

class NormaMRNBR:
    # nome ruim, pode ser que haja novas normas, mas eh isso
    def __init__(self, tipo_cabo:TipoCabo , maneira_instalar, metodo):
        self.tipo_cabo = tipo_cabo
        self.maneira_instalar = maneira_instalar
        if metodo not in MetodosNorma.METODOS_NORMA_MR_NBR410:
            raise ValueError("Método de norma não reconhecido")
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
        raise ValueError("Tipo de cabo ou maneira de instalar não reconhecido")

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
            raise ValueError("Temperatura não reconhecida")
        self.temperatura = temperatura
        self.isolacao_condutor = isolacao_condutor
        self.correcao_temperatura = None
        
    def __eq__(self, value):
        if self.tipo_medicao_temperatura == value.tipo_medicao_temperatura and self.temperatura == value.temperatura and self.isolacao_condutor == value.isolacao_condutor:
            return True

class ConjuntoCorrecaoTemperatura: 
    def __init__(self):
        self.fatores = []       

    def fator_correcao_temperatura(self, fator_correcao_temperatura: FatorCorrecaoTemperatura):
        for fator in self.fatores:
            if fator_correcao_temperatura == fator:
                return fator.correcao_temperatura
        raise ValueError("Fator de correção de temperatura não encontrado")

class QuantidadeCamadas:
    CAMADAS = [1,2,3]

class FatorAgrupamento:
    def __init__(self, quantidade_circuitos_ou_multipolares, quantidade_camadas_infraestrutura, metodo_norma, agrupamento=None):
        self.quantidade_circuitos_ou_multipolares = quantidade_circuitos_ou_multipolares
        self.quantidade_camadas_infraestrutura = quantidade_camadas_infraestrutura
        self.metodo_norma = metodo_norma
        self.agrupamento = agrupamento
        
    def __eq__(self, value):
        if self.quantidade_circuitos_ou_multipolares == value.quantidade_circuitos_ou_multipolares and self.quantidade_camadas_infraestrutura == value.quantidade_camadas_infraestrutura and self.metodo_norma == value.metodo_norma:
            return True

class ConjuntoFatoresAgrupamento:
    def __init__(self):
        self.fatores = []
        
    def fator_agrupamento(self, fator_agrupamento: FatorAgrupamento):
        for fator in self.fatores:
            if fator_agrupamento == fator:
                return fator.agrupamento
        raise ValueError("Fator de agrupamento não encontrado")


    """ If camadas = 1 Then
        ' Verificação se o circuito está preenchido
            If circuitos = "" Then
                wsCalculosBT.Cells(i, "T").Interior.Color = RGB(255, 0, 0) ' Pinte a célula de vermelho
                MsgBox "Preencha todos os campos!"
                GoTo ProximoRegistro
            Else
                wsCalculosBT.Cells(i, "T").Interior.ColorIndex = xlNone ' Limpe a formatação da célula
            End If
        
            Set foundCell = wsCorAgrup.Rows(2).Find(What:=circuitos, LookIn:=xlValues, LookAt:=xlWhole)
            'Descobrindo o fator de correção
            If metodo = "A1" Or metodo = "A2" Or metodo = "B1" Or metodo = "B2" Or metodo = "E" Or metodo = "F" And maneira_instalar = "bandeja não-perfurada" Then
                If Not foundCell Is Nothing Then
                    valorEncontrado = foundCell.Offset(1, 0).Value ' 1 coluna  a baixo
                    wsCalculosBT.Cells(i, "V").Value = valorEncontrado
                End If
            End If
            If metodo = "C" Then
                If Not foundCell Is Nothing Then
                    valorEncontrado = foundCell.Offset(2, 0).Value
                    wsCalculosBT.Cells(i, "V").Value = valorEncontrado
                End If
            End If
            If metodo = "E" Or metodo = "F" Then
                If Not foundCell Is Nothing And maneira_instalar = "bandeja perfurada" Then
                    valorEncontrado = foundCell.Offset(4, 0).Value
                    wsCalculosBT.Cells(i, "V").Value = valorEncontrado
                End If
                If Not foundCell Is Nothing And maneira_instalar = "leito" Then
                    valorEncontrado = foundCell.Offset(5, 0).Value
                    wsCalculosBT.Cells(i, "V").Value = valorEncontrado
                End If
            End If
        End If
        If camadas = 2 Or camad 
    """

class MetodosCalcularCabos:
    METODOS_CALCULAR_CABOS = [1, 2 ,3]

class NumeroCondutoresCarregados:
    NUMERO_CONDUTORES = [2, 3]

class FatorCabos:
    def __init__(self, metodo_calcular_cabos:MetodosCalcularCabos, metodo_norma, isolacao_condutor:TipoIsolacaoCondutor, numero_condutores_carregados, corrente):
        self.metodo_calcular_cabos = metodo_calcular_cabos
        self.metodo_norma = metodo_norma
        self.isolacao_condutor = isolacao_condutor
        self.fator = fator
        


def calcular_corrente_nominal(tensao, potencia_nominal, fp, sistema):
    """ 
    Calcula a corrente com base no sistema e atualiza a coluna 'H' na planilha.
    """  
    if sistema is "Trifásico com Neutro":
        corrente = (potencia_nominal * 1000) / (math.sqrt(3) * tensao * fp)
    elif sistema is "Trifásico com Terra":
        corrente = (potencia_nominal * 1000) / tensao
    elif sistema is "Trifásico":
        corrente = (potencia_nominal * 1000) / (tensao * fp)
    else:
        raise ValueError("Sistema não reconhecido")

    return corrente

def metodo_norma(normas_mrnbr:ConjuntoNormasMRNBR, tipo_cabo, maneira_instalar):
    #norma brasileira que estabelece regras para instalações elétricas de baixa tensao
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
        raise ValueError("Método de norma não reconhecido")
    
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
    if metodo_calcular_cabos not in MetodosCalcularCabos.METODOS_CALCULAR_CABOS:
        raise ValueError("Método de cálculo de cabos não reconhecido")
    
    if metodo_calcular_cabos == 1:
        
