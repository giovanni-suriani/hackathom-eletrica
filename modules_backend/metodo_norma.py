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

def get_metodo_norma(tipo_cabo, maneira_instalar): 
    # MR_NBR5410
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

    for norma in NormasMRNBR:
        if norma["tipo_cabo"] == tipo_cabo and norma["maneira_instalar"] == maneira_instalar:
            return norma["metodo"]
        
    
tipo_cabo = TipoCabo.CABO_UNIPOLAR_TRIFOLIO
maneira_instalar = ManeiraInstalar.LEITO
print(get_metodo_norma(tipo_cabo, maneira_instalar))