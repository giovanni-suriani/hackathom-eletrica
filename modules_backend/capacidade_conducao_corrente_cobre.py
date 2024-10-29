class NumeroCondutoresCarregados:
    NUMERO_CONDUTORES = [2, 3]

class TamanhosBitolaCabo:
    TAMANHOS_MAXIMO_BITOLA = [16 , 25, 35, 50, 70, 95, 120, 150, 185, 240, 300, 400, 500, 630, 800, 1000]
    TAMANHOS_MINIMO_BITOLA = [0.5, 0.75, 1, 1.5, 2.5, 4, 6, 10, 16, 25, 35]

def make_capacidade_conducao_correntes_cobre():
    capacidade_conducao_corrente_cobre_A1 = [
    (10, 9),
    (12, 11),
    (15, 13),
    (19, 17),
    (26, 23),
    (35, 31),
    (45, 40),
    (61, 54),
    (81, 73),
    (106, 95),
    (131, 117),
    (158, 141),
    (200, 179),
    (241, 216),
    (278, 249),
    (318, 285),
    (362, 324),
    (424, 380),
    (486, 435),
    (579, 519),
    (664, 595),
    (765, 685),
    (885, 792),
    (1014, 908)
    ]

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

#TODO: capacaidade_conducao_corrente_cobre_E e F


    capacidade_conducao_correntes_cobre_dict = {
        "A1": capacidade_conducao_corrente_cobre_A1,
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
     
    capacidade_conducao_correntes_cobre_final = {} 
     
    for metodo_norma, capacidades in capacidade_conducao_correntes_cobre_dict.items():
        for secao in secoes_nominais:
            for cap_2_cond, cap_3_cond in capacidades:
                chave = (metodo_norma, secao, 2)
                capacidade_conducao_correntes_cobre_final[chave] = cap_2_cond
                chave = (metodo_norma, secao, 3)
                capacidade_conducao_correntes_cobre_final[chave] = cap_3_cond
                
    return capacidade_conducao_correntes_cobre_final
    
    
# Referencia esta errada tem que ser a planilha 2 de captura de corrente
    
cap = make_capacidade_conducao_correntes_cobre()
print(cap[("A1", 0.5, 2)])
    
""" def metodo_capacidade_corrente(corrente, tamanho_bitola_min, tamanho_bitola_max, numero_condutores):
    numero_cabos = 1
    def condicao_corrente(corrente, numero_cabos):
        
    
    if tamanho_bitola_max not in TamanhosBitolaCabo.TAMANHOS_MAXIMO_BITOLA:
       raise ValueError("Tamanho da bitola maxima invalido")
    if tamanho_bitola_min not in TamanhosBitolaCabo.TAMANHOS_MINIMO_BITOLA:
        raise ValueError("Tamanho da bitola minima invalido")
    
    capacidade_conducao_correntes_cobre = make_capacidade_conducao_correntes_cobre()
    
    while corrente  """
    
