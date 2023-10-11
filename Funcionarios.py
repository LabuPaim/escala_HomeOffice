
class Funcionario:
    def __init__(self, nome, cargo, home):
        self.nome = nome
        self.cargo = cargo
        self.home = home

funcionarios = []
# 1 == "Home-Office" 2 == "Presencial"
dados_funcionarios = [
    {"nome": "Alana", "cargo": "Analista de Requisitos", "home": 1},
    {"nome": "Alexandre", "cargo": "Designer", "home": 1},
    {"nome": "Cândida", "cargo": "Analista de Requisitos", "home": 2},
    {"nome": "Hamilton", "cargo": "Desenvolvedor", "home": 1},
    {"nome": "Victor", "cargo": "Desenvolvedor", "home": 2},
]

for dados in dados_funcionarios:
    funcionarios.append(Funcionario(dados["nome"], dados["cargo"], dados["home"]))
