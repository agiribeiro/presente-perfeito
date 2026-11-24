# -*- coding: utf-8 -*-
"""Coleção O Homem-Cão (Dog Man) e derivado O Clube do Pepezinho (Cat Kid), Dav Pilkey.
Faixa 6 a 9 anos. O amazon_sync --descobrir acha ASIN/preço/foto; revisar os matches."""

_PQ_DM = "HQ de humor e aventura do Dav Pilkey — um dos maiores sucessos entre 6 e 9 anos. Fácil de ler, impossível de largar."
_PQ_CP = "Spin-off do Homem-Cão: o gatinho Pepezinho e a turma criam as próprias histórias em quadrinhos. Diverte e ainda anima a criança a desenhar e inventar."

def _dm(nome, busca, pq=_PQ_DM):
    return dict(nome=nome, autor="Dav Pilkey", busca=busca, idade_min=6, ideal=(6, 9), preco=50,
                perfil=["personagens", "criativo"], prio=dict(div=3, edu=2, cri=2, imp=2, cb=3), por_que=pq)

LIVROS_HOMEMCAO = [
    _dm("O Homem-Cão (Vol. 1)", "O Homem-Cão vol 1 Dav Pilkey"),
    _dm("O Homem-Cão Desgovernado (Vol. 2)", "O Homem-Cão Desgovernado Dav Pilkey"),
    _dm("O Homem-Cão: Um Conto de Dois Gatinhos (Vol. 3)", "O Homem-Cão Um Conto de Dois Gatinhos Dav Pilkey"),
    _dm("O Homem-Cão e o Supergatinho (Vol. 4)", "O Homem-Cão e o Supergatinho Dav Pilkey"),
    _dm("O Homem-Cão: O Senhor das Pulgas (Vol. 5)", "O Homem-Cão O Senhor das Pulgas Dav Pilkey"),
    _dm("O Homem-Cão: O Confronto Selvagem (Vol. 6)", "O Homem-Cão O Confronto Selvagem Dav Pilkey"),
    _dm("O Homem-Cão: Por Quem as Bolas Rolam (Vol. 7)", "O Homem-Cão Por Quem as Bolas Rolam Dav Pilkey"),
    _dm("O Homem-Cão: Latiu-22 (Vol. 8)", "O Homem-Cão Latiu-22 Dav Pilkey"),
    _dm("O Homem-Cão: Imundície e Castigo (Vol. 9)", "O Homem-Cão Imundície e Castigo Dav Pilkey"),
    _dm("O Homem-Cão: Mamãe dos Ventos Uivantes (Vol. 10)", "O Homem-Cão Mamãe dos Ventos Uivantes Dav Pilkey"),
    _dm("O Homem-Cão: Vinte Mil Pulgas Submarinas (Vol. 11)", "O Homem-Cão Vinte Mil Pulgas Submarinas Dav Pilkey"),
    _dm("O Homem-Cão: A Fera Escarlate (Vol. 12)", "O Homem-Cão A Fera Escarlate Dav Pilkey"),
    _dm("O Homem-Cão com Amor: Livro para Colorir", "O Homem-Cão com Amor Livro para Colorir Dav Pilkey"),
    _dm("O Clube do Pepezinho (Vol. 1)", "O Clube do Pepezinho vol 1 Dav Pilkey", _PQ_CP),
    _dm("O Clube do Pepezinho: Perspectivas (Vol. 2)", "O Clube do Pepezinho Perspectivas Dav Pilkey", _PQ_CP),
    _dm("O Clube do Pepezinho: Com Propósito (Vol. 3)", "O Clube do Pepezinho Com Propósito Dav Pilkey", _PQ_CP),
    _dm("O Clube do Pepezinho: Colaborações (Vol. 4)", "O Clube do Pepezinho Colaborações Dav Pilkey", _PQ_CP),
    _dm("O Clube do Pepezinho: Influenciadores (Vol. 5)", "O Clube do Pepezinho Influenciadores Dav Pilkey", _PQ_CP),
]
