# src/buscar_xmls.py
import os
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path

# As notas exatas do seu arquivo informacoes.txt
NOTAS_ALVO = [
    {"cnpj": "37511744000180", "nota": "1424", "fornecedor": "ONDE POUSO A"},
    {"cnpj": "49748035000100", "nota": "465",  "fornecedor": "BULGAR & HIS"},
    {"cnpj": "42130537000973", "nota": "1838", "fornecedor": "C. BLOCO SUL"}
]

def _get_node(raiz, tags_possiveis):
    for elem in raiz.iter():
        local_name = elem.tag.split('}')[-1]
        if local_name in tags_possiveis:
            return elem
    return None

def _get_text(raiz, tags_possiveis, default=None):
    nodo = _get_node(raiz, tags_possiveis)
    if nodo is not None and nodo.text:
        return nodo.text.strip()
    return default

def _normalizar_numero_nota(numero_bruto: str) -> str:
    if not numero_bruto: return None
    num_limpo = ''.join(filter(str.isdigit, str(numero_bruto)))
    if not num_limpo: return str(numero_bruto).strip()
    return str(int(num_limpo))

def buscar_notas_especificas():
    # Define os caminhos relativos ao projeto
    raiz_projeto = Path(__file__).resolve().parent.parent
    pasta_xmls = raiz_projeto / "dados" / "xml_servicos_nacional"
    pasta_destino = raiz_projeto / "dados" / "xmls_encontrados"
    
    if not pasta_xmls.exists():
        print(f"[ERRO] A pasta {pasta_xmls} não existe. Verifique o caminho.")
        return

    # Cria a pasta de destino se ela não existir
    pasta_destino.mkdir(parents=True, exist_ok=True)
    
    print(f"🔍 Iniciando busca por {len(NOTAS_ALVO)} notas específicas...")
    print(f"Pasta de origem: {pasta_xmls}\n")
    
    # Cria um dicionário com a chave "CNPJ_NOTA" para facilitar a busca
    chaves_busca = {f"{alvo['cnpj']}_{alvo['nota']}": alvo for alvo in NOTAS_ALVO}
    encontrados = 0

    # Varre todos os XMLs da pasta
    for caminho_xml in pasta_xmls.glob("*.xml"):
        try:
            arvore = ET.parse(caminho_xml)
            raiz = arvore.getroot()

            # Extrai Número
            numero_bruto = _get_text(raiz, ['nNFSe', 'Numero', 'nNFS-e'])
            numero_normalizado = _normalizar_numero_nota(numero_bruto)

            # Extrai CNPJ
            cnpj = None
            nodo_forn = _get_node(raiz, ['prest', 'PrestadorServico', 'emit', 'Emitente'])
            if nodo_forn is not None:
                cnpj_bruto = _get_text(nodo_forn, ['CNPJ', 'Cnpj'])
                if cnpj_bruto: 
                    cnpj = ''.join(filter(str.isdigit, str(cnpj_bruto))).zfill(14)

            # Verifica se a nota bate com algum dos alvos
            if cnpj and numero_normalizado:
                chave_atual = f"{cnpj}_{numero_normalizado}"
                
                if chave_atual in chaves_busca:
                    fornecedor = chaves_busca[chave_atual]['fornecedor']
                    nome_arquivo_novo = f"{fornecedor.strip().replace('.', '').replace('&', 'E')} - Nota {numero_normalizado}.xml"
                    
                    print(f"✅ ACHOU! Nota {numero_normalizado} de {fornecedor}")
                    print(f"   -> Copiando de: {caminho_xml.name} para {nome_arquivo_novo}")
                    
                    # Faz uma cópia limpa e renomeada do arquivo
                    shutil.copy2(caminho_xml, pasta_destino / nome_arquivo_novo)
                    encontrados += 1
                    
                    # Remove da lista de buscas para não procurar mais esta
                    del chaves_busca[chave_atual]
                    
                    # Se zerou a lista de busca, já achou todas
                    if not chaves_busca:
                        print("\n🎉 Todas as 3 notas foram encontradas!")
                        break

        except Exception:
            continue # Ignora XMLs corrompidos durante a varredura

    print("\n" + "="*50)
    print(f"RESUMO: {encontrados} arquivos encontrados e copiados para:")
    print(f"📂 {pasta_destino}")
    print("="*50)
    
    if chaves_busca:
        print("\n❌ NOTAS NÃO ENCONTRADAS (Elas não estão na pasta):")
        for chave, alvo in chaves_busca.items():
            print(f"   - Fornecedor: {alvo['fornecedor']} | CNPJ: {alvo['cnpj']} | Nota: {alvo['nota']}")

if __name__ == '__main__':
    buscar_notas_especificas()