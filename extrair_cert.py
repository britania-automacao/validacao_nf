from cryptography.hazmat.primitives.serialization import pkcs12
from cryptography.hazmat.primitives import serialization

# === COLOQUE SEUS DADOS AQUI ===
CAMINHO_PFX = r"C:\Users\angelo.jorge\Desktop\Projetos\validacao_notas\validacao_nf\dados\certificados\1 - Britânia Eletrodomésticos S.A. - 0001-57.pfx"
SENHA_PFX = b"Britania@2025"  # Mantenha o 'b' antes da string da senha
# ===============================

def extrair():
    print("Extraindo chaves do arquivo .pfx...")
    with open(CAMINHO_PFX, "rb") as f:
        pfx_data = f.read()

    # Descriptografa o arquivo
    private_key, certificate, additional_certificates = pkcs12.load_key_and_certificates(pfx_data, SENHA_PFX)

    # 1. Salva a Chave Privada (.key) - ATENÇÃO: NÃO COMPARTILHE ESSE ARQUIVO
    with open("chave_privada.key", "wb") as f:
        f.write(private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption()
        ))
    
    # 2. Salva o Certificado Público (.pem)
    with open("certificado.pem", "wb") as f:
        f.write(certificate.public_bytes(serialization.Encoding.PEM))
        
        # O governo às vezes exige a cadeia completa (certificados adicionais)
        if additional_certificates:
            for cert in additional_certificates:
                f.write(cert.public_bytes(serialization.Encoding.PEM))

    print("SUCESSO! Os arquivos 'chave_privada.key' e 'certificado.pem' foram criados na sua pasta.")

if __name__ == "__main__":
    extrair()